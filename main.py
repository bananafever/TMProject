import os
import re
import sys
from datetime import datetime, time, timedelta

import pytz
import win32com.client
from dateutil.relativedelta import relativedelta
from PySide6 import QtGui as qtg
from PySide6 import QtWidgets as qtw

ref_num_key = re.compile(
    r"[a-zA-Z][a-zA-Z][a-zA-Z]?-[tT]?2[a-zA-Z][0-9]{4}[a-zA-Z0-9-]{0,5}|"
    r"[a-zA-Z][a-zA-Z][a-zA-Z]?-[0-9]{5}"
)
name_key = re.compile(r"; [ㄱ-ㅎ|ㅏ-ㅣ|가-힣]+(?=\()")

# 경로 설정 (환경에 맞게 수정하세요)
parent_path = r"C:\Dropbox\1_Projects\Tasks"
target_path = r"C:\Dropbox\4_Archives\04_Work\02_Review"


def create_dir_if_not_exists(directory):
    """디렉토리가 없으면 생성합니다.
    반환값: 이미 존재하면 True, 새로 생성하면 False.
    """
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
            return False  # 새로 생성됨
        else:
            return True  # 이미 존재함
    except OSError:
        print("Error: Creating directory. " + directory)
        return False


def _extract_message_info(message):
    """이메일 제목에서 날짜, 시간 ID, ref 번호, 이름을 추출합니다.

    반환값: (delivery_date, time_id, ref_num, name) 튜플
    실패 시 ValueError 발생.
    """
    delivery_date = message.CreationTime.strftime("%Y-%m-%d")
    time_id = message.CreationTime.strftime("%Hhr%Mmn")

    ref_nums = ref_num_key.findall(message.Subject)
    names = name_key.findall(message.Subject)

    if not ref_nums:
        raise ValueError(f"제목에서 Ref 번호를 찾을 수 없습니다: {message.Subject}")
    if not names:
        raise ValueError(f"제목에서 이름을 찾을 수 없습니다: {message.Subject}")

    ref_num = ref_nums[0]
    name = names[0][2:]
    return delivery_date, time_id, ref_num, name


def _save_attachments_for_message(message, attach_list=None):
    """단일 메시지의 첨부파일을 처리하고 저장합니다."""
    try:
        delivery_date, time_id, ref_num, name = _extract_message_info(message)
    except ValueError as e:
        print(f"메시지 정보 추출 실패: {e}")
        return

    date_folder = os.path.join(parent_path, delivery_date)
    create_dir_if_not_exists(date_folder)

    ref_folder = ref_num + "_" + name
    task_folder = os.path.join(date_folder, time_id + "_" + ref_folder)

    if create_dir_if_not_exists(task_folder):  # True = 이미 존재
        print(ref_folder + ": Existing")
        if attach_list is not None:
            attach_list.append(time_id + "_" + ref_folder + ": Already Existing")
        return

    if attach_list is not None:
        attach_list.append(time_id + "_" + ref_folder + ": Downloaded")

    for attachment in message.Attachments:
        file_name = "SYS_" + attachment.FileName
        attachment.SaveAsFile(os.path.join(task_folder, file_name))


def save_attachments():
    inbox = outlook_old.GetDefaultFolder(6)
    messages = inbox.Items
    attach_list = []

    for message in messages:
        if (
            (message.SenderName == "DoNotReply")
            and ("위임" in message.Subject)
            and (message.Attachments.Count != 0)
        ):
            _save_attachments_for_message(message, attach_list)

    return attach_list


class Handler_Class(object):

    def OnNewMailEx(self, receivedItemsIDs):
        # receivedItemsIDs는 ","로 구분된 메일 ID 모음입니다.
        for ID in receivedItemsIDs.split(","):
            message = outlook_new.Session.GetItemFromID(ID)

            if (
                (message.SenderName == "DoNotReply")
                and ("위임" in message.Subject)
                and (message.Attachments.Count != 0)
            ):
                _save_attachments_for_message(message)

            # 이름 추출 - 조건 체크 전 실패 가능성 처리
            names = name_key.findall(message.Subject)
            if not names:
                continue
            name = names[0][2:]

            if (
                (message.SenderName in ["DoNotReply", "Yong-Sok SHIN [신용석]"])
                and ("위임" in message.Subject)
                and (name in ["김지원", "한송희"])
            ):
                if "거절결정서" in message.Subject:
                    rejection_type = "FR"
                else:
                    rejection_type = "OA"

                self.add_calendar_event(message, rejection_type)

    def add_calendar_event(self, message, rejection_type):
        # Outlook의 Calendar에 접근
        calendar = outlook_new.GetNamespace("MAPI").GetDefaultFolder(
            9
        )  # 9은 Calendar 폴더

        subfolder_name = "Mail"
        subfolder = calendar.Folders[subfolder_name]

        subfolder2_name = "No_Instructions"
        subfolder2 = calendar.Folders[subfolder2_name]

        appointment = subfolder.Items.Add()
        appointment2 = subfolder2.Items.Add()

        # Ref 번호 추출
        ref_nums = ref_num_key.findall(message.Subject)
        if not ref_nums:
            print(f"Ref 번호를 찾을 수 없습니다: {message.Subject}")
            return
        ref_num = ref_nums[0]
        appointment.Subject = f"{ref_num}_Comments"

        # 이메일 제목에서 날짜 추출
        date_pattern = r"(\d{4}-\d{2}-\d{2})"
        match = re.search(date_pattern, message.Subject)

        target_date = None
        response_date = None

        if match:
            target_date_str = match.group(1)
            target_date = datetime.strptime(target_date_str, "%Y-%m-%d")

            utc_timezone = pytz.utc
            target_date_with_time = datetime.combine(target_date, time(0, 0, 0))
            target_date_utc = utc_timezone.localize(target_date_with_time)
            appointment.Start = target_date_utc

            two_months_later = target_date + relativedelta(months=2)
            three_months_later = target_date + relativedelta(months=3)

            if rejection_type == "OA":
                appointment2.Subject = f"{ref_num}_의견서"
                response_date = two_months_later - timedelta(days=11)
            else:
                appointment2.Subject = f"{ref_num}_재심사"
                response_date = three_months_later - timedelta(days=10)

            response_date_with_time = datetime.combine(response_date, time(0, 0, 0))
            response_date_utc = utc_timezone.localize(response_date_with_time)
            appointment2.Start = response_date_utc

        else:
            # 날짜를 찾지 못한 경우 이메일 생성 시간으로 대체
            appointment.Start = message.CreationTime
            appointment2.Start = message.CreationTime  # appointment2도 Start 설정
            appointment2.Subject = f"{ref_num}_일정"

        appointment.AllDayEvent = True
        appointment.Body = f"Task details: {message.Body}"

        appointment2.AllDayEvent = True
        appointment2.Body = f"Task details: {message.Body}"

        appointment.Save()
        appointment2.Save()

        messageBox = qtw.QMessageBox()
        messageBox.setWindowTitle("Schedule Creation Complete")
        if target_date and response_date:
            messageBox.setText(
                f"{target_date}: {ref_num}\n{response_date}: {ref_num}\nSchedule Creation Complete"
            )
        else:
            messageBox.setText(f"{ref_num}\nSchedule Creation Complete (날짜 미지정)")
        messageBox.exec()


class MainWin(qtw.QWidget):

    def __init__(self, attach_list):
        super().__init__()
        self.attach_list = attach_list
        self.initUI()

    def initUI(self):
        notice = qtw.QGroupBox()
        notice.setTitle("Messages")
        notice_layout = qtw.QVBoxLayout()

        all_string = ""
        for attach in self.attach_list:  # self.attach_list 사용 (버그 수정)
            all_string += attach + "\n"

        label = qtw.QLabel(all_string)
        notice_layout.addWidget(label)
        notice.setLayout(notice_layout)

        transfer_btn = qtw.QPushButton("File Transfer")
        quit_btn = qtw.QPushButton("Quit")

        win_layout = qtw.QVBoxLayout()
        win_layout.addWidget(notice)
        win_layout.addWidget(transfer_btn)
        win_layout.addWidget(quit_btn)

        transfer_btn.clicked.connect(self.transfer_clicked)
        quit_btn.clicked.connect(self.quit_clicked)

        self.label = label

        self.setLayout(win_layout)
        self.setWindowTitle("Task Manager")
        self.setWindowIcon(qtg.QIcon("icon.png"))
        self.move(300, 300)
        self.resize(400, 200)
        self.show()

        tray_icon = qtw.QSystemTrayIcon(qtg.QIcon("icon.png"), self)
        menu = qtw.QMenu()

        action1 = menu.addAction("Show")
        action1.triggered.connect(lambda: self.show())
        action2 = menu.addAction("Quit")
        action2.triggered.connect(lambda: sys.exit())

        tray_icon.setContextMenu(menu)
        tray_icon.show()

    def closeEvent(self, event):
        event.ignore()
        self.hide()

    def transfer_clicked(self):
        if not os.path.exists(parent_path):
            raise FileNotFoundError(f"{parent_path} 경로가 존재하지 않습니다.")

        completion_messages = ""

        if not os.listdir(parent_path):
            self.label.setText("이전할 파일이 없습니다.")
            return

        for date_folder in os.listdir(parent_path):
            folder_path = os.path.join(parent_path, date_folder)

            ref_info = None
            person_info = None

            for case_folder in os.listdir(folder_path):
                try:
                    time_info, ref_info, person_info = case_folder.split("_")
                    target_folder_name = (
                        f"{ref_info}\\{date_folder}_{time_info}_{person_info}"
                    )
                except ValueError:  # bare except 대신 ValueError 명시
                    target_folder_name = f"{case_folder}"

                create_dir_if_not_exists(os.path.join(target_path, target_folder_name))

                origin_path = os.path.join(folder_path, case_folder)

                for file_name in os.listdir(origin_path):
                    file_path = os.path.join(origin_path, file_name)
                    os.rename(
                        file_path,
                        os.path.join(target_path, target_folder_name, file_name),
                    )

                os.rmdir(origin_path)

            os.rmdir(folder_path)

            if ref_info and person_info:
                completion_messages += f"{ref_info}_{person_info} 이전 작업 완료\n"
            else:
                completion_messages += f"{date_folder} 이전 작업 완료\n"

        self.label.setText(completion_messages)

        messageBox = qtw.QMessageBox()
        messageBox.setWindowTitle("File Transfer")
        messageBox.setText("파일 이전 작업이 완료되었습니다.")
        messageBox.exec()

    def quit_clicked(self):
        sys.exit()


if __name__ == "__main__":
    outlook_old = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    outlook_new = win32com.client.DispatchWithEvents(
        "Outlook.Application", Handler_Class
    )
    attach_list = save_attachments()
    app = qtw.QApplication(sys.argv)
    mw = MainWin(attach_list)
    sys.exit(app.exec())
