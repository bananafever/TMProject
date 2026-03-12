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

"""
current_folder = os.getcwd()
disc_name = current_folder.split("Dropbox")[0]
"""

parent_path = r"C:\Dropbox\1_Projects\Tasks"
target_path = r"C:\Dropbox\4_Archives\04_Work\02_Review"


def create_dir_if_not_exists(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
        else:
            isExisting = "Yes"
            return isExisting

    except OSError:
        print("Error: Creating directory. " + directory)


def save_attachments():

    inbox = outlook_old.GetDefaultFolder(6)
    messages = inbox.Items
    attach_list = []

    # To iterate through inbox emails using inbox.Items object.
    for message in messages:

        if (
            (message.SenderName == "DoNotReply")
            and ("위임" in message.Subject)
            and (message.Attachments.Count != 0)
        ):

            delivery_date = (
                "{:04d}".format(message.CreationTime.year)
                + "-"
                + "{:02d}".format(message.CreationTime.month)
                + "-"
                + "{:02d}".format(message.CreationTime.day)
            )

            date_folder = parent_path + r"\{0}".format(delivery_date)

            create_dir_if_not_exists(date_folder)

            time_id = (
                "{:02d}".format(message.CreationTime.hour)
                + "hr"
                + "{:02d}".format(message.CreationTime.minute)
                + "mn"
            )

            ref_num = ref_num_key.findall(message.Subject)[0]
            name = name_key.findall(message.Subject)[0][2:]
            ref_folder = ref_num + "_" + name
            task_folder = date_folder + "\\" + time_id + "_" + ref_folder

            if create_dir_if_not_exists(task_folder) == "Yes":
                print(ref_folder + ": Existing")
                attach_list.append(time_id + "_" + ref_folder + ": Already Existing")
                continue

            attach_list.append(time_id + "_" + ref_folder + ": Downloaded")

            # To iterate through email items using message.Attachments object.
            for attachment in message.Attachments:
                file_name = "SYS_" + attachment.FileName

                # To save the perticular attachment at the desired location in your hard disk.
                attachment.SaveAsFile(os.path.join(task_folder, file_name))

    return attach_list


class Handler_Class(object):

    def OnNewMailEx(self, receivedItemsIDs):
        # RecrivedItemIDs is a collection of mail IDs separated by a ",".
        # You know, sometimes more than 1 mail is received at the same moment.

        for ID in receivedItemsIDs.split(","):
            message = outlook_new.Session.GetItemFromID(ID)
            name = name_key.findall(message.Subject)[0][2:]

            if (
                (message.SenderName == "DoNotReply")
                and ("위임" in message.Subject)
                and (message.Attachments.Count != 0)
            ):

                delivery_date = (
                    "{:04d}".format(message.CreationTime.year)
                    + "-"
                    + "{:02d}".format(message.CreationTime.month)
                    + "-"
                    + "{:02d}".format(message.CreationTime.day)
                )

                date_folder = parent_path + r"\{0}".format(delivery_date)

                create_dir_if_not_exists(date_folder)

                time_id = (
                    "{:02d}".format(message.CreationTime.hour)
                    + "hr"
                    + "{:02d}".format(message.CreationTime.minute)
                    + "mn"
                )

                ref_num = ref_num_key.findall(message.Subject)[0]
                name = name_key.findall(message.Subject)[0][2:]
                ref_folder = ref_num + "_" + name
                task_folder = date_folder + "\\" + time_id + "_" + ref_folder

                if create_dir_if_not_exists(task_folder) == "Yes":
                    print(ref_folder + ": Existing")
                    continue

                # To iterate through email items using message.Attachments object.
                for attachment in message.Attachments:
                    file_name = "SYS_" + attachment.FileName

                    # To save the particular attachment at the desired location in your hard disk.
                    attachment.SaveAsFile(os.path.join(task_folder, file_name))

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

        # 일정의 세부사항 설정
        ref_num = ref_num_key.findall(message.Subject)[0]
        appointment.Subject = f"{ref_num}_Comments"

        # 카테고리 설정 (예: "Important" 카테고리)
        # appointment.Categories = "녹색 범주"

        # 이메일 제목에서 날짜 추출하기 위한 정규 표현식
        date_pattern = r"(\d{4}-\d{2}-\d{2})"  # "YYYY-MM-DD" 형식의 날짜 찾기
        match = re.search(date_pattern, message.Subject)

        if match:
            # 추출된 날짜를 datetime 객체로 변환
            target_date_str = match.group(1)
            target_date = datetime.strptime(target_date_str, "%Y-%m-%d")

            # 종일 일정이기 때문에 자정(00:00)으로 설정하고 UTC로 변환
            utc_timezone = pytz.utc  # UTC 시간대 설정
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
            # 기본적으로 이메일 생성 시간으로 일정 시작 시간 설정
            appointment.Start = message.CreationTime

        appointment.AllDayEvent = True  # 종일 이벤트로 설정
        appointment.Body = f"Task details: {message.Body}"

        appointment2.AllDayEvent = True  # 종일 이벤트로 설정
        appointment2.Body = f"Task details: {message.Body}"

        # appointment.ReminderSet = True
        # appointment.ReminderMinutesBeforeStart = 15  # 15분 전에 알림

        # 일정 저장
        appointment.Save()
        appointment2.Save()

        messageBox = qtw.QMessageBox()
        messageBox.setWindowTitle("Schedule Creation Complete")
        messageBox.setText(
            f"{target_date}: {ref_num}\n{response_date}: {ref_num}\nSchedule Creation Complete"
        )
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
        for attach in attach_list:
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

        # 상위 폴더의 하위 폴더를 순회합니다.
        for date_folder in os.listdir(parent_path):
            folder_path = os.path.join(parent_path, date_folder)

            for case_folder in os.listdir(
                folder_path
            ):  # case_folder: 시간_ref_담당자 폴더

                try:
                    time_info, ref_info, person_info = case_folder.split("_")
                    target_folder_name = (
                        f"{ref_info}\\{date_folder}_{time_info}_{person_info}"
                    )
                except:
                    target_folder_name = f"{case_folder}"

                create_dir_if_not_exists(os.path.join(target_path, target_folder_name))

                origin_path = os.path.join(folder_path, case_folder)

                for file_name in os.listdir(origin_path):
                    file_path = os.path.join(origin_path, file_name)
                    os.rename(
                        file_path,
                        os.path.join(target_path, target_folder_name, file_name),
                    )

                os.rmdir(origin_path)  # Ref. 폴더 삭제

            os.rmdir(folder_path)  # 날짜 폴더 삭제

            completion_messages += f"{ref_info}_{person_info} 이전 작업 완료\n"

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
