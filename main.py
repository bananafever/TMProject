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


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


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
        # 수정 #3: 저장 실패 시 예외를 잡아 나머지 첨부파일 저장 계속 진행
        try:
            attachment.SaveAsFile(os.path.join(task_folder, file_name))
        except Exception as e:
            print(f"첨부파일 저장 실패 ({file_name}): {e}")
            if attach_list is not None:
                attach_list.append(f"{file_name}: 저장 실패")


def save_attachments(outlook):
    """받은 편지함의 위임 메일 첨부파일을 저장합니다.

    Args:
        outlook: Outlook MAPI 네임스페이스 인스턴스 (개선 #8: 전역 변수 대신 인자로 전달)
    """
    inbox = outlook.GetDefaultFolder(6)
    # 수정 #2: list()로 스냅샷을 찍어 순회 중 Items 변경으로 인한 누락/중복 방지
    messages = list(inbox.Items)
    attach_list = []

    for message in messages:
        if (
            (message.SenderName == "DoNotReply")
            and ("위임" in message.Subject)
            and (message.Attachments.Count != 0)
        ):
            _save_attachments_for_message(message, attach_list)

    return attach_list


# 개선 #6: Python 3에서 (object) 상속 명시는 불필요
class Handler_Class:
    # 개선 #8: outlook 인스턴스를 클래스 변수로 관리 (DispatchWithEvents 특성상 생성자 인자 전달 불가)
    outlook = None
    main_win = None  # UI 업데이트를 위한 MainWin 인스턴스 참조

    def OnNewMailEx(self, receivedItemsIDs):
        # receivedItemsIDs는 ","로 구분된 메일 ID 모음입니다.
        for ID in receivedItemsIDs.split(","):
            message = Handler_Class.outlook.Session.GetItemFromID(ID)

            if (
                (message.SenderName == "DoNotReply")
                and ("위임" in message.Subject)
                and (message.Attachments.Count != 0)
            ):
                # 수정 #1: attach_list를 넘겨 저장 결과를 받고 QListWidget에 항목 추가
                new_items = []
                _save_attachments_for_message(message, new_items)
                if Handler_Class.main_win is not None and new_items:
                    for item in new_items:
                        Handler_Class.main_win.list_widget.addItem(item)

            # 이름 추출 - 조건 체크 전 실패 가능성 처리
            names = name_key.findall(message.Subject)
            if not names:
                continue
            name = names[0][2:]

            if (
                (message.SenderName in ["DoNotReply", "Yong-Sok SHIN [신용석]"])
                and ("위임" in message.Subject)
                and (name in ["이여름", "한송희"])
            ):
                if "거절결정서" in message.Subject:
                    rejection_type = "FR"
                else:
                    rejection_type = "OA"

                self.add_calendar_event(message, rejection_type)

    def add_calendar_event(self, message, rejection_type):
        # Outlook의 Calendar에 접근
        calendar = Handler_Class.outlook.GetNamespace("MAPI").GetDefaultFolder(
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
            # datetime.strptime은 datetime 객체를 반환하므로 .date()로 변환
            target_date = datetime.strptime(target_date_str, "%Y-%m-%d").date()

            utc_timezone = pytz.utc
            target_date_with_time = datetime.combine(target_date, time(0, 0, 0))
            target_date_utc = utc_timezone.localize(target_date_with_time)
            appointment.Start = target_date_utc

            four_months_later = target_date + relativedelta(months=4)

            if rejection_type == "OA":
                appointment2.Subject = f"{ref_num}_의견서"
                response_date = four_months_later - timedelta(days=11)
            else:
                appointment2.Subject = f"{ref_num}_재심사"
                # 이메일 수신 날짜(CreationTime) 기준 3개월 후로 설정
                receipt_date = message.CreationTime.date()
                response_date = receipt_date + relativedelta(months=3)

            response_date_with_time = datetime.combine(response_date, time(0, 0, 0))
            response_date_utc = utc_timezone.localize(response_date_with_time)
            appointment2.Start = response_date_utc

        else:
            # 날짜를 찾지 못한 경우 이메일 생성 시간으로 대체
            appointment.Start = message.CreationTime
            appointment2.Start = message.CreationTime

            # 개선 #5: else 브랜치에서도 rejection_type에 따라 제목 구분
            if rejection_type == "OA":
                appointment2.Subject = f"{ref_num}_의견서"
            else:
                appointment2.Subject = f"{ref_num}_재심사"

        appointment.AllDayEvent = True
        appointment.Body = f"Task details: {message.Body}"

        appointment2.AllDayEvent = True
        appointment2.Body = f"Task details: {message.Body}"

        appointment.Save()
        appointment2.Save()

        # UI 개선 #5: 활성 창을 부모로 지정 (Handler_Class에서 MainWin 직접 접근 불가)
        parent_win = qtw.QApplication.activeWindow()
        messageBox = qtw.QMessageBox(parent_win)
        messageBox.setWindowTitle("Schedule Creation Complete")
        if target_date and response_date:
            # 개선 #4: datetime 객체를 날짜 문자열로 포맷
            target_str = target_date.strftime("%Y-%m-%d")
            response_str = response_date.strftime("%Y-%m-%d")
            messageBox.setText(
                f"{target_str}: {ref_num}\n{response_str}: {ref_num}\nSchedule Creation Complete"
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

        # UI 개선 #1: QLabel 대신 QListWidget 사용 (스크롤 가능, 메시지 많아도 잘리지 않음)
        self.list_widget = qtw.QListWidget()
        for attach in self.attach_list:
            self.list_widget.addItem(attach)
        notice_layout.addWidget(self.list_widget)
        notice.setLayout(notice_layout)

        transfer_btn = qtw.QPushButton("File Transfer")
        quit_btn = qtw.QPushButton("Quit")

        win_layout = qtw.QVBoxLayout()
        win_layout.addWidget(notice)
        win_layout.addWidget(transfer_btn)
        win_layout.addWidget(quit_btn)

        transfer_btn.clicked.connect(self.transfer_clicked)
        quit_btn.clicked.connect(self.quit_clicked)

        self.setLayout(win_layout)
        self.setWindowTitle("Task Manager")
        self.setWindowIcon(qtg.QIcon(resource_path("icon.png")))
        self.resize(400, 300)

        # UI 개선 #2: 하드코딩된 move(300,300) 대신 화면 중앙에 배치
        screen_geometry = qtw.QApplication.primaryScreen().availableGeometry()
        self.move(
            screen_geometry.center().x() - self.width() // 2,
            screen_geometry.center().y() - self.height() // 2,
        )

        # UI 개선 #4: 트레이 아이콘 더블클릭으로 창 표시
        self.tray_icon = qtw.QSystemTrayIcon(qtg.QIcon(resource_path("icon.png")), self)
        menu = qtw.QMenu()

        action1 = menu.addAction("Show")
        action1.triggered.connect(lambda: self.show())
        action2 = menu.addAction("Quit")
        action2.triggered.connect(lambda: sys.exit())

        self.tray_icon.setContextMenu(menu)
        self.tray_icon.activated.connect(self._on_tray_activated)
        self.tray_icon.show()

        # UI 개선 #6: 레이아웃과 트레이 설정이 모두 끝난 후 show() 호출
        self.show()

    def _on_tray_activated(self, reason):
        """트레이 아이콘 더블클릭 시 창을 표시합니다."""
        if reason == qtw.QSystemTrayIcon.ActivationReason.DoubleClick:
            self.show()
            self.activateWindow()

    def update_list(self, messages):
        """메시지 목록을 갱신합니다."""
        self.list_widget.clear()
        for msg in messages:
            self.list_widget.addItem(msg)

    def closeEvent(self, event):
        event.ignore()
        self.hide()

    def transfer_clicked(self):
        if not os.path.exists(parent_path):
            # UI 개선 #3: raise 대신 QMessageBox.critical()로 사용자에게 에러 표시
            qtw.QMessageBox.critical(self, "경로 오류", f"{parent_path} 경로가 존재하지 않습니다.")
            return

        completion_messages = ""

        if not os.listdir(parent_path):
            self.update_list(["이전할 파일이 없습니다."])
            return

        for date_folder in os.listdir(parent_path):
            folder_path = os.path.join(parent_path, date_folder)

            # 개선 #3: 파일이 섞여 있을 경우를 대비해 폴더 여부 확인
            if not os.path.isdir(folder_path):
                continue

            for case_folder in os.listdir(folder_path):
                origin_path = os.path.join(folder_path, case_folder)

                # 수정 #1: case_folder가 파일이면 건너뜀 (NotADirectoryError 방지)
                if not os.path.isdir(origin_path):
                    continue

                try:
                    # 개선 #2: split("_", 2)로 이름에 _ 포함된 경우도 안전하게 처리
                    time_info, ref_info, person_info = case_folder.split("_", 2)
                    target_folder_name = (
                        f"{ref_info}\\{date_folder}_{time_info}_{person_info}"
                    )
                    label = f"{ref_info}_{person_info}"
                except ValueError:
                    target_folder_name = f"{case_folder}"
                    label = case_folder

                create_dir_if_not_exists(os.path.join(target_path, target_folder_name))

                failed = False
                for file_name in os.listdir(origin_path):
                    file_path = os.path.join(origin_path, file_name)
                    # 수정 #2: os.rename() 예외 처리 — 권한 오류·동일 파일명 충돌 대응
                    try:
                        os.rename(
                            file_path,
                            os.path.join(target_path, target_folder_name, file_name),
                        )
                    except OSError as e:
                        print(f"파일 이동 실패 ({file_name}): {e}")
                        failed = True

                if not failed:
                    os.rmdir(origin_path)
                    # 수정 #3: 완료 메시지를 case_folder 단위로 기록 (마지막 건만 표시되던 버그 수정)
                    completion_messages += f"{label} 이전 작업 완료\n"
                else:
                    completion_messages += f"{label} 이전 실패 (일부 파일 이동 오류)\n"

            # 비어있으면 날짜 폴더도 삭제
            if not os.listdir(folder_path):
                os.rmdir(folder_path)

        self.update_list(completion_messages.strip().splitlines())

        # UI 개선 #5: 부모 위젯 self 지정
        messageBox = qtw.QMessageBox(self)
        messageBox.setWindowTitle("File Transfer")
        messageBox.setText("파일 이전 작업이 완료되었습니다.")
        messageBox.exec()

    def quit_clicked(self):
        sys.exit()


if __name__ == "__main__":
    # 개선 #8: outlook 인스턴스를 함수/클래스에 명시적으로 전달
    outlook_mapi = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    outlook_new = win32com.client.DispatchWithEvents(
        "Outlook.Application", Handler_Class
    )
    Handler_Class.outlook = outlook_new  # 클래스 변수로 주입

    attach_list = save_attachments(outlook_mapi)
    app = qtw.QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    mw = MainWin(attach_list)
    Handler_Class.main_win = mw  # UI 업데이트를 위한 MainWin 인스턴스 주입
    sys.exit(app.exec())
