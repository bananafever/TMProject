import os
import re
import sys
import json
import logging
from datetime import datetime, time, timedelta

import pytz
import win32com.client
from dateutil.relativedelta import relativedelta
from PySide6 import QtGui as qtg
from PySide6 import QtWidgets as qtw

# 로깅 설정
logging.basicConfig(
    filename='app.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

ref_num_key = re.compile(
    r"[a-zA-Z][a-zA-Z][a-zA-Z]?-[tT]?2[a-zA-Z][0-9]{4}[a-zA-Z0-9-]{0,5}|"
    r"[a-zA-Z][a-zA-Z][a-zA-Z]?-[0-9]{5}"
)
name_key = re.compile(r"; [ㄱ-ㅎ|ㅏ-ㅣ|가-힣]+(?=\()")

CONFIG_FILE = "settings.json"
DEFAULT_CONFIG = {
    "parent_path": r"C:\OneDrive - WorkSpace\1_MyWork\1_Project_Work\Tasks",
    "target_path": r"C:\OneDrive - WorkSpace\1_MyWork\04_Archives_Work\02_Review"
}
config = DEFAULT_CONFIG.copy()

def load_config():
    global config
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
                config.update(loaded)
                logging.info("설정 파일을 성공적으로 불러왔습니다.")
        except Exception as e:
            logging.error(f"설정 파일 로드 실패: {e}")
    else:
        save_config()

def save_config():
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
            logging.info("설정 파일을 성공적으로 저장했습니다.")
    except Exception as e:
        logging.error(f"설정 파일 저장 실패: {e}")

load_config()


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
        logging.error("Error: Creating directory. " + directory)
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
        logging.error(f"메시지 정보 추출 실패: {e}")
        return

    date_folder = os.path.join(config["parent_path"], delivery_date)
    create_dir_if_not_exists(date_folder)

    ref_folder = ref_num + "_" + name
    task_folder = os.path.join(date_folder, time_id + "_" + ref_folder)

    time_str = f"{time_id[:2]}:{time_id[4:6]}"
    if create_dir_if_not_exists(task_folder):  # True = 이미 존재
        logging.info(ref_folder + ": Existing")
        if attach_list is not None:
            attach_list.append((f"{delivery_date} {time_str}", ref_num, name, "Already Existing"))
        return

    if attach_list is not None:
        attach_list.append((f"{delivery_date} {time_str}", ref_num, name, "Downloaded"))

    for attachment in message.Attachments:
        file_name = "SYS_" + attachment.FileName
        # 수정 #3: 저장 실패 시 예외를 잡아 나머지 첨부파일 저장 계속 진행
        try:
            attachment.SaveAsFile(os.path.join(task_folder, file_name))
            logging.info(f"첨부파일 저장 성공: {file_name}")
        except Exception as e:
            logging.error(f"첨부파일 저장 실패 ({file_name}): {e}")
            if attach_list is not None:
                attach_list.append((f"{delivery_date} {time_str}", ref_num, name, f"저장 실패: {file_name}"))


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


def check_duplicate_appointment(folder, subject, start_date):
    """지정된 폴더에 동일한 제목과 시작 날짜(Start.date())를 가진 일정이 있는지 확인"""
    items = folder.Items
    items.IncludeRecurrences = False
    try:
        restricted_items = items.Restrict(f"[Subject] = '{subject}'")
        for item in restricted_items:
            if hasattr(item, 'Start') and item.Start.date() == start_date:
                return True
    except Exception as e:
        logging.error(f"중복 체크 중 오류 발생: {e}")
    return False


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
                # 수정 #1: attach_list를 넘겨 저장 결과를 받고 QTableWidget에 항목 추가
                new_items = []
                _save_attachments_for_message(message, new_items)
                if Handler_Class.main_win is not None and new_items:
                    for item in new_items:
                        Handler_Class.main_win.add_table_row(*item)

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

        # Ref 번호 추출
        ref_nums = ref_num_key.findall(message.Subject)
        if not ref_nums:
            logging.warning(f"Ref 번호를 찾을 수 없습니다: {message.Subject}")
            return
        ref_num = ref_nums[0]
        
        subj_comments = f"{ref_num}_Comments"
        if rejection_type == "OA":
            subj_action = f"{ref_num}_의견서"
        else:
            subj_action = f"{ref_num}_재심사"

        # 이메일 제목에서 날짜 추출
        date_pattern = r"(\d{4}-\d{2}-\d{2})"
        match = re.search(date_pattern, message.Subject)

        target_date = None
        response_date = None
        utc_timezone = pytz.utc

        if match:
            target_date_str = match.group(1)
            target_date = datetime.strptime(target_date_str, "%Y-%m-%d").date()
            target_date_utc = utc_timezone.localize(datetime.combine(target_date, time(0, 0, 0)))

            four_months_later = target_date + relativedelta(months=4)

            if rejection_type == "OA":
                response_date = four_months_later - timedelta(days=11)
            else:
                receipt_date = message.CreationTime.date()
                response_date = receipt_date + relativedelta(months=3)
                
            response_date_utc = utc_timezone.localize(datetime.combine(response_date, time(0, 0, 0)))
        else:
            target_date = message.CreationTime.date()
            target_date_utc = message.CreationTime
            response_date = message.CreationTime.date()
            response_date_utc = message.CreationTime

        dup_comments = check_duplicate_appointment(subfolder, subj_comments, target_date)
        dup_action = check_duplicate_appointment(subfolder2, subj_action, response_date)
        
        msg_parts = []

        if not dup_comments:
            appointment = subfolder.Items.Add()
            appointment.Subject = subj_comments
            appointment.Start = target_date_utc
            appointment.AllDayEvent = True
            appointment.Body = f"Task details: {message.Body}"
            appointment.Save()
            msg_parts.append(f"추가됨: {subj_comments} ({target_date})")
            logging.info(f"캘린더 일정 추가됨: {subj_comments} ({target_date})")
        else:
            msg_parts.append(f"건너뜀(중복): {subj_comments} ({target_date})")
            logging.info(f"캘린더 일정 건너뜀(중복): {subj_comments} ({target_date})")

        if not dup_action:
            appointment2 = subfolder2.Items.Add()
            appointment2.Subject = subj_action
            appointment2.Start = response_date_utc
            appointment2.AllDayEvent = True
            appointment2.Body = f"Task details: {message.Body}"
            appointment2.Save()
            msg_parts.append(f"추가됨: {subj_action} ({response_date})")
            logging.info(f"캘린더 일정 추가됨: {subj_action} ({response_date})")
        else:
            msg_parts.append(f"건너뜀(중복): {subj_action} ({response_date})")
            logging.info(f"캘린더 일정 건너뜀(중복): {subj_action} ({response_date})")

        # UI 개선 #5: 활성 창을 부모로 지정 (Handler_Class에서 MainWin 직접 접근 불가)
        parent_win = qtw.QApplication.activeWindow()
        messageBox = qtw.QMessageBox(parent_win)
        messageBox.setWindowTitle("Schedule Creation Result")
        messageBox.setText("\n".join(msg_parts))
        messageBox.exec()


class MainWin(qtw.QWidget):

    def __init__(self, attach_list):
        super().__init__()
        self.attach_list = attach_list
        self.initUI()

    def initUI(self):
        self.tabs = qtw.QTabWidget()

        # 1. Tasks 탭 생성
        self.tab_tasks = qtw.QWidget()
        tasks_layout = qtw.QVBoxLayout()
        
        notice = qtw.QGroupBox("Messages")
        notice_layout = qtw.QVBoxLayout()

        # UI 개선 #1: QListWidget 대신 QTableWidget 사용
        self.table_widget = qtw.QTableWidget()
        self.table_widget.setColumnCount(4)
        self.table_widget.setHorizontalHeaderLabels(["일시", "Ref 번호", "담당자", "상태"])
        
        header = self.table_widget.horizontalHeader()
        header.setSectionResizeMode(0, qtw.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, qtw.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, qtw.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, qtw.QHeaderView.Stretch)
        
        self.table_widget.setEditTriggers(qtw.QAbstractItemView.NoEditTriggers)
        self.table_widget.setSelectionBehavior(qtw.QAbstractItemView.SelectRows)

        for attach in self.attach_list:
            self.add_table_row(*attach)
            
        notice_layout.addWidget(self.table_widget)
        notice.setLayout(notice_layout)

        transfer_btn = qtw.QPushButton("File Transfer")
        transfer_btn.clicked.connect(self.transfer_clicked)
        
        tasks_layout.addWidget(notice)
        tasks_layout.addWidget(transfer_btn)
        self.tab_tasks.setLayout(tasks_layout)

        # 2. History 탭 생성
        self.tab_history = qtw.QWidget()
        history_layout = qtw.QVBoxLayout()
        self.history_text = qtw.QTextEdit()
        self.history_text.setReadOnly(True)
        refresh_btn = qtw.QPushButton("Refresh History")
        refresh_btn.clicked.connect(self.load_history)
        history_layout.addWidget(self.history_text)
        history_layout.addWidget(refresh_btn)
        self.tab_history.setLayout(history_layout)
        self.load_history()

        # 3. Settings 탭 생성
        self.tab_settings = qtw.QWidget()
        settings_layout = qtw.QFormLayout()
        
        self.parent_path_edit = qtw.QLineEdit(config["parent_path"])
        parent_btn = qtw.QPushButton("Browse...")
        parent_btn.clicked.connect(self.browse_parent)
        parent_box = qtw.QHBoxLayout()
        parent_box.addWidget(self.parent_path_edit)
        parent_box.addWidget(parent_btn)
        
        self.target_path_edit = qtw.QLineEdit(config["target_path"])
        target_btn = qtw.QPushButton("Browse...")
        target_btn.clicked.connect(self.browse_target)
        target_box = qtw.QHBoxLayout()
        target_box.addWidget(self.target_path_edit)
        target_box.addWidget(target_btn)
        
        save_settings_btn = qtw.QPushButton("Save Settings")
        save_settings_btn.clicked.connect(self.save_settings_ui)
        
        settings_layout.addRow("Parent Path (Tasks):", parent_box)
        settings_layout.addRow("Target Path (Review):", target_box)
        settings_layout.addRow("", save_settings_btn)
        self.tab_settings.setLayout(settings_layout)

        # 탭 추가
        self.tabs.addTab(self.tab_tasks, "Tasks")
        self.tabs.addTab(self.tab_history, "History")
        self.tabs.addTab(self.tab_settings, "Settings")

        # 메인 레이아웃 구성
        win_layout = qtw.QVBoxLayout()
        win_layout.addWidget(self.tabs)
        
        quit_btn = qtw.QPushButton("Quit")
        quit_btn.clicked.connect(self.quit_clicked)
        win_layout.addWidget(quit_btn)

        self.setLayout(win_layout)
        self.setWindowTitle("Task Manager")
        self.setWindowIcon(qtg.QIcon(resource_path("icon.png")))
        self.resize(500, 400)

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

    def update_table(self, messages_data):
        """테이블 내용을 모두 지우고 새 데이터를 추가합니다."""
        self.table_widget.setRowCount(0)
        for data in messages_data:
            self.add_table_row(*data)

    def load_history(self):
        try:
            with open('app.log', 'r', encoding='utf-8') as f:
                self.history_text.setPlainText(f.read())
            self.history_text.moveCursor(qtg.QTextCursor.End)
        except Exception as e:
            self.history_text.setPlainText(f"로그 파일을 읽을 수 없습니다: {e}")

    def browse_parent(self):
        directory = qtw.QFileDialog.getExistingDirectory(self, "Select Parent Path")
        if directory:
            self.parent_path_edit.setText(directory)

    def browse_target(self):
        directory = qtw.QFileDialog.getExistingDirectory(self, "Select Target Path")
        if directory:
            self.target_path_edit.setText(directory)

    def save_settings_ui(self):
        config["parent_path"] = self.parent_path_edit.text()
        config["target_path"] = self.target_path_edit.text()
        save_config()
        logging.info(f"설정이 변경되었습니다. Parent: {config['parent_path']}, Target: {config['target_path']}")
        qtw.QMessageBox.information(self, "Settings", "설정이 저장되었습니다.")

    def add_table_row(self, dt, ref, name, status):
        """테이블에 새 행을 추가합니다."""
        row_idx = self.table_widget.rowCount()
        self.table_widget.insertRow(row_idx)
        self.table_widget.setItem(row_idx, 0, qtw.QTableWidgetItem(dt))
        self.table_widget.setItem(row_idx, 1, qtw.QTableWidgetItem(ref))
        self.table_widget.setItem(row_idx, 2, qtw.QTableWidgetItem(name))
        self.table_widget.setItem(row_idx, 3, qtw.QTableWidgetItem(status))

    def closeEvent(self, event):
        event.ignore()
        self.hide()

    def transfer_clicked(self):
        p_path = config["parent_path"]
        t_path = config["target_path"]
        
        if not os.path.exists(p_path):
            # UI 개선 #3: raise 대신 QMessageBox.critical()로 사용자에게 에러 표시
            qtw.QMessageBox.critical(self, "경로 오류", f"{p_path} 경로가 존재하지 않습니다.")
            logging.error(f"경로 오류: {p_path} 경로가 존재하지 않습니다.")
            return

        completion_messages = []

        if not os.listdir(p_path):
            self.update_table([("", "", "", "이전할 파일이 없습니다.")])
            return

        for date_folder in os.listdir(p_path):
            folder_path = os.path.join(p_path, date_folder)

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
                    label_ref = ref_info
                    label_name = person_info
                except ValueError:
                    target_folder_name = f"{case_folder}"
                    label_ref = case_folder
                    label_name = ""

                create_dir_if_not_exists(os.path.join(t_path, target_folder_name))

                failed = False
                for file_name in os.listdir(origin_path):
                    file_path = os.path.join(origin_path, file_name)
                    # 수정 #2: os.rename() 예외 처리 — 권한 오류·동일 파일명 충돌 대응
                    try:
                        os.rename(
                            file_path,
                            os.path.join(t_path, target_folder_name, file_name),
                        )
                    except OSError as e:
                        logging.error(f"파일 이동 실패 ({file_name}): {e}")
                        failed = True

                if not failed:
                    os.rmdir(origin_path)
                    # 수정 #3: 완료 메시지를 튜플로 기록
                    completion_messages.append((date_folder, label_ref, label_name, "이전 완료"))
                else:
                    completion_messages.append((date_folder, label_ref, label_name, "이전 실패(오류)"))

            # 비어있으면 날짜 폴더도 삭제
            if not os.listdir(folder_path):
                os.rmdir(folder_path)

        self.update_table(completion_messages)
        logging.info("파일 이전 작업 1회 사이클 완료")

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
