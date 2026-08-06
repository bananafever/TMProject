import os
import re
import sys
import json
import logging
import shutil
import stat
from datetime import datetime, time, timedelta
from time import sleep

import pytz
import win32com.client
from dateutil.relativedelta import relativedelta
from PySide6 import QtCore as qtc
from PySide6 import QtGui as qtg
from PySide6 import QtWidgets as qtw
from PySide6 import QtSvg as qts

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


NAV_ICONS = {
    "tasks": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20">
        <rect x="3" y="3" width="14" height="14" rx="3" fill="none" stroke="{color}" stroke-width="1.6"/>
        <path d="M6.5 10 L9 12.5 L14 7" fill="none" stroke="{color}" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>""",
    "history": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20">
        <circle cx="10" cy="10" r="7" fill="none" stroke="{color}" stroke-width="1.6"/>
        <path d="M10 6 V10 L13 12" fill="none" stroke="{color}" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>""",
    "settings": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20">
        <path d="M3 6 H17 M3 10 H17 M3 14 H17" stroke="{color}" stroke-width="1.6" stroke-linecap="round"/>
        <circle cx="7" cy="6" r="1.8" fill="{color}"/>
        <circle cx="13" cy="10" r="1.8" fill="{color}"/>
        <circle cx="9" cy="14" r="1.8" fill="{color}"/>
        </svg>""",
}


def make_icon(svg_key, color="#3A3A38", size=20):
    """NAV_ICONS의 인라인 SVG 문자열을 렌더링해 QIcon으로 변환합니다."""
    svg_data = NAV_ICONS[svg_key].format(color=color).encode("utf-8")
    renderer = qts.QSvgRenderer(qtc.QByteArray(svg_data))
    pixmap = qtg.QPixmap(size, size)
    pixmap.fill(qtc.Qt.transparent)
    painter = qtg.QPainter(pixmap)
    painter.setRenderHint(qtg.QPainter.Antialiasing, True)
    renderer.render(painter)
    painter.end()
    return qtg.QIcon(pixmap)


def setup_app_fonts(app):
    """assets/fonts에 번들된 Pretendard를 등록하고 앱 전역 글꼴로 지정합니다."""
    font_dir = resource_path(os.path.join("assets", "fonts"))
    if os.path.isdir(font_dir):
        for f in os.listdir(font_dir):
            if f.lower().endswith((".ttf", ".otf")):
                qtg.QFontDatabase.addApplicationFont(os.path.join(font_dir, f))

    font = app.font()
    font.setFamily("Pretendard")
    font.setStyleStrategy(qtg.QFont.PreferAntialias)
    font.setHintingPreference(qtg.QFont.PreferNoHinting)
    app.setFont(font)


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


def safe_remove_path(path, retries=5, delay=0.1):
    """파일 또는 디렉토리의 읽기 전용 속성을 해제하고 재시도 로직을 포함하여 안전하게 삭제합니다."""
    if not os.path.exists(path):
        return True

    def _remove_readonly(func, p, exc_info):
        try:
            os.chmod(p, stat.S_IWRITE)
            func(p)
        except Exception:
            pass

    for i in range(retries):
        try:
            os.chmod(path, stat.S_IWRITE)
            if os.path.isdir(path):
                for root, dirs, files in os.walk(path, topdown=False):
                    for name in files:
                        fp = os.path.join(root, name)
                        try:
                            os.chmod(fp, stat.S_IWRITE)
                        except Exception:
                            pass
                    for name in dirs:
                        dp = os.path.join(root, name)
                        try:
                            os.chmod(dp, stat.S_IWRITE)
                        except Exception:
                            pass
                shutil.rmtree(path, onerror=_remove_readonly)
            else:
                os.remove(path)
            return True
        except Exception as e:
            if i < retries - 1:
                sleep(delay)
            else:
                raise e


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


class ScheduleEditDialog(qtw.QDialog):
    """캘린더 일정 자동 등록 시 날짜 확인 및 수정을 위한 모달 다이얼로그"""

    def __init__(self, parent=None, ref_num="", subj_comments="", target_date=None, subj_action="", response_date=None):
        super().__init__(parent)
        self.setWindowTitle("Schedule Registration & Edit")
        self.resize(440, 280)

        self.ref_num = ref_num
        self.subj_comments = subj_comments
        self.target_date = target_date
        self.subj_action = subj_action
        self.response_date = response_date

        self.initUI()

    def initUI(self):
        layout = qtw.QVBoxLayout(self)

        info_label = qtw.QLabel(
            f"<b>[Ref: {self.ref_num}]</b> 캘린더 자동 등록 일정 확인<br>"
            "등록할 일정을 확인하고 날짜를 필요에 맞게 수정하세요."
        )
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        # 1. Comments 일정 그룹
        comments_group = qtw.QGroupBox("Comments 일정 (Folder: Mail)")
        comments_layout = qtw.QHBoxLayout()

        self.chk_comments = qtw.QCheckBox("등록")
        self.chk_comments.setChecked(True)

        lbl_comments = qtw.QLabel(self.subj_comments)
        lbl_comments.setMinimumWidth(140)

        self.date_comments = qtw.QDateEdit()
        self.date_comments.setCalendarPopup(True)
        self.date_comments.setDisplayFormat("yyyy-MM-dd")
        if self.target_date:
            self.date_comments.setDate(qtc.QDate(self.target_date.year, self.target_date.month, self.target_date.day))

        comments_layout.addWidget(self.chk_comments)
        comments_layout.addWidget(lbl_comments)
        comments_layout.addWidget(self.date_comments)
        comments_group.setLayout(comments_layout)
        layout.addWidget(comments_group)

        # 2. Action 일정 그룹
        action_group = qtw.QGroupBox("Action 일정 (Folder: No_Instructions)")
        action_layout = qtw.QHBoxLayout()

        self.chk_action = qtw.QCheckBox("등록")
        self.chk_action.setChecked(True)

        lbl_action = qtw.QLabel(self.subj_action)
        lbl_action.setMinimumWidth(140)

        self.date_action = qtw.QDateEdit()
        self.date_action.setCalendarPopup(True)
        self.date_action.setDisplayFormat("yyyy-MM-dd")
        if self.response_date:
            self.date_action.setDate(qtc.QDate(self.response_date.year, self.response_date.month, self.response_date.day))

        action_layout.addWidget(self.chk_action)
        action_layout.addWidget(lbl_action)
        action_layout.addWidget(self.date_action)
        action_group.setLayout(action_layout)
        layout.addWidget(action_group)

        self.chk_comments.toggled.connect(self.date_comments.setEnabled)
        self.chk_action.toggled.connect(self.date_action.setEnabled)

        btn_box = qtw.QHBoxLayout()
        btn_box.addStretch()

        self.btn_confirm = qtw.QPushButton("등록")
        self.btn_confirm.clicked.connect(self.accept)

        self.btn_cancel = qtw.QPushButton("취소")
        self.btn_cancel.setObjectName("quitButton")
        self.btn_cancel.clicked.connect(self.reject)

        btn_box.addWidget(self.btn_confirm)
        btn_box.addWidget(self.btn_cancel)
        layout.addLayout(btn_box)

    def get_results(self):
        """다이얼로그에서 설정한 결과 반환
        Returns:
            (enabled_comments, target_date, enabled_action, response_date)
        """
        qdate_comments = self.date_comments.date()
        qdate_action = self.date_action.date()

        py_target_date = qdate_comments.toPython()
        py_response_date = qdate_action.toPython()

        return (
            self.chk_comments.isChecked(),
            py_target_date,
            self.chk_action.isChecked(),
            py_response_date
        )


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

        if match:
            target_date_str = match.group(1)
            target_date = datetime.strptime(target_date_str, "%Y-%m-%d").date()

            four_months_later = target_date + relativedelta(months=4)

            if rejection_type == "OA":
                response_date = four_months_later - timedelta(days=11)
            else:
                receipt_date = message.CreationTime.date()
                response_date = receipt_date + relativedelta(months=3)
        else:
            target_date = message.CreationTime.date()
            response_date = message.CreationTime.date()

        # 알림창(대화상자)을 띄워 사용자에게 날짜 확인 및 수정 기회 제공
        parent_win = qtw.QApplication.activeWindow()
        dialog = ScheduleEditDialog(
            parent=parent_win,
            ref_num=ref_num,
            subj_comments=subj_comments,
            target_date=target_date,
            subj_action=subj_action,
            response_date=response_date
        )

        if dialog.exec() != qtw.QDialog.DialogCode.Accepted:
            logging.info(f"캘린더 일정 등록이 사용자에 의해 취소되었습니다. Ref: {ref_num}")
            return

        enabled_comments, final_target_date, enabled_action, final_response_date = dialog.get_results()

        utc_timezone = pytz.utc
        msg_parts = []

        if enabled_comments:
            target_date_utc = utc_timezone.localize(datetime.combine(final_target_date, time(0, 0, 0)))
            dup_comments = check_duplicate_appointment(subfolder, subj_comments, final_target_date)
            if not dup_comments:
                appointment = subfolder.Items.Add()
                appointment.Subject = subj_comments
                appointment.Start = target_date_utc
                appointment.AllDayEvent = True
                appointment.Body = f"Task details: {message.Body}"
                appointment.Save()
                msg_parts.append(f"추가됨: {subj_comments} ({final_target_date})")
                logging.info(f"캘린더 일정 추가됨: {subj_comments} ({final_target_date})")
            else:
                msg_parts.append(f"건너뜀(중복): {subj_comments} ({final_target_date})")
                logging.info(f"캘린더 일정 건너뜀(중복): {subj_comments} ({final_target_date})")

        if enabled_action:
            response_date_utc = utc_timezone.localize(datetime.combine(final_response_date, time(0, 0, 0)))
            dup_action = check_duplicate_appointment(subfolder2, subj_action, final_response_date)
            if not dup_action:
                appointment2 = subfolder2.Items.Add()
                appointment2.Subject = subj_action
                appointment2.Start = response_date_utc
                appointment2.AllDayEvent = True
                appointment2.Body = f"Task details: {message.Body}"
                appointment2.Save()
                msg_parts.append(f"추가됨: {subj_action} ({final_response_date})")
                logging.info(f"캘린더 일정 추가됨: {subj_action} ({final_response_date})")
            else:
                msg_parts.append(f"건너뜀(중복): {subj_action} ({final_response_date})")
                logging.info(f"캘린더 일정 건너뜀(중복): {subj_action} ({final_response_date})")

        if msg_parts:
            result_box = qtw.QMessageBox(parent_win)
            result_box.setWindowTitle("Schedule Creation Result")
            result_box.setText("\n".join(msg_parts))
            result_box.exec()


class MainWin(qtw.QWidget):

    def __init__(self, attach_list):
        super().__init__()
        self.attach_list = attach_list
        self.initUI()

    def initUI(self):
        # 1. Tasks 탭 생성 (카드형 대시보드)
        self.tab_tasks = qtw.QWidget()
        tasks_layout = qtw.QVBoxLayout(self.tab_tasks)
        tasks_layout.setContentsMargins(18, 18, 18, 18)
        tasks_layout.setSpacing(12)

        top_row = qtw.QHBoxLayout()
        self.tip_label = qtw.QLabel()
        self.tip_label.setObjectName("tipBox")
        self.tip_label.setWordWrap(True)

        transfer_btn = qtw.QPushButton("File Transfer")
        transfer_btn.clicked.connect(self.transfer_clicked)

        top_row.addWidget(self.tip_label, 1)
        top_row.addWidget(transfer_btn, 0, qtc.Qt.AlignTop)
        tasks_layout.addLayout(top_row)

        self.cards_container = qtw.QWidget()
        self.cards_container.setObjectName("cardsContainer")
        self.cards_layout = qtw.QVBoxLayout(self.cards_container)
        self.cards_layout.setContentsMargins(0, 0, 0, 0)
        self.cards_layout.setSpacing(8)
        self.cards_layout.addStretch()

        self.cards_scroll = qtw.QScrollArea()
        self.cards_scroll.setWidgetResizable(True)
        self.cards_scroll.setFrameShape(qtw.QFrame.NoFrame)
        self.cards_scroll.setWidget(self.cards_container)
        tasks_layout.addWidget(self.cards_scroll, 1)

        for attach in self.attach_list:
            self.add_table_row(*attach)
        self._update_tip()

        # 2. History 탭 생성
        self.tab_history = qtw.QWidget()
        history_layout = qtw.QVBoxLayout(self.tab_history)
        history_layout.setContentsMargins(18, 18, 18, 18)
        history_layout.setSpacing(12)

        refresh_btn = qtw.QPushButton("Refresh History")
        refresh_btn.clicked.connect(self.load_history)

        btn_layout_history = qtw.QHBoxLayout()
        btn_layout_history.addStretch()
        btn_layout_history.addWidget(refresh_btn)
        history_layout.addLayout(btn_layout_history)

        self.history_text = qtw.QTextEdit()
        self.history_text.setReadOnly(True)
        history_layout.addWidget(self.history_text)

        self.load_history()

        # 3. Settings 탭 생성
        self.tab_settings = qtw.QWidget()
        settings_main_layout = qtw.QVBoxLayout(self.tab_settings)
        settings_main_layout.setContentsMargins(18, 18, 18, 18)
        settings_main_layout.setSpacing(12)

        save_settings_btn = qtw.QPushButton("Save Settings")
        save_settings_btn.clicked.connect(self.save_settings_ui)

        btn_layout_settings = qtw.QHBoxLayout()
        btn_layout_settings.addStretch()
        btn_layout_settings.addWidget(save_settings_btn)
        settings_main_layout.addLayout(btn_layout_settings)

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

        settings_layout.addRow("Parent Path (Tasks):", parent_box)
        settings_layout.addRow("Target Path (Review):", target_box)

        settings_main_layout.addLayout(settings_layout)

        # 좌측 아이콘 사이드바 + 우측 콘텐츠 스택
        self.stack = qtw.QStackedWidget()
        self.stack.addWidget(self.tab_tasks)
        self.stack.addWidget(self.tab_history)
        self.stack.addWidget(self.tab_settings)

        self.sidebar = qtw.QWidget()
        self.sidebar.setObjectName("navBar")
        self.sidebar.setFixedWidth(128)
        side_layout = qtw.QVBoxLayout(self.sidebar)
        side_layout.setContentsMargins(8, 20, 8, 20)
        side_layout.setSpacing(10)

        self.nav_group = qtw.QButtonGroup(self)
        self.nav_group.setExclusive(True)

        nav_specs = [("tasks", "Tasks"), ("history", "History"), ("settings", "Settings")]
        for i, (icon_key, label) in enumerate(nav_specs):
            btn = qtw.QToolButton()
            btn.setObjectName("navBtn")
            btn.setCheckable(True)
            btn.setToolButtonStyle(qtc.Qt.ToolButtonTextUnderIcon)
            btn.setIcon(make_icon(icon_key))
            btn.setIconSize(qtc.QSize(30, 30))
            btn.setText(label)
            btn.setFixedSize(108, 76)
            side_layout.addWidget(btn, 0, qtc.Qt.AlignHCenter)
            self.nav_group.addButton(btn, i)
        side_layout.addStretch()

        self.nav_group.idClicked.connect(self.stack.setCurrentIndex)
        self.nav_group.button(0).setChecked(True)

        content_row = qtw.QHBoxLayout()
        content_row.setContentsMargins(0, 0, 0, 0)
        content_row.setSpacing(0)
        content_row.addWidget(self.sidebar)
        content_row.addWidget(self.stack, 1)

        # 메인 레이아웃 구성
        win_layout = qtw.QVBoxLayout()
        win_layout.setContentsMargins(0, 0, 0, 0)
        win_layout.setSpacing(0)
        win_layout.addLayout(content_row, 1)

        quit_row = qtw.QHBoxLayout()
        quit_row.setContentsMargins(18, 8, 18, 14)
        quit_btn = qtw.QPushButton("Quit")
        quit_btn.setObjectName("quitButton")
        quit_btn.clicked.connect(self.quit_clicked)
        quit_row.addStretch()
        quit_row.addWidget(quit_btn)
        win_layout.addLayout(quit_row)

        self.setLayout(win_layout)
        self.setWindowTitle("Task Manager")
        self.setWindowIcon(qtg.QIcon(resource_path("icon.png")))
        self.resize(560, 460)

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
        """카드 목록을 모두 지우고 새 데이터를 추가합니다."""
        self._clear_cards()
        for data in messages_data:
            self.add_table_row(*data)
        self._update_tip()

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
        """메시지 카드를 하나 추가합니다."""
        row_idx = self.cards_layout.count() - 1  # 마지막 stretch 앞에 삽입

        card = qtw.QFrame()
        card.setObjectName("msgCard")
        card_layout = qtw.QHBoxLayout(card)
        card_layout.setContentsMargins(14, 10, 14, 10)
        card_layout.setSpacing(12)

        badge = qtw.QLabel(str(row_idx + 1))
        badge.setObjectName("cardBadge")
        badge.setFixedSize(28, 28)
        badge.setAlignment(qtc.Qt.AlignCenter)
        card_layout.addWidget(badge)

        info_layout = qtw.QVBoxLayout()
        info_layout.setSpacing(2)
        title = " · ".join(part for part in (ref, name) if part)
        title_label = qtw.QLabel(title or "-")
        title_label.setObjectName("cardTitle")
        subtitle_label = qtw.QLabel(dt or "")
        subtitle_label.setObjectName("cardSubtitle")
        info_layout.addWidget(title_label)
        info_layout.addWidget(subtitle_label)
        card_layout.addLayout(info_layout, 1)

        bg, fg = self._status_pill_colors(status)
        status_label = qtw.QLabel()
        status_label.setFixedWidth(132)
        status_label.setAlignment(qtc.Qt.AlignCenter)
        status_label.setToolTip(status)
        metrics = status_label.fontMetrics()
        status_label.setText(metrics.elidedText(status, qtc.Qt.ElideRight, 132 - 16))
        status_label.setStyleSheet(
            f"background-color:{bg}; color:{fg}; padding:4px 6px; border-radius:8px; font-size:9pt;"
        )
        card_layout.addWidget(status_label, 0, qtc.Qt.AlignVCenter)

        self.cards_layout.insertWidget(row_idx, card)

    def _clear_cards(self):
        """stretch를 제외한 모든 메시지 카드를 제거합니다."""
        while self.cards_layout.count() > 1:
            item = self.cards_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def _update_tip(self):
        count = self.cards_layout.count() - 1
        if count > 0:
            self.tip_label.setText(f"자동으로 처리된 {count}건의 항목이 있습니다.")
        else:
            self.tip_label.setText("표시할 항목이 없습니다.")

    @staticmethod
    def _status_pill_colors(status):
        if any(keyword in status for keyword in ("실패", "오류")):
            return "#FAECE7", "#993C1D"
        if any(keyword in status for keyword in ("Existing", "완료", "Downloaded")):
            return "#E1EEE7", "#1B4B43"
        return "#F2F2F0", "#5F5E5A"

    def closeEvent(self, event):
        event.ignore()
        self.hide()

    def transfer_clicked(self):
        try:
            p_path = config.get("parent_path", "")
            t_path = config.get("target_path", "")
            
            if not p_path or not os.path.exists(p_path):
                qtw.QMessageBox.critical(self, "경로 오류", f"작업 경로(Parent Path)가 존재하지 않습니다:\n{p_path}")
                logging.error(f"경로 오류: {p_path} 경로가 존재하지 않습니다.")
                return

            if not t_path:
                qtw.QMessageBox.critical(self, "경로 오류", "이전 대상 경로(Target Path)가 설정되지 않았습니다.")
                logging.error("경로 오류: target_path가 설정되지 않았습니다.")
                return

            create_dir_if_not_exists(t_path)

            if not os.listdir(p_path):
                self.update_table([("", "", "", "이전할 파일이 없습니다.")])
                qtw.QMessageBox.information(self, "File Transfer", "이전할 파일이 없습니다.")
                return

            completion_messages = []

            for date_folder in os.listdir(p_path):
                folder_path = os.path.join(p_path, date_folder)

                if not os.path.isdir(folder_path):
                    continue

                for case_folder in os.listdir(folder_path):
                    origin_path = os.path.join(folder_path, case_folder)

                    if not os.path.isdir(origin_path):
                        continue

                    try:
                        time_info, ref_info, person_info = case_folder.split("_", 2)
                        target_folder_name = os.path.join(ref_info, f"{date_folder}_{time_info}_{person_info}")
                        label_ref = ref_info
                        label_name = person_info
                    except ValueError:
                        target_folder_name = case_folder
                        label_ref = case_folder
                        label_name = ""

                    dest_dir = os.path.join(t_path, target_folder_name)
                    create_dir_if_not_exists(dest_dir)

                    failed = False
                    files_in_origin = os.listdir(origin_path)
                    
                    for file_name in files_in_origin:
                        file_path = os.path.join(origin_path, file_name)
                        dst_file_path = os.path.join(dest_dir, file_name)
                        try:
                            if os.path.exists(file_path):
                                os.chmod(file_path, stat.S_IWRITE)
                            if os.path.exists(dst_file_path):
                                if os.path.isdir(dst_file_path):
                                    shutil.copytree(file_path, dst_file_path, dirs_exist_ok=True)
                                    safe_remove_path(file_path)
                                else:
                                    os.chmod(dst_file_path, stat.S_IWRITE)
                                    os.replace(file_path, dst_file_path)
                            else:
                                shutil.move(file_path, dest_dir)
                        except Exception as e:
                            logging.error(f"파일/폴더 이동 실패 ({file_name}): {e}")
                            failed = True

                    if not failed:
                        try:
                            safe_remove_path(origin_path)
                            completion_messages.append((date_folder, label_ref, label_name, "이전 완료"))
                        except Exception as e:
                            logging.error(f"폴더 삭제 실패 ({origin_path}): {e}")
                            completion_messages.append((date_folder, label_ref, label_name, "이전 완료(폴더 삭제 잔여)"))
                    else:
                        completion_messages.append((date_folder, label_ref, label_name, "이전 실패(오류)"))

                if os.path.exists(folder_path) and not os.listdir(folder_path):
                    try:
                        safe_remove_path(folder_path)
                    except Exception as e:
                        logging.error(f"날짜 폴더 삭제 실패 ({folder_path}): {e}")

            if completion_messages:
                self.update_table(completion_messages)
                logging.info(f"파일 이전 작업 1회 사이클 완료 ({len(completion_messages)}개 항목 처리)")
                qtw.QMessageBox.information(self, "File Transfer", f"파일 이전 작업이 완료되었습니다.\n(총 {len(completion_messages)}개 항목 처리)")
            else:
                self.update_table([("", "", "", "이전할 파일이 없습니다.")])
                logging.info("파일 이전 작업 완료: 이전할 항목 없음")
                qtw.QMessageBox.information(self, "File Transfer", "이전할 파일이 없습니다.")

        except Exception as e:
            logging.error(f"File Transfer 중 오류 발생: {e}", exc_info=True)
            qtw.QMessageBox.critical(self, "오류 발생", f"파일 이전 중 오류가 발생했습니다:\n{e}")

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
    setup_app_fonts(app)

    # pdfProject Clean Light 테마 참고 (웜 뉴트럴 캔버스 + 틸/오렌지 액센트)
    app.setStyleSheet("""
        QWidget {
            background-color: #F6F6F4;
            color: #333333;
            font-family: 'Pretendard', 'Malgun Gothic', sans-serif;
            font-size: 10pt;
        }
        #navBar {
            background-color: #F6F6F4;
            border-right: 1px solid #DCE3DF;
        }
        QToolButton#navBtn {
            background: transparent;
            border: 2px solid transparent;
            border-radius: 12px;
            padding: 6px;
            color: #4B4A47;
            font-size: 10pt;
            text-align: center;
        }
        QToolButton#navBtn:checked {
            background-color: white;
            border: 2px solid #E8611C;
            color: #232220;
            font-weight: bold;
        }
        QToolButton#navBtn:hover:!checked {
            background-color: #ECEAE5;
        }
        #tipBox {
            background-color: #E1EEE7;
            color: #1B4B43;
            padding: 10px 14px;
            border-radius: 10px;
            font-size: 9.5pt;
        }
        #msgCard {
            background-color: white;
            border: 1px solid #DCE3DF;
            border-radius: 12px;
        }
        #cardBadge {
            background-color: #26695E;
            color: white;
            border-radius: 14px;
            font-size: 9pt;
            font-weight: bold;
        }
        #cardTitle {
            color: #232220;
            font-size: 10pt;
            font-weight: bold;
        }
        #cardSubtitle {
            color: #8A877E;
            font-size: 9pt;
        }
        QPushButton {
            background-color: #26695E;
            color: white;
            border: none;
            border-radius: 10px;
            padding: 8px 16px;
            font-weight: bold;
        }
        QPushButton:hover {
            background-color: #2F7C6E;
        }
        QPushButton:pressed {
            background-color: #1B4B43;
        }
        QTextEdit, QLineEdit {
            background-color: white;
            border: 1px solid #DCE3DF;
            border-radius: 8px;
            padding: 4px;
        }
        QTextEdit:focus, QLineEdit:focus {
            border: 1px solid #26695E;
        }
        QGroupBox {
            font-weight: bold;
            border: 1px solid #DCE3DF;
            border-radius: 10px;
            margin-top: 10px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            subcontrol-position: top left;
            padding: 0 3px;
            left: 10px;
        }
        QScrollArea, #cardsContainer {
            background-color: transparent;
            border: none;
        }
        #quitButton {
            background-color: white;
            color: #993C1D;
            border: 1px solid #F0997B;
            border-radius: 10px;
            font-weight: normal;
        }
        #quitButton:hover {
            background-color: #FAECE7;
        }
        #quitButton:pressed {
            background-color: #F5C4B3;
        }
    """)

    app.setQuitOnLastWindowClosed(False)
    mw = MainWin(attach_list)
    Handler_Class.main_win = mw  # UI 업데이트를 위한 MainWin 인스턴스 주입
    sys.exit(app.exec())
