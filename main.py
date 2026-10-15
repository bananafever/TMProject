import os
import re
import sys
import json
import logging
import shutil
import stat
from datetime import date, datetime, timedelta
from time import sleep

try:
    import win32com.client
except Exception:
    # gen_py 캐시가 손상된 경우 (%LOCALAPPDATA%\Temp\gen_py 등) 자동 삭제 후 재시도
    import tempfile
    gen_py_dirs = [
        os.path.expandvars(r"%LOCALAPPDATA%\Temp\gen_py"),
        os.path.join(tempfile.gettempdir(), "gen_py"),
    ]
    for d in gen_py_dirs:
        if os.path.exists(d):
            shutil.rmtree(d, ignore_errors=True)
    for mod in list(sys.modules.keys()):
        if mod.startswith("win32com"):
            del sys.modules[mod]
    import win32com.client
from dateutil.relativedelta import relativedelta
from PySide6 import QtCore as qtc
from PySide6 import QtGui as qtg
from PySide6 import QtWidgets as qtw
from PySide6 import QtSvg as qts

import kipris
import naver_cal

# ── 데이터 폴더 ─────────────────────────────────────────────────────
# settings.json(경로·약칭표·담당자)·case_cache.json·calendar_state.json·app.log 를 한 곳에 둔다.
# exe 는 시작 프로그램 바로가기로 이 폴더에서 실행된다 — 소스로 실행해도 같은 파일을 쓴다.
DEFAULT_DATA_DIR = r"G:/내 드라이브/bananafever/5_APPS/TaskManager"


def data_dir():
    """환경변수 TASKMANAGER_DATA_DIR → exe 가 있는 폴더 → DEFAULT_DATA_DIR → 소스 폴더(드라이브가 없을 때)."""
    env = os.getenv("TASKMANAGER_DATA_DIR")
    if env:
        return env
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    if os.path.isdir(DEFAULT_DATA_DIR):
        return DEFAULT_DATA_DIR
    return os.path.dirname(os.path.abspath(__file__))


DATA_DIR = data_dir()


def data_path(name):
    return os.path.join(DATA_DIR, name)


LOG_FILE = data_path("app.log")

# 로깅 설정
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

ref_num_key = re.compile(
    r"[a-zA-Z][a-zA-Z][a-zA-Z]?-[tT]?2[a-zA-Z][0-9]{4}[a-zA-Z0-9-]{0,5}|"
    r"[a-zA-Z][a-zA-Z][a-zA-Z]?-[0-9]{5}"
)
name_key = re.compile(r"; [ㄱ-ㅎ|ㅏ-ㅣ|가-힣]+(?=\()")

CONFIG_FILE = data_path("settings.json")
DEFAULT_CONFIG = {
    "parent_path": r"C:/Dropbox/1_Project/Tasks",
    "target_path": r"G:/내 드라이브/bananafever/4_Archives/01_Work/02_Review",
    "prenotice_assignee": "신용석",  # 예고 메일 표에서 읽을 담당자
    # 위임 메일 일정을 등록할 네이버 캘린더 이름 (계정: python naver_cal.py --setup)
    "naver_calendars": {"Mail": "Mail", "No_Instructions": "No_Instructions"},
    # 일정 제목 끝에 붙일 출원인 약칭 (예고 메일의 국문 명칭 → 약칭). 없으면 자동으로 줄인다.
    "applicant_abbr": {},
    # 이 담당자 앞으로 온 위임 메일만 네이버 캘린더에 등록한다
    "calendar_assignees": ["이여름", "한송희"],
}
config = DEFAULT_CONFIG.copy()

def load_config():
    global config
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
                config.update(loaded)
                logging.info(f"설정 파일을 성공적으로 불러왔습니다. ({CONFIG_FILE})")
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


# Outlook Object Model Guard는 MailItem.SenderName 같은 주소 관련 속성을 외부 프로세스가
# 직접 읽을 때마다 "전자 메일 주소 정보에 액세스" 보안 경고를 띄운다.
# Items.Restrict 필터는 보호 대상이 아니므로, 보낸 사람 판정은 필터로 처리한다.
DELEGATION_SENDERS = ["DoNotReply"]  # 위임 메일 첨부 저장 대상 보낸 사람
CALENDAR_SENDERS = ["DoNotReply", "Yong-Sok SHIN [신용석]"]  # 캘린더 등록 대상 보낸 사람
PRENOTICE_SENDERS = ["Chang-Kyu KIM [김창규]"]  # 예고 메일(사건 담당자 지정) 보낸 사람
PRENOTICE_SUBJECT = "사건 담당자 지정"

_DASL_FROMNAME = '"urn:schemas:httpmail:fromname"'
_DASL_SUBJECT = '"urn:schemas:httpmail:subject"'


def _dasl_str(value):
    """DASL 문자열 리터럴로 변환합니다 (작은따옴표는 '' 로 이스케이프)."""
    return "'" + str(value).replace("'", "''") + "'"


def build_sender_filter(sender_names, subject=None):
    """보낸 사람 이름(선택적으로 제목까지)으로 Items.Restrict용 DASL 필터를 만듭니다."""
    if not sender_names:
        raise ValueError("sender_names가 비어 있습니다.")
    sender_cond = " OR ".join(f"{_DASL_FROMNAME} = {_dasl_str(n)}" for n in sender_names)
    query = f"({sender_cond})"
    if subject is not None:
        query += f" AND {_DASL_SUBJECT} = {_dasl_str(subject)}"
    return "@SQL=" + query


def sender_matches(message, sender_names):
    """SenderName을 직접 읽지 않고 message의 보낸 사람이 sender_names 중 하나인지 확인합니다.

    message가 들어 있는 폴더에 보낸 사람+제목 필터를 걸고, 결과에 같은 EntryID가 있는지 봅니다.
    필터가 실패하면 SenderName 직접 비교로 대체합니다 (이 경우 보안 경고가 뜰 수 있음).
    """
    try:
        entry_id = message.EntryID
        items = message.Parent.Items.Restrict(
            build_sender_filter(sender_names, message.Subject)
        )
        return any(item.EntryID == entry_id for item in items)
    except Exception as e:
        logging.error(f"보낸 사람 필터 확인 실패, SenderName 직접 비교로 대체합니다: {e}")
        try:
            return message.SenderName in sender_names
        except Exception as e2:
            logging.error(f"SenderName 확인 실패: {e2}")
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
    # SenderName을 직접 읽으면 Outlook 보안 경고가 뜨므로 Restrict 필터로 보낸 사람을 거른다.
    # 수정 #2: list()로 스냅샷을 찍어 순회 중 Items 변경으로 인한 누락/중복 방지
    messages = list(inbox.Items.Restrict(build_sender_filter(DELEGATION_SENDERS)))
    attach_list = []

    for message in messages:
        if ("위임" in message.Subject) and (message.Attachments.Count != 0):
            _save_attachments_for_message(message, attach_list)

    return attach_list


# ── 예고 메일(사건 담당자 지정) ─────────────────────────────────────────
# 본문 표 한 행: 접수일 / 출원번호 / Ref / 담당자 / 출원인 / 문서종류
# 예) 20261012	10-2025-7000694	PA-2T8084	신용석	에어 프로덕츠 앤드 케미칼스, 인코오포레이티드	의견제출통지서
# Outlook 의 MailItem.Body 는 HTML 표를 칸마다 한 줄(사이에 빈 줄)로 풀어 준다. 탭으로 이은 한 줄도 받는다.
_RECEIPT_CELL = re.compile(r"\d{8}")
_APP_NO_CELL = re.compile(r"\d{2}-\d{4}-\d{7}")
_PRENOTICE_CELLS = 6
OA_DOC_TYPE = "의견제출통지서"  # 「최후의견제출통지서」도 이 글자를 포함한다
FR_DOC_TYPE = "거절결정"        # 「거절결정서」·「특허거절결정서」 어느 쪽이든
CASE_CACHE_FILE = data_path("case_cache.json")
CASE_CACHE_KEEP_DAYS = 183
PRENOTICE_SCAN_DAYS = 14


def parse_prenotice_rows(body, assignee):
    """예고 메일 본문 표에서 담당자가 assignee 인 행만 읽습니다(중복 행 제거).

    반환값: [{"ref", "app_no", "receipt_date"(YYYY-MM-DD), "doc_type", "applicant"}, ...]
    """
    # 빈 칸을 버리고 칸을 한 줄로 늘어놓는다 — 「접수일, 출원번호」 가 이어지는 곳이 한 행의 시작
    cells = [c.strip() for line in (body or "").splitlines() for c in line.split("\t") if c.strip()]
    rows, seen = [], set()
    for i in range(len(cells) - _PRENOTICE_CELLS + 1):
        if not (_RECEIPT_CELL.fullmatch(cells[i]) and _APP_NO_CELL.fullmatch(cells[i + 1])):
            continue
        receipt, app_no, ref_token, name, applicant, doc_type = cells[i:i + _PRENOTICE_CELLS]
        if name != assignee:
            continue
        # 위임 메일 제목과 같은 규칙(ref_num_key)으로 Ref 를 맞춘다 — 캐시 키가 같아야 찾을 수 있다
        ref_m = ref_num_key.match(ref_token)
        if not ref_m:
            continue
        ref = ref_m.group(0)
        key = (ref, app_no, doc_type)
        if key in seen:
            continue
        seen.add(key)
        rows.append({
            "ref": ref,
            "app_no": app_no,
            "receipt_date": kipris.iso_date(receipt),
            "doc_type": doc_type,
            "applicant": applicant,
        })
    return rows


def _case_key(ref, doc_type):
    return f"{ref}|{doc_type}"


def load_case_cache():
    if os.path.exists(CASE_CACHE_FILE):
        try:
            with open(CASE_CACHE_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
        except Exception as e:
            logging.error(f"사건 캐시 로드 실패: {e}")
    return {}


def save_case_cache(cache, today=None):
    """사건 캐시를 저장합니다. 받은 지 CASE_CACHE_KEEP_DAYS 가 지난 항목은 지웁니다."""
    cutoff = ((today or date.today()) - timedelta(days=CASE_CACHE_KEEP_DAYS)).isoformat()
    kept = {k: v for k, v in cache.items() if (v.get("received_at") or "9999") >= cutoff}
    try:
        tmp = CASE_CACHE_FILE + ".tmp"
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(kept, f, indent=2, ensure_ascii=False)
        os.replace(tmp, CASE_CACHE_FILE)
    except Exception as e:
        logging.error(f"사건 캐시 저장 실패: {e}")


def find_case(cache, ref_num, doc_word):
    """Ref 의 항목 가운데 문서종류에 doc_word 가 들어간 것 중 접수일이 가장 늦은 것. 없으면 None."""
    cands = [
        v for k, v in cache.items()
        if k.split("|", 1)[0] == ref_num and doc_word in (v.get("doc_type") or "")
    ]
    return max(cands, key=lambda v: v.get("receipt_date") or "") if cands else None


def find_oa_case(cache, ref_num):
    """Ref 의 의견제출통지서 항목 가운데 접수일이 가장 늦은 것. 없으면 None."""
    return find_case(cache, ref_num, OA_DOC_TYPE)


def handle_prenotice(message):
    """예고 메일 한 통을 처리합니다. 의견제출통지서 행은 KIPRIS 마감기한을, 거절결정서 행은 재심사에 의한
    거절결정인지를 KIPRIS 에 1회 조회해 함께 저장합니다.

    이미 같은 내용(Ref·문서종류·출원번호·접수일)으로 저장된 행은 건너뜁니다(KIPRIS 다시 부르지 않음).
    단 의견제출통지서인데 마감기한을, 거절결정서인데 재심사 여부를 아직 못 받은 행은 다시 조회합니다.
    반환값: 새로 저장(또는 갱신)한 행 수
    """
    rows = parse_prenotice_rows(message.Body, config["prenotice_assignee"])
    if not rows:
        logging.info(f"예고 메일에서 담당자 행을 찾지 못했습니다: {message.Subject}")
        return 0

    cache = load_case_cache()
    received_at = message.CreationTime.strftime("%Y-%m-%d")
    added = 0
    for row in rows:
        key = _case_key(row["ref"], row["doc_type"])
        old = cache.get(key)
        is_oa = OA_DOC_TYPE in row["doc_type"]
        is_fr = FR_DOC_TYPE in row["doc_type"]
        same = old and old.get("app_no") == row["app_no"] and old.get("receipt_date") == row["receipt_date"]
        if same and (old.get("due_date") if is_oa else ("reexam" in old if is_fr else True)):
            if old.get("applicant") != row["applicant"]:  # 출원인 저장 전에 받은 행 — KIPRIS 는 다시 부르지 않는다
                old["applicant"] = row["applicant"]
                added += 1
            continue

        entry = {
            "app_no": row["app_no"],
            "doc_type": row["doc_type"],
            "receipt_date": row["receipt_date"],
            "applicant": row["applicant"],
            "due_date": "",
            "received_at": received_at,
        }
        if same:
            entry["received_at"] = old.get("received_at") or received_at
        if is_oa:
            try:
                notice = kipris.pick_oa_due(kipris.fetch_due_dates(row["app_no"]), row["doc_type"])
                if notice:
                    entry["due_date"] = notice.due_date
                    logging.info(
                        f"KIPRIS 마감기한: {row['ref']} {row['app_no']} {notice.title} "
                        f"발송번호 {notice.send_number} → {notice.due_date}"
                    )
                else:
                    logging.info(f"KIPRIS 에 {row['doc_type']} 마감기한이 아직 없습니다: {row['ref']} {row['app_no']}")
            except kipris.KiprisError as e:
                logging.error(f"KIPRIS 조회 실패 ({row['ref']} {row['app_no']}): {e}")
        elif is_fr:
            try:
                entry["reexam"] = kipris.fetch_reexam(row["app_no"])
                logging.info(f"KIPRIS 재심사 거절결정 여부: {row['ref']} {row['app_no']} → {entry['reexam']}")
            except kipris.KiprisError as e:
                logging.error(f"KIPRIS 조회 실패 ({row['ref']} {row['app_no']}): {e}")

        cache[key] = entry
        added += 1
        logging.info(f"예고 메일 저장: {key} → {entry}")

    if added:
        save_case_cache(cache)
    return added


def is_prenotice(message):
    return PRENOTICE_SUBJECT in message.Subject and sender_matches(message, PRENOTICE_SENDERS)


def scan_prenotices(outlook, days=PRENOTICE_SCAN_DAYS):
    """앱이 꺼져 있던 동안 받은 편지함에 도착한 예고 메일을 처리합니다 (최근 days 일).

    확인을 마치고 다른 폴더로 옮긴 예고 메일은 보지 않습니다.
    """
    inbox = outlook.GetDefaultFolder(6)
    query = (
        build_sender_filter(PRENOTICE_SENDERS)
        + f" AND {_DASL_SUBJECT} LIKE {_dasl_str('%' + PRENOTICE_SUBJECT + '%')}"
    )
    since = date.today() - timedelta(days=days)
    messages = [m for m in inbox.Items.Restrict(query) if m.CreationTime.date() >= since]
    for message in sorted(messages, key=lambda m: m.CreationTime):  # 오래된 것부터
        try:
            handle_prenotice(message)
        except Exception as e:
            logging.error(f"예고 메일 처리 실패: {e}")


ABBR_LEN = 6
# 이름 앞뒤에 단어로 붙는 법인 표기 — 약칭에서 뗀다
_CORP_PREFIXES = ("주식회사", "(주)", "유한회사", "가부시키가이샤", "카부시키가이샤")
_CORP_SUFFIXES = _CORP_PREFIXES + (
    "가이샤", "가부시키", "인코포레이티드", "인코오포레이티드", "인크", "인크.", "코포레이션", "코오포레이션",
    "컴퍼니", "앤드", "리미티드", "엘엘씨", "엘티디", "게엠베하", "악티엔게젤샤프트", "아게", "에스알엘",
    "에스에이", "에스.에이.", "엔.브이.", "비.브이.",
)
# 띄어 쓰지 않고 이름 끝에 붙는 것 — 「도쿄엘렉트론가부시키가이샤」
_CORP_GLUED = ("가부시키가이샤", "카부시키가이샤", "주식회사", "유한회사")


def squash(text):
    """앞뒤·연속 공백 정리."""
    return " ".join((text or "").split())


def applicant_abbr(name):
    """일정 제목에 붙일 출원인 약칭. 약칭표에 있으면 그것, 없으면 법인 표기·공백을 떼고 앞 ABBR_LEN 자."""
    name = squash(name)
    table = {squash(k): v for k, v in (config.get("applicant_abbr") or {}).items()}
    if table.get(name):
        return table[name]
    words = name.replace(",", " ").split()
    while words and words[0] in _CORP_PREFIXES:
        words.pop(0)
    while words:
        last = words[-1]
        if last in _CORP_SUFFIXES:
            words.pop()
            continue
        glued = next((s for s in _CORP_GLUED if last.endswith(s) and len(last) > len(s)), "")
        if not glued:
            break
        words[-1] = last[:-len(glued)]
    short = "".join(words).strip(".") or name.replace(" ", "")
    return short[:ABBR_LEN]


def with_applicant(subject, applicant):
    """출원인이 있으면 제목 끝에 「_약칭」."""
    return f"{subject}_{applicant_abbr(applicant)}" if squash(applicant) else subject


def abbr_table_from_rows(rows):
    """Settings 약칭표의 (명칭, 약칭) 행들 → dict. 빈 칸이 있는 행은 버리고, 같은 명칭은 뒤의 것."""
    table = {}
    for name, abbr in rows:
        name, abbr = squash(name), squash(abbr)
        if name and abbr:
            table[name] = abbr
    return table


def fr_action(ref_num, case):
    """거절결정서 Action 일정 제목을 정합니다. 예고 메일 때 저장한 재심사 여부만 씁니다.

    반환값: (subject, doc_label, source_text, warning)
    """
    reexam = (case or {}).get("reexam")
    if reexam is True:
        return f"{ref_num}_거결불", "재심사 거절결정서 (FR)", "재심사에 의한 거절결정 (KIPRIS)", ""
    if reexam is False:
        return f"{ref_num}_재심사", "거절결정서 (FR)", "일반 거절결정 (KIPRIS)", ""
    return (f"{ref_num}_재심사", "거절결정서 (FR)", "⚠ 재심사 여부 확인 불가",
            "재심사 여부를 확인하지 못했습니다 — 재심사에 의한 거절결정이면 네이버 캘린더에서 제목을 _거결불 로 고치세요.")


def resolve_oa_dates(mail_date, fallback_date, case, today):
    """OA 청서류 기한을 정합니다. KIPRIS 는 부르지 않고 예고 메일 때 저장한 값만 씁니다.

    반환값: (action_date, source, warning) — source 는 "KIPRIS" 또는 "계산식"
    """
    if case is None:
        return fallback_date, "계산식", "예고 메일 정보가 없어 계산식을 적용했습니다."
    due = case.get("due_date") or ""
    if not due:
        return fallback_date, "계산식", (
            f"KIPRIS 마감기한이 없어 계산식을 적용했습니다 (출원번호 {case.get('app_no', '')})."
        )
    try:
        due_date = date.fromisoformat(due)
    except ValueError:
        return fallback_date, "계산식", f"KIPRIS 마감기한 형식이 잘못되어 계산식을 적용했습니다 ({due})."
    if due_date < today or (mail_date is not None and due_date <= mail_date):
        return fallback_date, "계산식", (
            f"KIPRIS 마감기한 {due}이(가) 오늘 이전이거나 서신 기한보다 빨라 무시했습니다."
        )
    return due_date, "KIPRIS", ""


CASE_RECENT_DAYS = 60   # 예고 메일 접수일이 위임 메일보다 이만큼 넘게 앞서면 지난 통지서의 기록으로 본다
CASE_AHEAD_DAYS = 3


def recent_case(case, received):
    """예고 메일 기록이 이번 위임 메일의 것이면 그대로, 지난 통지서의 것이면 None."""
    try:
        receipt = date.fromisoformat((case or {}).get("receipt_date") or "")
    except ValueError:
        return None
    gap = (received - receipt).days
    return case if -CASE_AHEAD_DAYS <= gap <= CASE_RECENT_DAYS else None


def case_memo(case):
    """일정 설명칸 — 출원인·출원번호 가운데 있는 것만."""
    case = case or {}
    lines = []
    if squash(case.get("applicant")):
        lines.append(f"출원인: {squash(case.get('applicant'))}")
    if case.get("app_no"):
        lines.append(f"출원번호: {case['app_no']}")
    return "\n".join(lines)


def build_schedule(subject, received, rejection_type, cache, today):
    """위임 메일 한 통의 Comments·Action 일정을 정합니다. 등록은 하지 않습니다.

    subject: 메일 제목, received: 메일 수신일(date), cache: 예고 메일 사건 캐시.
    반환값: dict(ref, items, folder_names, memo, doc_label, source_text, warning, editable). Ref 가 없으면 None.
    """
    ref_nums = ref_num_key.findall(subject)
    if not ref_nums:
        return None
    ref_num = ref_nums[0]

    # 예고 메일 때 저장한 사건 정보 (KIPRIS 마감기한·재심사 여부·접수일·출원인·출원번호)
    case = find_case(cache, ref_num, OA_DOC_TYPE if rejection_type == "OA" else FR_DOC_TYPE)

    # 제목의 통지일 「[서 2024-12-23]」 — Comments 일정 날짜. 없으면 수신일
    match = re.search(r"(\d{4}-\d{2}-\d{2})", subject)
    notice_date = datetime.strptime(match.group(1), "%Y-%m-%d").date() if match else None
    target_date = notice_date or received
    warning = ""

    if rejection_type == "OA":
        subj_action = f"{ref_num}_의견서"
        doc_label = "의견제출통지서 (OA)"
        # 예고 메일 때 저장한 KIPRIS 마감기한을 쓴다. 없으면 통지일로 계산한 값(계산식).
        fallback = notice_date + relativedelta(months=4) - timedelta(days=11) if notice_date else received
        response_date, action_source, warning = resolve_oa_dates(notice_date, fallback, case, today)
        if action_source == "KIPRIS":
            source_text = "청서류 기한: KIPRIS 마감기한 사용"
        elif notice_date:
            source_text = "⚠ KIPRIS 마감기한 사용 불가 → 계산식 적용 (날짜를 확인하세요)"
        else:
            source_text = "⚠ 제목에 통지일이 없어 기한을 계산하지 못했습니다 — 날짜를 입력하세요"
        editable = dates_editable(rejection_type, action_source)
    else:
        fresh = recent_case(case, received)  # 지난 거절결정의 기록(재심사 여부·접수일)은 쓰지 않는다
        subj_action, doc_label, fr_source, warning = fr_action(ref_num, fresh)
        try:
            base, base_text = date.fromisoformat(fresh["receipt_date"]), "접수일"
        except (TypeError, KeyError, ValueError):
            base, base_text = received, "수신일"
        response_date = base + relativedelta(months=3)
        source_text = f"{fr_source} · 청서류 기한: {base_text} {base.isoformat()} + 3개월"
        editable = False

    applicant = (case or {}).get("applicant", "")
    return {
        "ref": ref_num,
        "items": [
            ("Comments 일정 (네이버: Mail)", with_applicant(f"{ref_num}_Comments", applicant), target_date),
            ("Action 일정 (네이버: No_Instructions)", with_applicant(subj_action, applicant), response_date),
        ],
        "folder_names": ["Mail", "No_Instructions"],
        "memo": case_memo(case),
        "doc_label": doc_label,
        "source_text": source_text,
        "warning": warning,
        "editable": editable,
    }


def dates_editable(rejection_type, action_source):
    """날짜를 고칠 수 있는가 — OA 에서 KIPRIS 마감기한을 쓸 수 없을 때만."""
    return rejection_type == "OA" and action_source != "KIPRIS"


STATUS_ADDED = "추가됨"
STATUS_DUPLICATE = "건너뜀(중복)"


def _register_naver(session, folder_name, subject, day, memo=""):
    """종일 일정 하나를 네이버 캘린더에 등록합니다. 실패해도 예외 대신 「오류: …」 를 돌려줍니다."""
    calendar_name = (config.get("naver_calendars") or {}).get(folder_name, "")
    try:
        if not calendar_name:
            raise naver_cal.NaverError(f"settings.json 의 naver_calendars 에 「{folder_name}」 캘린더 이름이 없습니다.")
        added = session.add_all_day(calendar_name, subject, day, memo)  # 메일 본문은 넣지 않는다
    except Exception as e:
        logging.error(f"네이버 캘린더 일정 등록 실패: {subject} ({day}) — {e}")
        return f"오류: {e}"
    status = STATUS_ADDED if added else STATUS_DUPLICATE
    logging.info(f"네이버 캘린더 일정 {status}: {subject} ({day}) → {calendar_name}")
    return status


def register_items(folder_names, items, selected, naver, memo=""):
    """고른 일정을 네이버 캘린더에 등록합니다. memo 는 일정 설명칸. 반환값: {index: 결과}"""
    return {i: _register_naver(naver, folder_names[i], items[i][1], day, memo) for i, day in selected}


def failed_indices(results):
    """「오류」로 끝난 일정의 index — [다시 시도] 대상."""
    return sorted(i for i, status in results.items() if status.startswith("오류"))


# ── 등록 기록 (처리한 위임 메일·등록 실패 일정) ──────────────────────────
CALENDAR_STATE_FILE = data_path("calendar_state.json")
CALENDAR_DONE_KEEP_DAYS = 30
DELEGATION_SCAN_DAYS = 14


def load_calendar_state():
    """{"done": {EntryID: 처리일}, "pending": [실패 일정]}. 파일이 없으면 None (처음 실행)."""
    if not os.path.exists(CALENDAR_STATE_FILE):
        return None
    try:
        with open(CALENDAR_STATE_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        logging.error(f"등록 기록 로드 실패: {e}")
        data = {}
    data = data if isinstance(data, dict) else {}
    return {"done": dict(data.get("done") or {}), "pending": list(data.get("pending") or [])}


def save_calendar_state(state, today=None):
    """등록 기록을 저장합니다. 처리한 지 CALENDAR_DONE_KEEP_DAYS 가 지난 메일 기록은 지웁니다."""
    cutoff = ((today or date.today()) - timedelta(days=CALENDAR_DONE_KEEP_DAYS)).isoformat()
    data = {"done": {k: v for k, v in state["done"].items() if v >= cutoff}, "pending": state["pending"]}
    try:
        tmp = CALENDAR_STATE_FILE + ".tmp"
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        os.replace(tmp, CALENDAR_STATE_FILE)
    except Exception as e:
        logging.error(f"등록 기록 저장 실패: {e}")


def _pending_key(p):
    return (p["folder"], p["subject"], p["day"])


def track_results(ref, folder_names, items, results, days, memo=""):
    """등록 결과를 기록에 반영 — 「오류」 일정은 pending 에 넣고, 등록됐거나 이미 있던 일정은 뺀다."""
    state = load_calendar_state() or {"done": {}, "pending": []}
    pending = {_pending_key(p): p for p in state["pending"]}
    for i, status in results.items():
        p = {"ref": ref, "folder": folder_names[i], "subject": items[i][1], "day": days[i].isoformat(), "memo": memo}
        if status.startswith("오류"):
            pending[_pending_key(p)] = p
        else:
            pending.pop(_pending_key(p), None)
    state["pending"] = list(pending.values())
    save_calendar_state(state)


def register_and_track(ref, folder_names, items, selected, naver, memo=""):
    """register_items + 결과를 등록 기록에 반영. 반환값: {index: 결과}"""
    results = register_items(folder_names, items, selected, naver, memo)
    track_results(ref, folder_names, items, results, dict(selected), memo)
    return results


def mark_done(entry_id, today=None):
    """위임 메일 하나를 처리했다고 기록 — 다음 시작 때 다시 등록하지 않는다."""
    state = load_calendar_state() or {"done": {}, "pending": []}
    state["done"][entry_id] = (today or date.today()).isoformat()
    save_calendar_state(state, today)


def retry_pending(naver):
    """등록에 실패했던 일정을 다시 등록합니다. 반환값: [(Ref, 제목, 날짜, 결과), ...]"""
    state = load_calendar_state()
    if not state or not state["pending"]:
        return []
    rows, still = [], []
    for p in state["pending"]:
        day = date.fromisoformat(p["day"])
        status = _register_naver(naver, p["folder"], p["subject"], day, p.get("memo", ""))
        rows.append((p["ref"], p["subject"], day, status))
        if status.startswith("오류"):
            still.append(p)
    state = load_calendar_state() or {"done": {}, "pending": []}
    state["pending"] = still
    save_calendar_state(state)
    return rows


class ScheduleDialog(qtw.QDialog):
    """캘린더 등록 알림창 — 날짜·출처(KIPRIS 사용 여부)·등록 결과를 한 창에 보여 줍니다.

    editable=False: 이미 등록한 결과(results)를 보여 주고 [닫기]만 둡니다.
    editable=True : 날짜를 고치고 일정별로 고를 수 있습니다. [등록]을 누르면 register 를 불러
                    등록한 뒤, 같은 창에 결과를 채우고 날짜를 잠급니다.
    등록에 실패한 일정이 있으면 [다시 시도]로 그 일정만 다시 등록합니다.
    """

    def __init__(self, parent=None, ref_num="", doc_label="", items=(), source_text="", warning="",
                 editable=False, register=None, results=None):
        super().__init__(parent)
        self.setWindowTitle("캘린더 등록")
        self.resize(460, 260)

        self.ref_num = ref_num
        self.items = list(items)        # [(그룹 제목, 일정 제목, 날짜), ...]
        self.editable = editable
        self.register = register        # [(index, 날짜), ...] -> {index: 결과}
        self.registered = bool(results)
        self.results = {}
        self.checks, self.date_edits, self.result_labels = [], [], []

        layout = qtw.QVBoxLayout(self)

        header = qtw.QLabel(f"<b>[Ref: {ref_num}]</b> {doc_label}")
        layout.addWidget(header)

        source = qtw.QLabel(source_text + (f"<br><small>{warning}</small>" if warning else ""))
        source.setWordWrap(True)
        if warning or editable:
            source.setStyleSheet("background-color: #FDF3E1; color: #8A5300; padding: 6px; border-radius: 4px;")
        else:
            source.setStyleSheet("background-color: #E7F4EC; color: #1E6B3A; padding: 6px; border-radius: 4px;")
        layout.addWidget(source)

        for i, (group_title, subject, day) in enumerate(self.items):
            group = qtw.QGroupBox(group_title)
            row = qtw.QHBoxLayout()

            if editable:
                chk = qtw.QCheckBox("등록")
                chk.setChecked(True)
                row.addWidget(chk)
                self.checks.append(chk)

            lbl = qtw.QLabel(subject)
            lbl.setMinimumWidth(140)
            row.addWidget(lbl)

            if editable:
                date_edit = qtw.QDateEdit()
                date_edit.setCalendarPopup(True)
                date_edit.setDisplayFormat("yyyy-MM-dd")
                date_edit.setDate(qtc.QDate(day.year, day.month, day.day))
                self.checks[-1].toggled.connect(date_edit.setEnabled)
                row.addWidget(date_edit)
                self.date_edits.append(date_edit)
            else:
                row.addWidget(qtw.QLabel(day.isoformat()))

            result = qtw.QLabel()
            result.setMinimumWidth(80)
            row.addWidget(result)
            self.result_labels.append(result)

            group.setLayout(row)
            layout.addWidget(group)

        btn_box = qtw.QHBoxLayout()
        btn_box.addStretch()
        self.btn_retry = qtw.QPushButton("다시 시도")
        self.btn_retry.clicked.connect(self._on_retry)
        self.btn_register = qtw.QPushButton("등록")
        self.btn_register.clicked.connect(self._on_register)
        self.btn_cancel = qtw.QPushButton("취소")
        self.btn_cancel.setObjectName("quitButton")
        self.btn_cancel.clicked.connect(self.reject)
        self.btn_close = qtw.QPushButton("닫기")
        self.btn_close.clicked.connect(self.accept)
        for btn in (self.btn_retry, self.btn_register, self.btn_cancel, self.btn_close):
            btn_box.addWidget(btn)
        layout.addLayout(btn_box)

        self.btn_retry.setVisible(False)
        self.btn_register.setVisible(editable)
        self.btn_cancel.setVisible(editable)
        self.btn_close.setVisible(not editable)

        if results:
            self._show_results(results)

    @staticmethod
    def _status_color(status):
        return "#1E6B3A" if status == STATUS_ADDED else "#C0392B" if status.startswith("오류") else "#6B6B6B"

    def _show_results(self, results):
        """results 에 든 일정만 결과를 바꾼다 — [다시 시도]는 실패한 일정만 넘긴다."""
        self.results.update(results)
        for i, label in enumerate(self.result_labels):
            status = self.results.get(i, "등록 안 함")
            label.setText(status)
            label.setToolTip(status)
            label.setStyleSheet(f"color: {self._status_color(status)};")
        self.btn_retry.setVisible(bool(failed_indices(self.results)))

    def _day(self, i):
        return self.date_edits[i].date().toPython() if self.editable else self.items[i][2]

    def _call_register(self, selected):
        try:
            return self.register(selected) if selected else {}
        except Exception as e:
            logging.error(f"캘린더 일정 등록 실패: {e}", exc_info=True)
            return {i: f"오류: {e}" for i, _day in selected}

    def _on_retry(self):
        self._show_results(self._call_register([(i, self._day(i)) for i in failed_indices(self.results)]))

    def done(self, result):
        failed = failed_indices(self.results)
        if failed:
            subjects = ", ".join(self.items[i][1] for i in failed)
            logging.warning(f"네이버 등록 실패 상태로 닫음: Ref {self.ref_num} — {subjects}")
        super().done(result)

    def _on_register(self):
        selected = [
            (i, self.date_edits[i].date().toPython())
            for i, chk in enumerate(self.checks) if chk.isChecked()
        ]
        results = self._call_register(selected)
        self.registered = True
        self._show_results(results)

        # 등록을 마쳤으면 날짜를 잠그고 [닫기]만 남긴다
        for widget in self.checks + self.date_edits:
            widget.setEnabled(False)
        self.btn_register.setVisible(False)
        self.btn_cancel.setVisible(False)
        self.btn_close.setVisible(True)


def parse_names(text):
    """「이여름, 한송희」 → ["이여름", "한송희"] (공백 정리, 빈 이름·중복 제거)."""
    return list(dict.fromkeys(n for n in (squash(t) for t in (text or "").split(",")) if n))


def calendar_assignees():
    """이 담당자 앞으로 온 위임 메일만 일정을 등록한다 (Settings 「위임 메일 담당자」)."""
    return config.get("calendar_assignees") or DEFAULT_CONFIG["calendar_assignees"]


def recent_applicants(cache, existing=()):
    """사건 캐시에 저장된 출원인 가운데 existing 에 없는 이름 (가나다순)."""
    have = {squash(n) for n in existing}
    names = {squash(v.get("applicant")) for v in cache.values()} - have - {""}
    return sorted(names)


def is_calendar_delegation(message):
    """일정을 등록할 위임 메일인가 — 제목 「위임」, 담당자 이름, 보낸 사람."""
    if "위임" not in message.Subject:
        return False
    names = name_key.findall(message.Subject)
    if not names or names[0][2:] not in calendar_assignees():
        return False
    return sender_matches(message, CALENDAR_SENDERS)


def rejection_type_of(message):
    return "FR" if "거절결정서" in message.Subject else "OA"


def register_delegation(message, always_show=True):
    """위임 메일 한 통의 일정을 등록하고 기록합니다.

    날짜를 고칠 수 있는 경우(OA + KIPRIS 사용 불가)는 늘 알림창에서 확인받습니다.
    always_show=False 면(시작할 때 스캔) 바로 등록한 결과는 창 없이 돌려줍니다.
    반환값: [(Ref, 제목, 날짜, 결과), ...]
    """
    rejection_type = rejection_type_of(message)
    plan = build_schedule(message.Subject, message.CreationTime.date(), rejection_type,
                          load_case_cache(), date.today())
    if plan is None:
        logging.warning(f"Ref 번호를 찾을 수 없습니다: {message.Subject}")
        mark_done(message.EntryID)
        return []
    ref, items, folder_names = plan["ref"], plan["items"], plan["folder_names"]
    logging.info(f"일정 계산: {ref} {rejection_type} → {[(s, d.isoformat()) for _g, s, d in items]} "
                 f"({plan['source_text']}) {plan['warning']}")

    def register(selected):
        return register_and_track(ref, folder_names, items, selected, naver_cal.Session(), plan["memo"])

    results = None if plan["editable"] else register([(0, items[0][2]), (1, items[1][2])])
    rows = []
    if plan["editable"] or always_show:
        dialog = ScheduleDialog(
            parent=qtw.QApplication.activeWindow(),
            ref_num=ref,
            doc_label=plan["doc_label"],
            items=items,
            source_text=plan["source_text"],
            warning=plan["warning"],
            editable=plan["editable"],
            register=register,
            results=results,
        )
        dialog.exec()
        if plan["editable"] and not dialog.registered:
            logging.info(f"캘린더 일정 등록이 사용자에 의해 취소되었습니다. Ref: {ref}")
    else:
        rows = [(ref, items[i][1], items[i][2], status) for i, status in sorted(results.items())]
    mark_done(message.EntryID)
    return rows


def scan_delegations(outlook, days=DELEGATION_SCAN_DAYS):
    """앱이 꺼져 있던 동안 받은 위임 메일의 일정을 등록합니다 (최근 days 일). 반환값: 요약 행 목록.

    처음 실행할 때(등록 기록 파일 없음)는 이미 등록했을 수 있으니 지금 있는 메일을 처리한 것으로만 기록합니다.
    """
    inbox = outlook.GetDefaultFolder(6)
    query = (
        build_sender_filter(CALENDAR_SENDERS)
        + f" AND {_DASL_SUBJECT} LIKE {_dasl_str('%위임%')}"
    )
    since = date.today() - timedelta(days=days)
    messages = sorted(
        (m for m in inbox.Items.Restrict(query) if m.CreationTime.date() >= since and is_calendar_delegation(m)),
        key=lambda m: m.CreationTime,
    )
    state = load_calendar_state()
    if state is None:
        state = {"done": {m.EntryID: date.today().isoformat() for m in messages}, "pending": []}
        save_calendar_state(state)
        logging.info(f"등록 기록을 새로 만들었습니다 — 위임 메일 {len(messages)}통은 처리한 것으로 봅니다.")
        return []
    rows = []
    for message in messages:
        if message.EntryID in state["done"]:
            continue
        try:
            rows += register_delegation(message, always_show=False)
        except Exception as e:
            logging.error(f"위임 메일 일정 등록 실패: {message.Subject} — {e}", exc_info=True)
    return rows


class StartupSummaryDialog(qtw.QDialog):
    """앱을 시작할 때 다시 등록한 일정(실패 일정·놓친 위임 메일)의 결과."""

    def __init__(self, rows, parent=None):
        super().__init__(parent)
        self.setWindowTitle("캘린더 등록 — 시작할 때 처리한 일정")
        self.resize(560, 300)
        layout = qtw.QVBoxLayout(self)
        layout.addWidget(qtw.QLabel("앱이 꺼져 있던 동안 받은 위임 메일과 지난번 실패한 일정을 등록했습니다."))
        table = qtw.QTableWidget(len(rows), 4)
        table.setHorizontalHeaderLabels(["Ref", "일정 제목", "날짜", "결과"])
        table.horizontalHeader().setSectionResizeMode(1, qtw.QHeaderView.Stretch)
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(qtw.QAbstractItemView.NoEditTriggers)
        for r, (ref, subject, day, status) in enumerate(rows):
            for c, text in enumerate((ref, subject, day.isoformat(), status)):
                table.setItem(r, c, qtw.QTableWidgetItem(text))
            table.item(r, 3).setForeground(qtg.QColor(ScheduleDialog._status_color(status)))
        layout.addWidget(table)
        if failed_indices(dict(enumerate(s for *_x, s in rows))):
            note = qtw.QLabel("실패한 일정은 다음에 앱을 시작할 때 다시 등록합니다.")
            note.setStyleSheet("color: #8A5300;")
            layout.addWidget(note)
        close_btn = qtw.QPushButton("닫기")
        close_btn.clicked.connect(self.accept)
        btn_box = qtw.QHBoxLayout()
        btn_box.addStretch()
        btn_box.addWidget(close_btn)
        layout.addLayout(btn_box)


def catch_up_calendar(outlook):
    """시작할 때: 실패했던 일정을 다시 등록하고, 놓친 위임 메일의 일정을 등록한 뒤 결과를 요약해 보여 줍니다."""
    rows = []
    try:
        rows += retry_pending(naver_cal.Session())
    except Exception as e:
        logging.error(f"실패 일정 다시 등록 실패: {e}", exc_info=True)
    try:
        rows += scan_delegations(outlook)
    except Exception as e:
        logging.error(f"위임 메일 스캔 실패: {e}", exc_info=True)
    if any(status != STATUS_DUPLICATE for *_x, status in rows):
        StartupSummaryDialog(rows, qtw.QApplication.activeWindow()).exec()
    return rows


# 개선 #6: Python 3에서 (object) 상속 명시는 불필요
class Handler_Class:
    # 개선 #8: outlook 인스턴스를 클래스 변수로 관리 (DispatchWithEvents 특성상 생성자 인자 전달 불가)
    outlook = None
    main_win = None  # UI 업데이트를 위한 MainWin 인스턴스 참조

    def OnNewMailEx(self, receivedItemsIDs):
        # receivedItemsIDs는 ","로 구분된 메일 ID 모음입니다.
        for ID in receivedItemsIDs.split(","):
            message = Handler_Class.outlook.Session.GetItemFromID(ID)

            # 보낸 사람 판정(sender_matches)은 비용이 있으므로 제목/첨부 조건을 먼저 확인
            if (
                ("위임" in message.Subject)
                and (message.Attachments.Count != 0)
                and sender_matches(message, DELEGATION_SENDERS)
            ):
                # 수정 #1: attach_list를 넘겨 저장 결과를 받고 QTableWidget에 항목 추가
                new_items = []
                _save_attachments_for_message(message, new_items)
                if Handler_Class.main_win is not None and new_items:
                    for item in new_items:
                        Handler_Class.main_win.add_table_row(*item)

            # 예고 메일(사건 담당자 지정): 출원번호·KIPRIS 마감기한을 Ref 별로 저장해 둔다
            if is_prenotice(message):
                try:
                    handle_prenotice(message)
                except Exception as e:
                    logging.error(f"예고 메일 처리 실패: {e}", exc_info=True)
                continue

            if is_calendar_delegation(message):
                try:
                    register_delegation(message)
                except Exception as e:
                    logging.error(f"위임 메일 일정 등록 실패: {message.Subject} — {e}", exc_info=True)


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

        def section(title):
            """제목이 있는 테두리 칸 — QGroupBox 는 굵은 글꼴이 안쪽 위젯까지 번져 쓰지 않는다."""
            frame = qtw.QFrame()
            frame.setObjectName("settingsSection")
            box = qtw.QVBoxLayout(frame)
            box.setContentsMargins(14, 12, 14, 12)
            box.setSpacing(8)
            label = qtw.QLabel(title)
            label.setObjectName("sectionTitle")
            box.addWidget(label)
            return frame, box

        def hint(text):
            label = qtw.QLabel(text)
            label.setObjectName("settingsHint")
            label.setWordWrap(True)
            return label

        # 3-1. 파일 경로
        path_frame, path_box = section("파일 경로")
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

        path_box.addLayout(settings_layout)
        settings_main_layout.addWidget(path_frame)

        # 3-2. 일정 등록 — 위임 메일 담당자
        cal_frame, cal_box = section("일정 등록")
        cal_form = qtw.QFormLayout()
        self.assignees_edit = qtw.QLineEdit(", ".join(calendar_assignees()))
        self.assignees_edit.setPlaceholderText("이여름, 한송희")
        cal_form.addRow("위임 메일 담당자:", self.assignees_edit)
        cal_box.addLayout(cal_form)
        cal_box.addWidget(hint("쉼표로 구분 · 이 담당자 앞으로 온 위임 메일만 네이버 캘린더에 등록합니다."))
        settings_main_layout.addWidget(cal_frame)

        # 3-3. 출원인 약칭표 — 일정 제목 끝에 붙일 약칭 (표에 없으면 자동으로 줄인다)
        abbr_frame, abbr_box = section("출원인 약칭")
        toolbar = qtw.QHBoxLayout()
        self.abbr_search = qtw.QLineEdit()
        self.abbr_search.setPlaceholderText("검색 (명칭·약칭)")
        self.abbr_search.setClearButtonEnabled(True)
        self.abbr_search.textChanged.connect(self.filter_abbr_rows)
        toolbar.addWidget(self.abbr_search, 1)
        for text, slot in (("최근 출원인 가져오기", self.import_applicants_clicked),
                           ("추가", self.add_abbr_clicked), ("삭제", self.delete_abbr_clicked)):
            btn = qtw.QPushButton(text)
            btn.setObjectName("secondaryButton")
            btn.clicked.connect(slot)
            toolbar.addWidget(btn)
        abbr_box.addLayout(toolbar)

        self.abbr_table = qtw.QTableWidget(0, 2)
        self.abbr_table.setObjectName("abbrTable")
        self.abbr_table.setHorizontalHeaderLabels(["출원인 국문 명칭", "약칭"])
        header = self.abbr_table.horizontalHeader()
        header.setSectionResizeMode(0, qtw.QHeaderView.Stretch)
        header.setSectionResizeMode(1, qtw.QHeaderView.Interactive)
        header.resizeSection(1, 200)
        header.setHighlightSections(False)
        header.setDefaultAlignment(qtc.Qt.AlignLeft | qtc.Qt.AlignVCenter)
        self.abbr_table.verticalHeader().setVisible(False)
        self.abbr_table.verticalHeader().setDefaultSectionSize(32)
        self.abbr_table.setShowGrid(False)
        self.abbr_table.setAlternatingRowColors(True)
        self.abbr_table.setSelectionBehavior(qtw.QAbstractItemView.SelectRows)
        self.abbr_table.setEditTriggers(qtw.QAbstractItemView.DoubleClicked | qtw.QAbstractItemView.EditKeyPressed
                                        | qtw.QAbstractItemView.AnyKeyPressed)
        for name, abbr in sorted((config.get("applicant_abbr") or {}).items()):
            self._add_abbr_row(name, abbr)
        abbr_box.addWidget(self.abbr_table, 1)
        abbr_box.addWidget(hint(
            f"표에 없거나 약칭을 비운 출원인은 자동 약칭(법인 표기·공백 제거 후 앞 {ABBR_LEN}자)을 씁니다. "
            "약칭을 비운 행은 저장하지 않습니다. 셀을 더블클릭해 고칩니다."))
        settings_main_layout.addWidget(abbr_frame, 1)

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
            with open(LOG_FILE, 'r', encoding='utf-8') as f:
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

    def _add_abbr_row(self, name="", abbr=""):
        row = self.abbr_table.rowCount()
        self.abbr_table.insertRow(row)
        self.abbr_table.setItem(row, 0, qtw.QTableWidgetItem(name))
        abbr_item = qtw.QTableWidgetItem(abbr)
        font = abbr_item.font()
        font.setBold(True)
        abbr_item.setFont(font)
        self.abbr_table.setItem(row, 1, abbr_item)
        return row

    def add_abbr_clicked(self):
        self.abbr_search.clear()  # 새 행이 검색에 가려지지 않게
        row = self._add_abbr_row()
        self.abbr_table.setCurrentCell(row, 0)
        self.abbr_table.editItem(self.abbr_table.item(row, 0))

    def delete_abbr_clicked(self):
        for row in sorted({i.row() for i in self.abbr_table.selectedIndexes()}, reverse=True):
            self.abbr_table.removeRow(row)

    def _abbr_cell(self, row, col):
        item = self.abbr_table.item(row, col)
        return item.text() if item else ""

    def filter_abbr_rows(self, text):
        needle = squash(text).lower()
        for row in range(self.abbr_table.rowCount()):
            hay = f"{self._abbr_cell(row, 0)} {self._abbr_cell(row, 1)}".lower()
            self.abbr_table.setRowHidden(row, bool(needle) and needle not in hay)

    def import_applicants_clicked(self):
        """예고 메일 때 저장한 출원인 가운데 표에 없는 이름을 약칭이 빈 행으로 더한다."""
        existing = [self._abbr_cell(r, 0) for r in range(self.abbr_table.rowCount())]
        names = recent_applicants(load_case_cache(), existing)
        if not names:
            qtw.QMessageBox.information(self, "출원인 약칭", "새로 가져올 출원인이 없습니다.")
            return
        self.abbr_search.clear()
        for name in names:
            row = self._add_abbr_row(name, "")
            self.abbr_table.item(row, 0).setToolTip("최근 예고 메일에서 가져옴")
        first = self.abbr_table.rowCount() - len(names)
        self.abbr_table.scrollToItem(self.abbr_table.item(first, 0))
        self.abbr_table.setCurrentCell(first, 1)

    def save_settings_ui(self):
        config["parent_path"] = self.parent_path_edit.text()
        config["target_path"] = self.target_path_edit.text()
        config["calendar_assignees"] = parse_names(self.assignees_edit.text()) or DEFAULT_CONFIG["calendar_assignees"]
        self.assignees_edit.setText(", ".join(config["calendar_assignees"]))
        config["applicant_abbr"] = abbr_table_from_rows(
            (self._abbr_cell(r, 0), self._abbr_cell(r, 1)) for r in range(self.abbr_table.rowCount())
        )
        # 저장한 그대로 다시 채운다 — 약칭을 비운 행은 빠지고 이름순으로 정리된다
        self.abbr_table.setRowCount(0)
        for name, abbr in sorted(config["applicant_abbr"].items()):
            self._add_abbr_row(name, abbr)
        self.filter_abbr_rows(self.abbr_search.text())
        save_config()
        logging.info(f"설정이 변경되었습니다. Parent: {config['parent_path']}, Target: {config['target_path']}, "
                     f"담당자: {config['calendar_assignees']}, 약칭 {len(config['applicant_abbr'])}개")
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
    try:
        scan_prenotices(outlook_mapi)
    except Exception as e:
        logging.error(f"예고 메일 스캔 실패: {e}", exc_info=True)
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
        #settingsSection {
            background-color: white;
            border: 1px solid #DCE3DF;
            border-radius: 10px;
        }
        #settingsSection QLabel {
            background-color: transparent;
        }
        #sectionTitle {
            font-weight: bold;
            font-size: 11pt;
        }
        #settingsHint {
            color: #8A877E;
            font-size: 9pt;
        }
        QPushButton#secondaryButton {
            background-color: white;
            color: #26695E;
            border: 1px solid #9FC9BF;
            border-radius: 8px;
            padding: 5px 12px;
            font-weight: normal;
        }
        QPushButton#secondaryButton:hover {
            background-color: #E1F5EE;
        }
        QTableWidget#abbrTable {
            background-color: white;
            alternate-background-color: #F6F6F4;
            border: 1px solid #DCE3DF;
            border-radius: 8px;
            selection-background-color: #E1F5EE;
            selection-color: #085041;
        }
        QTableWidget#abbrTable::item {
            padding: 0 8px;
        }
        QTableWidget#abbrTable QHeaderView::section {
            background-color: #F1EFE8;
            color: #5F5E5A;
            border: none;
            border-bottom: 1px solid #D3D1C7;
            padding: 6px 8px;
            font-weight: bold;
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
    # 창을 띄운 뒤: 실패했던 일정 다시 등록 + 꺼져 있던 동안 온 위임 메일 일정 등록 (예고 메일 스캔 다음)
    qtc.QTimer.singleShot(0, lambda: catch_up_calendar(outlook_mapi))
    sys.exit(app.exec())
