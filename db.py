"""NAS MariaDB `iplog` — 받은 위임(`assignments`)·받은 지시(`instructions`)를 쓰고, 차수(`projects`)를 읽고 고친다.
Qt·Outlook 을 모른다.

- 접속 정보: ChekPat 의 %APPDATA%\\ChekPat\\credentials.dat (DPAPI 봉인 JSON — `DB_HOST`·`DB_PORT`(비면 3307)·
  `DB_USERNAME`·`DB_PASSWORD`). kipris.py 가 KIPRIS 키를 읽는 것과 같은 엔트로피로 연다. DB 이름은 `iplog`.
- 앱 계정은 SELECT·INSERT·UPDATE·DELETE 만 있다 — 표를 만들지 않는다(ChekPat `scripts/iplog_case_columns.py`).
- 받은 위임은 `mail_entry_id`(Outlook EntryID, UNIQUE)로 `INSERT IGNORE` — 같은 메일을 두 번 써도 한 줄.
- NAS 에 닿지 않으면 못 쓴 줄을 상태 파일의 `pending` 에 쌓아 두고 다음에 다시 쓴다. 일정 등록은 이와 상관없이 돈다.
- `pymysql` 은 함수 안에서만 들인다 — 없어도 앱(과 시험)은 뜬다.
"""

import json
import logging
import os
import time
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Callable, Dict, List, Optional, Tuple

import kipris

DB_NAME = "iplog"
DEFAULT_PORT = 3307
CONNECT_TIMEOUT = 5                  # 초. GUI 스레드에서 부르므로 짧게
RETRY_AFTER = 300                    # 초. 한 번 못 닿으면 이만큼은 다시 붙어 보지 않고 바로 pending 에 쌓는다
RECORDED_KEEP_DAYS = 30

ASSIGNMENT_COLUMNS = ("mail_entry_id", "ref", "doc_type", "notice_date", "report_due", "assigner",
                      "app_no", "applicant", "due_date", "reexam", "received_at")


class DbError(Exception):
    """NAS DB 접속·쿼리 실패(접속 정보 없음·pymysql 없음·네트워크·SQL)."""


# ── 접속 ─────────────────────────────────────────────────────────────

def credentials() -> Optional[Dict[str, str]]:
    """ChekPat credentials.dat → {host, port, user, password}. 없거나 비었으면 None."""
    try:
        raw = kipris._unprotect((kipris._chekpat_dir() / "credentials.dat").read_bytes())
        values = json.loads(raw.decode("utf-8"))
    except Exception:
        return None
    host = str(values.get("DB_HOST") or "").strip()
    user = str(values.get("DB_USERNAME") or "").strip()
    password = str(values.get("DB_PASSWORD") or "")
    if not (host and user and password):
        return None
    try:
        port = int(str(values.get("DB_PORT") or "").strip() or DEFAULT_PORT)
    except ValueError:
        port = DEFAULT_PORT
    return {"host": host, "port": port, "user": user, "password": password}


_last_fail = 0.0


def connect(force: bool = False):
    """iplog 에 붙는다 (DictCursor, autocommit 끔). 실패하면 DbError.

    force=False 면 방금(RETRY_AFTER 안) 실패했을 때 붙어 보지 않고 바로 DbError — 위임 메일마다 기다리지 않는다.
    """
    global _last_fail
    if not force and _last_fail and time.monotonic() - _last_fail < RETRY_AFTER:
        raise DbError("NAS DB 에 조금 전 닿지 않아 잠시 쉬는 중입니다.")
    cred = credentials()
    if cred is None:
        raise DbError("NAS DB 접속 정보가 없습니다 (ChekPat 연결 설정 — credentials.dat).")
    try:
        import pymysql
        import pymysql.cursors
    except ImportError:
        raise DbError("pymysql 이 설치되어 있지 않습니다.")
    try:
        conn = pymysql.connect(host=cred["host"], port=cred["port"], user=cred["user"], password=cred["password"],
                               database=DB_NAME, charset="utf8mb4", cursorclass=pymysql.cursors.DictCursor,
                               connect_timeout=CONNECT_TIMEOUT, read_timeout=30, write_timeout=30, autocommit=False)
    except Exception as e:
        _last_fail = time.monotonic()
        raise DbError(f"NAS DB 에 연결하지 못했습니다: {type(e).__name__}: {e}")
    _last_fail = 0.0
    return conn


def test_connection() -> str:
    """설정 창 「시험」 — 붙어서 받은 위임 표를 세어 본다. 실패하면 DbError."""
    conn = connect(force=True)
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT COUNT(*) AS n, SUM(project_id IS NULL) AS waiting FROM assignments")
            row = cur.fetchone() or {}
    except Exception as e:
        raise DbError(f"assignments 표를 읽지 못했습니다: {e}")
    finally:
        conn.close()
    return f"연결됨 — {DB_NAME} · 받은 위임 {int(row.get('n') or 0)}줄 (등록 대기 {int(row.get('waiting') or 0)})"


# ── 받은 위임·받은 지시 쓰기 ──────────────────────────────────────────

@dataclass(frozen=True)
class Table:
    """`INSERT IGNORE` 로 쓰는 표 — 이름·칸·기록에 쓸 이름·한 줄 설명."""
    name: str
    columns: Tuple[str, ...]
    label: str
    describe: Callable[[dict], str]


ASSIGNMENTS = Table("assignments", ASSIGNMENT_COLUMNS, "받은 위임",
                    lambda r: f"{r.get('ref')} {r.get('doc_type') or ''} (보고 {r.get('report_due') or '-'})")
# 받은 지시(CPProject docs/INSTRUCTION_MAIL_PLAN.md 6절) — 원 발신인·고객 날짜는 ChekPat 이 본문에서 푼다
INSTRUCTION_COLUMNS = ("mail_entry_id", "message_id", "received_at", "from_name", "subject", "ref_hint",
                       "attachments", "body_text", "orig_text", "orig_head", "msg_path", "source")
INSTRUCTIONS = Table("instructions", INSTRUCTION_COLUMNS, "받은 지시",
                     lambda r: f"{r.get('ref_hint') or '-'} {(r.get('subject') or '')[:80]}")


def insert_rows(conn, rows: List[dict], table: Table = ASSIGNMENTS) -> int:
    """`INSERT IGNORE` 로 쓴다(같은 mail_entry_id 는 건너뜀). 커밋까지. 반환값: 새로 들어간 줄 수."""
    if not rows:
        return 0
    cols = ", ".join(table.columns)
    marks = ", ".join(["%s"] * len(table.columns))
    added = 0
    with conn.cursor() as cur:
        for row in rows:
            added += cur.execute(f"INSERT IGNORE INTO {table.name} ({cols}) VALUES ({marks})",
                                 tuple(row.get(c) for c in table.columns)) or 0
    conn.commit()
    return added


def insert_assignments(conn, rows: List[dict]) -> int:
    return insert_rows(conn, rows, ASSIGNMENTS)


# ── 출원인 약칭 (`applicant_abbr` — 일정 제목 꼬리, 두 PC·ChekPat 이 함께) ──────────────────────

def load_abbr(conn) -> Dict[str, str]:
    """NAS 약칭 표 {국문 명칭: 약칭}. 표가 없으면 DbError."""
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT name, abbr FROM applicant_abbr")
            return {r["name"]: r["abbr"] for r in cur.fetchall() if r.get("name") and r.get("abbr")}
    except Exception as e:
        raise DbError(f"applicant_abbr 표를 읽지 못했습니다: {e}")


def save_abbr(conn, rows: Dict[str, str], removed=()) -> Tuple[int, int]:
    """약칭을 쓰고(있으면 고침) `removed` 명칭은 지운다. 커밋까지. 반환값: (쓴 줄, 지운 줄)."""
    wrote = deleted = 0
    with conn.cursor() as cur:
        for name, abbr in rows.items():
            cur.execute("INSERT INTO applicant_abbr (name, abbr) VALUES (%s, %s) ON DUPLICATE KEY UPDATE abbr = VALUES(abbr)",
                        (name, abbr))
            wrote += 1
        for name in removed:
            deleted += cur.execute("DELETE FROM applicant_abbr WHERE name = %s", (name,)) or 0
    conn.commit()
    return wrote, deleted


def load_state(path: str) -> dict:
    """{"recorded": {EntryID: 쓴 날}, "pending": [못 쓴 줄]}."""
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError:
        data = {}
    except Exception as e:
        logging.error(f"NAS 기록 상태 로드 실패({os.path.basename(path)}): {e}")
        data = {}
    data = data if isinstance(data, dict) else {}
    return {"recorded": dict(data.get("recorded") or {}), "pending": list(data.get("pending") or [])}


def save_state(path: str, state: dict, today: Optional[date] = None) -> None:
    """쓴 지 RECORDED_KEEP_DAYS 가 지난 EntryID 는 지운다(위임 메일 스캔은 14일)."""
    cutoff = ((today or date.today()) - timedelta(days=RECORDED_KEEP_DAYS)).isoformat()
    data = {"recorded": {k: v for k, v in state["recorded"].items() if v >= cutoff}, "pending": state["pending"]}
    try:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        os.replace(tmp, path)
    except Exception as e:
        logging.error(f"NAS 기록 상태 저장 실패({os.path.basename(path)}): {e}")


def is_recorded(path: str, entry_id: str) -> bool:
    state = load_state(path)
    return entry_id in state["recorded"] or any(p.get("mail_entry_id") == entry_id for p in state["pending"])


def record(rows: List[dict], path: str, connect_fn=None, today: Optional[date] = None,
           table: Table = ASSIGNMENTS) -> dict:
    """받은 위임(또는 table 의) 줄을 NAS 에 쓴다 — 쌓여 있던 pending 도 함께. 예외를 내지 않는다.

    반환값: {"written": 새로 들어간 줄, "pending": 아직 못 쓴 줄 수, "error": 실패 사유 또는 ""}
    """
    state = load_state(path)
    todo = {p["mail_entry_id"]: p for p in state["pending"] if p.get("mail_entry_id")}
    for row in rows:
        if row.get("mail_entry_id") and row["mail_entry_id"] not in state["recorded"]:
            todo[row["mail_entry_id"]] = row
    if not todo:
        return {"written": 0, "pending": 0, "error": ""}
    error, written = "", 0
    try:
        conn = (connect_fn or connect)()
        try:
            written = insert_rows(conn, list(todo.values()), table)
        finally:
            conn.close()
    except Exception as e:
        error = str(e)
    day = (today or date.today()).isoformat()
    if error:
        state["pending"] = list(todo.values())
        logging.warning(f"{table.label} NAS 기록 실패 — {len(todo)}줄은 다음에 다시 씁니다: {error}")
    else:
        for entry_id, row in todo.items():
            state["recorded"][entry_id] = day
            logging.info(f"{table.label} NAS 기록: {table.describe(row)}")
        state["pending"] = []
        if written < len(todo):
            logging.info(f"{table.label} {len(todo) - written}줄은 이미 NAS 에 있었습니다.")
    save_state(path, state, today)
    return {"written": written, "pending": len(state["pending"]), "error": error}
