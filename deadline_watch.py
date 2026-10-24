"""기한 확인 — 트레이에 있는 동안 12:00(기한일 경고)·17:00(연장 확인)에 NAS `iplog` 차수를 본다. Qt·Outlook 을 모른다.

- 17:00 연장 확인: ChekPat 에서 기연을 의뢰하면 그 차수에 `ext_status='requested'`·`ext_requested_at`·`ext_months`
  가 적힌다. 회사 담당자는 보통 17:00 전에 기간연장신청서를 낸다. 그날 KIPRIS 서지 상세 서류 목록에 요청일 이후
  접수된 `[지정기간연장]`·`[법정기간연장]` 서류가 있으면(「접수중」 이어도) `confirmed` 를 적고, 연장 월수를 알면
  `due_extended`(대응기한 + 월수)를 적고 네이버 대응기한 일정을 옮긴다. 없으면 `unconfirmed` 를 적고 알린다(요청 뒤
  EXT_LOOKBACK_DAYS 동안 날마다 다시 본다). 마감기한(`ExpirationDuedateDate`)은 「수리」 뒤에야 늘어 그날은 못 쓴다.
  기연 요청 없이 KIPRIS 기한이 늦어진 진행 중 차수(대응기한이 가까운 것만 — KIPRIS 호출 수)도 고치고, 월수로 셈한
  연장 기한이 KIPRIS 에 반영된 기한과 다르면 KIPRIS 쪽으로 고친다.
- 12:00 기한일 경고: 오늘이 대응기한(`COALESCE(due_extended, due_date)`)인데 제출도 기연(확인)도 없는 진행 중 차수.
- 마지막으로 돈 날은 `deadline_state.json` — 앱이 꺼져 있어 놓친 확인은 켤 때 따라잡는다(「늦은 알림」).
- 알림: 트레이 말풍선(main 이 넘긴 함수) + 네이버 「알림」 캘린더 일정(VALARM) + 로그(History 탭).
"""

import json
import logging
import os
import re
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from typing import Callable, Dict, List, Optional

from dateutil.relativedelta import relativedelta

import db
import kipris

KINDS = ("warn", "ext")
KIND_NAMES = {"warn": "기한일 경고", "ext": "연장 확인"}
DEFAULT_TIMES = {"warn": "12:00", "ext": "17:00"}
CATCHUP_DAYS = 7                     # 이보다 오래 꺼져 있었으면 그 앞은 따라잡지 않는다
GRACE = timedelta(minutes=10)        # 정한 시각보다 이만큼 넘게 늦게 돌면 「늦은 알림」
RETRY_AFTER = timedelta(minutes=10)  # NAS 에 못 닿으면 이만큼 뒤에 다시
EXT_LOOKBACK_DAYS = 7                # 기연신청서가 안 보였거나 KIPRIS 오류였던 차수는 요청 뒤 이만큼 날마다 다시 본다
DRIFT_AHEAD_DAYS = 14                # 기연 요청 없이 기한이 바뀌었나 — 대응기한이 이 안에 드는 차수만
DRIFT_BEHIND_DAYS = 7                #   (기한 당일 기연 → KIPRIS 반영은 하루 이틀 늦다)
CLOSED_STATUSES = ("completed", "allowed", "abandoned")
LATE_PREFIX = "늦은 알림"
ALERT_LEAD = timedelta(minutes=1)    # 네이버 알림 일정은 지금보다 조금 뒤에 — 이미 지난 시각은 울리지 않는다

_OPEN = ("p.submitted_date IS NULL AND (p.status IS NULL OR p.status NOT IN (%s))"
         % ", ".join("'%s'" % s for s in CLOSED_STATUSES))
_ROUND_COLS = ("p.id, r.ref, r.app_no, p.stage, p.round_n, p.due_date, p.due_extended, p.ext_status, "
               "p.ext_requested_at, p.ext_months, COALESCE(p.due_extended, p.due_date) AS due")
_EXT_DOC = re.compile(r"\[(지정|법정)기간연장\]")   # 「[지정기간연장]기간 연장신청서·…」 — 「법정기간연장승인서」 는 아니다


@dataclass
class Alert:
    title: str
    lines: List[str] = field(default_factory=list)
    late: bool = False
    day: Optional[date] = None       # 확인한 날 (늦은 알림이면 그날을 적는다)

    @property
    def full_title(self) -> str:
        if not self.late:
            return self.title
        return f"{LATE_PREFIX}({self.day:%m-%d}) · {self.title}" if self.day else f"{LATE_PREFIX} · {self.title}"

    @property
    def text(self) -> str:
        return "\n".join(self.lines)


# ── 언제 도나 ─────────────────────────────────────────────────────────

def parse_times(times) -> Dict[str, time]:
    """{"warn": "12:00", "ext": "17:00"} → {kind: time}. 틀린 값은 기본 시각."""
    times = times if isinstance(times, dict) else {}
    out = {}
    for kind in KINDS:
        text = str(times.get(kind) or DEFAULT_TIMES[kind])
        m = re.fullmatch(r"\s*(\d{1,2}):(\d{2})\s*", text)
        try:
            out[kind] = time(int(m.group(1)), int(m.group(2))) if m else None
        except ValueError:
            out[kind] = None
        if out[kind] is None:
            h, mi = DEFAULT_TIMES[kind].split(":")
            out[kind] = time(int(h), int(mi))
    return out


def load_state(path: str) -> dict:
    """{kind: 마지막으로 확인을 마친 날 YYYY-MM-DD}. 파일이 없으면 {}."""
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except FileNotFoundError:
        return {}
    except Exception as e:
        logging.error(f"기한 확인 기록 로드 실패: {e}")
        return {}


def save_state(path: str, state: dict) -> None:
    try:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=2, ensure_ascii=False)
        os.replace(tmp, path)
    except Exception as e:
        logging.error(f"기한 확인 기록 저장 실패: {e}")


def due_runs(state: dict, now: datetime, times: Dict[str, time]):
    """지금 돌릴 확인 [(kind, 그날, 늦었나)] — 시각 순. 마지막으로 마친 날 다음날부터 오늘까지(최대 CATCHUP_DAYS).

    기록이 없으면(처음) 오늘 것만.
    """
    runs = []
    today = now.date()
    for kind in KINDS:
        try:
            start = date.fromisoformat(state.get(kind) or "") + timedelta(days=1)
        except ValueError:
            start = today
        start = max(start, today - timedelta(days=CATCHUP_DAYS))
        day = start
        while day <= today:
            slot = datetime.combine(day, times[kind])
            if slot <= now:
                runs.append((slot, kind, day, now - slot > GRACE))
            day += timedelta(days=1)
    return [(kind, day, late) for _slot, kind, day, late in sorted(runs)]


# ── DB 값 ────────────────────────────────────────────────────────────

def as_date(value) -> Optional[date]:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        return date.fromisoformat(str(value or "")[:10])
    except ValueError:
        return None


def round_label(row: dict) -> str:
    n = row.get("round_n")
    return " ".join(x for x in (row.get("ref") or "?", f"R{n}" if n else "", row.get("stage") or "") if x)


def round_due(notices: List[kipris.Notice], stage: str) -> Optional[date]:
    """KIPRIS 통지서 마감기한 가운데 이 차수의 것 — 차수 서류(OA·거절결정)와 맞는 가장 최신. 모르면 가장 최신."""
    stage = stage or ""
    if "거절" in stage:
        notice = kipris.pick_oa_due(notices, "거절결정")
    elif not stage or re.search(r"의견|OA|office", stage, re.I):
        notice = kipris.pick_oa_due(notices, "의견제출통지서")
    else:
        cands = [n for n in notices if n.due_date]
        notice = max(cands, key=lambda n: int(re.sub(r"\D", "", n.send_number) or 0)) if cands else None
    return as_date(notice.due_date) if notice else None


# ── 12:00 기한일 경고 ────────────────────────────────────────────────

def check_deadlines(conn, day: date, today: Optional[date] = None) -> List[Alert]:
    """그날이 대응기한인데 제출도 기연(확인)도 없는 진행 중 차수. 기연을 의뢰했으나 확인 안 된(unconfirmed) 차수도 알린다."""
    with conn.cursor() as cur:
        cur.execute(f"SELECT {_ROUND_COLS} FROM projects p JOIN refs r ON r.id = p.ref_id"
                    f" WHERE COALESCE(p.due_extended, p.due_date) = %s AND {_OPEN}"
                    " AND (p.ext_status IS NULL OR p.ext_status IN ('', 'unconfirmed')) ORDER BY r.ref, p.round_n",
                    (day.isoformat(),))
        rows = cur.fetchall()
    if not rows:
        logging.info(f"기한일 경고 {day}: 제출·기연 없는 차수 없음")
        return []
    lines = [round_label(r) + (" — 기연 의뢰했으나 기연신청서 제출 확인 안 됨" if r.get("ext_status") == "unconfirmed" else "")
             for r in rows]
    when = "오늘" if day == (today or date.today()) else day.isoformat()
    return [Alert(f"대응기한 {when} — 제출·기연 없음 {len(rows)}건",
                  lines, day=day)]


# ── 17:00 연장 확인 ──────────────────────────────────────────────────

def move_action_event(session, calendar_name: str, subjects: List[str], old_days, new_day: date) -> str:
    """네이버 대응기한 일정을 옮긴다 — (제목 후보 × 옛 날짜 후보) 가운데 있는 것. 없으면 첫 제목으로 새로 만든다."""
    olds = [d for d in dict.fromkeys(old_days) if d and d != new_day]
    for old in olds:
        for subject in subjects:
            if session.exists(calendar_name, subject, old):
                session.move_all_day(calendar_name, subject, old, new_day)
                return f"일정 옮김 {subject} {old} → {new_day}"
    if not subjects:
        return "일정 제목을 알 수 없어 옮기지 못함"
    if session.add_all_day(calendar_name, subjects[0], new_day, ""):
        return f"옛 일정을 찾지 못해 {subjects[0]} 을(를) {new_day} 에 새로 만듦 — 옛 일정을 확인하세요"
    return f"{subjects[0]} {new_day} 일정이 이미 있음"


class _KiprisCache:
    """한 번 도는 동안 출원번호마다 KIPRIS 를 한 번만 부른다."""

    def __init__(self, fetch):
        self.fetch, self.got = fetch, {}

    def get(self, row: dict):
        app_no = row.get("app_no") or ""
        if app_no not in self.got:
            try:
                self.got[app_no] = self.fetch(app_no)
            except kipris.KiprisError as e:
                self.got[app_no] = e
        got = self.got[app_no]
        if isinstance(got, Exception):
            raise got
        return got

    def due(self, row: dict) -> Optional[date]:
        return round_due(self.get(row), row.get("stage") or "")


def extension_filed(docs: List[kipris.Document], since: date) -> Optional[kipris.Document]:
    """since(기연 요청일) 이후 접수된 기간연장신청서 가운데 처음 것. 처리 상태(접수중·수리)는 따지지 않는다."""
    for doc in docs:
        received = as_date(doc.receipt_date)
        if _EXT_DOC.search(doc.name) and received and received >= since:
            return doc
    return None


def _move(mover, row, olds, new_day) -> str:
    try:
        return mover(row, olds, new_day)
    except Exception as e:
        logging.error(f"대응기한 일정 옮기기 실패: {round_label(row)} → {new_day} — {e}")
        return f"⚠ 일정 옮기기 실패({e}) — 네이버에서 직접 옮기세요"


def check_extensions(conn, day: date, fetch_docs=kipris.fetch_documents, fetch_due=kipris.fetch_due_dates,
                     mover=None, drift=True) -> List[Alert]:
    """기연을 의뢰한 차수에 기간연장신청서가 들어갔는지 KIPRIS 서류 목록으로 확인한다(+ drift 면 KIPRIS 기한이 바뀐 차수).

    mover(row, 옛 날짜들, 새 날짜) → 결과 글. 고친 값은 커밋하지 않는다(부르는 쪽이 커밋).
    """
    docs, kip = _KiprisCache(fetch_docs), _KiprisCache(fetch_due)
    confirmed, unconfirmed, failed, drifted, corrected = [], [], [], [], []
    with conn.cursor() as cur:
        cur.execute(f"SELECT {_ROUND_COLS} FROM projects p JOIN refs r ON r.id = p.ref_id"
                    " WHERE p.ext_status IN ('requested', 'unconfirmed') AND DATE(p.ext_requested_at) BETWEEN %s AND %s"
                    " ORDER BY r.ref, p.round_n",
                    ((day - timedelta(days=EXT_LOOKBACK_DAYS)).isoformat(), day.isoformat()))
        requested = cur.fetchall()
        for row in requested:
            label, current = round_label(row), as_date(row.get("due"))
            since = as_date(row.get("ext_requested_at")) or day
            try:
                filed = extension_filed(docs.get(row), since)
            except kipris.KiprisError as e:
                logging.error(f"연장 확인 KIPRIS 실패: {label} — {e}")
                failed.append(f"{label} — KIPRIS 조회 실패({e}), 내일 다시 확인")
                continue
            if not filed:
                if row.get("ext_status") != "unconfirmed":
                    cur.execute("UPDATE projects SET ext_status = 'unconfirmed' WHERE id = %s", (row["id"],))
                logging.warning(f"기연신청서 제출 확인 안 됨: {label} — {since} 이후 접수된 기간연장신청서 없음")
                unconfirmed.append(f"{label} — {since} 이후 낸 기간연장신청서가 KIPRIS 에 없음")
                continue
            months = row.get("ext_months")
            new = current + relativedelta(months=int(months)) if current and months else None
            if new:
                cur.execute("UPDATE projects SET due_extended = %s, ext_status = 'confirmed' WHERE id = %s",
                            (new.isoformat(), row["id"]))
                moved = _move(mover, row, [current, as_date(row.get("due_date"))], new) if mover else ""
                change = f"{current} → {new}" + (f" · {moved}" if moved else "")
            else:
                cur.execute("UPDATE projects SET ext_status = 'confirmed' WHERE id = %s", (row["id"],))
                change = "연장 월수를 몰라 대응기한·일정은 그대로 — KIPRIS 기한이 늘면 고칩니다"
            text = f"기연신청서 {filed.receipt_date} {filed.status} · {change}"
            logging.info(f"기연신청서 제출 확인: {label} — {text}")
            confirmed.append(f"{label}: {text}")

        if drift:
            done = {r["id"] for r in requested}
            cur.execute(f"SELECT {_ROUND_COLS} FROM projects p JOIN refs r ON r.id = p.ref_id"
                        f" WHERE {_OPEN} AND (p.ext_status IS NULL OR p.ext_status <> 'requested')"
                        " AND COALESCE(p.due_extended, p.due_date) BETWEEN %s AND %s"
                        " AND r.app_no IS NOT NULL AND r.app_no <> '' ORDER BY r.ref, p.round_n",
                        ((day - timedelta(days=DRIFT_BEHIND_DAYS)).isoformat(),
                         (day + timedelta(days=DRIFT_AHEAD_DAYS)).isoformat()))
            for row in cur.fetchall():
                if row["id"] in done:
                    continue
                label, current = round_label(row), as_date(row.get("due"))
                try:
                    new = kip.due(row)
                except kipris.KiprisError as e:
                    logging.error(f"기한 변경 확인 KIPRIS 실패: {label} — {e}")
                    continue
                # 늘었거나, 월수로 셈한 연장 기한이 KIPRIS 에 반영된 기한(처음 기한보다 늦음)과 다르면 고친다
                floor = min(d for d in (current, as_date(row.get("due_date"))) if d) if current else None
                if not (current and new and new != current and new > floor):
                    continue
                status = "confirmed" if row.get("ext_status") == "unconfirmed" else row.get("ext_status")
                cur.execute("UPDATE projects SET due_extended = %s, ext_status = %s WHERE id = %s",
                            (new.isoformat(), status or None, row["id"]))
                moved = _move(mover, row, [current, as_date(row.get("due_date"))], new) if mover else ""
                line = f"{label}: {current} → {new}" + (f" · {moved}" if moved else "")
                if row.get("ext_status") == "confirmed":
                    logging.info(f"연장 기한을 KIPRIS 반영 기한으로 고침: {label} {current} → {new} ({moved})")
                    corrected.append(line)
                else:
                    logging.info(f"기연 요청 없이 KIPRIS 기한이 바뀜: {label} {current} → {new} ({moved})")
                    drifted.append(line)

    if not (requested or drifted or corrected):
        logging.info(f"연장 확인 {day}: 기연 의뢰한 차수 없음")
    alerts = []
    for title, lines in ((f"기연신청서 제출 확인 안 됨 {len(unconfirmed)}건", unconfirmed),
                         (f"기간연장 확인 실패 {len(failed)}건", failed),
                         (f"기연신청서 제출 확인 {len(confirmed)}건", confirmed),
                         (f"기연 요청 없이 대응기한이 바뀜 {len(drifted)}건", drifted),
                         (f"연장 기한이 KIPRIS 와 달라 고침 {len(corrected)}건", corrected)):
        if lines:
            alerts.append(Alert(title, lines, day=day))
    return alerts


# ── 한 번 돌기 ───────────────────────────────────────────────────────

class Watcher:
    """tick(now) 을 자주(1분마다) 부르면 정한 시각이 지난 확인을 돌린다. 스레드 하나에서만 부른다."""

    def __init__(self, state_path: str, times_fn: Callable[[], dict], connect=None, fetch_due=None,
                 mover_factory: Optional[Callable[[], Callable]] = None, fetch_docs=None):
        self.state_path = state_path
        self.times_fn = times_fn
        self.connect = connect or db.connect
        self.fetch_due = fetch_due or kipris.fetch_due_dates
        self.fetch_docs = fetch_docs or kipris.fetch_documents
        self.mover_factory = mover_factory
        self._retry_at: Optional[datetime] = None

    def pending(self, now: Optional[datetime] = None):
        now = now or datetime.now()
        if self._retry_at and now < self._retry_at:
            return []
        return due_runs(load_state(self.state_path), now, parse_times(self.times_fn()))

    def tick(self, now: Optional[datetime] = None) -> List[Alert]:
        now = now or datetime.now()
        runs = self.pending(now)
        if not runs:
            return []
        last_ext = max((day for kind, day, _late in runs if kind == "ext"), default=None)
        movers = []

        def mover(row, olds, new_day):
            if not movers:
                movers.append(self.mover_factory())
            return movers[0](row, olds, new_day)

        alerts = []
        for kind, day, late in runs:
            try:
                conn = self.connect()
                try:
                    if kind == "warn":
                        got = check_deadlines(conn, day, now.date())
                    else:
                        got = check_extensions(conn, day, self.fetch_docs, self.fetch_due,
                                               mover if self.mover_factory else None, drift=(day == last_ext))
                    conn.commit()
                finally:
                    conn.close()
            except Exception as e:
                self._retry_at = now + RETRY_AFTER
                logging.warning(f"{KIND_NAMES[kind]}({day}) 을(를) 하지 못했습니다 — {RETRY_AFTER.seconds // 60}분 뒤 다시: {e}")
                break
            logging.info(f"{KIND_NAMES[kind]} {day} 마침" + (" (따라잡기)" if late else ""))
            for a in got:
                a.late = late
            alerts += got
            state = load_state(self.state_path)
            state[kind] = day.isoformat()
            save_state(self.state_path, state)
        return alerts


def deliver(alerts: List[Alert], tray: Optional[Callable[[str, str], None]], session, alert_calendar: str,
            now: Optional[datetime] = None) -> None:
    """알림 — 로그(History 탭) · 트레이 말풍선 · 네이버 알림 캘린더(조금 뒤 시각, VALARM). 예외를 내지 않는다."""
    now = now or datetime.now()
    for a in alerts:
        title, text = a.full_title, a.text
        logging.warning(f"[기한 확인] {title}\n  " + "\n  ".join(a.lines))
        if tray:
            try:
                tray(title, text)
            except Exception as e:
                logging.error(f"트레이 알림 실패: {e}")
        if alert_calendar and session is not None:
            try:
                session.add_alert(alert_calendar, title, (now + ALERT_LEAD).astimezone(), text)
            except Exception as e:
                logging.error(f"네이버 알림 일정 등록 실패: {title} — {e}")
