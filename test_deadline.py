"""기한 확인(deadline_watch.py)·네이버 일정 옮기기·알림(naver_cal.py) 테스트 — NAS·KIPRIS·네이버는 가짜."""

import os
import tempfile
import unittest
from datetime import date, datetime, time, timedelta, timezone
from unittest.mock import MagicMock, patch

# test_calendar 가 (없으면) win32com·PySide6 를 가짜로 끼우고 임시 데이터 폴더를 정한다 — main 보다 먼저 들인다
import test_calendar  # noqa: F401

import deadline_watch as dw  # noqa: E402
import kipris  # noqa: E402
import main  # noqa: E402
import naver_cal  # noqa: E402

TIMES = {"warn": time(12, 0), "ext": time(17, 0)}


class FakeCursor:
    def __init__(self, conn):
        self.conn = conn
        self._rows = []

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def execute(self, sql, params=None):
        self.conn.executed.append((sql, params))
        self._rows = self.conn.respond(sql, params) if sql.startswith("SELECT") else []
        return 1

    def fetchall(self):
        return list(self._rows)


class FakeConn:
    """SELECT 는 respond(sql, params) 가 돌려준 줄, UPDATE 는 기록만."""

    def __init__(self, respond):
        self.respond = respond
        self.executed = []
        self.commits = 0
        self.closed = False

    def cursor(self):
        return FakeCursor(self)

    def commit(self):
        self.commits += 1

    def close(self):
        self.closed = True

    def updates(self):
        return [(sql, params) for sql, params in self.executed if sql.startswith("UPDATE")]


def _round(id_, ref="PA-2T8084", stage="의견제출통지서", due_date=date(2026, 10, 24), due_extended=None,
           ext_status=None, app_no="10-2025-7000694", round_n=2, ext_requested_at=None, ext_months=None):
    return {"id": id_, "ref": ref, "app_no": app_no, "stage": stage, "round_n": round_n, "due_date": due_date,
            "due_extended": due_extended, "ext_status": ext_status, "ext_requested_at": ext_requested_at,
            "ext_months": ext_months, "due": due_extended or due_date}


def _notice(due, title="의견제출통지서", send="9520260000001"):
    return kipris.Notice(send_number=send, title=title, due_date=due)


EXT_NAME = "[지정기간연장]기간 연장신청서·기간 단축신청서·기간 경과 구제신청서·절차 계속신청서"


def _doc(received, name=EXT_NAME, status="접수중"):
    return kipris.Document(name=name, receipt_date=received, status=status, number="1-1-2026-1315776-24")


class TestDueRuns(unittest.TestCase):
    def test_first_run_only_today_passed(self):
        runs = dw.due_runs({}, datetime(2026, 10, 24, 12, 5), TIMES)
        self.assertEqual(runs, [("warn", date(2026, 10, 24), False)])

    def test_catch_up_after_pc_off(self):
        """10-22 연장 확인까지 마치고 꺼졌다가 10-24 09:00 에 켬 — 10-23 둘 다 늦은 알림, 10-24 는 아직"""
        state = {"warn": "2026-10-22", "ext": "2026-10-22"}
        runs = dw.due_runs(state, datetime(2026, 10, 24, 9, 0), TIMES)
        self.assertEqual(runs, [("warn", date(2026, 10, 23), True), ("ext", date(2026, 10, 23), True)])

    def test_grace_and_done(self):
        state = {"warn": "2026-10-23", "ext": "2026-10-23"}
        self.assertEqual(dw.due_runs(state, datetime(2026, 10, 24, 17, 9), TIMES),
                         [("warn", date(2026, 10, 24), True), ("ext", date(2026, 10, 24), False)])
        state = {"warn": "2026-10-24", "ext": "2026-10-24"}
        self.assertEqual(dw.due_runs(state, datetime(2026, 10, 24, 23, 0), TIMES), [])

    def test_catch_up_limited(self):
        runs = dw.due_runs({"warn": "2026-01-01", "ext": "2026-01-01"}, datetime(2026, 10, 24, 8, 0), TIMES)
        self.assertEqual(min(d for _k, d, _l in runs), date(2026, 10, 24) - timedelta(days=dw.CATCHUP_DAYS))

    def test_parse_times(self):
        self.assertEqual(dw.parse_times({"warn": "9:30", "ext": "25:00"}), {"warn": time(9, 30), "ext": time(17, 0)})
        self.assertEqual(dw.parse_times(None), TIMES)


class TestCheckDeadlines(unittest.TestCase):
    def test_alert_rows(self):
        rows = [_round(1), _round(2, ref="PA-2T9001", stage="거절결정서", round_n=1, ext_status="unconfirmed")]
        conn = FakeConn(lambda sql, params: rows)
        alerts = dw.check_deadlines(conn, date(2026, 10, 24), today=date(2026, 10, 24))
        sql, params = conn.executed[0]
        self.assertEqual(params, ("2026-10-24",))
        self.assertIn("COALESCE(p.due_extended, p.due_date) = %s", sql)
        self.assertIn("p.submitted_date IS NULL", sql)
        self.assertIn("NOT IN ('completed', 'allowed', 'abandoned')", sql)
        self.assertIn("p.status IS NULL OR", sql)  # status 가 비어 있는 옛 차수도
        self.assertEqual(len(alerts), 1)
        self.assertEqual(alerts[0].title, "대응기한 오늘 — 제출·기연 없음 2건")
        self.assertEqual(alerts[0].lines, ["PA-2T8084 R2 의견제출통지서",
                                           "PA-2T9001 R1 거절결정서 — 기연 의뢰했으나 기연신청서 제출 확인 안 됨"])

    def test_nothing_no_alert(self):
        self.assertEqual(dw.check_deadlines(FakeConn(lambda s, p: []), date(2026, 10, 24)), [])

    def test_late_title(self):
        a = dw.Alert("대응기한 2026-10-23 — 제출·기연 없음 1건", ["x"], late=True, day=date(2026, 10, 23))
        self.assertEqual(a.full_title, "늦은 알림(10-23) · 대응기한 2026-10-23 — 제출·기연 없음 1건")


class TestCheckExtensions(unittest.TestCase):
    DAY = date(2026, 10, 23)
    ASKED = datetime(2026, 10, 23, 10, 30)

    def _conn(self, requested, drift=()):
        def respond(sql, params):
            return list(requested) if "ext_status IN ('requested', 'unconfirmed')" in sql else list(drift)
        return FakeConn(respond)

    def _asked(self, id_=7, **kw):
        kw.setdefault("due_date", date(2026, 11, 23))
        kw.setdefault("ext_status", "requested")
        return _round(id_, ext_requested_at=self.ASKED, **kw)

    def test_filed_today_moves_event_by_months(self):
        """10-2021-7008131 처럼 — 옛 연장신청서(수리)는 빼고, 그날 낸 것(접수중)으로 확인. 마감기한은 아직 그대로"""
        conn = self._conn([self._asked(ext_months=1)])
        mover = MagicMock(return_value="일정 옮김")
        docs = [_doc("2024-11-18", status="수리"), _doc("2026-10-23")]
        alerts = dw.check_extensions(conn, self.DAY, fetch_docs=lambda app: docs, mover=mover, drift=False)
        self.assertEqual(conn.updates(), [("UPDATE rounds SET due_extended = %s, ext_status = 'confirmed' WHERE id = %s",
                                           ("2026-12-23", 7))])
        row, olds, new = mover.call_args.args
        self.assertEqual((row["id"], olds, new), (7, [date(2026, 11, 23), date(2026, 11, 23)], date(2026, 12, 23)))
        self.assertEqual([a.title for a in alerts], ["기연신청서 제출 확인 1건"])
        self.assertIn("기연신청서 2026-10-23 접수중 · 2026-11-23 → 2026-12-23 · 일정 옮김", alerts[0].lines[0])
        sql, params = conn.executed[0]
        self.assertIn("DATE(p.ext_requested_at) BETWEEN %s AND %s", sql)
        self.assertIn("p.ext_months", sql)
        self.assertEqual(params, ("2026-10-16", "2026-10-23"))

    def test_not_filed_is_unconfirmed(self):
        """요청일 전의 연장신청서·연장승인서는 이번 기연이 아니다"""
        conn = self._conn([self._asked(ext_months=1)])
        mover = MagicMock()
        docs = [_doc("2026-10-20"), _doc("2026-10-23", name="법정기간연장승인서", status="발송처리완료")]
        alerts = dw.check_extensions(conn, self.DAY, fetch_docs=lambda app: docs, mover=mover, drift=False)
        self.assertEqual(conn.updates(), [("UPDATE rounds SET ext_status = 'unconfirmed' WHERE id = %s", (7,))])
        mover.assert_not_called()
        self.assertEqual(alerts[0].title, "기연신청서 제출 확인 안 됨 1건")
        self.assertIn("2026-10-23 이후 낸 기간연장신청서가 KIPRIS 에 없음", alerts[0].lines[0])

    def test_unconfirmed_rechecked_next_day(self):
        """17:00 뒤에 냈으면 다음 날 확인된다. 아직 없으면 다시 알리되 DB 는 그대로"""
        row = self._asked(ext_status="unconfirmed", ext_months=2)
        conn = self._conn([row])
        alerts = dw.check_extensions(conn, date(2026, 10, 24), fetch_docs=lambda app: [_doc("2026-10-23")],
                                     mover=None, drift=False)
        self.assertEqual(conn.updates()[0][1], ("2027-01-23", 7))
        self.assertEqual(alerts[0].title, "기연신청서 제출 확인 1건")
        conn = self._conn([row])
        alerts = dw.check_extensions(conn, date(2026, 10, 24), fetch_docs=lambda app: [], mover=None, drift=False)
        self.assertEqual(conn.updates(), [])
        self.assertEqual(alerts[0].title, "기연신청서 제출 확인 안 됨 1건")

    def test_months_unknown_confirms_without_moving(self):
        conn = self._conn([self._asked()])
        mover = MagicMock()
        alerts = dw.check_extensions(conn, self.DAY, fetch_docs=lambda app: [_doc("2026-10-23")], mover=mover,
                                     drift=False)
        self.assertEqual(conn.updates(), [("UPDATE rounds SET ext_status = 'confirmed' WHERE id = %s", (7,))])
        mover.assert_not_called()
        self.assertIn("연장 월수를 몰라", alerts[0].lines[0])

    def test_statutory_extension_counts(self):
        """거절결정 뒤 기간은 「[법정기간연장]…」"""
        conn = self._conn([self._asked(stage="거절결정서", ext_months=1)])
        name = "[법정기간연장]기간 연장신청서·기간 단축신청서·기간 경과 구제신청서·절차 계속신청서"
        dw.check_extensions(conn, self.DAY, fetch_docs=lambda app: [_doc("2026-10-23", name=name, status="수리")],
                            mover=None, drift=False)
        self.assertEqual(conn.updates()[0][1], ("2026-12-23", 7))

    def test_kipris_failure_keeps_requested(self):
        conn = self._conn([self._asked()])

        def fail(app):
            raise kipris.KiprisError("연결 실패")
        alerts = dw.check_extensions(conn, self.DAY, fetch_docs=fail, mover=MagicMock(), drift=False)
        self.assertEqual(conn.updates(), [])
        self.assertEqual(alerts[0].title, "기간연장 확인 실패 1건")

    def test_second_extension_adds_to_extended(self):
        conn = self._conn([self._asked(due_extended=date(2026, 12, 23), ext_months=1)])
        mover = MagicMock(return_value="")
        dw.check_extensions(conn, self.DAY, fetch_docs=lambda app: [_doc("2026-10-23")], mover=mover, drift=False)
        self.assertEqual(conn.updates()[0][1], ("2027-01-23", 7))
        self.assertEqual(mover.call_args.args[1], [date(2026, 12, 23), date(2026, 11, 23)])

    def test_drift_without_request(self):
        requested = [self._asked(due_date=date(2026, 11, 2))]
        drift = [_round(7, due_date=date(2026, 11, 2)),           # 위에서 본 차수 — 다시 보지 않는다
                 _round(8, ref="PA-2T9001", app_no="10-2024-0000001", due_date=date(2026, 10, 30)),
                 _round(9, ref="PA-2T9002", app_no="10-2024-0000002", due_date=date(2026, 10, 30),
                        ext_status="unconfirmed"),
                 _round(10, ref="PA-2T9003", app_no="10-2024-0000003", due_date=date(2026, 10, 30))]
        dues = {"10-2024-0000001": "2026-11-30", "10-2024-0000002": "2026-11-30", "10-2024-0000003": "2026-10-30"}
        calls = []

        def fetch(app):
            calls.append(app)
            return [_notice(dues[app])]
        conn = self._conn(requested, drift)
        alerts = dw.check_extensions(conn, self.DAY, fetch_docs=lambda app: [], fetch_due=fetch,
                                     mover=MagicMock(return_value="옮김"))
        self.assertEqual(calls, ["10-2024-0000001", "10-2024-0000002", "10-2024-0000003"])
        ups = conn.updates()
        self.assertEqual(ups[0], ("UPDATE rounds SET ext_status = 'unconfirmed' WHERE id = %s", (7,)))
        self.assertEqual(ups[1], ("UPDATE rounds SET due_extended = %s, ext_status = %s WHERE id = %s",
                                  ("2026-11-30", None, 8)))
        self.assertEqual(ups[2][1], ("2026-11-30", "confirmed", 9))
        self.assertEqual(len(ups), 3)
        drift_sql, drift_params = conn.executed[2]
        self.assertIn("p.ext_status <> 'requested'", drift_sql)
        self.assertEqual(drift_params, ("2026-10-16", "2026-10-30"))
        self.assertEqual(alerts[-1].title, "기연 요청 없이 대응기한이 바뀜 2건")

    def test_drift_corrects_counted_extension(self):
        """월수로 셈한 12-23 이 KIPRIS 수리 뒤 12-20 이면 고친다. KIPRIS 가 아직 처음 기한(11-23)이면 두고 본다"""
        row = _round(7, due_date=date(2026, 11, 23), due_extended=date(2026, 12, 23), ext_status="confirmed")
        for kipris_due, expect in (("2026-12-20", [("2026-12-20", "confirmed", 7)]), ("2026-11-23", []),
                                   ("2026-12-23", [])):
            conn = self._conn([], [row])
            mover = MagicMock(return_value="옮김")
            alerts = dw.check_extensions(conn, date(2026, 12, 10), fetch_docs=lambda app: [],
                                         fetch_due=lambda app: [_notice(kipris_due)], mover=mover)
            self.assertEqual([p for _s, p in conn.updates()], expect)
            if expect:
                self.assertEqual(mover.call_args.args[1:], ([date(2026, 12, 23), date(2026, 11, 23)],
                                                            date(2026, 12, 20)))
                self.assertEqual(alerts[-1].title, "연장 기한이 KIPRIS 와 달라 고침 1건")
            else:
                self.assertEqual(alerts, [])

    def test_one_kipris_call_per_app_no(self):
        rows = [self._asked(1), self._asked(2, round_n=3)]
        fetch = MagicMock(return_value=[_doc("2026-10-23")])
        dw.check_extensions(self._conn(rows), self.DAY, fetch_docs=fetch, mover=None, drift=False)
        fetch.assert_called_once_with("10-2025-7000694")

    def test_parse_documents(self):
        xml = ("<response><header><successYN>Y</successYN><resultCode>00</resultCode></header><body><item>"
               "<biblioSummaryInfo><applicationNumber>1020217008131</applicationNumber>"
               "<finalDisposal>거절결정 후 재심사중</finalDisposal></biblioSummaryInfo>"
               "<legalStatusInfoArray><legalStatusInfo><commonCodeName>수리 (Accepted)</commonCodeName>"
               "<documentEngName/><documentName>[지정기간연장]기간 연장신청서</documentName>"
               "<receiptDate>2024.11.18</receiptDate><receiptNumber>1-1-2024-1262771-06</receiptNumber>"
               "</legalStatusInfo><legalStatusInfo><commonCodeName>접수중 (On receiving)</commonCodeName>"
               "<documentEngName/><documentName>" + EXT_NAME + "</documentName>"
               "<receiptDate>2026.10.23</receiptDate><receiptNumber>1-1-2026-1315776-24</receiptNumber>"
               "</legalStatusInfo></legalStatusInfoArray></item></body></response>")
        docs = kipris.parse_documents(xml)
        self.assertEqual(docs, [kipris.Document("[지정기간연장]기간 연장신청서", "2024-11-18", "수리", "1-1-2024-1262771-06"),
                                kipris.Document(EXT_NAME, "2026-10-23", "접수중", "1-1-2026-1315776-24")])
        self.assertEqual(dw.extension_filed(docs, date(2026, 10, 23)), docs[1])
        self.assertIsNone(dw.extension_filed(docs, date(2026, 10, 24)))

    def test_round_due_by_stage(self):
        notices = [_notice("2026-11-02", "의견제출통지서", "9520260000001"),
                   _notice("2027-01-05", "거절결정서", "9520260000009")]
        self.assertEqual(dw.round_due(notices, "의견제출통지서"), date(2026, 11, 2))
        self.assertEqual(dw.round_due(notices, "OA"), date(2026, 11, 2))
        self.assertEqual(dw.round_due(notices, "재심사 거절결정서"), date(2027, 1, 5))
        self.assertEqual(dw.round_due(notices, "보정요구서"), date(2027, 1, 5))
        self.assertIsNone(dw.round_due([], "OA"))


class TestMoveActionEvent(unittest.TestCase):
    def test_moves_first_existing_candidate(self):
        session = MagicMock()
        session.exists.side_effect = lambda cal, subj, day: (subj, day) == ("PA-2T8084_의견서", date(2026, 11, 2))
        text = dw.move_action_event(session, "기한", ["PA-2T8084_의견서_아우디", "PA-2T8084_의견서"],
                                    [date(2026, 12, 2), date(2026, 11, 2)], date(2027, 1, 2))
        session.move_all_day.assert_called_once_with("기한", "PA-2T8084_의견서", date(2026, 11, 2), date(2027, 1, 2))
        self.assertIn("일정 옮김", text)

    def test_creates_when_missing(self):
        session = MagicMock()
        session.exists.return_value = False
        session.add_all_day.return_value = True
        text = dw.move_action_event(session, "기한", ["A", "B"], [date(2026, 11, 2), None], date(2026, 12, 2))
        session.add_all_day.assert_called_once_with("기한", "A", date(2026, 12, 2), "")
        self.assertIn("새로 만듦", text)

    def test_action_subjects(self):
        with patch.object(main, "load_case_cache", return_value={
                "PA-2T8084|의견제출통지서": {"doc_type": "의견제출통지서", "receipt_date": "2026-10-12",
                                       "applicant": "아우디 아게"}}), \
                patch.dict(main.config, {"applicant_abbr": {}}):
            self.assertEqual(main.action_subjects("PA-2T8084", "의견제출통지서"),
                             ["PA-2T8084_의견서_아우디", "PA-2T8084_의견서"])
            self.assertEqual(main.action_subjects("PA-2T9001", "재심사 거절결정서"),
                             ["PA-2T9001_거결불", "PA-2T9001_재심사"])
            self.assertEqual(main.action_subjects("PA-2T9001", "거절결정서")[0], "PA-2T9001_재심사")


class TestWatcher(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.path = os.path.join(self.tmp.name, "deadline_state.json")

    def test_runs_and_records_late(self):
        dw.save_state(self.path, {"warn": "2026-10-22", "ext": "2026-10-22"})
        conns = []

        def connect():
            conns.append(FakeConn(lambda sql, params: [_round(1, due_date=date(2026, 10, 23))]
                                  if "= %s AND" in sql else []))
            return conns[-1]
        w = dw.Watcher(self.path, lambda: {}, connect=connect, fetch_due=MagicMock(), mover_factory=MagicMock())
        alerts = w.tick(datetime(2026, 10, 24, 9, 0))
        self.assertEqual(dw.load_state(self.path), {"warn": "2026-10-23", "ext": "2026-10-23"})
        self.assertEqual(len(alerts), 1)
        self.assertTrue(alerts[0].late)
        self.assertTrue(alerts[0].full_title.startswith("늦은 알림(10-23)"))
        self.assertTrue(all(c.closed and c.commits == 1 for c in conns))
        self.assertEqual(w.tick(datetime(2026, 10, 24, 9, 1)), [])

    def test_db_down_retries_later_without_marking(self):
        connect = MagicMock(side_effect=main.db.DbError("NAS 꺼짐"))
        w = dw.Watcher(self.path, lambda: {"warn": "12:00"}, connect=connect)
        self.assertEqual(w.tick(datetime(2026, 10, 24, 12, 1)), [])
        self.assertEqual(dw.load_state(self.path), {})
        self.assertEqual(w.pending(datetime(2026, 10, 24, 12, 5)), [])        # 10분 쉬는 중
        self.assertEqual(len(w.pending(datetime(2026, 10, 24, 12, 12))), 1)
        self.assertEqual(connect.call_count, 1)

    def test_drift_only_on_last_ext_day(self):
        dw.save_state(self.path, {"warn": "2026-10-24", "ext": "2026-10-21"})
        seen = []
        with patch.object(dw, "check_extensions", side_effect=lambda conn, day, docs, due, mover, drift: seen.append(
                (day, drift)) or []):
            dw.Watcher(self.path, lambda: {}, connect=lambda: FakeConn(lambda s, p: [])).tick(
                datetime(2026, 10, 24, 18, 0))
        self.assertEqual(seen, [(date(2026, 10, 22), False), (date(2026, 10, 23), False), (date(2026, 10, 24), True)])

    def test_deliver(self):
        tray = MagicMock()
        session = MagicMock()
        now = datetime(2026, 10, 24, 12, 0)
        dw.deliver([dw.Alert("대응기한 오늘 — 제출·기연 없음 1건", ["PA-2T8084 R2"])], tray, session, "알림", now)
        tray.assert_called_once_with("대응기한 오늘 — 제출·기연 없음 1건", "PA-2T8084 R2")
        cal, title, start, body = session.add_alert.call_args.args
        self.assertEqual((cal, body), ("알림", "PA-2T8084 R2"))
        self.assertEqual(start.replace(tzinfo=None), now + dw.ALERT_LEAD)
        session.add_alert.reset_mock()
        session.add_alert.side_effect = naver_cal.NaverError("끊김")
        dw.deliver([dw.Alert("x", ["y"])], None, session, "알림", now)   # 예외를 내지 않는다
        dw.deliver([dw.Alert("x", ["y"])], None, session, "", now)
        self.assertEqual(session.add_alert.call_count, 1)


class TestNaverMoveAndAlert(unittest.TestCase):
    """naver_cal — 일정 옮기기(DELETE + PUT)·알림 일정(VALARM)"""

    def _session(self, existing=()):
        s = naver_cal.Session("id", "pw")
        s._calendars = {"기한": "https://caldav.calendar.naver.com/calendars/me/a/",
                        "알림": "https://caldav.calendar.naver.com/calendars/me/b/"}
        old_ics = naver_cal.build_ics("u", "PA-2T8084_의견서", date(2026, 11, 2), "출원인: 아우디 아게\n출원번호: 10-1")
        store = {naver_cal.event_url_for(s._calendars[c], naver_cal.make_uid(c, subj, d)): old_ics
                 for c, subj, d in existing}
        calls = []

        def request(method, url, body="", headers=None):
            calls.append((method, url, body))
            if method == "GET":
                return (200, url, store[url]) if url in store else (404, url, "")
            if method == "DELETE":
                return (204, url, "") if store.pop(url, None) else (404, url, "")
            store[url] = body
            return 201, url, ""
        s._request = request
        return s, calls, store

    def test_move_puts_new_then_deletes_old_uid(self):
        s, calls, store = self._session([("기한", "PA-2T8084_의견서", date(2026, 11, 2))])
        self.assertEqual(s.move_all_day("기한", "PA-2T8084_의견서", date(2026, 11, 2), date(2026, 12, 2)), (True, True))
        methods = [m for m, _u, _b in calls]
        self.assertEqual(methods, ["GET", "GET", "PUT", "DELETE"])
        old_url = naver_cal.event_url_for(s._calendars["기한"], naver_cal.make_uid("기한", "PA-2T8084_의견서",
                                                                                   date(2026, 11, 2)))
        self.assertEqual(calls[-1][1], old_url)
        self.assertNotIn(old_url, store)
        put_body = calls[2][2].replace("\r\n ", "")
        self.assertIn("DTSTART;VALUE=DATE:20261202", put_body)
        self.assertIn("DESCRIPTION:출원인: 아우디 아게\\n출원번호: 10-1", put_body)  # 옛 설명칸을 옮긴다

    def test_move_missing_old_just_creates(self):
        s, calls, _store = self._session()
        self.assertEqual(s.move_all_day("기한", "x", date(2026, 11, 2), date(2026, 12, 2)), (True, False))
        self.assertNotIn("DELETE", [m for m, _u, _b in calls])

    def test_delete(self):
        s, _calls, _store = self._session([("기한", "x", date(2026, 11, 2))])
        self.assertTrue(s.delete_all_day("기한", "x", date(2026, 11, 2)))
        self.assertFalse(s.delete_all_day("기한", "x", date(2026, 11, 2)))

    def test_alert_ics_has_valarm(self):
        start = datetime(2026, 10, 24, 12, 1, tzinfo=timezone(timedelta(hours=9)))
        ics = naver_cal.build_alert_ics("u@tmproject", "대응기한 오늘", start, "PA-2T8084 R2",
                                        now=datetime(2026, 10, 24, 3, 0, tzinfo=timezone.utc))
        self.assertIn("DTSTART:20261024T030100Z\r\n", ics)
        self.assertIn("DTEND:20261024T031100Z\r\n", ics)
        self.assertIn("BEGIN:VALARM\r\nACTION:DISPLAY\r\n", ics)
        self.assertIn("TRIGGER:PT0S\r\n", ics)
        self.assertEqual(naver_cal.ics_description(ics), "PA-2T8084 R2")

    def test_add_alert_once_per_day(self):
        s, calls, _store = self._session()
        start = datetime(2026, 10, 24, 12, 1).astimezone()
        self.assertTrue(s.add_alert("알림", "대응기한 오늘", start, "x"))
        self.assertFalse(s.add_alert("알림", "대응기한 오늘", start, "x"))
        self.assertEqual([m for m, _u, _b in calls].count("PUT"), 1)

    def test_ics_description_unescape(self):
        ics = naver_cal.build_ics("u", "s", date(2026, 1, 1), "a;b,c\\d\n" + "가" * 60)
        self.assertEqual(naver_cal.ics_description(ics), "a;b,c\\d\n" + "가" * 60)
        self.assertEqual(naver_cal.ics_description(naver_cal.build_ics("u", "s", date(2026, 1, 1), "")), "")


if __name__ == "__main__":
    unittest.main(verbosity=2)
