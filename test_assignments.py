"""받은 위임 → NAS assignments (db.py·main.assignment_row) 테스트 — 실제 NAS 에는 붙지 않는다."""

import os
import sys
import tempfile
import unittest
from datetime import date
from unittest.mock import MagicMock, patch

# test_calendar 가 (없으면) win32com·PySide6 를 가짜로 끼우고 임시 데이터 폴더를 정한다 — main 보다 먼저 들인다
from test_calendar import FR_SUBJECT, OA_SUBJECT, make_mock_message  # noqa: E402

import db  # noqa: E402
import main  # noqa: E402


class FakeCursor:
    def __init__(self, conn):
        self.conn = conn

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def execute(self, sql, params=None):
        self.conn.executed.append((sql, params))
        if sql.startswith("INSERT IGNORE"):
            entry_id = params[0]
            if entry_id in self.conn.rows:
                return 0
            self.conn.rows[entry_id] = params
            return 1
        return 0


class FakeConn:
    def __init__(self):
        self.executed, self.rows = [], {}
        self.commits = 0
        self.closed = False

    def cursor(self):
        return FakeCursor(self)

    def commit(self):
        self.commits += 1

    def close(self):
        self.closed = True


def _row(entry_id, ref="PA-2T8084"):
    return {"mail_entry_id": entry_id, "ref": ref, "doc_type": "OA", "report_due": "2026-10-01"}


class TestRecord(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.path = os.path.join(self.tmp.name, "assignment_state.json")

    def test_insert_ignore_and_recorded(self):
        conn = FakeConn()
        res = db.record([_row("A"), _row("B")], self.path, connect_fn=lambda: conn)
        self.assertEqual(res, {"written": 2, "pending": 0, "error": ""})
        self.assertTrue(all(sql.startswith("INSERT IGNORE INTO assignments (mail_entry_id, ref,")
                            for sql, _p in conn.executed))
        self.assertEqual(conn.commits, 1)
        self.assertTrue(conn.closed)
        self.assertEqual(set(db.load_state(self.path)["recorded"]), {"A", "B"})
        # 이미 쓴 메일은 다시 보내지 않는다
        conn2 = FakeConn()
        self.assertEqual(db.record([_row("A")], self.path, connect_fn=lambda: conn2)["written"], 0)
        self.assertEqual(conn2.executed, [])

    def test_offline_queues_then_flushes(self):
        def offline():
            raise db.DbError("NAS DB 에 연결하지 못했습니다")
        res = db.record([_row("A")], self.path, connect_fn=offline)
        self.assertEqual(res["pending"], 1)
        self.assertIn("연결하지 못했습니다", res["error"])
        res = db.record([_row("B")], self.path, connect_fn=offline)
        self.assertEqual(res["pending"], 2)
        self.assertTrue(db.is_recorded(self.path, "A"))  # pending 도 「알고 있는」 메일

        conn = FakeConn()
        conn.rows["A"] = ("A",)  # 다른 PC 가 이미 쓴 줄 — IGNORE
        res = db.record([], self.path, connect_fn=lambda: conn)
        self.assertEqual(res, {"written": 1, "pending": 0, "error": ""})
        state = db.load_state(self.path)
        self.assertEqual(state["pending"], [])
        self.assertEqual(set(state["recorded"]), {"A", "B"})

    def test_recorded_pruned(self):
        conn = FakeConn()
        db.record([_row("old")], self.path, connect_fn=lambda: conn, today=date(2026, 9, 1))
        db.record([_row("new")], self.path, connect_fn=lambda: conn, today=date(2026, 10, 24))
        self.assertEqual(set(db.load_state(self.path)["recorded"]), {"new"})

    def test_credentials_from_chekpat_file(self):
        blob = b'{"DB_HOST": "nas.local", "DB_PORT": "", "DB_USERNAME": "u", "DB_PASSWORD": "p"}'
        with patch.object(db.kipris, "_unprotect", return_value=blob) as unprotect, \
                patch.object(db.kipris, "_chekpat_dir", return_value=MagicMock()):
            self.assertEqual(db.credentials(), {"host": "nas.local", "port": 3307, "user": "u", "password": "p"})
        unprotect.assert_called_once()
        with patch.object(db.kipris, "_unprotect", side_effect=OSError):
            self.assertIsNone(db.credentials())

    def test_tests_never_see_real_credentials(self):
        """시험은 실제 ChekPat 접속 정보를 읽지 않는다 — 윈도에서 시험이 NAS 에 받은 위임을 썼다(2026-10-24)."""
        self.assertTrue(str(db.kipris._chekpat_dir()).startswith(os.environ["APPDATA"]))
        self.assertNotIn("TaskManager", os.environ["APPDATA"])
        self.assertIsNone(db.credentials())
        import pymysql
        with self.assertRaises(OSError):
            pymysql.connect(host="nas")

    def test_connect_without_credentials(self):
        with patch.object(db, "credentials", return_value=None):
            with self.assertRaises(db.DbError):
                db.connect(force=True)

    def test_connect_backoff_after_failure(self):
        fake = MagicMock()
        fake.connect.side_effect = OSError("timed out")
        cred = {"host": "h", "port": 3307, "user": "u", "password": "p"}
        with patch.object(db, "credentials", return_value=cred), \
                patch.dict(sys.modules, {"pymysql": fake, "pymysql.cursors": fake.cursors}):
            with self.assertRaises(db.DbError):
                db.connect()
            with self.assertRaisesRegex(db.DbError, "쉬는 중"):
                db.connect()
            self.assertEqual(fake.connect.call_count, 1)
            fake.connect.side_effect = None
            db.connect(force=True)
            kwargs = fake.connect.call_args.kwargs
            self.assertEqual((kwargs["database"], kwargs["charset"], kwargs["port"]), ("iplog", "utf8mb4", 3307))
        db._last_fail = 0.0


class TestAssignmentRow(unittest.TestCase):
    def _message(self, subject, entry_id="E1"):
        m = make_mock_message(subject, 2026, 10, 16, hour=14, minute=5)
        m.EntryID = entry_id
        return m

    def test_oa_row(self):
        cache = {"PA-2T8084|의견제출통지서": {"doc_type": "의견제출통지서", "receipt_date": "2026-10-12",
                                         "due_date": "2027-02-08", "app_no": "10-2025-7000694",
                                         "applicant": "아우디  아게", "received_at": "2026-10-13"}}
        row = main.assignment_row(self._message(OA_SUBJECT), cache)
        self.assertEqual(row, {
            "mail_entry_id": "E1", "ref": "PA-2T8084", "doc_type": "OA", "notice_date": "2026-10-12",
            "report_due": "2026-10-01", "assigner": "한송희", "app_no": "10-2025-7000694", "applicant": "아우디 아게",
            "due_date": "2027-02-08", "reexam": None, "received_at": "2026-10-16 14:05:00"})
        self.assertEqual(set(row), set(db.ASSIGNMENT_COLUMNS))

    def test_fr_row_reexam_and_due_from_receipt(self):
        cache = {"PA-2T9001|거절결정서": {"doc_type": "거절결정서", "receipt_date": "2026-10-06", "reexam": True,
                                       "app_no": "10-2024-0000001", "due_date": ""}}
        row = main.assignment_row(self._message(FR_SUBJECT), cache)
        self.assertEqual((row["doc_type"], row["reexam"], row["due_date"], row["report_due"], row["assigner"]),
                         ("거절결정서", 1, "2027-01-06", "2026-10-05", "이여름"))
        row_no_cache = main.assignment_row(self._message(FR_SUBJECT), {})
        self.assertEqual(row_no_cache["due_date"], "2027-01-16")

    def test_stale_case_and_no_cache(self):
        cache = {"PA-2T8084|의견제출통지서": {"doc_type": "의견제출통지서", "receipt_date": "2026-03-02",
                                         "due_date": "2026-07-02", "app_no": "10-2025-7000694"}}
        row = main.assignment_row(self._message(OA_SUBJECT), cache)
        self.assertEqual((row["notice_date"], row["app_no"], row["due_date"]), (None, None, None))
        row = main.assignment_row(self._message("[위임] PA-2T8084; 한송희(팀); 의견제출통지서; 001"), {})
        self.assertEqual((row["doc_type"], row["report_due"]), ("의견제출통지서", None))

    def test_no_ref(self):
        self.assertIsNone(main.assignment_row(self._message("[위임] 한송희(팀); OA"), {}))


class TestRecordAssignmentsInFlow(unittest.TestCase):
    """일정 등록 길에 끼워 넣은 NAS 기록 — 실패해도 등록은 그대로"""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        for name, value in (("CALENDAR_STATE_FILE", "calendar_state.json"), ("CASE_CACHE_FILE", "case_cache.json"),
                            ("ASSIGNMENT_STATE_FILE", "assignment_state.json")):
            patcher = patch.object(main, name, os.path.join(self.tmp.name, value))
            patcher.start()
            self.addCleanup(patcher.stop)
        patcher = patch.dict(main.config, {"naver_calendars": {"Mail": "Mail", "No_Instructions": "No_Instructions"},
                                           "applicant_abbr": {}})
        patcher.start()
        self.addCleanup(patcher.stop)

    def _delegation(self, entry_id):
        m = make_mock_message("[위임] PA-2T9001; 이여름(팀); 거절결정서 [서 2026-10-05]; 001",
                              date.today().year, date.today().month, date.today().day)
        m.EntryID = entry_id
        return m

    def test_db_down_still_registers(self):
        with patch.object(db, "connect", side_effect=db.DbError("NAS 꺼짐")), \
                patch.object(main.naver_cal, "Session") as session_cls, patch.object(main, "ScheduleDialog"):
            session_cls.return_value.add_all_day.return_value = True
            rows = main.register_delegation(self._delegation("C"), always_show=False)
        self.assertEqual([r[3] for r in rows], [main.STATUS_ADDED, main.STATUS_ADDED])
        state = db.load_state(main.ASSIGNMENT_STATE_FILE)
        self.assertEqual([p["mail_entry_id"] for p in state["pending"]], ["C"])

    def test_scan_records_all_once_and_flushes(self):
        """시작할 때 스캔: 처음 실행(일정 기록 없음)이어도 NAS 에는 쓴다. 쌓인 줄도 함께."""
        conn = FakeConn()
        outlook = MagicMock()
        outlook.GetDefaultFolder.return_value.Items.Restrict.return_value = [self._delegation("A"),
                                                                             self._delegation("B")]
        with patch.object(main, "is_calendar_delegation", return_value=True), \
                patch.object(db, "connect", return_value=conn):
            self.assertEqual(main.scan_delegations(outlook), [])
        self.assertEqual(set(conn.rows), {"A", "B"})
        with patch.object(db, "connect", side_effect=AssertionError("다시 붙으면 안 된다")):
            self.assertIsNone(main.record_assignments([self._delegation("A")]))


    def test_doc_kind_of_delegation(self):
        """위임 서류 종류 — 의견제출통지서·거절결정서만 일정, 등록결정·보정요구서는 「OTHER」(10-24 사용자: 일정 없음)."""
        kind = lambda s: main.rejection_type_of(MagicMock(Subject=s))      # noqa: E731
        self.assertEqual(kind("[위임] PA-2T8084; 한송희(팀); OA [서 2026-10-01]; 001"), "OA")
        self.assertEqual(kind("[위임] PA-2T8084; 한송희(팀); 의견제출통지서 [서 2026-10-01]; 001"), "OA")
        self.assertEqual(kind("[위임] PA-2T8084; 한송희(팀); 최후의견제출통지서; 001"), "OA")
        self.assertEqual(kind("[위임] PA-2T9001; 이여름(팀); 거절결정서 [서 2026-10-05]; 001"), "FR")
        self.assertEqual(kind("[위임] PA-2T9001; 이여름(팀); 특허거절결정서; 001"), "FR")
        self.assertEqual(kind("[위임] PA-2T9002; 이여름(팀); 특허결정서 [서 2026-10-05]; 001"), main.OTHER_DOC)
        self.assertEqual(kind("[위임] PA-T2K6503; 이예진(팀); 보정요구서; 001"), main.OTHER_DOC)
        self.assertEqual(kind("[위임] PA-2T8084; 한송희(팀); 001"), "OA")                 # 서류 이름 없음 — 예전처럼

    def test_other_delegation_records_but_no_schedule(self):
        """등록결정 위임 — 네이버 일정은 만들지 않고(Session 을 부르지 않음) 처리 표시·받은 위임 기록만."""
        m = make_mock_message("[위임] PA-2T9002; 이여름(팀); 특허결정서 [서 2026-10-05]; 001",
                              date.today().year, date.today().month, date.today().day)
        m.EntryID = "D"
        conn = FakeConn()
        with patch.object(db, "connect", return_value=conn), \
                patch.object(main.naver_cal, "Session", side_effect=AssertionError("일정을 만들면 안 된다")), \
                patch.object(main, "ScheduleDialog", side_effect=AssertionError("창을 띄우면 안 된다")):
            self.assertEqual(main.register_delegation(m, always_show=True), [])
        self.assertIn("D", main.load_calendar_state()["done"])
        self.assertEqual(set(conn.rows), {"D"})
        self.assertEqual(conn.rows["D"][db.ASSIGNMENT_COLUMNS.index("doc_type")], "특허결정서")



class AbbrConn(FakeConn):
    """약칭 표 가짜 — SELECT 는 self.table, INSERT … ON DUPLICATE·DELETE 는 self.table 을 고친다."""

    def __init__(self, table=None):
        super().__init__()
        self.table = dict(table or {})

    def cursor(self):
        conn = self

        class _Cur(FakeCursor):
            def execute(self, sql, params=None):
                conn.executed.append((sql, params))
                if sql.startswith("SELECT name, abbr"):
                    self._rows = [{"name": k, "abbr": v} for k, v in conn.table.items()]
                elif sql.startswith("INSERT INTO applicant_abbr"):
                    conn.table[params[0]] = params[1]
                elif sql.startswith("DELETE FROM applicant_abbr"):
                    return 1 if conn.table.pop(params[0], None) is not None else 0
                return 1

            def fetchall(self):
                return getattr(self, "_rows", [])
        return _Cur(self)


class TestApplicantAbbr(unittest.TestCase):
    """출원인 약칭 — NAS applicant_abbr 가 기준(10-24 사용자), settings.json 은 사본."""

    def setUp(self):
        patcher = patch.dict(main.config, {"applicant_abbr": {"가공 주식회사": "가공", "이피씨만": "이피"}})
        patcher.start()
        self.addCleanup(patcher.stop)
        patcher = patch.object(main, "save_config")
        self.save_config = patcher.start()
        self.addCleanup(patcher.stop)

    def test_db_load_and_save(self):
        conn = AbbrConn({"가공 주식회사": "가공"})
        self.assertEqual(db.load_abbr(conn), {"가공 주식회사": "가공"})
        self.assertEqual(db.save_abbr(conn, {"새 회사": "새"}, ["가공 주식회사", "없는 이름"]), (1, 1))
        self.assertEqual(conn.table, {"새 회사": "새"})
        self.assertEqual(conn.commits, 1)
        self.assertIn("ON DUPLICATE KEY UPDATE", conn.executed[1][0])                # [0] 은 앞의 SELECT

    def test_merge_nas_wins_and_uploads_local_only(self):
        merged, upload = main.merge_abbr({"가공  주식회사": "가공", "이피씨만": "이피"}, {"가공 주식회사": "가공사", "엔에이에스": "나스"})
        self.assertEqual(merged, {"가공 주식회사": "가공사", "이피씨만": "이피", "엔에이에스": "나스"})
        self.assertEqual(upload, {"이피씨만": "이피"})

    def test_sync_at_start(self):
        conn = AbbrConn({"가공 주식회사": "가공사", "엔에이에스": "나스"})
        got = main.sync_abbr(connect_fn=lambda: conn)
        self.assertEqual(got, {"가공 주식회사": "가공사", "이피씨만": "이피", "엔에이에스": "나스"})
        self.assertEqual(main.config["applicant_abbr"], got)
        self.assertEqual(conn.table["이피씨만"], "이피")              # 이 PC 에만 있던 것은 NAS 로
        self.save_config.assert_called_once()

    def test_sync_without_nas_keeps_file(self):
        self.assertIsNone(main.sync_abbr(connect_fn=lambda: (_ for _ in ()).throw(db.DbError("NAS 꺼짐"))))
        self.assertEqual(main.config["applicant_abbr"], {"가공 주식회사": "가공", "이피씨만": "이피"})
        self.save_config.assert_not_called()

    def test_settings_save_pushes_changes_and_deletes(self):
        loaded = {"가공 주식회사": "가공", "이피씨만": "이피"}
        saved = {"가공 주식회사": "가공2", "새 회사": "새"}
        changed, removed = main.abbr_changes(loaded, saved)
        self.assertEqual((changed, removed), ({"가공 주식회사": "가공2", "새 회사": "새"}, ["이피씨만"]))
        conn = AbbrConn(loaded)
        self.assertTrue(main.push_abbr(changed, removed, connect_fn=lambda: conn))
        self.assertEqual(conn.table, saved)
        self.assertFalse(main.push_abbr(changed, removed, connect_fn=lambda: (_ for _ in ()).throw(db.DbError("x"))))


if __name__ == "__main__":
    unittest.main(verbosity=2)
