"""KIPRIS 호출 셈·속도 문턱·막힘 표시 — ChekPat 과 공유하는 규약(kipris_ledger·gate·BLOCKED·발신자 표).
실제 KIPRIS·NAS 에는 닿지 않는다(가짜 연결·가짜 시계)."""

import json
import os
import tempfile
import unittest
import urllib.error
from datetime import date
from pathlib import Path
from unittest.mock import patch

import test_calendar  # noqa: F401  (%APPDATA% 를 빈 임시 폴더로, 접속 환경변수는 지운다)

import db  # noqa: E402
import kipris  # noqa: E402

KEY = "same-key"


class _NoTable(Exception):
    def __init__(self):
        super().__init__(1146, "Table 'iplog.kipris_usage_src' doesn't exist")


class _Nas:
    """메모리 NAS — kipris_usage·kipris_usage_src. `has_src=False` 면 예전 NAS(1146), `lose_ack` 면 커밋은 되고 응답만 잃는다."""

    def __init__(self):
        self.usage, self.src = {}, {}
        self.has_src, self.down, self.lose_ack = True, False, False
        self.connects = 0


class _Cur:
    def __init__(self, nas):
        self.nas, self.rowcount, self._row = nas, 0, None
        self.log = []

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def execute(self, sql, a=()):
        n = self.nas
        if sql.startswith("INSERT IGNORE INTO kipris_usage_src"):
            if not n.has_src:
                raise _NoTable()
            n.src.setdefault((a[0], a[1], a[2]), a[3])
        elif sql.startswith("INSERT IGNORE INTO kipris_usage "):
            n.usage.setdefault((a[0], a[1]), 0)
        elif "FROM kipris_usage_src" in sql:
            self._row = {"calls": n.src[(a[0], a[1], a[2])]}
        elif sql.startswith("SELECT calls FROM kipris_usage"):
            self._row = {"calls": n.usage[(a[0], a[1])]}
        elif sql.startswith("UPDATE kipris_usage_src"):
            k = (a[1], a[2], a[3])
            n.src[k] = max(n.src[k], a[0])
        elif "GREATEST" in sql:
            k = (a[1], a[2])
            n.usage[k] = max(n.usage[k], a[0])
        elif sql.startswith("UPDATE kipris_usage") and "calls < %s" in sql:
            k = (a[0], a[1])
            ok = n.usage[k] < a[2]
            if ok:
                n.usage[k] += 1
            self.rowcount = 1 if ok else 0
        elif sql.startswith("UPDATE kipris_usage"):
            n.usage[(a[1], a[2])] += a[0]
        else:
            raise AssertionError(sql)

    def fetchone(self):
        return self._row


class _Conn:
    def __init__(self, nas):
        self.nas = nas

    def cursor(self):
        return _Cur(self.nas)

    def commit(self):
        if self.nas.lose_ack:
            raise OSError("연결이 끊겼습니다")

    def close(self):
        pass


class _Base(unittest.TestCase):
    def setUp(self):
        self._dir = tempfile.TemporaryDirectory()
        self.appdata = Path(self._dir.name)
        self.guard = self.appdata / "guard"
        self._env = patch.dict(os.environ, {"APPDATA": self._dir.name, "KIPRIS_API_KEY": KEY})
        self._env.start()
        self._home = patch.object(kipris, "_guard_dir", lambda: self.guard)
        self._home.start()
        self.nas = _Nas()
        db._last_fail = 0.0
        kipris._known_total.clear()

        def connect(force=False):
            self.nas.connects += 1
            if self.nas.down:
                raise db.DbError("끊김")
            return _Conn(self.nas)
        self._conn = patch.object(db, "connect", connect)
        self._conn.start()

    def tearDown(self):
        self._conn.stop()
        self._home.stop()
        self._env.stop()
        self._dir.cleanup()

    @property
    def ym(self):
        return date.today().strftime("%Y-%m")

    def ledger(self):
        with open(self.appdata / "ChekPat" / "kipris_ledger.json", encoding="utf-8") as f:
            return json.load(f)

    def entry(self, key=KEY, ym=None):
        return self.ledger()["entries"][f"{kipris.key_id(key)}|{ym or self.ym}"]

    def nas_total(self, key=KEY, ym=None):
        return self.nas.usage[(ym or self.ym, kipris.key_id(key))]

    def put_ledger(self, entries, src="0123456789abcdef"):
        d = self.appdata / "ChekPat"
        d.mkdir(parents=True, exist_ok=True)
        (d / "kipris_ledger.json").write_text(json.dumps({"version": 1, "src": src, "entries": entries}), encoding="utf-8")


class TestSharedCount(_Base):
    def test_counts_in_nas_and_ledger_with_hashed_key(self):
        kipris._count_call()
        kipris._count_call()
        self.assertEqual(self.nas_total(), 2)
        self.assertNotIn(KEY, "".join(k[1] for k in self.nas.usage))
        self.assertEqual(self.entry(), {"calls": 2, "reported": 2, "full": False})
        self.assertRegex(self.ledger()["src"], r"^[0-9a-f]{16}$")
        self.assertEqual(list(self.nas.src.values()), [2])            # 발신자 줄은 누적값

    def test_other_pcs_calls_count_toward_the_cap(self):
        with patch.object(kipris, "MONTHLY_CAP", 3):
            self.nas.usage[(self.ym, kipris.key_id(KEY))] = 3          # ChekPat(맥·윈도)이 센 것
            with self.assertRaises(kipris.KiprisError):
                kipris._count_call()
        self.assertEqual(self.entry()["calls"], 0)                      # 보내지 않았으니 세지도 않는다

    def test_nas_down_continues_locally_then_reports(self):
        self.nas.down = True
        kipris._count_call()
        kipris._count_call()
        self.assertEqual(self.nas.usage, {})
        self.assertEqual(self.entry(), {"calls": 2, "reported": 0, "full": False})
        self.nas.down = False
        db._last_fail = 0.0
        kipris._count_call()
        self.assertEqual(self.nas_total(), 3)                          # 못 올린 2회 + 이번 1회
        self.assertEqual(self.entry()["reported"], 3)

    def test_nas_down_still_stops_at_the_cap(self):
        self.nas.down = True
        with patch.object(kipris, "MONTHLY_CAP", 2):
            kipris._count_call()
            kipris._count_call()
            with self.assertRaises(kipris.KiprisError):
                kipris._count_call()

    def test_pending_is_settled_then_new_call_checked_against_cap(self):
        """x1 4 — NAS 1 + 미보고 2 = 3 이 이미 상한 3: 미보고분은 올리되 새 호출은 막는다."""
        self.nas.usage[(self.ym, kipris.key_id(KEY))] = 1
        self.put_ledger({f"{kipris.key_id(KEY)}|{self.ym}": {"calls": 2, "reported": 0, "full": False}})
        with patch.object(kipris, "MONTHLY_CAP", 3):
            with self.assertRaises(kipris.KiprisError):
                kipris._count_call()
        self.assertEqual(self.nas_total(), 3)
        self.assertEqual(self.entry(), {"calls": 2, "reported": 2, "full": False})

    def test_pending_below_cap_still_grants_the_new_call(self):
        self.nas.usage[(self.ym, kipris.key_id(KEY))] = 1
        self.put_ledger({f"{kipris.key_id(KEY)}|{self.ym}": {"calls": 2, "reported": 0, "full": False}})
        kipris._count_call()
        self.assertEqual(self.nas_total(), 4)
        self.assertEqual(self.entry(), {"calls": 3, "reported": 3, "full": False})

    def test_full_flag_blocks_and_raises_nas_to_cap(self):
        self.put_ledger({f"{kipris.key_id(KEY)}|{self.ym}": {"calls": 5, "reported": 5, "full": True}})
        with self.assertRaises(kipris.KiprisError):
            kipris._count_call()
        self.assertEqual(self.nas_total(), kipris.MONTHLY_CAP)

    def test_lost_commit_ack_is_not_counted_twice(self):
        """x1 5 — 커밋은 됐는데 응답을 잃으면 오프라인으로 넘어가도, 다음 정산이 같은 호출을 다시 더하지 않는다."""
        self.nas.lose_ack = True
        kipris._count_call()                                           # NAS 에는 1 이 들어갔고 이 PC 는 모른다
        self.assertEqual(self.nas_total(), 1)
        self.assertEqual(self.entry(), {"calls": 1, "reported": 0, "full": False})
        self.nas.lose_ack = False
        db._last_fail = 0.0
        kipris._count_call()
        self.assertEqual(self.nas_total(), 2)                          # 실제 호출 2회

    def test_old_nas_without_sender_table_falls_back(self):
        self.nas.has_src = False
        kipris._count_call()
        kipris._count_call()
        self.assertEqual(self.nas_total(), 2)
        self.assertEqual(self.entry()["reported"], 2)
        with patch.object(kipris, "MONTHLY_CAP", 2):
            with self.assertRaises(kipris.KiprisError):
                kipris._count_call()

    def test_two_keys_are_kept_apart(self):
        """x1 9 — 키 A 의 미보고분이 키 B 에 올라가지 않는다."""
        self.nas.down = True
        kipris._count_call()
        kipris._count_call()
        self.nas.down = False
        db._last_fail = 0.0
        with patch.dict(os.environ, {"KIPRIS_API_KEY": "other-key"}):
            kipris._count_call()
        self.assertEqual(self.nas_total("other-key"), 1)
        self.assertEqual(self.nas_total(KEY), 2)                        # A 의 2회는 A 로 정산됐다

    def test_last_months_pending_goes_to_that_month(self):
        """x1 10 — 지난 달의 미보고분은 이번 달에 닿아도 지난 달 셈으로."""
        kid = kipris.key_id(KEY)
        self.put_ledger({f"{kid}|2026-10": {"calls": 2, "reported": 0, "full": False}})
        kipris._count_call(date(2026, 11, 1))
        self.assertEqual(self.nas.usage[("2026-10", kid)], 2)
        self.assertEqual(self.nas.usage[("2026-11", kid)], 1)
        self.assertEqual(self.ledger()["entries"][f"{kid}|2026-10"]["reported"], 2)

    def test_old_unreported_entries_are_not_pruned(self):
        kid = kipris.key_id(KEY)
        self.nas.down = True
        self.put_ledger({f"{kid}|2026-01": {"calls": 2, "reported": 0, "full": False},
                         f"{kid}|2026-02": {"calls": 2, "reported": 2, "full": False}})
        kipris._count_call(date(2026, 11, 1))
        self.assertIn(f"{kid}|2026-01", self.ledger()["entries"])
        self.nas.down = False
        db._last_fail = 0.0
        kipris._count_call(date(2026, 11, 1))
        self.assertEqual(self.nas.usage[("2026-01", kid)], 2)
        self.assertNotIn(f"{kid}|2026-01", self.ledger()["entries"])    # 정산이 끝나면 지워진다
        self.assertNotIn(f"{kid}|2026-02", self.ledger()["entries"])

    def test_known_total_never_shrinks(self):
        """x1 6 — 늦게 끝난 옛 값이 기억을 낮추지 않는다."""
        kid = kipris.key_id(KEY)
        kipris._remember(kid, self.ym, 899)
        kipris._remember(kid, self.ym, 898)
        self.assertEqual(kipris._known_total[(kid, self.ym)], 899)

    def test_first_ledger_migrates_settings_json(self):
        d = self.appdata / "ChekPat"
        d.mkdir(parents=True)
        (d / "settings.json").write_text(json.dumps({"kipris_calls": {self.ym: 5}, "kipris_reported": {self.ym: 9}}),
                                         encoding="utf-8")
        self.nas.down = True
        kipris._count_call()
        self.assertEqual(self.entry(), {"calls": 6, "reported": 5, "full": False})   # reported = min(reported, calls)
        self.assertEqual(json.loads((d / "settings.json").read_text(encoding="utf-8"))["kipris_calls"], {self.ym: 5})

    def test_two_apps_share_one_ledger_without_double_reporting(self):
        """x1 1 — 같은 PC 의 다른 앱이 이미 올린 미보고분(reported 가 올라 있음)을 다시 올리지 않는다."""
        self.nas.down = True
        kipris._count_call()                                            # 앱 A: 오프라인 1
        self.nas.down = False
        db._last_fail = 0.0
        kipris._count_call()                                            # 앱 B: 정산 + 1
        kipris._count_call()                                            # 앱 A: 이미 정산됨
        self.assertEqual(self.nas_total(), 3)

    def test_corrupt_ledger_restarts_without_remigrating(self):
        """ChekPat 과 같다 — 못 읽는 셈 파일은 새 발신자 번호로 새로 시작하고 옛 settings.json 을 다시 옮기지 않는다."""
        d = self.appdata / "ChekPat"
        d.mkdir(parents=True)
        (d / "settings.json").write_text(json.dumps({"kipris_calls": {self.ym: 5}}), encoding="utf-8")
        (d / "kipris_ledger.json").write_text("{깨진", encoding="utf-8")
        kipris._count_call()
        self.assertEqual(self.entry(), {"calls": 1, "reported": 1, "full": False})
        self.assertRegex(self.ledger()["src"], r"^[0-9a-f]{16}$")

    def test_offline_exit_still_prunes_settled_old_entries(self):
        kid = kipris.key_id(KEY)
        self.nas.down = True
        self.put_ledger({f"{kid}|2026-01": {"calls": 2, "reported": 2, "full": False}})
        kipris._count_call(date(2026, 11, 1))
        self.assertNotIn(f"{kid}|2026-01", self.ledger()["entries"])

    def test_file_lock_is_exclusive(self):
        path = self.appdata / "x.lock"
        with kipris._locked_file(path):
            with open(path, "a+b") as other:
                self.assertFalse(kipris._try_flock(other))
        with open(path, "a+b") as other:
            self.assertTrue(kipris._try_flock(other))

    def test_lock_timeout_does_not_send(self):
        path = self.appdata / "y.lock"
        with open(path, "a+b") as held:
            self.assertTrue(kipris._try_flock(held))
            with self.assertRaises(kipris.KiprisError):
                with kipris._locked_file(path, timeout=0.1):
                    pass


class TestGate(_Base):
    def reserve(self, now):
        return kipris._gate_reserve(lambda: now)

    def test_spaced_one_fortieth_and_queue_not_cut(self):
        ats = [self.reserve(1000.0) for _ in range(300)]
        self.assertAlmostEqual(ats[1] - ats[0], 1 / 40, places=5)
        self.assertAlmostEqual(ats[-1] - ats[0], 299 / 40, places=3)    # 5초 절단 없음 — 모두 다른 시각
        self.assertEqual(len(set(ats)), 300)

    def test_clock_rewind_drops_old_reservations(self):
        for _ in range(100):
            self.reserve(1000.0)                                        # 예약이 1002.5 까지
        at = self.reserve(500.0)                                        # 시계가 되감겼다
        self.assertLess(at, 501.0)

    def test_legacy_single_value_is_clipped_to_five_seconds(self):
        (self.guard).mkdir(parents=True)
        (self.guard / "gate").write_bytes(b"99999.0")
        self.assertLessEqual(self.reserve(1000.0), 1005.1)

    def test_wait_sleeps_again_if_woken_early(self):
        t = [1000.0]
        slept = []

        def sleep(s):
            slept.append(s)
            t[0] += s / 2 if len(slept) == 1 else s                     # 처음에만 일찍 깬다
        kipris._gate_reserve(lambda: t[0])
        kipris._gate_wait(clock=lambda: t[0], sleep=sleep)
        self.assertGreater(len(slept), 1)
        self.assertGreaterEqual(t[0], 1000.0 + 1 / 40)

    def test_acquire_goes_through_the_shared_gate(self):
        with patch.object(kipris, "_gate_wait") as gw:
            kipris._acquire()
        gw.assert_called_once()


class _Res:
    def __init__(self, text):
        self.text = text

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def read(self):
        return self.text.encode("utf-8")


class TestBlocked(_Base):
    def setUp(self):
        super().setUp()
        self.sent = []
        self.reply = "<response><header><resultCode>00</resultCode></header></response>"

        def opener(url, timeout=0):
            self.sent.append(url)
            return _Res(self.reply)
        self._op = patch.object(kipris._opener, "open", opener)
        self._op.start()
        self._gate = patch.object(kipris, "_gate_wait")
        self._gate.start()

    def tearDown(self):
        self._gate.stop()
        self._op.stop()
        super().tearDown()

    def test_blocked_file_stops_sending_and_counting(self):
        self.guard.mkdir(parents=True)
        (self.guard / "BLOCKED").write_text("2026-11-01 10:00:00\n접근 거부\n", encoding="utf-8")
        with self.assertRaises(kipris.KiprisError) as cm:
            kipris._send("x", {})
        self.assertIn("접근 거부", str(cm.exception))
        self.assertEqual(self.sent, [])
        self.assertEqual(self.nas.usage, {})

    def test_blocked_appearing_while_waiting_stops_the_send(self):
        def gate():
            self.guard.mkdir(parents=True, exist_ok=True)
            (self.guard / "BLOCKED").write_text("2026-11-01 10:00:00\n접근 거부\n", encoding="utf-8")
        with patch.object(kipris, "_gate_wait", gate):
            with self.assertRaises(kipris.KiprisError):
                kipris._send("x", {})
        self.assertEqual(self.sent, [])

    def test_result_code_20_leaves_the_block_marker(self):
        self.reply = "<response><header><resultCode>20</resultCode><resultMsg>접근 거부</resultMsg></header></response>"
        text = kipris._send("path", {})
        stamp, why = (self.guard / "BLOCKED").read_text(encoding="utf-8").strip().split("\n", 1)
        self.assertRegex(stamp, r"^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d$")
        self.assertIn("결과 코드 20", why)
        with self.assertRaises(kipris.KiprisError):
            kipris.parse_root(text)
        with self.assertRaises(kipris.KiprisError):                     # 다음 호출은 보내지 않는다
            kipris._send("path", {})
        self.assertEqual(len(self.sent), 1)

    def test_result_code_30_is_not_a_block(self):
        """규약 6절 — 키·상품 미등록은 막힘이 아니다: 표시를 남기지 않고, 다음 조회도 보낸다."""
        self.reply = "<response><header><resultCode>30</resultCode><resultMsg>SERVICE KEY IS NOT REGISTERED ERROR</resultMsg></header></response>"
        text = kipris._send("path", {})
        self.assertFalse((self.guard / "BLOCKED").exists())
        self.assertEqual(list(self.guard.glob("KEYERR_*")) if self.guard.exists() else [], [])
        with self.assertRaises(kipris.KiprisError):
            kipris.parse_root(text)
        kipris._send("path", {})
        self.assertEqual(len(self.sent), 2)

    def test_old_block_marker_for_code_30_is_ignored(self):
        self.guard.mkdir(parents=True)
        for why in ("KIPRIS Plus 키가 맞지 않거나 이 상품을 신청하지 않았습니다 (결과 코드 30: 등록된 서비스키를 입력해 주십시오"
                    "(Access key & Service key is not registerd error)).", "SERVICE_KEY_IS_NOT_REGISTERED_ERROR"):
            (self.guard / "BLOCKED").write_text("2026-11-01 06:54:06\n%s\n" % why, encoding="utf-8")
            kipris._send("path", {})
        self.assertTrue((self.guard / "BLOCKED").exists())              # 지우지 않는다
        self.assertEqual(len(self.sent), 2)

    def test_other_codes_do_not_block(self):
        self.reply = "<response><header><resultCode>00</resultCode></header></response>"
        kipris._send("path", {})
        self.assertFalse((self.guard / "BLOCKED").exists())

    def test_retry_counts_each_attempt(self):
        calls = []

        def opener(url, timeout=0):
            calls.append(1)
            if len(calls) == 1:
                raise urllib.error.URLError("reset")
            return _Res(self.reply)
        with patch.object(kipris._opener, "open", opener), patch.object(kipris.time, "sleep"):
            kipris._send("path", {})
        self.assertEqual(self.nas_total(), 2)


if __name__ == "__main__":
    unittest.main()
