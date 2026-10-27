"""KIPRIS 호출 수 NAS 공유 셈 — ChekPat `kipris_usage` 와 같은 표·같은 key_id. 실제 NAS 에는 닿지 않는다."""

import os
import tempfile
import unittest
from unittest.mock import patch

import test_calendar  # noqa: F401  (%APPDATA% 를 빈 임시 폴더로, 접속 환경변수는 지운다)

import db  # noqa: E402
import kipris  # noqa: E402


class _Cur:
    def __init__(self, rows):
        self.rows, self.rowcount, self._row = rows, 0, None

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def execute(self, sql, args=()):
        if sql.startswith("INSERT IGNORE"):
            self.rows.setdefault((args[0], args[1]), 0)
        elif sql.startswith("UPDATE"):
            delta, ym, kid, cap = args
            ok = self.rows[(ym, kid)] < cap
            if ok:
                self.rows[(ym, kid)] += delta
            self.rowcount = 1 if ok else 0
        else:
            ym, kid = args
            self._row = {"calls": self.rows[(ym, kid)]}

    def fetchone(self):
        return self._row


class _Conn:
    def __init__(self, rows):
        self.rows = rows

    def cursor(self):
        return _Cur(self.rows)

    def commit(self):
        pass

    def close(self):
        pass


class TestSharedCount(unittest.TestCase):
    def setUp(self):
        self._dir = tempfile.TemporaryDirectory()
        self._env = patch.dict(os.environ, {"APPDATA": self._dir.name, "KIPRIS_API_KEY": "same-key"})
        self._env.start()
        self.rows = {}
        self.down = False
        db._last_fail = 0.0
        kipris._known_total.clear()

        def connect(force=False):
            if self.down:
                raise db.DbError("끊김")
            return _Conn(self.rows)
        self._conn = patch.object(db, "connect", connect)
        self._conn.start()

    def tearDown(self):
        self._conn.stop()
        self._env.stop()
        self._dir.cleanup()

    def _settings(self):
        import json
        with open(os.path.join(self._dir.name, "ChekPat", "settings.json"), encoding="utf-8") as f:
            return json.load(f)

    def test_counts_in_nas_with_hashed_key(self):
        kipris._count_call()
        kipris._count_call()
        (ym, kid), n = next(iter(self.rows.items()))
        self.assertEqual(n, 2)
        self.assertEqual(kid, kipris.key_id("same-key"))
        self.assertNotIn("same-key", kid)
        self.assertEqual(self._settings()["kipris_calls"], {ym: 2})
        self.assertEqual(self._settings()["kipris_reported"], {ym: 2})

    def test_other_pcs_calls_count_toward_the_cap(self):
        with patch.object(kipris, "MONTHLY_CAP", 3):
            from datetime import date
            ym = date.today().strftime("%Y-%m")
            self.rows[(ym, kipris.key_id("same-key"))] = 3           # ChekPat(맥·윈도)이 센 것
            with self.assertRaises(kipris.KiprisError):
                kipris._count_call()

    def test_nas_down_continues_locally_then_reports(self):
        self.down = True
        kipris._count_call()
        kipris._count_call()
        self.assertEqual(self.rows, {})
        self.assertEqual(list(self._settings()["kipris_calls"].values()), [2])
        self.assertNotIn("kipris_reported", self._settings())
        self.down = False
        kipris._count_call()
        self.assertEqual(list(self.rows.values()), [3])               # 못 올린 2회 + 이번 1회

    def test_nas_down_still_stops_at_the_cap(self):
        self.down = True
        with patch.object(kipris, "MONTHLY_CAP", 2):
            kipris._count_call()
            kipris._count_call()
            with self.assertRaises(kipris.KiprisError):
                kipris._count_call()


if __name__ == "__main__":
    unittest.main()
