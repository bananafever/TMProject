"""History 탭 — app.log 를 날짜별 HTML 로 바꾸는 history_html."""

import unittest

# test_calendar 가 (없으면) win32com·PySide6 를 가짜로 끼우고 임시 데이터 폴더를 정한다 — main 보다 먼저 들인다
import test_calendar  # noqa: F401

import main  # noqa: E402

LOG = (
    "2026-10-24 17:07:48,921 - INFO - 연장 확인 2026-10-24: 기연 의뢰한 차수 없음\n"
    "2026-10-25 09:56:50,450 - ERROR - 첨부파일 저장 실패 <a&b>\n"
    "Traceback (most recent call last):\n"
    "  File \"main.py\", line 1\n"
    "2026-10-25 20:53:13,190 - WARNING - 출원인 약칭을 NAS 에 적지 못했습니다\n"
)


class HistoryHtmlTest(unittest.TestCase):
    def test_groups_by_day_with_weekday(self):
        html = main.history_html(LOG)
        self.assertIn("2026-10-24 (토)", html)
        self.assertIn("2026-10-25 (일)", html)
        self.assertEqual(html.count("<table"), 2)

    def test_time_without_seconds_and_level_tags(self):
        html = main.history_html(LOG)
        self.assertIn(">17:07<", html)
        self.assertNotIn("17:07:48", html)
        self.assertIn("오류", html)
        self.assertIn("경고", html)
        self.assertNotIn(" - INFO - ", html)

    def test_escapes_and_attaches_traceback(self):
        html = main.history_html(LOG)
        self.assertIn("&lt;a&amp;b&gt;", html)
        self.assertIn("Traceback (most recent call last):", html)
        self.assertLess(html.index("첨부파일 저장 실패"), html.index("Traceback"))

    def test_empty_log(self):
        self.assertIn("기록이 없습니다", main.history_html(""))


if __name__ == "__main__":
    unittest.main()
