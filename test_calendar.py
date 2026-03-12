"""
캘린더 기일 설정 기능 테스트
테스트 대상:
  - _extract_message_info(): 이메일 제목에서 ref_num, 이름, 날짜 추출
  - OA 기일 계산: target_date + 4개월 - 11일
  - FR 기일 계산: receipt_date(수신일) + 3개월
  - ref_num_key 정규식 패턴
"""

import sys
import unittest
from datetime import date, datetime, timedelta
from unittest.mock import MagicMock

# -----------------------------------------------------------------------
# win32com, PySide6는 실제 Outlook/GUI 없이도 테스트할 수 있도록 mock 처리
# -----------------------------------------------------------------------
sys.modules["win32com"] = MagicMock()
sys.modules["win32com.client"] = MagicMock()
sys.modules["PySide6"] = MagicMock()
sys.modules["PySide6.QtWidgets"] = MagicMock()
sys.modules["PySide6.QtGui"] = MagicMock()

from dateutil.relativedelta import relativedelta  # noqa: E402

import main  # noqa: E402  (mock 설정 후 import)
from main import _extract_message_info, ref_num_key  # noqa: E402


# -----------------------------------------------------------------------
# 헬퍼: Outlook message 객체를 흉내내는 Mock 생성
# -----------------------------------------------------------------------

class MockCreationTime:
    """win32com pywintypes.datetime을 모방하는 클래스."""

    def __init__(self, year, month, day, hour=9, minute=0):
        self._dt = datetime(year, month, day, hour, minute)

    def strftime(self, fmt):
        return self._dt.strftime(fmt)

    def date(self):
        return self._dt.date()

    @property
    def year(self):   return self._dt.year
    @property
    def month(self):  return self._dt.month
    @property
    def day(self):    return self._dt.day
    @property
    def hour(self):   return self._dt.hour
    @property
    def minute(self): return self._dt.minute


def make_mock_message(subject, creation_year=2025, creation_month=1, creation_day=15,
                      hour=9, minute=30):
    """테스트용 Outlook 메시지 Mock 객체를 반환합니다."""
    message = MagicMock()
    message.Subject = subject
    message.CreationTime = MockCreationTime(
        creation_year, creation_month, creation_day, hour, minute
    )
    message.SenderName = "DoNotReply"
    return message


# -----------------------------------------------------------------------
# 날짜 계산 로직 (add_calendar_event 내부와 동일)
# -----------------------------------------------------------------------

def calc_response_date(rejection_type, target_date, receipt_date):
    """기일을 계산해 반환합니다.

    Args:
        rejection_type: "OA" 또는 "FR"
        target_date: 이메일 제목에서 추출한 날짜 (date 객체)
        receipt_date: 이메일 수신일 (date 객체)
    Returns:
        response_date (date 객체)
    """
    four_months_later = target_date + relativedelta(months=4)
    if rejection_type == "OA":
        return four_months_later - timedelta(days=11)
    else:  # FR
        return receipt_date + relativedelta(months=3)


# -----------------------------------------------------------------------
# 테스트 케이스
# -----------------------------------------------------------------------

class TestExtractMessageInfo(unittest.TestCase):
    """_extract_message_info() 함수 테스트"""

    def test_OA_subject(self):
        """OA 이메일 제목에서 날짜·시간·ref_num·이름 정상 추출"""
        subject = "[위임] PJA-2R8603; 한송희(고객지원팀-책임); OA [서 2024-12-23]; 2412130993"
        message = make_mock_message(subject, 2025, 1, 15, hour=9, minute=30)

        delivery_date, time_id, ref_num, name = _extract_message_info(message)

        self.assertEqual(delivery_date, "2025-01-15")
        self.assertEqual(time_id, "09hr30mn")
        self.assertEqual(ref_num, "PJA-2R8603")
        self.assertEqual(name, "한송희")

    def test_FR_subject(self):
        """FR(거절결정서) 이메일 제목에서 정상 추출"""
        subject = "[위임] ABC-T2X1234; 이여름(팀-담당); 거절결정서 [서 2025-03-10]; 2503100001"
        message = make_mock_message(subject, 2025, 3, 12, hour=14, minute=5)

        delivery_date, time_id, ref_num, name = _extract_message_info(message)

        self.assertEqual(delivery_date, "2025-03-12")
        self.assertEqual(time_id, "14hr05mn")
        self.assertEqual(ref_num, "ABC-T2X1234")
        self.assertEqual(name, "이여름")

    def test_missing_ref_num_raises(self):
        """Ref 번호가 없는 제목 → ValueError 발생"""
        subject = "[위임] 한송희(고객지원팀); OA; 2412130993"
        message = make_mock_message(subject)

        with self.assertRaises(ValueError):
            _extract_message_info(message)

    def test_missing_name_raises(self):
        """이름이 없는 제목 → ValueError 발생"""
        subject = "[위임] PJA-2R8603; OA [서 2024-12-23]; 2412130993"
        message = make_mock_message(subject)

        with self.assertRaises(ValueError):
            _extract_message_info(message)

    def test_time_format_zero_padding(self):
        """시간 ID가 0으로 채워지는지 확인 (08hr05mn 등)"""
        subject = "[위임] PJA-2R8603; 한송희(팀); OA [서 2024-12-23]; 001"
        message = make_mock_message(subject, hour=8, minute=5)

        _, time_id, _, _ = _extract_message_info(message)

        self.assertEqual(time_id, "08hr05mn")


class TestResponseDateCalculation(unittest.TestCase):
    """기일 계산 로직 테스트"""

    def test_OA_basic(self):
        """OA: target_date + 4개월 - 11일 기본 케이스"""
        target_date  = date(2024, 12, 23)
        receipt_date = date(2025, 1, 15)

        result = calc_response_date("OA", target_date, receipt_date)

        # 2024-12-23 + 4개월 = 2025-04-23, - 11일 = 2025-04-12
        self.assertEqual(result, date(2025, 4, 12))

    def test_FR_basic(self):
        """FR: receipt_date + 3개월 기본 케이스"""
        target_date  = date(2025, 3, 10)
        receipt_date = date(2025, 3, 12)

        result = calc_response_date("FR", target_date, receipt_date)

        # 2025-03-12 + 3개월 = 2025-06-12
        self.assertEqual(result, date(2025, 6, 12))

    def test_OA_month_end(self):
        """OA: 월말(10/31) 처리 — relativedelta가 말일을 자동 조정"""
        target_date  = date(2025, 10, 31)
        receipt_date = date(2025, 11, 1)

        result = calc_response_date("OA", target_date, receipt_date)

        # 2025-10-31 + 4개월 = 2026-02-28(2월 말일), - 11일 = 2026-02-17
        self.assertEqual(result, date(2026, 2, 17))

    def test_FR_month_end(self):
        """FR: 월말(11/30) 처리 — relativedelta가 말일을 자동 조정"""
        target_date  = date(2025, 1, 1)
        receipt_date = date(2025, 11, 30)

        result = calc_response_date("FR", target_date, receipt_date)

        # 2025-11-30 + 3개월 = 2026-02-28(2월 말일)
        self.assertEqual(result, date(2026, 2, 28))

    def test_OA_year_boundary(self):
        """OA: 연말 날짜가 다음 해로 넘어가는 케이스"""
        target_date  = date(2025, 11, 1)
        receipt_date = date(2025, 11, 5)

        result = calc_response_date("OA", target_date, receipt_date)

        # 2025-11-01 + 4개월 = 2026-03-01, - 11일 = 2026-02-18
        self.assertEqual(result, date(2026, 2, 18))

    def test_FR_year_boundary(self):
        """FR: 연말 수신일이 다음 해로 넘어가는 케이스"""
        target_date  = date(2025, 1, 1)
        receipt_date = date(2025, 12, 15)

        result = calc_response_date("FR", target_date, receipt_date)

        # 2025-12-15 + 3개월 = 2026-03-15
        self.assertEqual(result, date(2026, 3, 15))


class TestRefNumRegex(unittest.TestCase):
    """ref_num_key 정규식 패턴 테스트"""

    def test_pattern_first_type(self):
        """첫 번째 패턴 (예: PJA-2R8603) 매칭"""
        cases = [
            ("[위임] PJA-2R8603; ...", "PJA-2R8603"),
            ("[위임] AB-2X1234; ...",  "AB-2X1234"),
            ("[위임] ABC-T2X1234AB; ...", "ABC-T2X1234AB"),
        ]
        for subject, expected in cases:
            with self.subTest(subject=subject):
                result = ref_num_key.findall(subject)
                self.assertTrue(len(result) > 0, f"패턴 미검출: {subject}")
                self.assertEqual(result[0], expected)

    def test_pattern_second_type(self):
        """두 번째 패턴 (예: KR-12345) 매칭"""
        cases = [
            ("[위임] KR-12345; ...", "KR-12345"),
            ("[위임] AB-99999; ...", "AB-99999"),
        ]
        for subject, expected in cases:
            with self.subTest(subject=subject):
                result = ref_num_key.findall(subject)
                self.assertTrue(len(result) > 0)
                self.assertEqual(result[0], expected)

    def test_no_match(self):
        """Ref 번호가 없는 제목 → 빈 리스트 반환"""
        subject = "[위임] 한송희(팀); OA 안내"
        result = ref_num_key.findall(subject)
        self.assertEqual(result, [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
