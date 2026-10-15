"""
캘린더 기일 설정 기능 테스트
테스트 대상:
  - _extract_message_info(): 이메일 제목에서 ref_num, 이름, 날짜 추출
  - OA 기일 계산: target_date + 4개월 - 11일
  - FR 기일 계산: receipt_date(수신일) + 3개월
  - ref_num_key 정규식 패턴
"""

import os
import sys
import unittest
from datetime import date, datetime, timedelta
from unittest.mock import MagicMock, PropertyMock, patch

# -----------------------------------------------------------------------
# win32com, PySide6는 실제 Outlook/GUI 없이도 테스트할 수 있도록 mock 처리
# -----------------------------------------------------------------------
sys.modules["win32com"] = MagicMock()
sys.modules["win32com.client"] = MagicMock()
sys.modules["PySide6"] = MagicMock()
sys.modules["PySide6.QtWidgets"] = MagicMock()
sys.modules["PySide6.QtGui"] = MagicMock()
sys.modules["PySide6.QtCore"] = MagicMock()

from dateutil.relativedelta import relativedelta  # noqa: E402

import kipris  # noqa: E402
import naver_cal  # noqa: E402
import naver_sync  # noqa: E402
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
# 날짜 계산 로직 (add_calendar_event 내부와 동일 — KIPRIS 값이 없을 때의 계산식)
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


class TestScheduleEditDialogLogic(unittest.TestCase):
    """ScheduleDialog 날짜 수정 흐름 단위 테스트"""

    def test_edited_dates_flow(self):
        """기본 계산된 날짜와 사용자가 다이얼로그에서 변경한 날짜의 흐름 검증"""
        target_date = date(2024, 12, 23)
        receipt_date = date(2025, 1, 15)

        default_response = calc_response_date("OA", target_date, receipt_date)
        self.assertEqual(default_response, date(2025, 4, 12))

        # 사용자가 알림창에서 날짜를 2025-04-20으로 수정한 케이스 시뮬레이션
        user_modified_response_date = date(2025, 4, 20)
        self.assertNotEqual(default_response, user_modified_response_date)
        self.assertEqual(user_modified_response_date.strftime("%Y-%m-%d"), "2025-04-20")


class TestSenderFilter(unittest.TestCase):
    """SenderName 직접 접근 없이 보낸 사람을 판정하는 로직 테스트 (Outlook 보안 경고 회피)"""

    def test_single_sender(self):
        self.assertEqual(
            main.build_sender_filter(["DoNotReply"]),
            "@SQL=(\"urn:schemas:httpmail:fromname\" = 'DoNotReply')",
        )

    def test_multiple_senders_with_brackets(self):
        f = main.build_sender_filter(["DoNotReply", "Yong-Sok SHIN [신용석]"])
        self.assertEqual(
            f,
            "@SQL=(\"urn:schemas:httpmail:fromname\" = 'DoNotReply'"
            " OR \"urn:schemas:httpmail:fromname\" = 'Yong-Sok SHIN [신용석]')",
        )

    def test_subject_and_quote_escape(self):
        f = main.build_sender_filter(["O'Brien"], subject="[위임] it's")
        self.assertIn("= 'O''Brien')", f)
        self.assertTrue(
            f.endswith(" AND \"urn:schemas:httpmail:subject\" = '[위임] it''s'")
        )

    def test_empty_senders_raises(self):
        with self.assertRaises(ValueError):
            main.build_sender_filter([])

    @staticmethod
    def _make_message(entry_id, restricted_ids):
        message = MagicMock()
        message.EntryID = entry_id
        message.Subject = "[위임] PJA-2R8603; 한송희(팀); OA [서 2024-12-23]; 001"
        items = []
        for rid in restricted_ids:
            item = MagicMock()
            item.EntryID = rid
            items.append(item)
        message.Parent.Items.Restrict.return_value = items
        # SenderName에 접근하면 실패하도록 설정 (보안 경고 유발 속성)
        type(message).SenderName = PropertyMock(side_effect=AssertionError("SenderName 접근"))
        return message

    def test_sender_matches_found(self):
        message = self._make_message("ID1", ["ID0", "ID1"])
        self.assertTrue(main.sender_matches(message, ["DoNotReply"]))
        used_filter = message.Parent.Items.Restrict.call_args[0][0]
        self.assertEqual(used_filter, main.build_sender_filter(["DoNotReply"], message.Subject))

    def test_sender_matches_not_found(self):
        message = self._make_message("ID9", ["ID0", "ID1"])
        self.assertFalse(main.sender_matches(message, ["DoNotReply"]))

    def test_sender_matches_fallback_on_restrict_error(self):
        message = MagicMock()
        message.Subject = "[위임] test"
        message.Parent.Items.Restrict.side_effect = Exception("Restrict 실패")
        message.SenderName = "DoNotReply"
        self.assertTrue(main.sender_matches(message, ["DoNotReply"]))
        message.SenderName = "Someone"
        self.assertFalse(main.sender_matches(message, ["DoNotReply"]))

    def test_save_attachments_uses_restrict_without_sender_name(self):
        delegated = MagicMock()
        delegated.Subject = "[위임] PJA-2R8603; 한송희(팀); OA [서 2024-12-23]; 001"
        delegated.Attachments.Count = 1
        type(delegated).SenderName = PropertyMock(side_effect=AssertionError("SenderName 접근"))

        no_attach = MagicMock()
        no_attach.Subject = "[위임] 첨부 없음"
        no_attach.Attachments.Count = 0
        type(no_attach).SenderName = PropertyMock(side_effect=AssertionError("SenderName 접근"))

        outlook = MagicMock()
        restrict = outlook.GetDefaultFolder.return_value.Items.Restrict
        restrict.return_value = [delegated, no_attach]

        with patch.object(main, "_save_attachments_for_message") as save_one:
            main.save_attachments(outlook)

        restrict.assert_called_once_with(main.build_sender_filter(main.DELEGATION_SENDERS))
        save_one.assert_called_once()
        self.assertIs(save_one.call_args[0][0], delegated)


PRENOTICE_BODY = (
    "아래 사건의 담당자를 지정합니다.\r\n"
    "20261012\t10-2025-7009898\tPA-T2K6640\t윤민우\t테크닙 에너지스 프랑스\t거절결정서\r\n"
    "20261012\t10-2025-0070677\tTEL-20035KR-1\t이윤지\t도쿄엘렉트론가부시키가이샤\t거절결정서\r\n"
    "20261012\t10-2025-7002565\tPA-T2K6503\t이예진\t누보 피그노네 에스알엘\t보정요구서\r\n"
    "20261012\t10-2025-7002565\tPA-T2K6503\t이예진\t누보 피그노네 에스알엘\t보정요구서\r\n"
    "20261012\t10-2025-7000694\tPA-2T8084\t신용석\t에어 프로덕츠 앤드 케미칼스, 인코오포레이티드\t의견제출통지서\r\n"
)

DUE_DATE_XML = """<?xml version="1.0" encoding="UTF-8"?>
<response><header><successYN>Y</successYN><resultCode>00</resultCode><resultMsg>NORMAL SERVICE.</resultMsg></header>
<body><items>
<dueDateApplicationNoticeApplnoInfo><ApplicationNumber>1020200000001</ApplicationNumber>
 <Mailingnumber>952020049798050</Mailingnumber><DocNm>의견제출통지서</DocNm><ExpirationDuedateDate>20200807</ExpirationDuedateDate>
</dueDateApplicationNoticeApplnoInfo>
<dueDateApplicationNoticeApplnoInfo><ApplicationNumber>1020200000001</ApplicationNumber>
 <Mailingnumber>952020069827303</Mailingnumber><DocNm>최후의견제출통지서</DocNm><ExpirationDuedateDate>20201212</ExpirationDuedateDate>
</dueDateApplicationNoticeApplnoInfo>
<dueDateApplicationNoticeApplnoInfo><ApplicationNumber>1020200000001</ApplicationNumber>
 <Mailingnumber>952020012159905</Mailingnumber><DocNm>의견제출통지서</DocNm><ExpirationDuedateDate>20200531</ExpirationDuedateDate>
</dueDateApplicationNoticeApplnoInfo>
</items></body></response>"""


class TestPrenoticeRows(unittest.TestCase):
    """예고 메일 본문 표 파싱"""

    def test_only_assignee_rows(self):
        rows = main.parse_prenotice_rows(PRENOTICE_BODY, "신용석")
        self.assertEqual(rows, [{
            "ref": "PA-2T8084",
            "app_no": "10-2025-7000694",
            "receipt_date": "2026-10-12",
            "doc_type": "의견제출통지서",
            "applicant": "에어 프로덕츠 앤드 케미칼스, 인코오포레이티드",
        }])

    def test_duplicate_rows_removed(self):
        rows = main.parse_prenotice_rows(PRENOTICE_BODY, "이예진")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["doc_type"], "보정요구서")

    def test_ref_normalized_like_subject(self):
        """「TEL-20035KR-1」 → 위임 메일 제목과 같은 규칙으로 「TEL-20035」"""
        rows = main.parse_prenotice_rows(PRENOTICE_BODY, "이윤지")
        self.assertEqual(rows[0]["ref"], ref_num_key.findall("[위임] TEL-20035KR-1; ...")[0])

    def test_cell_per_line_body(self):
        """실제 Outlook Body: HTML 표의 칸마다 한 줄, 사이에 빈 줄"""
        body = "\r\n\r\n".join(
            cell for line in PRENOTICE_BODY.splitlines()[1:] for cell in line.split("\t")
        )
        self.assertEqual(main.parse_prenotice_rows(body, "신용석"),
                         main.parse_prenotice_rows(PRENOTICE_BODY, "신용석"))
        self.assertEqual(len(main.parse_prenotice_rows(body, "이예진")), 1)

    def test_no_rows(self):
        self.assertEqual(main.parse_prenotice_rows("본문 없음", "신용석"), [])
        self.assertEqual(main.parse_prenotice_rows(None, "신용석"), [])


class TestResolveOaDates(unittest.TestCase):
    """OA 청서류 기한 결정 (KIPRIS 호출 없음)"""

    MAIL = date(2026, 11, 20)
    FALLBACK = date(2027, 3, 9)
    TODAY = date(2026, 10, 15)

    def test_kipris_due_used(self):
        case = {"app_no": "10-2025-7000694", "due_date": "2027-02-08"}
        self.assertEqual(main.resolve_oa_dates(self.MAIL, self.FALLBACK, case, self.TODAY),
                         (date(2027, 2, 8), "KIPRIS", ""))

    def test_no_case_uses_formula(self):
        d, src, warn = main.resolve_oa_dates(self.MAIL, self.FALLBACK, None, self.TODAY)
        self.assertEqual((d, src), (self.FALLBACK, "계산식"))
        self.assertTrue(warn)

    def test_missing_due_uses_formula(self):
        case = {"app_no": "10-2025-7000694", "due_date": ""}
        d, src, warn = main.resolve_oa_dates(self.MAIL, self.FALLBACK, case, self.TODAY)
        self.assertEqual((d, src), (self.FALLBACK, "계산식"))
        self.assertIn("10-2025-7000694", warn)

    def test_due_before_today_ignored(self):
        case = {"due_date": "2026-10-01"}
        d, src, warn = main.resolve_oa_dates(self.MAIL, self.FALLBACK, case, self.TODAY)
        self.assertEqual((d, src), (self.FALLBACK, "계산식"))
        self.assertIn("2026-10-01", warn)

    def test_due_not_after_mail_date_ignored(self):
        case = {"due_date": "2026-11-20"}
        d, src, _ = main.resolve_oa_dates(self.MAIL, self.FALLBACK, case, self.TODAY)
        self.assertEqual((d, src), (self.FALLBACK, "계산식"))


class TestFindOaCase(unittest.TestCase):

    def test_latest_oa_entry_for_ref(self):
        cache = {
            "PA-2T8084|의견제출통지서": {"doc_type": "의견제출통지서", "receipt_date": "2026-03-02"},
            "PA-2T8084|최후의견제출통지서": {"doc_type": "최후의견제출통지서", "receipt_date": "2026-10-12"},
            "PA-2T8084|보정요구서": {"doc_type": "보정요구서", "receipt_date": "2026-11-01"},
            "PA-2T8085|의견제출통지서": {"doc_type": "의견제출통지서", "receipt_date": "2026-12-01"},
        }
        self.assertEqual(main.find_oa_case(cache, "PA-2T8084")["receipt_date"], "2026-10-12")
        self.assertIsNone(main.find_oa_case(cache, "PA-0000000"))


class TestHandlePrenotice(unittest.TestCase):
    """예고 메일 처리 — 캐시 저장, KIPRIS 1회 조회, 같은 메일 재처리 시 호출 안 함"""

    def setUp(self):
        import tempfile
        self.tmpdir = tempfile.TemporaryDirectory()
        self.cache_patch = patch.object(main, "CASE_CACHE_FILE",
                                        os.path.join(self.tmpdir.name, "case_cache.json"))
        self.cache_patch.start()

    def tearDown(self):
        self.cache_patch.stop()
        self.tmpdir.cleanup()

    def _message(self):
        message = make_mock_message("FW: 사건 담당자 지정", date.today().year,
                                    date.today().month, date.today().day)
        message.Body = PRENOTICE_BODY
        return message

    def test_stores_due_date_and_skips_repeat(self):
        notices = [kipris.Notice("952026093034069", "의견제출통지서", "2027-02-08")]
        with patch.object(kipris, "fetch_due_dates", return_value=notices) as fetch:
            self.assertEqual(main.handle_prenotice(self._message()), 1)
            self.assertEqual(main.handle_prenotice(self._message()), 0)
        fetch.assert_called_once_with("10-2025-7000694")

        case = main.find_oa_case(main.load_case_cache(), "PA-2T8084")
        self.assertEqual(case["due_date"], "2027-02-08")
        self.assertEqual(case["receipt_date"], "2026-10-12")

    def test_kipris_failure_still_stores_app_no(self):
        with patch.object(kipris, "fetch_due_dates", side_effect=kipris.KiprisError("키 없음")):
            main.handle_prenotice(self._message())
        case = main.find_oa_case(main.load_case_cache(), "PA-2T8084")
        self.assertEqual((case["app_no"], case["due_date"]), ("10-2025-7000694", ""))

    def test_missing_due_retried_on_next_prenotice_pass(self):
        """지난번 조회가 실패한 행은 예고 메일을 다시 처리할 때 다시 조회한다"""
        with patch.object(kipris, "fetch_due_dates", side_effect=kipris.KiprisError("연결 끊김")):
            main.handle_prenotice(self._message())
        notices = [kipris.Notice("952026093034069", "의견제출통지서", "2027-02-08")]
        with patch.object(kipris, "fetch_due_dates", return_value=notices) as fetch:
            self.assertEqual(main.handle_prenotice(self._message()), 1)
        fetch.assert_called_once()
        case = main.find_oa_case(main.load_case_cache(), "PA-2T8084")
        self.assertEqual(case["due_date"], "2027-02-08")


class TestScheduleRegistration(unittest.TestCase):
    """알림창 날짜 편집 규칙과 일정 등록"""

    def test_dates_editable_only_for_oa_without_kipris(self):
        self.assertFalse(main.dates_editable("OA", "KIPRIS"))
        self.assertTrue(main.dates_editable("OA", "계산식"))
        self.assertFalse(main.dates_editable("FR", ""))

    def test_register_adds_when_not_duplicate(self):
        folder = MagicMock()
        appointment = folder.Items.Add.return_value
        with patch.object(main, "check_duplicate_appointment", return_value=False):
            status = main._register_appointment(folder, "PA-2T8084_의견서", date(2027, 2, 8), "본문")
        self.assertEqual(status, main.STATUS_ADDED)
        self.assertEqual(appointment.Subject, "PA-2T8084_의견서")
        self.assertTrue(appointment.AllDayEvent)
        self.assertEqual(appointment.Start.date(), date(2027, 2, 8))
        appointment.Save.assert_called_once()

    def test_register_skips_duplicate(self):
        folder = MagicMock()
        with patch.object(main, "check_duplicate_appointment", return_value=True):
            status = main._register_appointment(folder, "PA-2T8084_의견서", date(2027, 2, 8), "본문")
        self.assertEqual(status, main.STATUS_DUPLICATE)
        folder.Items.Add.assert_not_called()


class TestNaverRegistration(unittest.TestCase):
    """Outlook 등록과 함께 네이버 캘린더에도 등록"""

    ITEMS = [("Comments", "PA-2T8084_Comments", date(2026, 10, 1)),
             ("Action", "PA-2T8084_의견서", date(2027, 2, 8))]
    NAMES = ["Mail", "No_Instructions"]
    SELECTED = [(0, date(2026, 10, 1)), (1, date(2027, 2, 8))]

    def setUp(self):
        patcher = patch.dict(main.config, {"naver_calendars": {"Mail": "코멘트", "No_Instructions": "기한"}})
        patcher.start()
        self.addCleanup(patcher.stop)
        patcher = patch.object(main, "_register_appointment", return_value=main.STATUS_ADDED)
        self.outlook = patcher.start()
        self.addCleanup(patcher.stop)

    def _run(self, naver):
        return main.register_items([MagicMock(), MagicMock()], self.NAMES, self.ITEMS, "본문", self.SELECTED, naver)

    def test_naver_off_only_outlook(self):
        results = self._run(None)
        self.assertEqual(results, {0: (main.STATUS_ADDED, None), 1: (main.STATUS_ADDED, None)})

    def test_naver_each_folder_to_its_calendar(self):
        naver = MagicMock()
        naver.add_all_day.side_effect = [True, False]
        results = self._run(naver)
        self.assertEqual(results, {0: (main.STATUS_ADDED, main.STATUS_ADDED),
                                   1: (main.STATUS_ADDED, main.STATUS_DUPLICATE)})
        calls = [c.args[0] for c in naver.add_all_day.call_args_list]
        self.assertEqual(calls, ["코멘트", "기한"])

    def test_naver_failure_keeps_outlook(self):
        naver = MagicMock()
        naver.add_all_day.side_effect = naver_cal.NaverError("네이버 인증 실패")
        results = self._run(naver)
        self.assertEqual(self.outlook.call_count, 2)
        for outlook, naver_status in results.values():
            self.assertEqual(outlook, main.STATUS_ADDED)
            self.assertTrue(naver_status.startswith("오류"))

    def test_missing_calendar_name_is_error(self):
        main.config["naver_calendars"] = {"Mail": "코멘트"}
        naver = MagicMock()
        naver.add_all_day.return_value = True
        results = self._run(naver)
        self.assertEqual(results[0][1], main.STATUS_ADDED)
        self.assertTrue(results[1][1].startswith("오류"))


class TestNaverCal(unittest.TestCase):
    """naver_cal — iCalendar 만들기와 CalDAV 흐름 (네트워크는 가짜)"""

    def test_uid_deterministic(self):
        a = naver_cal.make_uid("기한", "PA-2T8084_의견서", date(2027, 2, 8))
        self.assertEqual(a, naver_cal.make_uid("기한", "PA-2T8084_의견서", date(2027, 2, 8)))
        self.assertNotEqual(a, naver_cal.make_uid("기한", "PA-2T8084_의견서", date(2027, 2, 9)))

    def test_event_url_uses_full_uid(self):
        """네이버는 일정을 UID 전체(@ 포함)로 저장한다 — 실제 서버에서 확인."""
        url = naver_cal.event_url_for("https://caldav.calendar.naver.com/caldav/me/calendar/1/", "abc@tmproject")
        self.assertEqual(url, "https://caldav.calendar.naver.com/caldav/me/calendar/1/abc%40tmproject.ics")

    def test_all_day_ics(self):
        ics = naver_cal.build_ics("u@tmproject", "PA-2T8084_의견서", date(2027, 12, 31), "a;b,c\\d\r\n줄",
                                  now=datetime(2026, 10, 15, 3, 0, 0))
        unfolded = ics.replace("\r\n ", "")
        self.assertIn("DTSTART;VALUE=DATE:20271231\r\n", unfolded)
        self.assertIn("DTEND;VALUE=DATE:20280101\r\n", unfolded)
        self.assertIn("SUMMARY:PA-2T8084_의견서\r\n", unfolded)
        self.assertIn("DESCRIPTION:a\\;b\\,c\\\\d\\n줄\r\n", unfolded)

    def test_empty_body_has_no_description(self):
        ics = naver_cal.build_ics("u@tmproject", "제목", date(2027, 1, 1), "")
        self.assertNotIn("DESCRIPTION", ics)

    def test_long_lines_folded_by_octets(self):
        ics = naver_cal.build_ics("u@tmproject", "제목", date(2027, 1, 1), "가" * 200)
        for line in ics.split("\r\n"):
            self.assertLessEqual(len(line.encode("utf-8")), 75)
        self.assertIn("가" * 200, ics.replace("\r\n ", ""))

    MULTI = """<?xml version="1.0"?><d:multistatus xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">{}</d:multistatus>"""

    def _fake(self, session, put_status=201, existing=False):
        principal = self.MULTI.format('<d:response><d:href>/</d:href><d:propstat><d:prop>'
                                      '<d:current-user-principal><d:href>/principals/me/</d:href></d:current-user-principal>'
                                      '</d:prop></d:propstat></d:response>')
        home = self.MULTI.format('<d:response><d:href>/principals/me/</d:href><d:propstat><d:prop>'
                                 '<c:calendar-home-set><d:href>/calendars/me/</d:href></c:calendar-home-set>'
                                 '</d:prop></d:propstat></d:response>')
        cals = self.MULTI.format(
            '<d:response><d:href>/calendars/me/</d:href><d:propstat><d:prop><d:resourcetype><d:collection/></d:resourcetype>'
            '</d:prop></d:propstat></d:response>'
            '<d:response><d:href>/calendars/me/a</d:href><d:propstat><d:prop><d:displayname>기한</d:displayname>'
            '<d:resourcetype><d:collection/><c:calendar/></d:resourcetype>'
            '<c:supported-calendar-component-set><c:comp name="VEVENT"/></c:supported-calendar-component-set>'
            '</d:prop></d:propstat></d:response>'
            '<d:response><d:href>/calendars/me/todo/</d:href><d:propstat><d:prop><d:displayname>할 일</d:displayname>'
            '<d:resourcetype><d:collection/><c:calendar/></d:resourcetype>'
            '<c:supported-calendar-component-set><c:comp name="VTODO"/></c:supported-calendar-component-set>'
            '</d:prop></d:propstat></d:response>')
        calls = []

        def request(method, url, body="", headers=None):
            calls.append((method, url))
            if method == "PROPFIND":
                text = {naver_cal.BASE_URL: principal}.get(url) or (home if "principals" in url else cals)
                return 207, url, text
            if method == "GET":
                return (200 if existing else 404), url, ""
            return put_status, url, ""

        session._request = request
        return calls

    def test_calendars_only_vevent(self):
        s = naver_cal.Session("id", "pw")
        self._fake(s)
        self.assertEqual(s.calendars(), {"기한": "https://caldav.calendar.naver.com/calendars/me/a/"})

    def test_add_puts_event(self):
        s = naver_cal.Session("id", "pw")
        calls = self._fake(s)
        self.assertTrue(s.add_all_day("기한", "PA-2T8084_의견서", date(2027, 2, 8), "본문"))
        method, url = calls[-1]
        self.assertEqual(method, "PUT")
        self.assertTrue(url.startswith("https://caldav.calendar.naver.com/calendars/me/a/"))
        self.assertTrue(url.endswith(".ics"))

    def test_exists_reads_only(self):
        s = naver_cal.Session("id", "pw")
        calls = self._fake(s, existing=True)
        self.assertTrue(s.exists("기한", "x", date(2027, 2, 8)))
        s = naver_cal.Session("id", "pw")
        calls = self._fake(s)
        self.assertFalse(s.exists("기한", "x", date(2027, 2, 8)))
        self.assertEqual([m for m, _u in calls if m not in ("PROPFIND",)], ["GET"])

    def test_existing_or_412_is_duplicate(self):
        s = naver_cal.Session("id", "pw")
        calls = self._fake(s, existing=True)
        self.assertFalse(s.add_all_day("기한", "x", date(2027, 2, 8), ""))
        self.assertNotIn("PUT", [m for m, _u in calls])
        s = naver_cal.Session("id", "pw")
        self._fake(s, put_status=412)
        self.assertFalse(s.add_all_day("기한", "x", date(2027, 2, 8), ""))

    def test_unknown_calendar_raises(self):
        s = naver_cal.Session("id", "pw")
        self._fake(s)
        with self.assertRaises(naver_cal.NaverError):
            s.add_all_day("없는 캘린더", "x", date(2027, 2, 8), "")

    def test_discovery_error_cached(self):
        s = naver_cal.Session("id", "pw")
        s._request = MagicMock(side_effect=naver_cal.NaverError("연결 실패"))
        for _ in range(2):
            with self.assertRaises(naver_cal.NaverError):
                s.add_all_day("기한", "x", date(2027, 2, 8), "")
        self.assertEqual(s._request.call_count, 1)

    def test_no_credentials_raises(self):
        with patch.object(naver_cal, "credentials", return_value=None):
            with self.assertRaises(naver_cal.NaverError):
                naver_cal.Session().calendars()


class TestNaverSync(unittest.TestCase):
    """Outlook 하위 캘린더 → 같은 이름의 네이버 캘린더"""

    @staticmethod
    def _item(subject, start, all_day=True):
        item = MagicMock()
        item.Subject, item.Start, item.AllDayEvent = subject, start, all_day
        return item

    def test_outlook_events_since_and_folder(self):
        folders = {
            "Mail": [self._item("A_Comments", datetime(2026, 10, 15)), self._item("B_Comments", datetime(2026, 10, 16))],
            "Complete": [self._item("C_의견서", datetime(2026, 11, 1, 10, 30), all_day=False)],
        }
        calendar = MagicMock()
        calendar.Folders.__getitem__.side_effect = lambda name: MagicMock(Items=MagicMock(
            __iter__=lambda _self: iter(folders[name])))
        events = naver_sync.outlook_events(calendar, date(2026, 10, 16), folders=["Mail", "Complete"])
        self.assertEqual(events, [("Mail", "B_Comments", date(2026, 10, 16), True),
                                  ("Complete", "C_의견서", date(2026, 11, 1), False)])


class TestKipris(unittest.TestCase):
    """KIPRIS 마감기한 응답 해석 (CPProject 응답 모양)"""

    def test_parse_and_pick_latest(self):
        notices = kipris.parse_due_dates(DUE_DATE_XML)
        self.assertEqual(len(notices), 3)
        picked = kipris.pick_oa_due(notices, "의견제출통지서")
        # 「최후의견제출통지서」도 「의견제출통지서」를 포함 — 발송번호가 가장 큰 것
        self.assertEqual((picked.send_number, picked.due_date), ("952020069827303", "2020-12-12"))

    def test_pick_none(self):
        self.assertIsNone(kipris.pick_oa_due(kipris.parse_due_dates(DUE_DATE_XML), "거절결정서"))
        self.assertIsNone(kipris.pick_oa_due([], "의견제출통지서"))

    def test_result_codes_raise(self):
        for code, msg in (("30", "SERVICE_KEY_IS_NOT_REGISTERED_ERROR"),
                          ("31", "DEADLINE_HAS_EXPIRED_ERROR"),
                          ("101", ""), ("11", "OTHER")):
            xml = f"<response><header><resultCode>{code}</resultCode><resultMsg>{msg}</resultMsg></header></response>"
            with self.subTest(code=code), self.assertRaises(kipris.KiprisError):
                kipris.parse_due_dates(xml)

    def test_not_xml_raises(self):
        with self.assertRaises(kipris.KiprisError):
            kipris.parse_due_dates("<html>code400")

    def test_appln_number(self):
        self.assertEqual(kipris.appln_number("10-2025-7000694"), "1020257000694")
        self.assertEqual(kipris.appln_number("10-2025-70"), "")


if __name__ == "__main__":
    unittest.main(verbosity=2)
