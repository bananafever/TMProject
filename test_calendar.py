"""
캘린더 기일 설정 기능 테스트
테스트 대상:
  - _extract_message_info(): 이메일 제목에서 ref_num, 이름, 날짜 추출
  - OA 기일 계산: target_date + 4개월 - 11일
  - FR 기일 계산: receipt_date(수신일) + 3개월
  - ref_num_key 정규식 패턴
"""

import logging
import os
import shutil
import sys
import unittest
from datetime import date, datetime, timedelta
from unittest.mock import MagicMock, PropertyMock, patch

# -----------------------------------------------------------------------
# win32com, PySide6는 실제 Outlook/GUI 없이도 테스트할 수 있도록 mock 처리
# 진짜를 들일 수 있으면 그대로 쓴다 — 진짜 PySide6 가 이미 들어와 있을 때(pytest-qt 등) 그 밑의 모듈을
# 가짜로 바꾸면 shiboken 이 Segmentation fault 로 죽는다. 한 묶음이라도 못 들이면 묶음 전체를 가짜로.
# -----------------------------------------------------------------------
def _fake_unless_importable(names):
    try:
        for name in names:
            __import__(name)
    except Exception:
        for name in names:
            sys.modules[name] = MagicMock()


_fake_unless_importable(["win32com", "win32com.client"])
_fake_unless_importable(["PySide6", "PySide6.QtWidgets", "PySide6.QtGui", "PySide6.QtCore", "PySide6.QtSvg"])

# main 의 logging.basicConfig(filename='app.log') 가 걸리지 않게 — 테스트가 실제 앱 로그에 쓰지 않는다
logging.getLogger().addHandler(logging.NullHandler())
# 실제 데이터 폴더(settings.json·case_cache.json …)를 읽거나 쓰지 않게 빈 임시 폴더를 쓴다
import tempfile  # noqa: E402
_TEST_DATA_DIR = tempfile.TemporaryDirectory()
os.environ["TASKMANAGER_DATA_DIR"] = _TEST_DATA_DIR.name
# 실제 NAS DB·네이버 캘린더·KIPRIS 에 닿지 않게 — ChekPat credentials.dat(NAS 접속)·kipris_key.dat 와
# TMProject naver_cred.dat 는 %APPDATA% 아래에서 읽는다. 윈도에서 시험을 돌리면 진짜 접속 정보로 NAS 에
# 받은 위임 줄(A·B·C)이 실제로 써졌다(2026-10-24). 빈 임시 폴더를 %APPDATA% 로, 접속 환경변수는 지운다.
_TEST_APPDATA = tempfile.TemporaryDirectory()
os.environ["APPDATA"] = _TEST_APPDATA.name
for _key in ("KIPRIS_API_KEY", "NAVER_ID", "NAVER_APP_PASSWORD"):
    os.environ.pop(_key, None)
try:  # 그래도 진짜 pymysql 로 붙으려 하면 시험이 실패하게 — 가짜 pymysql 을 끼우는 시험은 그대로 돈다
    import pymysql as _pymysql

    def _no_real_nas(*_a, **_k):
        raise OSError("시험 — 실제 NAS DB 에는 붙지 않는다")
    _pymysql.connect = _no_real_nas
except ImportError:
    pass

from dateutil.relativedelta import relativedelta  # noqa: E402

import kipris  # noqa: E402
# 실제 맥 키체인도 시험 중에는 읽지 않는다
kipris._mac_credentials = lambda: None
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

    def __lt__(self, other):  # pywintypes.datetime 처럼 정렬할 수 있게
        return self._dt < other._dt

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


    def test_transferred_attachments_are_not_downloaded_again(self):
        """File Transfer 로 02_Review 에 옮긴 검수 첨부는 메일이 받은편지함에 남아 있어도 다시 받지 않는다(10-24)."""
        with tempfile.TemporaryDirectory() as tmp:
            parent, target = os.path.join(tmp, "Tasks"), os.path.join(tmp, "02_Review")
            msg = MagicMock()
            msg.Subject = "[위임] LA-26999; 김가공(팀); 검토 요청"
            msg.CreationTime = datetime(2026, 8, 28, 13, 53)
            msg.Attachments = [MagicMock(FileName="a.docx")]
            with patch.dict(main.config, {"parent_path": parent, "target_path": target}):
                os.makedirs(os.path.join(target, main.transfer_name("2026-08-28", "13hr53mn", "LA-26999", "김가공")))
                got = []
                main._save_attachments_for_message(msg, got)
                self.assertEqual(got, [("2026-08-28 13:53", "LA-26999", "김가공", "Already Transferred")])
                msg.Attachments[0].SaveAsFile.assert_not_called()
                self.assertFalse(os.path.exists(parent))                 # 작업 폴더도 만들지 않는다
                shutil.rmtree(target)
                got = []
                main._save_attachments_for_message(msg, got)              # 옮긴 자리가 없으면 예전처럼 받는다
                self.assertEqual(got[0][3], "Downloaded")
                msg.Attachments[0].SaveAsFile.assert_called_once()
            with patch.dict(main.config, {"parent_path": parent, "target_path": ""}):
                self.assertFalse(main.already_transferred("2026-08-28", "13hr53mn", "LA-26999", "김가공"))

    def test_transfer_name_matches_task_folder(self):
        """작업 폴더 이름(시각_Ref_이름)을 transfer_clicked 처럼 쪼개면 already_transferred 가 보는 자리와 같다."""
        time_id, ref, name = "13hr53mn", "LA-26999", "김가공"
        time_info, ref_info, person_info = (time_id + "_" + ref + "_" + name).split("_", 2)
        self.assertEqual(main.transfer_name("2026-08-28", time_info, ref_info, person_info),
                         os.path.join("LA-26999", "2026-08-28_13hr53mn_김가공"))


PRENOTICE_BODY = (
    "아래 사건의 담당자를 지정합니다.\r\n"
    "20261012\t10-2025-7009898\tPA-T2K6640\t윤민우\t테크닙 에너지스 프랑스\t거절결정서\r\n"
    "20261012\t10-2025-0070677\tTEL-20035KR-1\t이윤지\t도쿄엘렉트론가부시키가이샤\t거절결정서\r\n"
    "20261012\t10-2025-7002565\tPA-T2K6503\t이예진\t누보 피그노네 에스알엘\t보정요구서\r\n"
    "20261012\t10-2025-7002565\tPA-T2K6503\t이예진\t누보 피그노네 에스알엘\t보정요구서\r\n"
    "20261012\t10-2025-7000694\tPA-2T8084\t신용석\t에어 프로덕츠 앤드 케미칼스, 인코오포레이티드\t의견제출통지서\r\n"
)

FR_PRENOTICE_BODY = (
    "아래 사건의 담당자를 지정합니다.\r\n"
    "20261012\t10-2014-0049963\tPA-2T9001\t신용석\t폭스바겐 악티엔게젤샤프트\t거절결정서\r\n"
)


def biblio_xml(disposal="", documents=()):
    """서지 상세 응답 (CPProject records.parse_biblio 가 읽는 모양)."""
    docs = "".join(f"<legalStatusInfo><documentName>{d}</documentName></legalStatusInfo>" for d in documents)
    return ("<response><header><successYN>Y</successYN><resultCode>00</resultCode></header><body><item>"
            f"<biblioSummaryInfoArray><biblioSummaryInfo><finalDisposal>{disposal}</finalDisposal>"
            f"</biblioSummaryInfo></biblioSummaryInfoArray><legalStatusInfoArray>{docs}</legalStatusInfoArray>"
            "</item></body></response>")

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

    def test_stores_applicant(self):
        with patch.object(kipris, "fetch_due_dates", return_value=[]):
            main.handle_prenotice(self._message())
        case = main.find_oa_case(main.load_case_cache(), "PA-2T8084")
        self.assertEqual(case["applicant"], "에어 프로덕츠 앤드 케미칼스, 인코오포레이티드")

    def test_fills_applicant_on_old_entry_without_kipris(self):
        """출원인 저장 전에 받은 행은 KIPRIS 를 부르지 않고 출원인만 채운다"""
        notices = [kipris.Notice("952026093034069", "의견제출통지서", "2027-02-08")]
        with patch.object(kipris, "fetch_due_dates", return_value=notices):
            main.handle_prenotice(self._message())
        cache = main.load_case_cache()
        for v in cache.values():
            v.pop("applicant")
        main.save_case_cache(cache)
        with patch.object(kipris, "fetch_due_dates") as fetch:
            self.assertEqual(main.handle_prenotice(self._message()), 1)
        fetch.assert_not_called()
        case = main.find_oa_case(main.load_case_cache(), "PA-2T8084")
        self.assertEqual((case["applicant"], case["due_date"]),
                         ("에어 프로덕츠 앤드 케미칼스, 인코오포레이티드", "2027-02-08"))

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


class TestHandlePrenoticeFr(TestHandlePrenotice):
    """거절결정서 행 — 재심사에 의한 거절결정인지 KIPRIS 서지 상세 1회 조회"""

    def _message(self):
        message = super()._message()
        message.Body = FR_PRENOTICE_BODY
        return message

    def _case(self):
        return main.find_case(main.load_case_cache(), "PA-2T9001", main.FR_DOC_TYPE)

    def test_stores_reexam_and_skips_repeat(self):
        with patch.object(kipris, "fetch_reexam", return_value=True) as fetch,                 patch.object(kipris, "fetch_due_dates") as due:
            self.assertEqual(main.handle_prenotice(self._message()), 1)
            self.assertEqual(main.handle_prenotice(self._message()), 0)
        fetch.assert_called_once_with("10-2014-0049963")
        due.assert_not_called()
        self.assertIs(self._case()["reexam"], True)

    def test_failure_retried_next_pass(self):
        with patch.object(kipris, "fetch_reexam", side_effect=kipris.KiprisError("연결 끊김")):
            main.handle_prenotice(self._message())
        self.assertNotIn("reexam", self._case())
        with patch.object(kipris, "fetch_reexam", return_value=False) as fetch:
            self.assertEqual(main.handle_prenotice(self._message()), 1)
        fetch.assert_called_once()
        self.assertIs(self._case()["reexam"], False)

    def test_delegation_requeries_missing_reexam(self):
        """예고 메일 때 못 받은 재심사 여부(KIPRIS 상한)는 위임 메일 때 한 번 더 묻는다 — 받으면 다시 묻지 않는다"""
        with patch.object(kipris, "fetch_reexam", side_effect=kipris.KiprisError("상한")):
            main.handle_prenotice(self._message())
        subject = "[위임] PA-2T9001; 이여름(팀); 거절결정서; 001 [서 2026-10-22]"
        with patch.object(kipris, "fetch_reexam", return_value=True) as fetch:
            self.assertIs(main.refresh_reexam(subject, date(2026, 10, 13)), True)
            self.assertIsNone(main.refresh_reexam(subject, date(2026, 10, 13)))
        fetch.assert_called_once_with("10-2014-0049963")
        self.assertIs(self._case()["reexam"], True)

    def test_delegation_requery_failure_leaves_unknown(self):
        with patch.object(kipris, "fetch_reexam", side_effect=kipris.KiprisError("상한")):
            main.handle_prenotice(self._message())
            self.assertIsNone(main.refresh_reexam("[위임] PA-2T9001; 거절결정서", date(2026, 10, 13)))
        self.assertNotIn("reexam", self._case())

    test_stores_due_date_and_skips_repeat = None
    test_stores_applicant = None
    test_fills_applicant_on_old_entry_without_kipris = None
    test_kipris_failure_still_stores_app_no = None
    test_missing_due_retried_on_next_prenotice_pass = None


class TestApplicant(unittest.TestCase):
    """출원인 약칭·제목"""

    def setUp(self):
        patcher = patch.dict(main.config, {"applicant_abbr": {}})
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_auto_abbr(self):
        cases = {
            "도쿄엘렉트론가부시키가이샤": "도쿄엘렉트론",
            "에어 프로덕츠 앤드 케미칼스, 인코오포레이티드": "에어프로덕츠",
            "테크닙 에너지스 프랑스": "테크닙에너지",
            "폭스바겐 악티엔게젤샤프트": "폭스바겐",
            "아우디 아게": "아우디",
            "주식회사 엘지화학": "엘지화학",
            "가부시키가이샤 덴소": "덴소",
            "혼다 기켄 고교 가부시키가이샤": "혼다기켄고교",
        }
        for name, want in cases.items():
            with self.subTest(name=name):
                self.assertEqual(main.applicant_abbr(name), want)

    def test_table_wins(self):
        main.config["applicant_abbr"] = {"도쿄엘렉트론가부시키가이샤": "TEL",
                                         "에이엠에스-오스람  인터내셔널 게엠베하": "오스람"}
        self.assertEqual(main.applicant_abbr("도쿄엘렉트론가부시키가이샤"), "TEL")
        self.assertEqual(main.applicant_abbr(" 에이엠에스-오스람 인터내셔널 게엠베하"), "오스람")

    def test_with_applicant(self):
        self.assertEqual(main.with_applicant("PA-2T8084_의견서", "에어 프로덕츠 앤드 케미칼스, 인코오포레이티드"),
                         "PA-2T8084_의견서_에어프로덕츠")
        self.assertEqual(main.with_applicant("PA-2T8084_의견서", ""), "PA-2T8084_의견서")
        self.assertEqual(main.with_applicant("PA-2T8084_의견서", None), "PA-2T8084_의견서")

    def test_data_dir_order(self):
        with patch.dict(os.environ, {"TASKMANAGER_DATA_DIR": "X:/data"}):
            self.assertEqual(main.data_dir(), "X:/data")
        with patch.dict(os.environ, {}, clear=True):
            with patch.object(sys, "frozen", True, create=True), \
                    patch.object(sys, "executable", "D:/apps/TaskManager/TaskManager.exe"):
                self.assertEqual(os.path.normpath(main.data_dir()), os.path.normpath("D:/apps/TaskManager"))
            with patch.object(os.path, "isdir", return_value=True):
                self.assertEqual(main.data_dir(), main.DEFAULT_DATA_DIR)
            with patch.object(os.path, "isdir", return_value=False):
                self.assertEqual(main.data_dir(), os.path.dirname(os.path.abspath(main.__file__)))
        self.assertEqual(main.DATA_DIR, os.environ["TASKMANAGER_DATA_DIR"])  # 테스트는 임시 폴더

    def test_parse_names(self):
        self.assertEqual(main.parse_names(" 이여름 ,한송희,, 이여름 "), ["이여름", "한송희"])
        self.assertEqual(main.parse_names(""), [])

    def test_calendar_assignees_from_config(self):
        with patch.dict(main.config, {"calendar_assignees": ["신용석"]}):
            self.assertEqual(main.calendar_assignees(), ["신용석"])
            message = make_mock_message("[위임] PA-2T8084; 한송희(팀); OA [서 2026-10-01]; 001")
            with patch.object(main, "sender_matches", return_value=True):
                self.assertFalse(main.is_calendar_delegation(message))
                message.Subject = "[위임] PA-2T8084; 신용석(팀); OA [서 2026-10-01]; 001"
                self.assertTrue(main.is_calendar_delegation(message))
        with patch.dict(main.config, {"calendar_assignees": []}):
            self.assertEqual(main.calendar_assignees(), ["이여름", "한송희"])

    def test_recent_applicants(self):
        cache = {"a": {"applicant": "아우디  아게"}, "b": {"applicant": "도쿄엘렉트론가부시키가이샤"},
                 "c": {"applicant": ""}, "d": {}}
        self.assertEqual(main.recent_applicants(cache, ["아우디 아게"]), ["도쿄엘렉트론가부시키가이샤"])

    def test_abbr_table_from_rows(self):
        rows = [(" 도쿄엘렉트론가부시키가이샤 ", "TEL"), ("빈 약칭", ""), ("", "X"),
                ("아우디  아게", "AUDI"), ("아우디 아게", "아우디")]
        self.assertEqual(main.abbr_table_from_rows(rows),
                         {"도쿄엘렉트론가부시키가이샤": "TEL", "아우디 아게": "아우디"})


class TestFrAction(unittest.TestCase):
    """거절결정서 Action 일정 제목"""

    def test_reexam(self):
        subj, label, _src, warn = main.fr_action("PA-2T9001", {"reexam": True})
        self.assertEqual((subj, label, warn), ("PA-2T9001_거결불", "재심사 거절결정서 (FR)", ""))

    def test_general(self):
        subj, _label, _src, warn = main.fr_action("PA-2T9001", {"reexam": False})
        self.assertEqual((subj, warn), ("PA-2T9001_재심사", ""))

    def test_unknown_warns(self):
        for case in (None, {"app_no": "10-2014-0049963"}):
            with self.subTest(case=case):
                subj, _label, _src, warn = main.fr_action("PA-2T9001", case)
                self.assertEqual(subj, "PA-2T9001_재심사")
                self.assertIn("_거결불", warn)


class TestScheduleRegistration(unittest.TestCase):
    """알림창 날짜 편집 규칙과 일정 등록"""

    def test_dates_editable_only_for_oa_without_kipris(self):
        self.assertFalse(main.dates_editable("OA", "KIPRIS"))
        self.assertTrue(main.dates_editable("OA", "계산식"))
        self.assertFalse(main.dates_editable("FR", ""))

    def test_failed_indices_for_retry(self):
        results = {0: main.STATUS_ADDED, 1: "오류: 연결 실패", 2: main.STATUS_DUPLICATE, 3: "오류: 인증 실패"}
        self.assertEqual(main.failed_indices(results), [1, 3])
        self.assertEqual(main.failed_indices({0: main.STATUS_ADDED}), [])


class TestNaverRegistration(unittest.TestCase):
    """위임 메일 일정은 네이버 캘린더에만 등록"""

    ITEMS = [("Comments", "PA-2T8084_Comments", date(2026, 10, 1)),
             ("Action", "PA-2T8084_의견서", date(2027, 2, 8))]
    NAMES = ["Mail", "No_Instructions"]
    SELECTED = [(0, date(2026, 10, 1)), (1, date(2027, 2, 8))]

    def setUp(self):
        patcher = patch.dict(main.config, {"naver_calendars": {"Mail": "코멘트", "No_Instructions": "기한"}})
        patcher.start()
        self.addCleanup(patcher.stop)

    def _run(self, naver, selected=None):
        return main.register_items(self.NAMES, self.ITEMS, selected or self.SELECTED, naver)

    def test_each_folder_to_its_calendar(self):
        naver = MagicMock()
        naver.add_all_day.side_effect = [True, False]
        results = self._run(naver)
        self.assertEqual(results, {0: main.STATUS_ADDED, 1: main.STATUS_DUPLICATE})
        calls = [c.args[:3] for c in naver.add_all_day.call_args_list]
        self.assertEqual(calls, [("코멘트", "PA-2T8084_Comments", date(2026, 10, 1)),
                                 ("기한", "PA-2T8084_의견서", date(2027, 2, 8))])

    def test_failure_is_error_status(self):
        naver = MagicMock()
        naver.add_all_day.side_effect = naver_cal.NaverError("네이버 인증 실패")
        results = self._run(naver)
        self.assertEqual(main.failed_indices(results), [0, 1])

    def test_retry_only_selected(self):
        naver = MagicMock()
        naver.add_all_day.return_value = True
        results = self._run(naver, selected=[(1, date(2027, 2, 9))])
        self.assertEqual(results, {1: main.STATUS_ADDED})
        naver.add_all_day.assert_called_once_with("기한", "PA-2T8084_의견서", date(2027, 2, 9), "")

    def test_memo_passed_as_description(self):
        naver = MagicMock()
        naver.add_all_day.return_value = True
        main.register_items(self.NAMES, self.ITEMS, [(0, date(2026, 10, 1))], naver, "출원인: 아우디 아게")
        naver.add_all_day.assert_called_once_with("코멘트", "PA-2T8084_Comments", date(2026, 10, 1),
                                                  "출원인: 아우디 아게")

    def test_missing_calendar_name_is_error(self):
        main.config["naver_calendars"] = {"Mail": "코멘트"}
        naver = MagicMock()
        naver.add_all_day.return_value = True
        results = self._run(naver)
        self.assertEqual(results[0], main.STATUS_ADDED)
        self.assertTrue(results[1].startswith("오류"))


class TestNaverCal(unittest.TestCase):
    """naver_cal — iCalendar 만들기와 CalDAV 흐름 (네트워크는 가짜)"""

    def test_mac_credentials_from_keychain(self):
        """맥은 DPAPI 가 없어 키체인에서 읽는다 (예전에는 늘 「네이버 계정이 없습니다」)"""
        found = MagicMock(returncode=0, stdout='{"user": "me", "password": "pw"}\n')
        with patch.dict(os.environ, {"NAVER_ID": "", "NAVER_APP_PASSWORD": ""}), \
                patch.object(naver_cal.sys, "platform", "darwin"), \
                patch.object(naver_cal.subprocess, "run", return_value=found) as run:
            self.assertEqual(naver_cal.credentials(), ("me", "pw"))
            run.return_value = MagicMock(returncode=0, stdout='{"user": "me", "password": "한"}'.encode().hex())
            self.assertEqual(naver_cal.credentials(), ("me", "한"))
            run.return_value = MagicMock(returncode=44, stdout="")
            self.assertIsNone(naver_cal.credentials())
        self.assertEqual(run.call_args.args[0][:2], ["security", "find-generic-password"])

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
        self.assertEqual([e[:4] for e in events], [("Mail", "B_Comments", date(2026, 10, 16), True),
                                                   ("Complete", "C_의견서", date(2026, 11, 1), False)])
        self.assertIs(events[0][4], folders["Mail"][1])

    def _delete(self, on_naver, apply, recurring=False):
        session = MagicMock()
        session.exists.return_value = on_naver
        item = self._item("PA-2T8084_의견서", datetime(2027, 2, 8))
        item.IsRecurring = recurring
        status = naver_sync.delete_status(session, "No_Instructions", item.Subject, date(2027, 2, 8), item, apply)
        return status, item

    def test_delete_only_when_on_naver(self):
        status, item = self._delete(on_naver=True, apply=True)
        self.assertEqual(status, "삭제됨")
        item.Delete.assert_called_once()
        status, item = self._delete(on_naver=False, apply=True)
        self.assertEqual(status, "네이버에 없음 — 유지")
        item.Delete.assert_not_called()

    def test_delete_preview_and_recurring_never_delete(self):
        status, item = self._delete(on_naver=True, apply=False)
        self.assertEqual(status, "삭제 예정")
        item.Delete.assert_not_called()
        status, item = self._delete(on_naver=True, apply=True, recurring=True)
        self.assertEqual(status, "반복 일정 — 유지")
        item.Delete.assert_not_called()


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

    def test_parse_reexam(self):
        cases = [
            (biblio_xml("거절결정(일반)", ["거절결정서"]), False),
            (biblio_xml("거절결정(재심사)", ["거절결정서", "[명세서등 보정]보정서(재심사)", "거절결정서"]), True),
            (biblio_xml("거절결정 후 재심사중", ["거절결정서", "[명세서등 보정]보정서(재심사)"]), True),
            (biblio_xml("", ["의견제출통지서", "[명세서등 보정]보정서(재심사)"]), True),
            (biblio_xml("", ["의견제출통지서", "[명세서등 보정]보정서"]), False),
        ]
        for xml, want in cases:
            with self.subTest(xml=xml):
                self.assertIs(kipris.parse_reexam(xml), want)

    def test_parse_reexam_result_code_names_product(self):
        xml = "<response><header><resultCode>101</resultCode></header></response>"
        with self.assertRaisesRegex(kipris.KiprisError, "서지 상세"):
            kipris.parse_reexam(xml)

    def test_appln_number(self):
        self.assertEqual(kipris.appln_number("10-2025-7000694"), "1020257000694")
        self.assertEqual(kipris.appln_number("10-2025-70"), "")


OA_SUBJECT = "[위임] PA-2T8084; 한송희(고객지원팀-책임); OA [서 2026-10-01]; 2610010001"
FR_SUBJECT = "[위임] PA-2T9001; 이여름(팀-담당); 거절결정서 [서 2026-10-05]; 2610050001"


class TestBuildSchedule(unittest.TestCase):
    """위임 메일 한 통의 일정 계산 (등록 전)"""

    RECEIVED = date(2026, 10, 16)
    TODAY = date(2026, 10, 16)

    def setUp(self):
        patcher = patch.dict(main.config, {"applicant_abbr": {}})
        patcher.start()
        self.addCleanup(patcher.stop)

    def _build(self, subject, kind, cache=None):
        return main.build_schedule(subject, self.RECEIVED, kind, cache or {}, self.TODAY)

    def _days(self, plan):
        return [d for _g, _s, d in plan["items"]]

    def test_no_ref(self):
        self.assertIsNone(self._build("[위임] 한송희(팀); OA", "OA"))

    def test_oa_kipris(self):
        cache = {"PA-2T8084|의견제출통지서": {"doc_type": "의견제출통지서", "receipt_date": "2026-10-12",
                                         "due_date": "2027-02-08", "app_no": "10-2025-7000694",
                                         "applicant": "아우디 아게"}}
        plan = self._build(OA_SUBJECT, "OA", cache)
        self.assertEqual(self._days(plan), [date(2026, 10, 1), date(2027, 2, 8)])
        self.assertEqual([s for _g, s, _d in plan["items"]], ["PA-2T8084_Comments_아우디", "PA-2T8084_의견서_아우디"])
        self.assertFalse(plan["editable"])
        self.assertEqual(plan["memo"], "출원인: 아우디 아게\n출원번호: 10-2025-7000694")

    def test_oa_formula_editable(self):
        plan = self._build(OA_SUBJECT, "OA")
        self.assertEqual(self._days(plan), [date(2026, 10, 1), date(2027, 1, 21)])  # 10/1 + 4개월 - 11일
        self.assertTrue(plan["editable"])
        self.assertEqual(plan["memo"], "")

    def test_oa_no_date_editable_with_notice(self):
        plan = self._build("[위임] PA-2T8084; 한송희(팀); OA; 001", "OA")
        self.assertEqual(self._days(plan), [self.RECEIVED, self.RECEIVED])
        self.assertTrue(plan["editable"])
        self.assertIn("통지일이 없어", plan["source_text"])

    def test_fr_no_date_uses_received_plus_3_months(self):
        """제목에 날짜가 없어도 FR 기한은 수신일 + 3개월 (예전에는 수신일 그대로 등록됐다)"""
        plan = self._build("[위임] PA-2T9001; 이여름(팀); 거절결정서; 001", "FR")
        self.assertEqual(self._days(plan), [self.RECEIVED, date(2027, 1, 16)])
        self.assertFalse(plan["editable"])
        self.assertIn("수신일", plan["source_text"])

    def test_fr_uses_receipt_date_and_reexam(self):
        cache = {"PA-2T9001|거절결정서": {"doc_type": "거절결정서", "receipt_date": "2026-10-06", "reexam": True}}
        plan = self._build(FR_SUBJECT, "FR", cache)
        self.assertEqual(self._days(plan), [date(2026, 10, 5), date(2027, 1, 6)])
        self.assertEqual(plan["items"][1][1], "PA-2T9001_거결불")
        self.assertEqual(plan["warning"], "")
        self.assertIn("접수일 2026-10-06", plan["source_text"])

    def test_fr_stale_case_ignored(self):
        """지난 거절결정의 기록(재심사 아님)은 쓰지 않는다 — 경고를 낸다"""
        cache = {"PA-2T9001|거절결정서": {"doc_type": "거절결정서", "receipt_date": "2026-05-01", "reexam": False,
                                       "applicant": "아우디 아게"}}
        plan = self._build(FR_SUBJECT, "FR", cache)
        self.assertEqual(plan["items"][1][1], "PA-2T9001_재심사_아우디")
        self.assertIn("_거결불", plan["warning"])
        self.assertEqual(self._days(plan)[1], date(2027, 1, 16))  # 수신일 기준

    def test_recent_case_bounds(self):
        received = date(2026, 10, 16)
        ok = lambda d: main.recent_case({"receipt_date": d}, received) is not None
        self.assertTrue(ok("2026-08-17"))   # 60일 전
        self.assertFalse(ok("2026-08-16"))  # 61일 전
        self.assertTrue(ok("2026-10-19"))   # 3일 뒤
        self.assertFalse(ok("2026-10-20"))
        self.assertIsNone(main.recent_case(None, received))
        self.assertIsNone(main.recent_case({"receipt_date": ""}, received))


class TestCalendarState(unittest.TestCase):
    """등록 기록 — 실패 일정 다시 시도, 처리한 위임 메일, 시작할 때 스캔"""

    ITEMS = [("Comments", "PA-2T8084_Comments", date(2026, 10, 1)),
             ("Action", "PA-2T8084_의견서", date(2027, 2, 8))]
    NAMES = ["Mail", "No_Instructions"]

    def setUp(self):
        import tempfile
        self.tmpdir = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmpdir.cleanup)
        for name, value in (("CALENDAR_STATE_FILE", "calendar_state.json"), ("CASE_CACHE_FILE", "case_cache.json")):
            patcher = patch.object(main, name, os.path.join(self.tmpdir.name, value))
            patcher.start()
            self.addCleanup(patcher.stop)
        patcher = patch.dict(main.config, {"naver_calendars": {"Mail": "Mail", "No_Instructions": "No_Instructions"},
                                           "applicant_abbr": {}})
        patcher.start()
        self.addCleanup(patcher.stop)

    def _selected(self):
        return [(0, self.ITEMS[0][2]), (1, self.ITEMS[1][2])]

    def test_failure_pending_then_retry_clears(self):
        naver = MagicMock()
        naver.add_all_day.side_effect = [True, naver_cal.NaverError("연결 끊김")]
        main.register_and_track("PA-2T8084", self.NAMES, self.ITEMS, self._selected(), naver, "메모")
        pending = main.load_calendar_state()["pending"]
        self.assertEqual([(p["subject"], p["day"], p["memo"]) for p in pending],
                         [("PA-2T8084_의견서", "2027-02-08", "메모")])

        naver = MagicMock()
        naver.add_all_day.return_value = True
        rows = main.retry_pending(naver)
        self.assertEqual(rows, [("PA-2T8084", "PA-2T8084_의견서", date(2027, 2, 8), main.STATUS_ADDED)])
        naver.add_all_day.assert_called_once_with("No_Instructions", "PA-2T8084_의견서", date(2027, 2, 8), "메모")
        self.assertEqual(main.load_calendar_state()["pending"], [])

    def test_retry_keeps_still_failing(self):
        naver = MagicMock()
        naver.add_all_day.side_effect = naver_cal.NaverError("연결 끊김")
        main.register_and_track("PA-2T8084", self.NAMES, self.ITEMS, self._selected(), naver)
        rows = main.retry_pending(naver)
        self.assertEqual(len(rows), 2)
        self.assertEqual(len(main.load_calendar_state()["pending"]), 2)

    def test_dialog_retry_success_clears_pending(self):
        naver = MagicMock()
        naver.add_all_day.side_effect = [True, naver_cal.NaverError("연결 끊김"), True]
        main.register_and_track("PA-2T8084", self.NAMES, self.ITEMS, self._selected(), naver)
        main.register_and_track("PA-2T8084", self.NAMES, self.ITEMS, [(1, self.ITEMS[1][2])], naver)
        self.assertEqual(main.load_calendar_state()["pending"], [])

    def test_done_pruned_after_30_days(self):
        main.mark_done("old", today=date(2026, 9, 1))
        main.mark_done("new", today=date(2026, 10, 16))
        self.assertEqual(set(main.load_calendar_state()["done"]), {"new"})

    def _outlook(self, *messages):
        outlook = MagicMock()
        outlook.GetDefaultFolder.return_value.Items.Restrict.return_value = list(messages)
        return outlook

    def _delegation(self, entry_id, subject=OA_SUBJECT):
        message = make_mock_message(subject, date.today().year, date.today().month, date.today().day)
        message.EntryID = entry_id
        return message

    def test_first_run_only_records(self):
        outlook = self._outlook(self._delegation("A"), self._delegation("B"))
        with patch.object(main, "is_calendar_delegation", return_value=True), \
                patch.object(main, "register_delegation") as reg:
            self.assertEqual(main.scan_delegations(outlook), [])
        reg.assert_not_called()
        self.assertEqual(set(main.load_calendar_state()["done"]), {"A", "B"})

    def test_scan_registers_only_new(self):
        main.mark_done("A")
        outlook = self._outlook(self._delegation("A"), self._delegation("B"))
        row = ("PA-2T8084", "PA-2T8084_Comments", date(2026, 10, 1), main.STATUS_ADDED)
        with patch.object(main, "is_calendar_delegation", return_value=True), \
                patch.object(main, "register_delegation", return_value=[row]) as reg:
            self.assertEqual(main.scan_delegations(outlook), [row])
        reg.assert_called_once()
        self.assertEqual(reg.call_args.args[0].EntryID, "B")
        self.assertEqual(reg.call_args.kwargs, {"always_show": False})

    def test_register_delegation_without_dialog_records_done(self):
        """시작할 때 스캔: 날짜를 고칠 필요 없는 메일은 창 없이 등록하고 처리 기록을 남긴다"""
        message = self._delegation("C", "[위임] PA-2T9001; 이여름(팀); 거절결정서 [서 2026-10-05]; 001")
        with patch.object(naver_cal, "Session") as session_cls, patch.object(main, "ScheduleDialog") as dlg:
            session_cls.return_value.add_all_day.return_value = True
            rows = main.register_delegation(message, always_show=False)
        dlg.assert_not_called()
        self.assertEqual([r[3] for r in rows], [main.STATUS_ADDED, main.STATUS_ADDED])
        self.assertIn("C", main.load_calendar_state()["done"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
