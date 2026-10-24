"""받은 지시 → .msg + NAS instructions (instruction_mail.py·db.record·main 의 고르기) 테스트.

메일은 모두 가공한 것이다(10-24 윈도 PC 의 실제 지시 메일과 같은 꼴 — 대표 메일함 FW:, 첫 줄 LEEDM:, 우리 서신 BCC 사본은
FW: 없이 LEEDM:…#F). 실제 Outlook·NAS 에는 닿지 않는다(test_calendar 가 %APPDATA%·pymysql 을 막는다).
"""

import os
import tempfile
import unittest
from unittest.mock import MagicMock, patch

# test_calendar 가 (없으면) win32com·PySide6 를 가짜로 끼우고 임시 데이터 폴더를 정한다 — main 보다 먼저 들인다
from test_calendar import MockCreationTime  # noqa: E402
from test_assignments import FakeConn  # noqa: E402

import db  # noqa: E402
import instruction_mail as IM  # noqa: E402
import main  # noqa: E402

SUBJECT = ("FW: Your Ref.: 2099P00181WOKR (6446); Lee Ref.: PA-2T0001/SYS; Korean Patent Application No. "
           "10-2099-7000001; RESPONSE DEADLINE: December 4, 2099")
BODY = """LEEDM:PA-2T0001
From: Berger, Anna <anna@client.example>
Sent: Thursday, October 22, 2099 11:03 PM
Subject: RE: Lee Ref.: PA-2T0001/SYS

Please proceed with filing a response within the set deadline.

From: Office IP & Law <info@office.example>
Sent: Tuesday, October 20, 2099 8:45 AM
Subject: Lee Ref.: PA-2T0001/SYS

LEEDM:PA-2T0001#F
We may request an extension of time if you need more time.
"""
BCC_SUBJECT = "Your Ref.: 2099P00181WOKR (6446); Lee Ref.: PA-2T0001/SYS; RESPONSE DEADLINE: December 4, 2099"
BCC_BODY = "LEEDM:PA-2T0001#F\nRe: Korean Patent Application No. 10-2099-7000001\nDear Ms. Berger:\n"


def _att(name, kind=1, cid="", hidden=False):
    a = MagicMock()
    a.FileName, a.Type = name, kind
    props = {IM.PR_ATTACH_CONTENT_ID: cid, IM.PR_ATTACHMENT_HIDDEN: hidden}
    a.PropertyAccessor.GetProperty.side_effect = lambda tag: props[tag]
    return a


def _message(subject=SUBJECT, body=BODY, entry_id="00000000ABCDEF0123456789", attachments=(), html=""):
    m = MagicMock()
    m.Subject, m.Body, m.HTMLBody, m.EntryID = subject, body, html, entry_id
    m.ReceivedTime = MockCreationTime(2099, 10, 22, 23, 17)
    m.Attachments = list(attachments)
    return m


class TestChoose(unittest.TestCase):
    def test_forward_with_leedm(self):
        self.assertEqual(IM.choose(SUBJECT, BODY), "PA-2T0001")
        self.assertEqual(IM.leedm_ref(BODY), ("PA-2T0001", ""))

    def test_bcc_copy_of_our_letter_is_not_chosen(self):
        """우리 서신 사본 — 같은 보낸 사람·같은 꼴 제목이지만 FW: 가 없고 LEEDM:…#F"""
        self.assertEqual(IM.leedm_ref(BCC_BODY), ("PA-2T0001", "#F"))
        self.assertEqual(IM.choose(BCC_SUBJECT, BCC_BODY), "")

    def test_forward_without_leedm_uses_subject_label(self):
        self.assertEqual(IM.choose("FW: Your Ref. X1; Our Ref.: PJA-2U1331/SYS", "hello"), "PJA-2U1331")
        self.assertEqual(IM.choose("FW: Office Action Due Date / JTI Ref. 1477/WO-KR / Lee Ref. PA-2T8837/SYS", ""),
                         "PA-2T8837")
        self.assertEqual(IM.choose("전달: PA-2T0002 instructions", ""), "PA-2T0002")
        self.assertEqual(IM.choose("FW: Bank Details", "no ref here"), "")            # Ref 가 없으면 고르지 않는다

    def test_leedm_without_fw_subject(self):
        self.assertEqual(IM.choose("Lee Ref.: PA-2T0003/SYS", "LEEDM:PA-2T0003\nFrom: a\nSent: b\n"), "PA-2T0003")

    def test_skipped(self):
        for subject in ("FW: [위임] PA-2T0001; 한송희(팀); 의견서 [서 2099-10-27]", "FW: 사건 담당자 지정 PA-2T0001",
                        "Automatic reply: Lee Ref.: PA-2T0001", "FW: Out of Office — PA-2T0001", "자동 회신: PA-2T0001"):
            self.assertEqual(IM.choose(subject, BODY), "", subject)

    def test_ref_boundaries(self):
        self.assertEqual(IM.choose("FW: see XPJA-2T0001 only", ""), "")                # 앞 글자 넷은 Ref 가 아니다
        self.assertEqual(IM.choose("FW: see PA-2T0001X only", ""), "PA-2T0001X")       # 꼬리 0~5자는 Ref 규칙대로

    def test_alert_only_for_extension_in_client_part(self):
        """연장 낱말이 인용된 우리 서신에만 있으면 알리지 않는다"""
        self.assertFalse(IM.wants_alert(SUBJECT, BODY))
        ext = BODY.replace("Please proceed with filing", "Please request an extension of two months for filing")
        self.assertTrue(IM.wants_alert(SUBJECT, ext))
        self.assertTrue(IM.wants_alert("FW: x", "기간연장 부탁드립니다"))                  # 머리가 없으면 본문 전체


class TestRowAndFiles(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def test_msg_file(self):
        p = IM.msg_file(self.tmp.name, MockCreationTime(2099, 10, 22, 23, 17), "PA-2T0001/SYS", "0000ABCDEF0123456789")
        self.assertEqual(p, os.path.join(self.tmp.name, "instructions", "2099-10",
                                         "20991022-2317_PA-2T0001SYS_23456789.msg"))

    def test_attachments_skip_inline_and_hidden(self):
        html = '<p>x</p><img src="cid:image001.png@01DD">'
        m = _message(html=html, attachments=[_att("image001.png", cid="image001.png@01DD"),
                                             _att("image002.png", hidden=True), _att("claims.docx"),
                                             _att("orig.msg", kind=IM.OL_EMBEDDED_ITEM)])
        names, embedded = IM.attachments(m)
        self.assertEqual(names, ["claims.docx", "orig.msg"])
        self.assertEqual([a.FileName for a in embedded], ["orig.msg"])

    def test_build_row(self):
        m = _message(body="b" * 70000, attachments=[_att("claims.docx")])
        row = IM.build_row(m, "PA-2T0001", r"G:\x.msg", "Lee International IP & Law")
        self.assertEqual(set(row), set(db.INSTRUCTION_COLUMNS))
        self.assertEqual((row["received_at"], row["ref_hint"], row["attachments"], row["source"]),
                         ("2099-10-22 23:17:00", "PA-2T0001", "claims.docx", "outlook"))
        self.assertEqual(len(row["body_text"]), IM.BODY_LIMIT)
        self.assertIsNone(row["message_id"])                    # 보안 경고를 확인하기 전까지 읽지 않는다
        self.assertIsNone(row["orig_text"])

    def test_handle_saves_msg_and_writes_once(self):
        conn = FakeConn()
        m = _message()
        got = IM.handle(m, "PA-2T0001", self.tmp.name, "Lee International IP & Law", connect_fn=lambda: conn)
        self.assertEqual(got["result"], {"written": 1, "pending": 0, "error": ""})
        self.assertFalse(got["alert"])
        path, kind = m.SaveAs.call_args.args
        self.assertEqual(kind, IM.OL_MSG_UNICODE)
        self.assertTrue(path.endswith(os.path.join("2099-10", "20991022-2317_PA-2T0001_23456789.msg")))
        sql, params = conn.executed[0]
        self.assertTrue(sql.startswith("INSERT IGNORE INTO instructions (mail_entry_id, message_id, received_at,"))
        self.assertEqual(params[0], m.EntryID)
        self.assertEqual(params[db.INSTRUCTION_COLUMNS.index("msg_path")], path)
        # 같은 메일은 다시 적지도 저장하지도 않는다
        m.SaveAs.reset_mock()
        self.assertIsNone(IM.handle(m, "PA-2T0001", self.tmp.name, "x", connect_fn=lambda: conn))
        m.SaveAs.assert_not_called()

    def test_handle_offline_queues_then_flushes(self):
        def offline():
            raise db.DbError("instructions 표가 없습니다")
        got = IM.handle(_message(), "PA-2T0001", self.tmp.name, "x", connect_fn=offline)
        self.assertEqual(got["result"]["pending"], 1)
        state = os.path.join(self.tmp.name, IM.STATE_NAME)
        self.assertEqual(db.load_state(state)["pending"][0]["ref_hint"], "PA-2T0001")
        conn = FakeConn()
        res = db.record([], state, connect_fn=lambda: conn, table=db.INSTRUCTIONS)    # 다음에 — 쌓인 것을 쓴다
        self.assertEqual(res["written"], 1)
        self.assertTrue(conn.executed[0][0].startswith("INSERT IGNORE INTO instructions"))

    def test_handle_msg_save_failure_still_writes(self):
        conn = FakeConn()
        m = _message()
        m.SaveAs.side_effect = OSError("drive offline")
        got = IM.handle(m, "PA-2T0001", self.tmp.name, "x", connect_fn=lambda: conn)
        self.assertEqual(got["result"]["written"], 1)
        self.assertIsNone(conn.executed[0][1][db.INSTRUCTION_COLUMNS.index("msg_path")])

    def test_assignments_still_written_the_same(self):
        """db.record 를 넓혀도 받은 위임은 예전 그대로"""
        conn = FakeConn()
        path = os.path.join(self.tmp.name, "assignment_state.json")
        db.record([{"mail_entry_id": "A", "ref": "PA-2T8084"}], path, connect_fn=lambda: conn)
        self.assertTrue(conn.executed[0][0].startswith("INSERT IGNORE INTO assignments (mail_entry_id, ref,"))


class TestMainChoosing(unittest.TestCase):
    def _outlook(self, inbox_id="INBOX"):
        outlook = MagicMock()
        outlook.GetDefaultFolder.return_value.EntryID = inbox_id
        return outlook

    def test_instruction_ref_checks_inbox_and_sender(self):
        m = _message()
        m.Parent.EntryID = "INBOX"
        with patch.object(main, "sender_matches", return_value=True) as sm:
            self.assertEqual(main.instruction_ref(m, self._outlook()), "PA-2T0001")
            self.assertEqual(sm.call_args.args[1], ["Lee International IP & Law"])
            m.Parent.EntryID = "01_Instructions"                       # 받은편지함만 본다
            self.assertEqual(main.instruction_ref(m, self._outlook()), "")
        m.Parent.EntryID = "INBOX"
        with patch.object(main, "sender_matches", return_value=False):   # 다른 보낸 사람
            self.assertEqual(main.instruction_ref(m, self._outlook()), "")

    def test_delegation_subject_never_reaches_sender_filter(self):
        m = _message(subject="[위임] PA-2T0001; 한송희(팀); 의견서 [서 2099-10-27]")
        with patch.object(main, "sender_matches") as sm:
            self.assertEqual(main.instruction_ref(m, self._outlook()), "")
            sm.assert_not_called()

    def test_record_instruction_tray_only_for_extension(self):
        tray = MagicMock()
        with patch.object(main.instruction_mail, "handle", return_value={"ref": "PA-2T0001", "alert": False}):
            main.record_instruction(_message(), "PA-2T0001", None, tray)
        tray.showMessage.assert_not_called()
        with patch.object(main.instruction_mail, "handle", return_value={"ref": "PA-2T0001", "alert": True}):
            main.record_instruction(_message(), "PA-2T0001", None, tray)
        self.assertIn("기연 지시일 수 있음 — PA-2T0001", tray.showMessage.call_args.args[0])

    def test_scan_instructions(self):
        old = _message(entry_id="OLD")
        old.ReceivedTime = MockCreationTime(2000, 1, 1)
        bcc = _message(subject=BCC_SUBJECT, body=BCC_BODY, entry_id="BCC")
        fresh = _message(entry_id="NEW")
        for m in (old, bcc, fresh):
            m.ReceivedTime = MockCreationTime(2000, 1, 1) if m is old else MockCreationTime(
                *(lambda d: (d.year, d.month, d.day))(main.date.today()))
        outlook = self._outlook()
        outlook.GetDefaultFolder.return_value.Items.Restrict.return_value = [old, bcc, fresh]
        with patch.object(main, "record_instruction", return_value={"alert": False}) as rec:
            main.scan_instructions(outlook)
        self.assertEqual([c.args[0].EntryID for c in rec.call_args_list], ["NEW"])
        flt = outlook.GetDefaultFolder.return_value.Items.Restrict.call_args.args[0]
        self.assertIn("'Lee International IP & Law'", flt)


if __name__ == "__main__":
    unittest.main(verbosity=2)
