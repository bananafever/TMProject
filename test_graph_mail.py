"""새 Outlook(Graph) 메일 읽기 — graph_mail.py 와 main·instruction_mail 의 Graph 갈래 테스트.

Graph·Entra 에는 닿지 않는다 — HTTP 는 가짜 urlopen 또는 가짜 Client, 로그인은 가짜 msal 앱이다.
예고 메일 표 HTML 은 실제 꼴(Outlook 이 보내는 <td><p class=MsoNormal>…</p></td>)을 본떠 가공한 것이다.
"""

import base64
import io
import json
import os
import tempfile
import unittest
import urllib.error
from pathlib import Path
from datetime import date, datetime, timezone
from unittest.mock import MagicMock, patch

# test_calendar 가 (없으면) win32com·PySide6 를 가짜로 끼우고 임시 데이터 폴더를 정한다 — main 보다 먼저 들인다
import test_calendar  # noqa: F401,E402
from test_assignments import FakeConn  # noqa: E402

import db  # noqa: E402
import graph_mail as G  # noqa: E402
import instruction_mail as IM  # noqa: E402
import main  # noqa: E402

PRENOTICE_HTML = """<html><head><style>p.MsoNormal{margin:0}</style></head><body>
<p class=MsoNormal>안녕하세요. 아래 사건의 담당자를 지정합니다.</p>
<table class=MsoNormalTable border=1>
<tr><td><p class=MsoNormal><b>접수일</b></p></td><td><p class=MsoNormal><b>출원번호</b></p></td>
<td><p class=MsoNormal>Ref</p></td><td><p class=MsoNormal>담당자</p></td><td><p class=MsoNormal>출원인</p></td>
<td><p class=MsoNormal>문서종류</p></td></tr>
<tr><td><p class=MsoNormal>20261012</p></td><td><p class=MsoNormal>10-2025-7000694</p></td>
<td><p class=MsoNormal>PA-2T8084</p></td><td><p class=MsoNormal>신용석</p></td>
<td><p class=MsoNormal>에어 프로덕츠 앤드&nbsp;케미칼스, 인코오포레이티드</p></td>
<td><p class=MsoNormal>의견제출통지서</p></td></tr>
<tr><td>20261013</td><td>10-2024-0001234</td><td>PA-2T9000</td><td>한송희</td><td>주식회사 가나</td><td>거절결정서</td></tr>
</table></body></html>"""


def _graph_id(raw: bytes) -> str:
    """translateExchangeIds 가 내는 꼴 — URL-safe base64, 끝 = 를 지우고 지운 개수를 숫자로."""
    b = base64.urlsafe_b64encode(raw).decode()
    n = b.count("=")
    return b.rstrip("=") + str(n)


def _row(mid, subject, sender, received="2099-10-22T14:17:00Z", **extra):
    row = {"id": mid, "subject": subject, "from": {"emailAddress": {"name": sender, "address": "x@y"}},
           "receivedDateTime": received, "createdDateTime": received, "hasAttachments": False}
    row.update(extra)
    return row


class FakeClient:
    """graph_mail.Client 자리 — 경로(쿼리 앞)별로 응답을 돌려주고, 부른 것을 적는다."""

    def __init__(self, routes=None, page_routes=None):
        self.routes, self.page_routes, self.calls = dict(routes or {}), dict(page_routes or {}), []

    def call(self, method, url, body=None, raw=False, headers=None):
        self.calls.append((method, url, body))
        if url == "/me/translateExchangeIds":
            return {"value": [{"sourceId": i, "targetId": _graph_id(("EID-" + i).encode())} for i in body["inputIds"]]}
        return self.routes[url.split("?")[0]]

    def get(self, url, **kw):
        return self.call("GET", url, **kw)

    def pages(self, url, headers=None):
        self.calls.append(("PAGES", url, None))
        pages = self.page_routes[url.split("?")[0]]
        if callable(pages):
            pages = pages()
        for page in pages:
            yield from page.get("value", [])


def _eid(mid):
    return ("EID-" + mid).encode().hex().upper()


class TestIds(unittest.TestCase):
    def test_entry_id_hex_all_paddings(self):
        for size in (30, 31, 32):                      # base64 패딩 0·2·1 개
            raw = bytes(range(1, size + 1))
            gid = _graph_id(raw)
            self.assertEqual(G.entry_id_hex(gid), raw.hex().upper(), gid)
        self.assertEqual({_graph_id(bytes(range(n)))[-1] for n in (30, 31, 32)}, {"0", "1", "2"})

    def test_entry_id_rejects_other_shapes(self):
        with self.assertRaises(G.GraphError):
            G.entry_id_hex("AAAA==")

    def test_local_time_is_naive_local(self):
        t = G.local_time("2099-10-22T14:17:00Z")
        self.assertIsNone(t.tzinfo)
        self.assertEqual(t, datetime(2099, 10, 22, 14, 17, tzinfo=timezone.utc).astimezone().replace(tzinfo=None))


class TestHtmlText(unittest.TestCase):
    def test_table_cells_become_tabs_and_prenotice_rows_parse(self):
        text = G.html_text(PRENOTICE_HTML)
        self.assertNotIn("margin", text)                                  # <style> 는 버린다
        self.assertIn("10-2025-7000694\tPA-2T8084\t신용석", text)
        rows = main.parse_prenotice_rows(text, "신용석")
        self.assertEqual(rows, [{"ref": "PA-2T8084", "app_no": "10-2025-7000694", "receipt_date": "2026-10-12",
                                 "doc_type": "의견제출통지서", "applicant": "에어 프로덕츠 앤드 케미칼스, 인코오포레이티드"}])
        self.assertEqual([r["ref"] for r in main.parse_prenotice_rows(text, "한송희")], ["PA-2T9000"])

    def test_paragraphs_keep_lines_for_leedm(self):
        html = "<div><p>LEEDM:PA-2T0001</p><p>From: Berger, Anna</p><br>Sent: x</div>"
        body = G.html_text(html)
        self.assertEqual(IM.leedm_ref(body), ("PA-2T0001", ""))
        self.assertEqual(body.splitlines()[:2], ["LEEDM:PA-2T0001", "From: Berger, Anna"])

    def test_text_body_untouched(self):
        self.assertEqual(G.body_text({"contentType": "text", "content": "a\tb"}), "a\tb")


class TestClient(unittest.TestCase):
    def _auth(self):
        auth = MagicMock()
        auth.token.side_effect = lambda force_refresh=False: "T2" if force_refresh else "T1"
        return auth

    def _resp(self, payload):
        r = MagicMock()
        r.read.return_value = payload if isinstance(payload, bytes) else json.dumps(payload).encode()
        r.__enter__.return_value = r
        return r

    def _http_error(self, code, retry_after=None):
        headers = {"Retry-After": retry_after} if retry_after else {}
        return urllib.error.HTTPError("u", code, "x", headers, io.BytesIO(b'{"error":{"message":"busy"}}'))

    def test_retry_after_then_success(self):
        seen, sleeps = [], []
        answers = [self._http_error(429, "3"), self._resp({"ok": 1})]

        def urlopen(req, timeout):
            seen.append(req)
            a = answers.pop(0)
            if isinstance(a, Exception):
                raise a
            return a
        c = G.Client(self._auth(), urlopen=urlopen, sleep=sleeps.append)
        self.assertEqual(c.get("/me"), {"ok": 1})
        self.assertEqual(sleeps, [3])
        self.assertEqual(seen[0].full_url, G.GRAPH + "/me")
        self.assertEqual(seen[0].get_header("Prefer"), 'IdType="ImmutableId"')
        self.assertEqual(seen[0].get_header("Authorization"), "Bearer T1")

    def test_401_refreshes_token_once(self):
        answers = [self._http_error(401), self._resp(b"MIME")]
        seen = []

        def urlopen(req, timeout):
            seen.append(req.get_header("Authorization"))
            a = answers.pop(0)
            if isinstance(a, Exception):
                raise a
            return a
        c = G.Client(self._auth(), urlopen=urlopen, sleep=lambda s: None)
        self.assertEqual(c.get("/me/messages/1/$value", raw=True), b"MIME")
        self.assertEqual(seen, ["Bearer T1", "Bearer T2"])

    def test_other_errors_raise(self):
        def urlopen(req, timeout):
            raise self._http_error(403)
        with self.assertRaisesRegex(G.GraphError, "403 busy"):
            G.Client(self._auth(), urlopen=urlopen, sleep=lambda s: None).get("/me")

    def test_pages_follow_next_link(self):
        answers = [self._resp({"value": [1, 2], "@odata.nextLink": G.GRAPH + "/next"}), self._resp({"value": [3]})]
        urls = []

        def urlopen(req, timeout):
            urls.append(req.full_url)
            return answers.pop(0)
        got = list(G.Client(self._auth(), urlopen=urlopen, sleep=lambda s: None).pages("/a"))
        self.assertEqual(got, [1, 2, 3])
        self.assertEqual(urls, [G.GRAPH + "/a", G.GRAPH + "/next"])


class TestMailbox(unittest.TestCase):
    INBOX = "/me/mailFolders/inbox/messages"

    def test_inbox_messages_filters_and_translates(self):
        client = FakeClient(page_routes={self.INBOX: [{"value": [
            _row("A", "[위임] PA-2T0001; 한송희(팀); 의견서", "DoNotReply"),
            _row("B", "[위임] PA-2T0002; 한송희(팀)", "Someone Else"),
            _row("C", "주간 회의", "DoNotReply")]}]})
        box = G.GraphMailbox(client)
        got = box.inbox_messages(["DoNotReply"], "위임", since=date(2099, 10, 8))
        self.assertEqual([(m.id, m.EntryID) for m in got], [("A", _eid("A"))])
        url = client.calls[0][1]
        self.assertIn("receivedDateTime+ge+", url)
        translate = [c for c in client.calls if c[1] == "/me/translateExchangeIds"][0][2]
        self.assertEqual((translate["inputIds"], translate["sourceIdType"], translate["targetIdType"]),
                         (["A"], "restImmutableEntryId", "entryId"))

    def test_no_match_no_translate_call(self):
        client = FakeClient(page_routes={self.INBOX: [{"value": [_row("C", "x", "Other")]}]})
        self.assertEqual(G.GraphMailbox(client).inbox_messages(["DoNotReply"]), [])
        self.assertFalse([c for c in client.calls if c[1] == "/me/translateExchangeIds"])

    def test_untranslated_id_is_an_error(self):
        client = FakeClient()
        client.call = lambda *a, **k: {"value": []}
        with self.assertRaises(G.GraphError):
            G.GraphMailbox(client).translate(["A"])

    def test_watch_sees_rule_moved_mail_not_sent_or_old(self):
        """클래식 OnNewMailEx 처럼 — 규칙이 (A)_NewTask 로 옮긴 새 메일도 잡고, 보낸 편지함·옛 메일은 뺀다."""
        folders = {f"/me/mailFolders/{n}": {"id": n.upper()} for n in G.NOT_RECEIVED_FOLDERS}
        folders["/me/mailFolders/inbox"] = {"id": "INBOX"}
        now = datetime(2099, 10, 22, 14, 0, tzinfo=timezone.utc)
        rounds = [
            [_row("EARLY", "x", "S", "2099-10-22T13:58:00Z", parentFolderId="INBOX")],          # 켜기 전 — 처리 안 함
            [_row("NEW2", "b", "S", "2099-10-22T14:03:00Z", parentFolderId="NEWTASK"),
             _row("SENT", "re", "me", "2099-10-22T14:02:30Z", parentFolderId="SENTITEMS"),
             _row("NEW1", "a", "S", "2099-10-22T14:02:00Z", parentFolderId="INBOX"),
             _row("EARLY", "x", "S", "2099-10-22T13:58:00Z", parentFolderId="INBOX")],
            [_row("NEW2", "b", "S", "2099-10-22T14:03:00Z", parentFolderId="NEWTASK")],
        ]
        urls = []
        client = FakeClient(routes=folders)
        client.pages = lambda url, headers=None: (urls.append(url), iter(rounds.pop(0)))[1]
        box = G.GraphMailbox(client)
        with self.assertRaises(G.GraphError):
            box.poll()                                                            # start_watch 전
        box.start_watch(now)
        self.assertTrue(urls[0].startswith("/me/messages?"))
        self.assertIn("receivedDateTime+ge+2099-10-22T13%3A55%3A00Z", urls[0])
        got = box.poll()
        self.assertEqual([m.id for m in got], ["NEW1", "NEW2"])                    # 오래된 순, 새로 받은 것만
        self.assertEqual(got[1].EntryID, _eid("NEW2"))
        self.assertEqual(box.poll(), [])                                           # 이미 낸 메일은 다시 안 낸다
        self.assertIn("2099-10-22T13%3A58%3A00Z", urls[2])                          # 본 것 가운데 가장 늦은 시각 − 5분부터
        self.assertTrue(box.is_inbox(got[0]))
        self.assertFalse(box.is_inbox(got[1]))

    def test_open_mailbox_needs_ids(self):
        with self.assertRaises(G.LoginRequired):
            G.open_mailbox({"graph_client_id": "", "graph_tenant_id": ""})

    def test_open_mailbox_without_saved_login(self):
        app = MagicMock()
        app.get_accounts.return_value = []
        with self.assertRaises(G.LoginRequired):
            G.open_mailbox({"graph_client_id": "c", "graph_tenant_id": "t"},
                           auth_factory=lambda c, t: G.Auth(c, t, app=app))
        app.get_accounts.return_value = [{"username": "me@x"}]
        app.acquire_token_silent.return_value = {"access_token": "T"}
        box = G.open_mailbox({"graph_client_id": "c", "graph_tenant_id": "t"},
                             auth_factory=lambda c, t: G.Auth(c, t, app=app))
        self.assertIsInstance(box, G.GraphMailbox)
        self.assertEqual(app.acquire_token_silent.call_args.args[0], ["Mail.Read", "User.Read"])


class TestMessage(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def _client(self):
        mime = (b"Subject: Original from client\r\nDate: Tue, 20 Oct 2099 08:45:00 +0000\r\n"
                b"Content-Type: text/plain; charset=utf-8\r\n\r\nPlease extend.\r\n")
        return FakeClient(
            routes={"/me/messages/M1": {"body": {"contentType": "html",
                                                 "content": '<p>LEEDM:PA-2T0001</p><img src="cid:image001.png@01DD">'}},
                    "/me/messages/M1/$value": b"RAW-MIME",
                    "/me/messages/M1/attachments/ITEM/$value": mime},
            page_routes={"/me/messages/M1/attachments": [{"value": [
                {"@odata.type": "#microsoft.graph.fileAttachment", "id": "IMG", "name": "image001.png",
                 "contentId": "image001.png@01DD", "isInline": True, "contentBytes": base64.b64encode(b"png").decode()},
                {"@odata.type": "#microsoft.graph.fileAttachment", "id": "DOC", "name": "claims.docx",
                 "contentBytes": base64.b64encode(b"docx").decode()},
                {"@odata.type": "#microsoft.graph.itemAttachment", "id": "ITEM", "name": "orig"}]}]})

    def test_lazy_body_attachments_and_save(self):
        client = self._client()
        m = G.GraphMessage(client, _row("M1", "FW: Lee Ref.: PA-2T0001", "Lee International IP & Law",
                                        hasAttachments=True), "ENTRY")
        self.assertEqual(client.calls, [])                                  # 본문·첨부는 읽을 때
        self.assertEqual(m.Body, "LEEDM:PA-2T0001")
        self.assertEqual(m.Attachments.Count, 3)
        names, embedded = IM.attachments(m)                                 # 인라인 그림은 빼고, 붙은 메일은 따로
        self.assertEqual(names, ["claims.docx", "orig"])
        self.assertEqual([a.FileName for a in embedded], ["orig"])
        path = os.path.join(self.tmp.name, "c.docx")
        m.Attachments[1].SaveAsFile(path)
        self.assertEqual(Path(path).read_bytes(), b"docx")
        m.SaveAs(os.path.join(self.tmp.name, "x.eml"), IM.OL_MSG_UNICODE)
        self.assertEqual(Path(os.path.join(self.tmp.name, "x.eml")).read_bytes(), b"RAW-MIME")
        text, head = IM.read_embedded(embedded[0], G.GraphMailbox(client))
        self.assertEqual(text.strip(), "Please extend.")
        self.assertIn("제목: Original from client", head)
        self.assertIn("보낸 날짜: ", head)

    def test_no_attachments_no_request(self):
        client = FakeClient()
        m = G.GraphMessage(client, _row("M2", "x", "y"))
        self.assertEqual(m.Attachments.Count, 0)
        self.assertEqual(client.calls, [])

    def test_handle_saves_eml_and_marks_source(self):
        client = self._client()
        m = G.GraphMessage(client, _row("M1", "FW: Lee Ref.: PA-2T0001", "Lee International IP & Law",
                                        hasAttachments=True), "00000000ABCDEF0123456789")
        conn = FakeConn()
        got = IM.handle(m, "PA-2T0001", self.tmp.name, "Lee International IP & Law", session=G.GraphMailbox(client),
                        connect_fn=lambda: conn, source="graph")
        self.assertEqual(got["result"]["written"], 1)
        params = conn.executed[0][1]
        path = params[db.INSTRUCTION_COLUMNS.index("msg_path")]
        self.assertTrue(path.endswith("_PA-2T0001_23456789.eml"), path)
        self.assertEqual(Path(path).read_bytes(), b"RAW-MIME")
        self.assertEqual(params[db.INSTRUCTION_COLUMNS.index("source")], "graph")
        self.assertEqual(params[db.INSTRUCTION_COLUMNS.index("orig_head")].splitlines()[0], "제목: Original from client")


class TestMainGraph(unittest.TestCase):
    def _message(self, subject, sender):
        return G.GraphMessage(FakeClient(), _row("M", subject, sender), "E")

    def test_sender_matches_reads_name_directly(self):
        m = self._message("[위임] PA-2T0001; 한송희(팀)", "DoNotReply")
        self.assertTrue(main.sender_matches(m, ["DoNotReply"]))
        self.assertFalse(main.sender_matches(m, ["Someone"]))

    def test_instruction_only_from_inbox(self):
        box = G.GraphMailbox(FakeClient(routes={"/me/mailFolders/inbox": {"id": "INBOX"}}))
        inbox = G.GraphMessage(FakeClient(), _row("M", "x", "y", parentFolderId="INBOX"), "E")
        moved = G.GraphMessage(FakeClient(), _row("M", "x", "y", parentFolderId="NEWTASK"), "E")
        self.assertTrue(main.in_inbox(inbox, box))
        self.assertFalse(main.in_inbox(moved, box))

    def test_inbox_items_routes_to_graph(self):
        box = G.GraphMailbox(FakeClient())
        with patch.object(box, "inbox_messages", return_value=["m"]) as im:
            self.assertEqual(main.inbox_items(box, ["DoNotReply"], "위임", date(2099, 1, 1)), ["m"])
        im.assert_called_once_with(["DoNotReply"], "위임", date(2099, 1, 1))

    def test_inbox_items_classic_filter(self):
        outlook = MagicMock()
        main.inbox_items(outlook, ["Chang-Kyu KIM [김창규]"], main.PRENOTICE_SUBJECT)
        flt = outlook.GetDefaultFolder.return_value.Items.Restrict.call_args.args[0]
        self.assertEqual(flt, main.build_sender_filter(["Chang-Kyu KIM [김창규]"])
                         + " AND \"urn:schemas:httpmail:subject\" LIKE '%사건 담당자 지정%'")

    def test_record_instruction_passes_graph_source(self):
        m = self._message("FW: x", "Lee International IP & Law")
        with patch.object(main.instruction_mail, "handle", return_value=None) as h:
            main.record_instruction(m, "PA-2T0001", None)
        self.assertEqual(h.call_args.kwargs["source"], "graph")

    def test_scan_instructions_on_graph(self):
        today = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        client = FakeClient(page_routes={"/me/mailFolders/inbox/messages": [{"value": [
            _row("I1", "FW: Lee Ref.: PA-2T0001", "Lee International IP & Law", today,
                 body={"contentType": "text", "content": "LEEDM:PA-2T0001\nFrom: a"})]}]})
        box = G.GraphMailbox(client)
        with patch.object(main, "record_instruction", return_value={"alert": False}) as rec:
            main.scan_instructions(box)
        self.assertEqual([c.args[0].EntryID for c in rec.call_args_list], [_eid("I1")])
        self.assertIs(rec.call_args.args[2], box)

    def test_open_mailbox_graph_login_required(self):
        with patch.dict(main.config, {"mail_backend": "graph", "graph_client_id": "", "graph_tenant_id": ""}):
            box, err = main.open_mailbox()
        self.assertIsNone(box)
        self.assertIn("Microsoft 로그인", err)

    def test_catch_up_without_mailbox_skips_scan(self):
        with patch.object(main, "retry_pending", return_value=[]), \
                patch.object(main.naver_cal, "Session"), \
                patch.object(main, "scan_delegations") as scan:
            self.assertEqual(main.catch_up_calendar(None), [])
        scan.assert_not_called()


if __name__ == "__main__":
    unittest.main()
