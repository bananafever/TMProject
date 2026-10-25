"""받은편지함을 Microsoft Graph 로 읽는다 — 새 Outlook(olk.exe)은 COM·MAPI 를 지원하지 않는다. Qt·네이버를 모른다.

settings.json `mail_backend: "graph"` 일 때 main 이 클래식 Outlook COM 대신 쓴다. 메일을 보내지·옮기지·고치지 않는다(읽기만).

- 앱: 회사 Entra 의 「TaskManager」(사용자 소유, 단일 테넌트, public client). `graph_client_id`·`graph_tenant_id` 는
  settings.json 에만 둔다. 리디렉션 URI `http://localhost` 가 있어야 MSAL 대화형 로그인이 돌아온다(10-25 추가).
- 권한: `Mail.Read` + `User.Read` — `translateExchangeIds` 가 회사 계정 위임에서 `User.ReadBasic.All` 이상을 요구한다.
- 토큰 캐시는 msal-extensions(Windows DPAPI · macOS Keychain)로 `%APPDATA%\\TMProject` / `~/.tmproject` 에 둔다(naver_cal 과 같은 자리).
- `GraphMessage` 는 main·instruction_mail 이 쓰는 COM MailItem 속성 이름(`Subject`, `Body`, `CreationTime`, `EntryID`,
  `Attachments`, `SaveAs` …)을 그대로 내준다 — 아래쪽 처리 코드는 메일이 어디서 왔는지 모른다.
- `EntryID` 는 클래식 Outlook 과 같은 값(대문자 hex)으로 바꿔 둔다 — NAS `mail_entry_id` 와 상태 파일(calendar_state·
  assignment_state·instruction_state)이 백엔드를 바꿔도 이어져 같은 메일을 두 번 등록하지 않는다.
"""

import base64
import email
import email.policy
import email.utils
import json
import logging
import os
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone
from html.parser import HTMLParser
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.Read", "User.Read"]
TIMEOUT = 20
RETRIES = 4
TRANSLATE_BATCH = 1000                                  # translateExchangeIds 한 번에 넣을 수 있는 수
LIST_FIELDS = "id,subject,from,receivedDateTime,createdDateTime,hasAttachments"
NEW_MAIL_SLACK = timedelta(minutes=5)                   # 감시를 켜기 조금 전에 받은 메일도 새 메일로 본다
OL_EMBEDDED_ITEM = 5                                    # instruction_mail 과 같은 COM 값 — 항목으로 붙은 메일
_PROP = "http://schemas.microsoft.com/mapi/proptag/"
PR_ATTACH_CONTENT_ID = _PROP + "0x3712001F"
PR_ATTACHMENT_HIDDEN = _PROP + "0x7FFE000B"


class GraphError(Exception):
    pass


class LoginRequired(GraphError):
    """저장된 토큰이 없거나 만료 — 설정 창 「Microsoft 로그인」 이 필요하다."""


# ── 로그인 ───────────────────────────────────────────────────────────

def cache_path() -> Path:
    base = os.getenv("APPDATA")
    return (Path(base) / "TMProject" if base else Path.home() / ".tmproject") / "graph_token.bin"


def _token_cache(path: Path):
    """암호화 토큰 캐시 — 못 만들면(Keychain·DPAPI 없음) 메모리 캐시: 켤 때마다 로그인해야 한다."""
    import msal
    try:
        from msal_extensions import PersistedTokenCache, build_encrypted_persistence
        path.parent.mkdir(parents=True, exist_ok=True)
        return PersistedTokenCache(build_encrypted_persistence(str(path)))
    except Exception as e:                              # noqa: BLE001
        logging.error(f"Graph 토큰 캐시를 암호화해 둘 수 없습니다 — 메모리에만 둡니다: {e}")
        return msal.SerializableTokenCache()


class Auth:
    def __init__(self, client_id: str, tenant_id: str, app=None):
        if not client_id or not tenant_id:
            raise LoginRequired("graph_client_id·graph_tenant_id 가 설정에 없습니다.")
        if app is None:
            import msal
            app = msal.PublicClientApplication(client_id, authority=f"https://login.microsoftonline.com/{tenant_id}",
                                               token_cache=_token_cache(cache_path()))
        self.app = app

    def token(self, force_refresh=False) -> str:
        for account in self.app.get_accounts():
            got = self.app.acquire_token_silent(SCOPES, account=account, force_refresh=force_refresh)
            if got and "access_token" in got:
                return got["access_token"]
        raise LoginRequired("Microsoft 로그인이 필요합니다.")

    def login(self, timeout=300) -> str:
        """시스템 브라우저로 로그인(MFA 포함). 돌아온 계정 이름. 실패하면 GraphError."""
        got = self.app.acquire_token_interactive(SCOPES, prompt="select_account", timeout=timeout)
        if not got or "access_token" not in got:
            raise GraphError((got or {}).get("error_description") or "Microsoft 로그인에 실패했습니다.")
        return (got.get("id_token_claims") or {}).get("preferred_username", "")

    def signed_in(self) -> str:
        accounts = self.app.get_accounts()
        return accounts[0].get("username", "") if accounts else ""


# ── HTTP ─────────────────────────────────────────────────────────────

class Client:
    """Graph 요청 — 429·503·504 는 Retry-After 만큼 쉬고 다시, 401 은 토큰을 새로 받아 한 번 더."""

    def __init__(self, auth, urlopen=urllib.request.urlopen, sleep=time.sleep):
        self.auth, self._urlopen, self._sleep = auth, urlopen, sleep

    def call(self, method: str, url: str, body=None, raw=False, headers=None):
        if not url.startswith("http"):
            url = GRAPH + url
        data = json.dumps(body).encode("utf-8") if body is not None else None
        refreshed = False
        for attempt in range(RETRIES + 1):
            req = urllib.request.Request(url, data=data, method=method)
            req.add_header("Authorization", "Bearer " + self.auth.token(force_refresh=refreshed))
            req.add_header("Prefer", 'IdType="ImmutableId"')          # 폴더를 옮겨도 id 가 안 바뀐다
            if data is not None:
                req.add_header("Content-Type", "application/json")
            for k, v in (headers or {}).items():
                req.add_header(k, v)
            try:
                with self._urlopen(req, timeout=TIMEOUT) as resp:
                    payload = resp.read()
                return payload if raw else json.loads(payload.decode("utf-8") or "{}")
            except urllib.error.HTTPError as e:
                if e.code == 401 and not refreshed:
                    refreshed = True
                    continue
                if e.code in (429, 503, 504) and attempt < RETRIES:
                    wait = e.headers.get("Retry-After") if e.headers else None
                    self._sleep(min(int(wait) if wait and wait.isdigit() else 2 ** attempt, 60))
                    continue
                detail = ""
                try:
                    detail = json.loads(e.read().decode("utf-8")).get("error", {}).get("message", "")
                except Exception:                       # noqa: BLE001
                    pass
                raise GraphError(f"Graph {method} {url.split('?')[0]} → HTTP {e.code} {detail}".strip()) from e
            except urllib.error.URLError as e:
                if attempt < RETRIES:
                    self._sleep(2 ** attempt)
                    continue
                raise GraphError(f"Graph 에 닿지 않습니다: {e.reason}") from e
        raise GraphError(f"Graph {method} {url.split('?')[0]} — 다시 시도해도 실패")

    def get(self, url, **kw):
        return self.call("GET", url, **kw)

    def pages(self, url: str, headers=None, tail: Optional[dict] = None) -> Iterable[dict]:
        """`value` 를 차례로 — `@odata.nextLink` 를 따라간다. 마지막 응답(deltaLink 포함)은 tail 에(주면).
        시작 스캔(GUI 스레드)과 새 메일 감시(뒤 스레드)가 함께 부르므로 공유 속성에 두지 않는다."""
        while url:
            page = self.get(url, headers=headers)
            if tail is not None:
                tail.clear()
                tail.update(page)
            yield from page.get("value", [])
            url = page.get("@odata.nextLink")


# ── 본문 HTML → 텍스트 ───────────────────────────────────────────────

_BLOCK = {"p", "div", "br", "tr", "li", "h1", "h2", "h3", "h4", "h5", "h6", "table", "blockquote", "pre", "hr"}
_SKIP = {"script", "style", "head", "title"}


class _Text(HTMLParser):
    """줄바꿈은 미뤄 두었다가 다음 글자 앞에 한 번만 — 칸 안의 <p> 가 행을 끊지 않고, 문단 사이에 빈 줄이 겹치지 않는다."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out, self.skip, self.cells, self.pending = [], 0, [0], 0

    def _row_break(self):
        if self.out and not self.out[-1].endswith("\n"):
            self.out.append("\n")
        self.pending = 0

    def handle_starttag(self, tag, attrs):
        if tag in _SKIP:
            self.skip += 1
        elif tag == "tr":
            self.cells.append(0)
            self._row_break()
        elif tag in ("td", "th"):
            if self.cells[-1]:
                self.out.append("\t")
            self.cells[-1] += 1
            self.pending = 0
        elif tag == "br":
            self.pending = min(self.pending + 1, 2)
        elif tag in _BLOCK:
            self.pending = max(self.pending, 1)

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)

    def handle_endtag(self, tag):
        if tag in _SKIP:
            self.skip = max(0, self.skip - 1)
        elif tag == "tr":
            if len(self.cells) > 1:
                self.cells.pop()
            self._row_break()
        elif tag in ("td", "th"):
            self.pending = 0
        elif tag in _BLOCK:
            self.pending = max(self.pending, 1)

    def handle_data(self, data):
        if self.skip:
            return
        text = re.sub(r"\s+", " ", data.replace("\xa0", " "))
        if not text.strip():
            if self.out and not self.out[-1].endswith((" ", "\t", "\n")):
                self.out.append(" ")
            return
        if self.pending and self.out and not self.out[-1].endswith("\t"):
            self.out.append("\n" * self.pending)
        self.pending = 0
        self.out.append(text)


def html_text(html: str) -> str:
    """HTML 본문 → 텍스트. 표 칸은 탭, 행·문단은 줄바꿈 — `main.parse_prenotice_rows` 가 탭으로 이은 행을 받는다.
    (Graph 의 text 변환은 표 모양이 보장되지 않아 쓰지 않는다.)"""
    p = _Text()
    p.feed(html or "")
    p.close()
    lines = [re.sub(r" *\t *", "\t", ln).strip(" ") for ln in "".join(p.out).split("\n")]
    text, blank = [], 0
    for ln in lines:
        blank = blank + 1 if not ln.strip() else 0
        if blank <= 1:
            text.append(ln)
    return "\n".join(text).strip("\n")


def body_text(body: dict) -> str:
    body = body or {}
    content = body.get("content") or ""
    return html_text(content) if (body.get("contentType") or "").lower() == "html" else content


# ── 시각·EntryID ─────────────────────────────────────────────────────

def local_time(value: str) -> Optional[datetime]:
    """Graph UTC ISO 시각 → 이 PC 시각(naive) — COM 의 CreationTime·ReceivedTime 과 같은 꼴."""
    if not value:
        return None
    stamp = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if stamp.tzinfo is None:
        stamp = stamp.replace(tzinfo=timezone.utc)
    return stamp.astimezone().replace(tzinfo=None)


def entry_id_hex(value: str) -> str:
    """translateExchangeIds 의 `entryId` → 클래식 Outlook `MailItem.EntryID`(대문자 hex).

    Graph 의 이진 id 는 URL-safe base64(`-`·`_`)에서 끝 `=` 를 지우고, 지운 개수(0·1·2)를 숫자 한 자리로 붙인 꼴이다.
    """
    if not value or value[-1] not in "012":
        raise GraphError(f"entryId 꼴이 아닙니다: {value[:20]}…")
    pad = int(value[-1])
    return base64.urlsafe_b64decode(value[:-1] + "=" * pad).hex().upper()


# ── 메시지 어댑터 ─────────────────────────────────────────────────────

class _Accessor:
    """COM `Attachment.PropertyAccessor.GetProperty` 흉내 — instruction_mail.attachments 가 그대로 돈다."""

    def __init__(self, props):
        self.props = props

    def GetProperty(self, tag):
        return self.props.get(tag)


class GraphAttachment:
    def __init__(self, client, message_id: str, data: dict):
        self._client, self._message_id, self.data = client, message_id, data
        self.FileName = data.get("name") or "attachment"
        self.is_item = data.get("@odata.type") == "#microsoft.graph.itemAttachment"
        self.Type = OL_EMBEDDED_ITEM if self.is_item else 1
        self.PropertyAccessor = _Accessor({PR_ATTACH_CONTENT_ID: data.get("contentId") or "",
                                           PR_ATTACHMENT_HIDDEN: False})

    def content(self) -> bytes:
        """파일 첨부는 받은 바이트, 항목 첨부(붙은 메일)는 MIME."""
        if not self.is_item and self.data.get("contentBytes"):
            return base64.b64decode(self.data["contentBytes"])
        return self._client.get(f"/me/messages/{self._message_id}/attachments/{self.data['id']}/$value", raw=True)

    def SaveAsFile(self, path):
        Path(path).write_bytes(self.content())


class GraphAttachments(list):
    @property
    def Count(self):
        return len(self)


class GraphMessage:
    """COM MailItem 과 같은 이름의 속성 — 본문·첨부는 처음 읽을 때 받아 온다."""

    def __init__(self, client, data: dict, entry_id: str = ""):
        self._client, self.data = client, data
        self.id = data["id"]
        self.Subject = data.get("subject") or ""
        self.SenderName = ((data.get("from") or {}).get("emailAddress") or {}).get("name") or ""
        self.ReceivedTime = local_time(data.get("receivedDateTime"))
        self.CreationTime = local_time(data.get("createdDateTime")) or self.ReceivedTime
        self.EntryID = entry_id
        self._body, self._attachments = data.get("body"), None

    def _load_body(self):
        if self._body is None:
            self._body = self._client.get(f"/me/messages/{self.id}?$select=body").get("body") or {}
        return self._body

    @property
    def Body(self) -> str:
        return body_text(self._load_body())

    @property
    def HTMLBody(self) -> str:
        body = self._load_body()
        return body.get("content", "") if (body.get("contentType") or "").lower() == "html" else ""

    @property
    def Attachments(self) -> GraphAttachments:
        if self._attachments is None:
            items = []
            if self.data.get("hasAttachments", True):
                items = list(self._client.pages(f"/me/messages/{self.id}/attachments"))
            self._attachments = GraphAttachments(GraphAttachment(self._client, self.id, a) for a in items)
        return self._attachments

    def SaveAs(self, path, *_ignored):
        """메일 원본(MIME)을 `.eml` 로 — ChekPat 은 .eml 을 읽는다."""
        Path(path).write_bytes(self._client.get(f"/me/messages/{self.id}/$value", raw=True))


class SharedItem:
    """COM `Namespace.OpenSharedItem` 흉내 — 붙은 원 메일(MIME)의 제목·보낸 날짜·본문."""

    def __init__(self, raw: bytes):
        msg = email.message_from_bytes(raw, policy=email.policy.default)
        self.Subject = str(msg.get("subject", "") or "")
        try:
            self.SentOn = email.utils.parsedate_to_datetime(msg["date"]).astimezone().replace(tzinfo=None)
        except Exception:                               # noqa: BLE001
            self.SentOn = None
        part = msg.get_body(preferencelist=("plain", "html"))
        text = part.get_content() if part is not None else ""
        self.Body = html_text(text) if part is not None and part.get_content_type() == "text/html" else text

    def Close(self, *_ignored):
        pass


# ── 메일함 ───────────────────────────────────────────────────────────

def _matches(message: GraphMessage, senders, subject_like) -> bool:
    if senders and message.SenderName not in senders:
        return False
    return not subject_like or subject_like in message.Subject


class GraphMailbox:
    """main 이 클래식 Outlook 네임스페이스 대신 받는 메일함. 받은편지함만 본다."""

    source = "graph"

    def __init__(self, client: Client):
        self.client = client
        self._delta = None
        self._seen = set()
        self._watch_since = None

    def translate(self, ids: List[str]) -> Dict[str, str]:
        """Graph immutable id → 클래식 EntryID(hex)."""
        out = {}
        for i in range(0, len(ids), TRANSLATE_BATCH):
            got = self.client.call("POST", "/me/translateExchangeIds", body={
                "inputIds": ids[i:i + TRANSLATE_BATCH], "sourceIdType": "restImmutableEntryId", "targetIdType": "entryId"})
            for row in got.get("value", []):
                if row.get("targetId"):
                    out[row["sourceId"]] = entry_id_hex(row["targetId"])
        missing = [x for x in ids if x not in out]
        if missing:
            raise GraphError(f"EntryID 로 바꾸지 못한 메일 {len(missing)}통")
        return out

    def _wrap(self, rows: List[dict], senders=None, subject_like=None) -> List[GraphMessage]:
        messages = [m for m in (GraphMessage(self.client, r) for r in rows) if _matches(m, senders, subject_like)]
        if messages:
            ids = self.translate([m.id for m in messages])
            for m in messages:
                m.EntryID = ids[m.id]
        return messages

    def inbox_messages(self, senders=None, subject_like=None, since=None) -> List[GraphMessage]:
        """받은편지함에서 senders(보낸 사람 표시 이름) 중 하나가 보낸 메일 — 제목에 subject_like 가 들어간 것만(주면).
        보낸 사람·제목은 받아 온 뒤 거른다 — $filter 의 from/name 조합은 믿지 않는다."""
        query = {"$select": LIST_FIELDS, "$top": "500", "$orderby": "receivedDateTime desc"}
        if since is not None:
            start = datetime.combine(since, datetime.min.time()).astimezone(timezone.utc)
            query["$filter"] = "receivedDateTime ge " + start.strftime("%Y-%m-%dT%H:%M:%SZ")
        url = "/me/mailFolders/inbox/messages?" + urllib.parse.urlencode(query)
        return self._wrap(list(self.client.pages(url)), senders, subject_like)

    def start_watch(self, now: Optional[datetime] = None):
        """새 메일 감시 시작 — 지금 받은편지함 상태를 delta 로 받아 두고(처리하지 않음), 이후 poll 이 새로 온 것만 낸다."""
        now = now or datetime.now(timezone.utc)
        self._watch_since = now - NEW_MAIL_SLACK
        query = {"$select": LIST_FIELDS, "$filter": "receivedDateTime ge " +
                 (now - timedelta(days=1)).strftime("%Y-%m-%dT%H:%M:%SZ")}
        url = "/me/mailFolders/inbox/messages/delta?" + urllib.parse.urlencode(query)
        tail = {}
        self._seen = {row["id"] for row in self.client.pages(url, headers={"Prefer": "odata.maxpagesize=100"}, tail=tail)}
        self._delta = tail.get("@odata.deltaLink")
        if not self._delta:
            raise GraphError("delta 링크를 받지 못했습니다.")

    def poll(self) -> List[GraphMessage]:
        """지난 poll 뒤로 받은편지함에 새로 들어온 메일(오래된 순). 읽음 표시·삭제·옛 메일을 옮겨 온 것은 빼고."""
        if not self._delta:
            raise GraphError("start_watch 를 먼저 불러야 합니다.")
        rows, tail = [], {}
        for row in self.client.pages(self._delta, headers={"Prefer": "odata.maxpagesize=100"}, tail=tail):
            if "@removed" in row or row["id"] in self._seen or "receivedDateTime" not in row:
                continue
            self._seen.add(row["id"])
            received = datetime.fromisoformat(row["receivedDateTime"].replace("Z", "+00:00"))
            if received >= self._watch_since:
                rows.append(row)
        self._delta = tail.get("@odata.deltaLink") or self._delta
        rows.sort(key=lambda r: r["receivedDateTime"])
        return self._wrap(rows)

    # instruction_mail.read_embedded 가 부르는 COM 이름
    def OpenSharedItem(self, path):
        return SharedItem(Path(path).read_bytes())


def open_mailbox(config: dict, auth_factory: Callable = Auth) -> GraphMailbox:
    """settings.json 값으로 메일함을 연다 — 저장된 토큰으로만(창 없이). 로그인이 필요하면 LoginRequired."""
    auth = auth_factory(config.get("graph_client_id", ""), config.get("graph_tenant_id", ""))
    auth.token()
    return GraphMailbox(Client(auth))
