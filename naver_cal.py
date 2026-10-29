"""네이버 캘린더에 종일 일정 등록 (CalDAV). Qt·Outlook 을 모른다.

- 서버: https://caldav.calendar.naver.com — 네이버 2단계 인증을 켠 뒤 발급한 「애플리케이션 비밀번호」로 접속한다
  (네이버 로그인 비밀번호로는 안 된다).
- 계정: 환경변수 `NAVER_ID`·`NAVER_APP_PASSWORD` → 윈도는 %APPDATA%\\TMProject\\naver_cred.dat (DPAPI 봉인 JSON),
  맥은 키체인 「TMProject Naver」 항목(JSON, `security` 명령으로 읽고 쓴다 — 맥에는 DPAPI 가 없다, 10-29).
  `python naver_cal.py --setup` 으로 저장하고, 접속해서 캘린더 이름 목록을 보여 준다.
- 중복: UID 를 (캘린더·제목·날짜)로 정해 `{캘린더}/{UID}.ics` 에 PUT 한다. 이미 있으면 건너뛴다.
- 옮기기(기한 연장): 새 날짜로 만들고 옛 날짜의 UID 를 DELETE.
- 표준 라이브러리만 쓴다(kipris.py 와 같이 urllib).
"""

import base64
import hashlib
import json
import os
import re
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, Optional, Tuple

BASE_URL = "https://caldav.calendar.naver.com/"
TIMEOUT = 10
MAX_BODY = 2000                      # DESCRIPTION 에 넣을 본문 최대 글자 수

_ENTROPY = b"TMProject/naver/v1"     # DPAPI 보조 엔트로피 (비밀 아님)
_MAC_SERVICE = "TMProject Naver"      # 맥 키체인 항목
_MAC_ACCOUNT = "credentials"
_NS = {"d": "DAV:", "c": "urn:ietf:params:xml:ns:caldav"}


class NaverError(Exception):
    """네이버 캘린더 등록 실패(계정 없음·인증·네트워크·캘린더 없음·응답 형식)."""


# ── 계정 ─────────────────────────────────────────────────────────────

def _cred_path() -> Path:
    base = os.getenv("APPDATA")
    return (Path(base) / "TMProject" if base else Path.home() / ".tmproject") / "naver_cred.dat"


def _security(*args: str, stdin: str = "") -> subprocess.CompletedProcess:
    return subprocess.run(["security", *args], input=stdin, capture_output=True, text=True, timeout=10)


def credentials() -> Optional[Tuple[str, str]]:
    """환경변수 → 맥 키체인 / 윈도 naver_cred.dat. 없으면 None."""
    user = (os.getenv("NAVER_ID") or "").strip()
    password = (os.getenv("NAVER_APP_PASSWORD") or "").strip()
    if user and password:
        return user, password
    try:
        if sys.platform == "darwin":
            done = _security("find-generic-password", "-s", _MAC_SERVICE, "-a", _MAC_ACCOUNT, "-w")
            if done.returncode != 0:
                return None
            text = done.stdout.strip()
            raw = bytes.fromhex(text) if re.fullmatch(r"(?:[0-9a-f]{2})+", text) else text.encode("utf-8")  # ASCII 가 아니면 16진수로 나온다
        else:
            import win32crypt
            _desc, raw = win32crypt.CryptUnprotectData(_cred_path().read_bytes(), _ENTROPY, None, None, 0)
        values = json.loads(raw.decode("utf-8"))
        user, password = values.get("user", "").strip(), values.get("password", "").strip()
        return (user, password) if user and password else None
    except Exception:
        return None


def save_credentials(user: str, password: str) -> str:
    """저장한 곳을 돌려준다. 맥은 비밀번호가 명령줄(ps)에 보이지 않게 `security -i` 의 표준 입력으로,
    따옴표 걱정 없게 16진수(-X)로 넘긴다."""
    data = json.dumps({"user": user, "password": password})
    if sys.platform == "darwin":
        hexed = data.encode("utf-8").hex()
        done = _security("-i", stdin=f'add-generic-password -U -s "{_MAC_SERVICE}" -a {_MAC_ACCOUNT} -X {hexed}\n')
        if done.returncode != 0 or credentials() != (user, password):
            raise NaverError(f"키체인에 저장하지 못했습니다: {(done.stderr or done.stdout).strip()}")
        return f"키체인 「{_MAC_SERVICE}」"
    import win32crypt
    data = data.encode("utf-8")
    path = _cred_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(win32crypt.CryptProtectData(data, "TMProject Naver", _ENTROPY, None, None, 0))
    return str(path)


# ── iCalendar ────────────────────────────────────────────────────────

def make_uid(calendar_name: str, subject: str, day: date) -> str:
    """같은 캘린더·제목·날짜면 늘 같은 UID — 다시 등록하면 중복으로 알아챈다."""
    digest = hashlib.sha1(f"{calendar_name}|{subject}|{day.isoformat()}".encode("utf-8")).hexdigest()
    return f"{digest}@tmproject"


def event_url_for(calendar_url: str, uid: str) -> str:
    """네이버는 PUT 한 경로와 상관없이 일정을 `{UID 전체}.ics` 로 저장한다 — 같은 경로를 써야 다시 찾을 수 있다."""
    return calendar_url + urllib.parse.quote(uid, safe="") + ".ics"


def _escape(text: str) -> str:
    return (text.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,")
            .replace("\r\n", "\\n").replace("\r", "\\n").replace("\n", "\\n"))


def _fold(line: str) -> str:
    """RFC 5545: 한 줄 75 옥텟 — UTF-8 글자를 쪼개지 않고 접는다."""
    out, cur, size = [], "", 0
    for ch in line:
        n = len(ch.encode("utf-8"))
        if size + n > 75:
            out.append(cur)
            cur, size = " ", 1
        cur += ch
        size += n
    out.append(cur)
    return "\r\n".join(out)


def build_ics(uid: str, subject: str, day: date, body: str, now: Optional[datetime] = None) -> str:
    """종일 일정 하나 (DTEND 는 다음날 — 종일 일정의 끝은 배타적)."""
    now = now or datetime.now(timezone.utc)
    if len(body) > MAX_BODY:
        body = body[:MAX_BODY] + "…"
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//TMProject//TaskManager//KO",
        "CALSCALE:GREGORIAN",
        "BEGIN:VEVENT",
        f"UID:{uid}",
        f"DTSTAMP:{now.astimezone(timezone.utc):%Y%m%dT%H%M%SZ}",
        f"DTSTART;VALUE=DATE:{day:%Y%m%d}",
        f"DTEND;VALUE=DATE:{day + timedelta(days=1):%Y%m%d}",
        f"SUMMARY:{_escape(subject)}",
        *([f"DESCRIPTION:{_escape(body)}"] if body else []),
        "TRANSP:TRANSPARENT",
        "END:VEVENT",
        "END:VCALENDAR",
    ]
    return "\r\n".join(_fold(line) for line in lines) + "\r\n"


def ics_description(ics: str) -> str:
    """일정 iCalendar 의 (VEVENT) DESCRIPTION 을 풀어 돌려준다. 없으면 ""."""
    unfolded = re.sub(r"\r?\n[ \t]", "", ics or "")
    in_alarm = False
    for line in unfolded.splitlines():
        if line == "BEGIN:VALARM":
            in_alarm = True
        elif line == "END:VALARM":
            in_alarm = False
        elif not in_alarm and re.match(r"DESCRIPTION[;:]", line):
            value = line.split(":", 1)[1] if ":" in line else ""
            return re.sub(r"\\([\\;,nN])", lambda m: "\n" if m.group(1) in "nN" else m.group(1), value)
    return ""


# ── CalDAV ───────────────────────────────────────────────────────────

class _NoRedirect(urllib.request.HTTPRedirectHandler):
    """PROPFIND·PUT 의 리다이렉트는 직접 따라간다(urllib 은 GET 만 따라감)."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


_opener = urllib.request.build_opener(_NoRedirect)


def _parse_multistatus(text: str) -> ET.Element:
    try:
        return ET.fromstring(text)
    except ET.ParseError:
        raise NaverError("네이버 캘린더 응답이 XML 이 아닙니다.")


class Session:
    """한 번의 등록(일정 여러 개) 동안 접속 정보·캘린더 목록을 재사용한다.

    만들 때는 아무것도 하지 않는다. 첫 호출에서 계정을 읽고 캘린더를 찾는다.
    한 번 실패하면 같은 오류를 그대로 내서, 네트워크가 끊겼을 때 일정마다 기다리지 않는다.
    """

    def __init__(self, user: str = "", password: str = ""):
        self._cred = (user, password) if user and password else None
        self._calendars: Optional[Dict[str, str]] = None
        self._error: Optional[NaverError] = None

    # 요청 ─────────────────────────────────────────────
    def _auth(self) -> str:
        if self._cred is None:
            self._cred = credentials()
        if self._cred is None:
            raise NaverError("네이버 계정이 없습니다 (python naver_cal.py --setup 또는 환경변수 NAVER_ID·NAVER_APP_PASSWORD).")
        token = base64.b64encode(f"{self._cred[0]}:{self._cred[1]}".encode("utf-8")).decode("ascii")
        return f"Basic {token}"

    def _request(self, method: str, url: str, body: str = "", headers: Optional[dict] = None):
        """(상태 코드, 최종 URL, 응답 본문). 2xx·404·412 가 아니면 NaverError."""
        headers = {"Authorization": self._auth(), **(headers or {})}
        data = body.encode("utf-8") if body else None
        for _hop in range(4):
            req = urllib.request.Request(url, data=data, headers=headers, method=method)
            try:
                with _opener.open(req, timeout=TIMEOUT) as res:
                    return res.status, url, res.read().decode("utf-8", errors="replace")
            except urllib.error.HTTPError as e:
                location = e.headers.get("Location") if e.headers else None
                if e.code in (301, 302, 303, 307, 308) and location:
                    nxt = urllib.parse.urljoin(url, location)
                    if urllib.parse.urlparse(nxt).netloc != urllib.parse.urlparse(BASE_URL).netloc:
                        raise NaverError(f"네이버 캘린더가 다른 서버로 보냈습니다: {nxt}")
                    url = nxt
                    continue
                if e.code in (404, 412):
                    return e.code, url, ""
                if e.code == 401:
                    raise NaverError("네이버 인증 실패 (아이디·애플리케이션 비밀번호 확인).")
                raise NaverError(f"네이버 캘린더가 HTTP {e.code} 로 답했습니다.")
            except (urllib.error.URLError, OSError) as e:
                raise NaverError(f"네이버 캘린더에 연결하지 못했습니다: {type(e).__name__}")
        raise NaverError("네이버 캘린더 리다이렉트가 너무 많습니다.")

    def _propfind(self, url: str, depth: str, prop_xml: str):
        body = ('<?xml version="1.0" encoding="utf-8"?>'
                f'<d:propfind xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav"><d:prop>{prop_xml}</d:prop></d:propfind>')
        status, final, text = self._request("PROPFIND", url, body, {
            "Depth": depth, "Content-Type": "application/xml; charset=utf-8"})
        if status != 207:
            raise NaverError(f"네이버 캘린더 PROPFIND 가 {status} 로 답했습니다.")
        return final, _parse_multistatus(text)

    def _href(self, url: str, prop_xml: str, path: str) -> str:
        final, root = self._propfind(url, "0", prop_xml)
        el = root.find(f".//d:prop/{path}/d:href", _NS)
        if el is None or not (el.text or "").strip():
            raise NaverError(f"네이버 캘린더에서 {path} 를 찾지 못했습니다.")
        return urllib.parse.urljoin(final, el.text.strip())

    # 캘린더 ───────────────────────────────────────────
    def calendars(self) -> Dict[str, str]:
        """{캘린더 이름: URL} — 일정(VEVENT)을 담을 수 있는 캘린더만."""
        if self._error:
            raise self._error
        if self._calendars is not None:
            return self._calendars
        try:
            principal = self._href(BASE_URL, "<d:current-user-principal/>", "d:current-user-principal")
            home = self._href(principal, "<c:calendar-home-set/>", "c:calendar-home-set")
            final, root = self._propfind(
                home, "1", "<d:displayname/><d:resourcetype/><c:supported-calendar-component-set/>")
            found = {}
            for resp in root.findall("d:response", _NS):
                href = resp.find("d:href", _NS)
                if href is None or resp.find(".//d:resourcetype/c:calendar", _NS) is None:
                    continue
                comps = resp.findall(".//c:supported-calendar-component-set/c:comp", _NS)
                if comps and not any(c.get("name") == "VEVENT" for c in comps):
                    continue
                name_el = resp.find(".//d:displayname", _NS)
                url = urllib.parse.urljoin(final, href.text.strip())
                name = (name_el.text or "").strip() if name_el is not None else ""
                found[name or url] = url if url.endswith("/") else url + "/"
            self._calendars = found
            return found
        except NaverError as e:
            self._error = e
            raise

    def _event_url(self, calendar_name: str, uid: str) -> str:
        url = self.calendars().get(calendar_name)
        if not url:
            raise NaverError(f"네이버 캘린더 「{calendar_name}」 가 없습니다.")
        return event_url_for(url, uid)

    def exists(self, calendar_name: str, subject: str, day: date) -> bool:
        """이 앱이 같은 캘린더·제목·날짜로 등록한 일정이 이미 있는가 (읽기만 함)."""
        status, _final, _text = self._request("GET", self._event_url(calendar_name, make_uid(calendar_name, subject, day)))
        return status != 404

    def add_all_day(self, calendar_name: str, subject: str, day: date, body: str) -> bool:
        """종일 일정 하나. 추가했으면 True, 이미 있으면 False. 실패하면 NaverError."""
        if self.exists(calendar_name, subject, day):
            return False
        uid = make_uid(calendar_name, subject, day)
        event_url = self._event_url(calendar_name, uid)
        status, _final, _text = self._request("PUT", event_url, build_ics(uid, subject, day, body), {
            "Content-Type": "text/calendar; charset=utf-8", "If-None-Match": "*"})
        if status == 412:
            return False
        if status == 404:
            raise NaverError(f"네이버 캘린더 「{calendar_name}」 에 일정을 쓰지 못했습니다 (HTTP 404).")
        return True

    # 옮기기 (기한 확인) ───────────────────────────────
    def delete_all_day(self, calendar_name: str, subject: str, day: date) -> bool:
        """이 앱이 (캘린더·제목·날짜)로 등록한 일정을 지운다. 지웠으면 True, 없었으면 False."""
        status, _final, _text = self._request("DELETE", self._event_url(calendar_name, make_uid(calendar_name, subject, day)))
        return status != 404

    def move_all_day(self, calendar_name: str, subject: str, old_day: date, new_day: date,
                     body: Optional[str] = None) -> Tuple[bool, bool]:
        """일정을 다른 날로 옮긴다 — 새 날짜로 먼저 만들고(실패하면 옛 일정은 그대로) 옛 UID 를 DELETE.

        body=None 이면 옛 일정의 설명칸을 그대로 옮긴다. 반환값: (새 일정을 만들었나, 옛 일정을 지웠나).
        """
        if old_day == new_day:
            return False, False
        old_url = self._event_url(calendar_name, make_uid(calendar_name, subject, old_day))
        status, _final, text = self._request("GET", old_url)
        if body is None:
            body = ics_description(text) if status != 404 else ""
        added = self.add_all_day(calendar_name, subject, new_day, body)
        deleted = False
        if status != 404:
            status, _final, _text = self._request("DELETE", old_url)
            deleted = status != 404
        return added, deleted


# ── 설정 도우미 ──────────────────────────────────────────────────────

def _setup():
    import getpass
    user = input("네이버 아이디: ").strip()
    password = getpass.getpass("애플리케이션 비밀번호(입력이 보이지 않음): ").strip()
    if not user or not password:
        print("아이디·비밀번호가 비어 있습니다.")
        return 1
    try:
        names = sorted(Session(user, password).calendars())
    except NaverError as e:
        print(f"접속 실패: {e}")
        return 1
    try:
        print(f"저장: {save_credentials(user, password)}")
    except NaverError as e:
        print(e)
        return 1
    print("\n네이버 캘린더 목록:")
    for name in names:
        print(f"  - {name}")
    print('\nsettings.json 에 다음을 넣으세요 (이름은 위 목록에서):')
    print('  "naver_calendars": {"Mail": "<캘린더 이름>", "No_Instructions": "<캘린더 이름>"}')
    return 0


if __name__ == "__main__":
    import sys
    if "--setup" in sys.argv[1:]:
        sys.exit(_setup())
    print("사용법: python naver_cal.py --setup")
