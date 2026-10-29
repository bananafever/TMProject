"""KIPRIS Plus 통지서 마감기한 조회. Qt·Outlook 을 모른다.

CPProject(ChekPat)의 `app/modules/patent/kipris_client.py`·`kipris_case.py` 에서
「특허·실용 통지서 마감기한」 상품(DueDateService) 하나만 떼어 왔다.

- 키: ChekPat 과 같은 키를 쓴다. 환경변수 `KIPRIS_API_KEY` → %APPDATA%\\ChekPat\\kipris_key.dat
  → %APPDATA%\\ChekPat\\credentials.dat 의 `KIPRIS_API_KEY` 순서. 두 파일은 DPAPI 로 봉인되어 있다.
- 호출 수: 무료 월 1,000회가 키 단위라 NAS `iplog.kipris_usage`(키의 SHA-256 앞 16자 × 달)에 ChekPat(윈도·맥)과 같이 센다
  (ChekPat 상한 900, 10-27). NAS 에 닿지 않으면 ChekPat 의 settings.json `kipris_calls` 로 잇고, 닿을 때 못 올린 만큼 올린다
  (`kipris_reported`). ChekPat `app/modules/patent/kipris_usage.py` 와 같은 표·같은 규칙이다 — 바꾸면 양쪽을 함께.
  다만 이 앱의 상한은 950 이다 — 이 앱의 호출은 한 달 몇 번뿐인데 빠지면 일정이 틀린다. ChekPat 이 900 을 다 쓴 달
  (10-28)에 재심사 여부를 못 받은 일이 있어, ChekPat 이 멈춘 뒤에도 이 앱은 50회를 더 쓴다(10-29).
- 응답 필드: `Mailingnumber`(발송번호)·`DocNm`(서류명)·`ExpirationDuedateDate`(연장 반영 마감기한).
- 「특허·실용 서지 상세」(`/kipo-api/kipi/` 게이트, 키는 `ServiceKey`)로 거절결정이 재심사에 의한 것인지 가린다
  — `finalDisposal`(「거절결정(일반)」·「거절결정(재심사)」)과 서류 목록(「…보정서(재심사)」).
- 같은 서지 상세의 서류 목록(`documentName`·`receiptDate`·`commonCodeName`)으로 기간연장신청서를 냈는지 본다.
  낸 날 바로 「접수중」 으로 실리고, 마감기한(`ExpirationDuedateDate`)은 「수리」 뒤에야 늘어난다(10-24 확인).
"""

import hashlib
import json
import logging
import os
import re
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Dict, List, Optional

BASE_URL = "https://plus.kipris.or.kr"
_GATEWAYS = {"openapi": ("openapi/rest", "accessKey"), "kipo": ("kipo-api/kipi", "ServiceKey")}
DUE_DATE_PATH = "DueDateService/dueDateApplicationNoticeApplnoInfo"
BIBLIO_PATH = "patUtiModInfoSearchSevice/getBibliographyDetailInfoSearch"   # kipo 게이트에만 있다
MONTHLY_CAP = 950                    # 무료 1,000회에서 손으로 찾아볼 여유 50회. ChekPat 은 900 에서 멈춘다
MIN_INTERVAL = 0.1                   # 초. 초당 50회를 넘기면 IP 가 막힌다
TIMEOUT = 10

_ENTROPY = b"ChekPat/credentials/v1"  # ChekPat DPAPI 보조 엔트로피 (비밀 아님)


class KiprisError(Exception):
    """KIPRIS 호출 실패(키 없음·네트워크·응답 형식·결과 코드)."""


@dataclass
class Notice:
    send_number: str                 # 발송번호
    title: str                       # 서류명 — 「의견제출통지서」·「최후의견제출통지서」 …
    due_date: str                    # 마감기한 YYYY-MM-DD — 연장했으면 연장 뒤 기한


@dataclass
class Document:
    name: str                        # 서류명 — 「[지정기간연장]기간 연장신청서·기간 단축신청서·…」
    receipt_date: str                # 접수일자 YYYY-MM-DD
    status: str                      # 처리 상태 — 「접수중」·「수리」·「발송처리완료」 …
    number: str                      # 접수번호


# ── 키 ───────────────────────────────────────────────────────────────

_MAC_SERVICE = "ChekPat"
_MAC_ACCOUNT = "credentials"


def _mac_credentials() -> Optional[dict]:
    """맥 키체인에서 ChekPat 자격증명 JSON 읽기."""
    if sys.platform != "darwin":
        return None
    try:
        import keyring
        text = keyring.get_password(_MAC_SERVICE, _MAC_ACCOUNT)
        if text:
            data = json.loads(text)
            return data if isinstance(data, dict) else None
    except Exception:
        pass
    try:
        done = subprocess.run(
            ["security", "find-generic-password", "-s", _MAC_SERVICE, "-a", _MAC_ACCOUNT, "-w"],
            capture_output=True, text=True, timeout=5,
        )
        if done.returncode == 0 and done.stdout.strip():
            data = json.loads(done.stdout.strip())
            return data if isinstance(data, dict) else None
    except Exception:
        pass
    return None


def _chekpat_dir() -> Path:
    base = os.getenv("APPDATA")
    return Path(base) / "ChekPat" if base else Path.home() / ".chekpat"


def _unprotect(data: bytes) -> bytes:
    import win32crypt
    _desc, raw = win32crypt.CryptUnprotectData(data, _ENTROPY, None, None, 0)
    return raw


def chekpat_credentials() -> Optional[dict]:
    """ChekPat 자격증명 딕셔너리. 맥은 키체인, 윈도는 credentials.dat (DPAPI)."""
    if sys.platform == "darwin":
        data = _mac_credentials()
        if data is not None:
            return data
    try:
        raw = _unprotect((_chekpat_dir() / "credentials.dat").read_bytes())
        data = json.loads(raw.decode("utf-8"))
        return data if isinstance(data, dict) else None
    except Exception:
        return None


def api_key() -> str:
    """환경변수 → ChekPat kipris_key.dat → ChekPat credentials.dat. 없으면 ""."""
    key = (os.getenv("KIPRIS_API_KEY") or "").strip()
    if key:
        return key
    d = _chekpat_dir()
    try:
        return _unprotect((d / "kipris_key.dat").read_bytes()).decode("utf-8").strip()
    except Exception:
        pass
    values = chekpat_credentials()
    if values and isinstance(values, dict):
        return (values.get("KIPRIS_API_KEY") or "").strip()
    return ""


# ── 호출 수 (NAS 공유 셈 + ChekPat settings.json 사본) ────────────────
_count_lock = threading.Lock()


def key_id(key: str) -> str:
    """키의 SHA-256 앞 16자 — NAS 에는 키 자체를 올리지 않는다 (ChekPat `kipris_usage.key_id` 와 같다)."""
    return hashlib.sha256((key or "").strip().encode("utf-8")).hexdigest()[:16]


class _Full(Exception):
    def __init__(self, total: int):
        self.total = total
        super().__init__(str(total))


_known_total: Dict[tuple, int] = {}      # (key_id, 달) → 마지막으로 본 공유 값


def _shared_charge(key: str, month: str, delta: int, cap: int) -> Optional[int]:
    """NAS 공유 셈에 `delta` 를 더한다. 이미 cap 이상이면 `_Full`. 닿지 않거나 표가 없으면 None."""
    import db                                       # db 가 kipris 를 들이므로 여기서
    kid = key_id(key)
    try:
        conn = db.connect()
    except db.DbError:
        return None
    try:
        with conn.cursor() as cur:
            cur.execute("INSERT IGNORE INTO kipris_usage (ym, key_id, calls) VALUES (%s, %s, 0)", (month, kid))
            cur.execute("UPDATE kipris_usage SET calls = calls + %s WHERE ym = %s AND key_id = %s AND calls < %s",
                        (int(delta), month, kid, int(cap)))
            added = cur.rowcount
            cur.execute("SELECT calls FROM kipris_usage WHERE ym = %s AND key_id = %s", (month, kid))
            total = int(list(cur.fetchone().values())[0])
        conn.commit()
    except Exception as e:                          # 표 없음·쓰기 거부·네트워크 — 이 PC 셈으로 잇는다
        logging.warning(f"KIPRIS 공유 셈 쓰기 실패: {type(e).__name__}: {e}")
        db._last_fail = time.monotonic()            # 한동안 다시 붙어 보지 않는다
        return None
    finally:
        try:
            conn.close()
        except Exception:
            pass
    _known_total[(kid, month)] = total
    if not added:
        raise _Full(total)
    return total


def _count_call(today: Optional[date] = None) -> None:
    """한 번 보낸다고 적는다. 이번 달 상한이면 보내기 전에 막는다.

    NAS 공유 셈에 먼저 더하고, ChekPat settings.json 의 이 PC 셈(`kipris_calls`)도 늘 함께 적는다.
    `kipris_reported` 는 그 가운데 공유 셈에 이미 올린 만큼 — NAS 에 못 닿은 사이의 호출은 다음에 닿을 때 올린다."""
    path = _chekpat_dir() / "settings.json"
    month = (today or date.today()).strftime("%Y-%m")
    with _count_lock:
        try:
            with open(path, encoding="utf-8") as f:
                settings = json.load(f)
        except (OSError, ValueError):
            settings = {}
        calls = settings.get("kipris_calls")
        calls = dict(calls) if isinstance(calls, dict) else {}
        reported = settings.get("kipris_reported")
        reported = dict(reported) if isinstance(reported, dict) else {}
        own = int(calls.get(month, 0))
        pending = max(0, own - int(reported.get(month, 0)))
        key = api_key()
        shared = None
        if key:
            try:
                shared = _shared_charge(key, month, pending + 1, MONTHLY_CAP)
            except _Full as e:
                raise KiprisError("이번 달 KIPRIS 호출이 상한에 닿았습니다 (%d/%d)." % (e.total, MONTHLY_CAP))
        if shared is None:
            base = _known_total.get((key_id(key), month)) if key else None
            used = own if base is None else base + pending
            if used >= MONTHLY_CAP:
                raise KiprisError("이번 달 KIPRIS 호출이 상한에 닿았습니다 (%d/%d)." % (used, MONTHLY_CAP))
        # 지난 달 기록은 석 달치만 남긴다 (ChekPat 과 같은 규칙)
        keep = sorted(k for k in calls if k != month)[-2:]
        settings["kipris_calls"] = {**{k: calls[k] for k in keep}, month: own + 1}
        if shared is not None:
            settings["kipris_reported"] = {**{k: reported[k] for k in keep if k in reported}, month: own + 1}
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            tmp = path.with_suffix(".tmp")
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            os.replace(tmp, path)
        except OSError as e:
            logging.error(f"KIPRIS 호출 수 기록 실패: {e}")


# ── 응답 해석 ─────────────────────────────────────────────────────────

def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _text(el) -> str:
    return (el.text or "").strip() if el is not None else ""


def _find(root, name: str):
    """이름공간·대소문자를 가리지 않고 첫 자손."""
    name = name.lower()
    for el in root.iter():
        if _local(el.tag).lower() == name:
            return el
    return None


def parse_root(xml_text, product="통지서 마감기한"):
    """응답 XML → 뿌리 요소. 결과 코드가 실패면 KiprisError."""
    try:
        root = ET.fromstring(xml_text.encode("utf-8") if isinstance(xml_text, str) else xml_text)
    except ET.ParseError as e:
        raise KiprisError(f"KIPRIS 응답을 읽지 못했습니다(XML 아님): {e}")
    code = _text(_find(root, "resultCode"))
    msg = _text(_find(root, "resultMsg"))
    success = _text(_find(root, "successYN")).upper()
    if code == "101":
        raise KiprisError(f"KIPRIS Plus 에서 「{product}」 상품을 신청하지 않았습니다 (결과 코드 101).")
    if code == "31" or "DEADLINE_HAS_EXPIRED" in msg.upper():
        raise KiprisError(f"KIPRIS Plus 「{product}」 상품의 이용기간이 끝났습니다 (결과 코드 31).")
    if code == "30" or re.search(r"SERVICE[\s_]*KEY[\s_]*IS[\s_]*NOT[\s_]*REGISTER", msg, re.I):
        raise KiprisError(f"KIPRIS Plus 키가 맞지 않거나 상품을 신청하지 않았습니다 (결과 코드 {code or '?'}: {msg}).")
    if success == "N" or code not in ("", "0", "00"):
        raise KiprisError(f"KIPRIS 결과 코드 {code or '?'}: {msg or '알 수 없는 오류'}")
    return root


def records(body) -> List[Dict[str, str]]:
    """`body` 아래에서 「자식이 모두 잎인 요소」 하나를 한 건으로 본다."""
    out = []
    if body is None:
        return out
    for el in body.iter():
        kids = list(el)
        if not kids or any(len(k) for k in kids):
            continue
        rec = {_local(k.tag): _text(k) for k in kids}
        if any(rec.values()):
            out.append(rec)
    return out


def field(rec: Dict[str, str], *names: str) -> str:
    """대소문자를 가리지 않고 처음 값이 있는 칸."""
    low = {k.lower(): v for k, v in rec.items()}
    for n in names:
        v = low.get(n.lower())
        if v:
            return v
    return ""


def appln_number(text: str) -> str:
    """「10-2025-7000694」·「1020257000694」 → 13자리. 알아볼 수 없으면 ""."""
    digits = re.sub(r"\D", "", text or "")
    return digits if len(digits) == 13 else ""


def iso_date(text: str) -> str:
    """「20260121」·「2026.01.21」·「2026-01-21」 → 「2026-01-21」. 아니면 ""."""
    m = re.search(r"(\d{4})\D?(\d{2})\D?(\d{2})", text or "")
    if not m:
        return ""
    try:
        return date(*(int(x) for x in m.groups())).isoformat()
    except ValueError:
        return ""


def parse_due_dates(xml_text) -> List[Notice]:
    """마감기한 응답 XML → Notice 목록."""
    out = []
    for r in records(_find(parse_root(xml_text), "body")):
        send = field(r, "Mailingnumber")
        if send:
            out.append(Notice(send_number=send, title=field(r, "DocNm"),
                              due_date=iso_date(field(r, "ExpirationDuedateDate"))))
    return out


def parse_reexam(xml_text) -> bool:
    """서지 상세 응답 XML → 이 출원의 거절결정이 재심사에 의한 것인가.

    최종처분이 「거절결정(일반)」 이면 아니다. 「거절결정(재심사)」 나 KIPRIS 반영이 늦어 아직
    「거절결정 후 재심사중」 이면 그렇다. 최종처분으로 알 수 없으면 재심사 보정서를 낸 적이 있는지로 본다.
    """
    root = parse_root(xml_text, "특허·실용 서지 상세")
    disposal = _text(_find(root, "finalDisposal"))
    if "(일반)" in disposal:
        return False
    if "재심사" in disposal:
        return True
    return any("보정서(재심사)" in _text(el) for el in root.iter() if _local(el.tag).lower() == "documentname")


def parse_documents(xml_text) -> List[Document]:
    """서지 상세 응답 XML → 서류 목록 (접수 순). 처리 상태의 영문 괄호는 뗀다 — 「접수중 (On receiving)」 → 「접수중」."""
    out = []
    for r in records(_find(parse_root(xml_text, "특허·실용 서지 상세"), "body")):
        name = field(r, "documentName")
        if name:
            out.append(Document(name=name, receipt_date=iso_date(field(r, "receiptDate")),
                                status=re.sub(r"\s*\(.*\)\s*$", "", field(r, "commonCodeName")),
                                number=field(r, "receiptNumber")))
    return out


def pick_oa_due(notices: List[Notice], doc_type: str) -> Optional[Notice]:
    """서류명에 문서종류가 들어간 통지서 가운데 가장 최신(발송번호 숫자가 가장 큰) 것."""
    cands = [n for n in notices if doc_type and doc_type in n.title and n.due_date]
    if not cands:
        return None
    return max(cands, key=lambda n: int(re.sub(r"\D", "", n.send_number) or 0))


# ── 요청 ─────────────────────────────────────────────────────────────
_gate_lock = threading.Lock()
_last_ts = 0.0


def _acquire():
    global _last_ts
    while True:
        with _gate_lock:
            now = time.monotonic()
            wait = _last_ts + MIN_INTERVAL - now
            if wait <= 0:
                _last_ts = now
                return
        time.sleep(wait)


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    """없는 경로는 404 가 아니라 302(→ code400.jsp)로 온다. 따라가지 않는다."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


_opener = urllib.request.build_opener(_NoRedirect)


def _send(path: str, params: dict, gateway: str = "openapi") -> str:
    key = api_key()
    if not key:
        raise KiprisError("KIPRIS Plus 키가 없습니다 (ChekPat 연결 설정 또는 환경변수 KIPRIS_API_KEY).")
    prefix, key_param = _GATEWAYS[gateway]
    url = f"{BASE_URL}/{prefix}/{path}?" + urllib.parse.urlencode({**params, key_param: key})
    for attempt in range(2):  # 연결이 끊기는 일시 오류(ConnectionResetError 등)는 한 번 더 보낸다
        _count_call()
        _acquire()
        try:
            with _opener.open(url, timeout=TIMEOUT) as res:
                return res.read().decode("utf-8", errors="replace")
        except urllib.error.HTTPError as e:
            if 300 <= e.code < 400 or e.code == 404:
                raise KiprisError(f"KIPRIS 에 「{path}」 서비스가 없습니다 (HTTP {e.code}).")
            raise KiprisError(f"KIPRIS 가 HTTP {e.code} 로 답했습니다.")
        except (urllib.error.URLError, OSError) as e:
            if attempt == 0:
                time.sleep(1)
                continue
            raise KiprisError(f"KIPRIS 에 연결하지 못했습니다: {type(e).__name__}")


def fetch_due_dates(appln: str) -> List[Notice]:
    """출원번호의 통지서 마감기한 목록 (호출 1회). 실패하면 KiprisError."""
    num = appln_number(appln)
    if not num:
        raise KiprisError(f"출원번호를 알아볼 수 없습니다: {appln}")
    return parse_due_dates(_send(DUE_DATE_PATH, {"applicationNumber": num}))


def fetch_reexam(appln: str) -> bool:
    """출원번호의 거절결정이 재심사에 의한 것인가 (서지 상세 호출 1회). 실패하면 KiprisError."""
    num = appln_number(appln)
    if not num:
        raise KiprisError(f"출원번호를 알아볼 수 없습니다: {appln}")
    return parse_reexam(_send(BIBLIO_PATH, {"applicationNumber": num}, gateway="kipo"))


def fetch_documents(appln: str) -> List[Document]:
    """출원번호의 서류 목록 (서지 상세 호출 1회). 실패하면 KiprisError."""
    num = appln_number(appln)
    if not num:
        raise KiprisError(f"출원번호를 알아볼 수 없습니다: {appln}")
    return parse_documents(_send(BIBLIO_PATH, {"applicationNumber": num}, gateway="kipo"))
