"""KIPRIS Plus 통지서 마감기한 조회. Qt·Outlook 을 모른다.

CPProject(ChekPat)의 `app/modules/patent/kipris_client.py`·`kipris_case.py` 에서
「특허·실용 통지서 마감기한」 상품(DueDateService) 하나만 떼어 왔다.

- 키: ChekPat 과 같은 키를 쓴다. 환경변수 `KIPRIS_API_KEY` → %APPDATA%\\ChekPat\\kipris_key.dat
  → %APPDATA%\\ChekPat\\credentials.dat 의 `KIPRIS_API_KEY` 순서. 두 파일은 DPAPI 로 봉인되어 있다.
- 호출 수·속도·막힘: ChekPat 과 같은 규약을 지킨다(ChekPat `docs/KIPRIS_PROTOCOL.md` — 바꾸면 양쪽을 함께, 11-01 x1 검토).
  ① 이 PC 셈 `kipris_ledger.json`(키·달마다, 파일 잠금 안에서 — settings.json 의 `kipris_calls` 는 더 쓰지 않는다)
  ② NAS `iplog.kipris_usage`(키의 SHA-256 앞 16자 × 달)와 발신자 표 `kipris_usage_src` 로 정산 — 미보고분은 조건 없이,
  새 호출 1회만 정산 뒤 값이 상한 미만일 때 ③ 모든 전송이 `~/.chekpat/kipris_bulk/gate` 속도 문턱(초당 40)을 거친다
  ④ `BLOCKED` 표시가 있으면 보내지 않고, 응답 결과 코드 20(접근 거부)이면 남긴다(지우는 것은 ChekPat).
  이 앱의 상한은 950 이다(ChekPat 900) — 이 앱의 호출은 한 달 몇 번뿐인데 빠지면 일정이 틀린다. ChekPat 이 900 을 다 쓴 달
  (10-28)에 재심사 여부를 못 받은 일이 있어, ChekPat 이 멈춘 뒤에도 이 앱은 50회를 더 쓴다(10-29).
- 응답 필드: `Mailingnumber`(발송번호)·`DocNm`(서류명)·`ExpirationDuedateDate`(연장 반영 마감기한).
- 「특허·실용 서지 상세」(`/kipo-api/kipi/` 게이트, 키는 `ServiceKey`)로 거절결정이 재심사에 의한 것인지 가린다
  — `finalDisposal`(「거절결정(일반)」·「거절결정(재심사)」)과 서류 목록(「…보정서(재심사)」).
- 같은 서지 상세의 서류 목록(`documentName`·`receiptDate`·`commonCodeName`)으로 기간연장신청서를 냈는지 본다.
  낸 날 바로 「접수중」 으로 실리고, 마감기한(`ExpirationDuedateDate`)은 「수리」 뒤에야 늘어난다(10-24 확인).
"""

import contextlib
import hashlib
import json
import logging
import os
import re
import secrets
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


# ── 막힘 방지·호출 수 (ChekPat 과 공유 규약 — ChekPat `docs/KIPRIS_PROTOCOL.md`, 바꾸면 양쪽을 함께) ──────────
# 규약: 이 PC 셈 `kipris_ledger.json`(잠금 `kipris_ledger.lock`) · 속도 문턱 `~/.chekpat/kipris_bulk/gate` ·
# 막힘 표시 `~/.chekpat/kipris_bulk/BLOCKED`. 정산은 NAS `kipris_usage`(+발신자 표 `kipris_usage_src`).
SAFE_MAX_RATE = 40.0                 # 초당 — 50 을 넘기면 IP 가 막힌다. ChekPat·수집기와 같은 문턱을 함께 거친다
LOCK_TIMEOUT = 30.0                  # 초. 이 PC 셈 잠금을 이만큼 못 잡으면 보내지 않는다
REWIND_SLACK = 1.0                   # 초. gate 의 「적은 때」 가 지금보다 이만큼 넘게 뒤면 시계가 되감긴 것
_count_lock = threading.RLock()


def _guard_dir() -> Path:
    return Path.home() / ".chekpat" / "kipris_bulk"      # 시험은 이것을 바꾼다


def key_id(key: str) -> str:
    """키의 SHA-256 앞 16자 — NAS 에는 키 자체를 올리지 않는다 (ChekPat `kipris_usage.key_id` 와 같다)."""
    return hashlib.sha256((key or "").strip().encode("utf-8")).hexdigest()[:16]


# 파일 잠금 — 0번 바이트 1개(규약 1절). ChekPat `kipris_guard` 와 같은 방식이라 두 앱이 서로를 막는다.

def _try_flock(fh) -> bool:
    if sys.platform == "win32":
        import msvcrt
        fh.seek(0)
        try:
            msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
            return True
        except OSError:
            return False
    import fcntl
    try:
        fcntl.flock(fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
        return True
    except (BlockingIOError, PermissionError):
        return False


def _flock(fh) -> None:
    """기다리는 배타 잠금 — gate 파일."""
    if sys.platform == "win32":
        import msvcrt
        fh.seek(0)
        msvcrt.locking(fh.fileno(), msvcrt.LK_LOCK, 1)
    else:
        import fcntl
        fcntl.flock(fh, fcntl.LOCK_EX)


def _funlock(fh) -> None:
    if sys.platform == "win32":
        import msvcrt
        fh.seek(0)
        msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
    else:
        import fcntl
        fcntl.flock(fh, fcntl.LOCK_UN)


@contextlib.contextmanager
def _locked_file(path: Path, timeout: float = None):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "a+b") as fh:
        deadline = time.monotonic() + (LOCK_TIMEOUT if timeout is None else timeout)
        while not _try_flock(fh):
            if time.monotonic() > deadline:
                raise KiprisError("KIPRIS 호출 수 기록을 잠그지 못했습니다 — 다른 프로그램이 쥐고 있습니다.")
            time.sleep(0.05)
        try:
            yield
        finally:
            _funlock(fh)


# 막힘 표시 — 읽고 남기기만 한다. 지우는 것은 사람(ChekPat).

def blocked_reason() -> Optional[str]:
    try:
        return (_guard_dir() / "BLOCKED").read_text(encoding="utf-8").strip() or "막힘"
    except OSError:
        return None


def _mark_blocked(reason: str) -> None:
    try:
        path = _guard_dir() / "BLOCKED"
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("%s\n%s\n" % (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()), reason), encoding="utf-8")
    except OSError as e:
        logging.error(f"KIPRIS 막힘 표시를 남기지 못했습니다: {e}")


def _check_not_blocked() -> None:
    reason = blocked_reason()
    if reason is not None:
        why = reason.split("\n", 1)[-1].strip() or reason
        raise KiprisError(f"KIPRIS 호출이 막혀 있어 보내지 않습니다 ({why}). 지우려면 ChekPat 연결 설정에서 KIPRIS 상품을 점검하세요.")


# 이 PC 셈 — `D/kipris_ledger.json` (규약 2절)

def _ledger_path() -> Path:
    return _chekpat_dir() / "kipris_ledger.json"


def _month_index(ym: str) -> int:
    y, m = ym.split("-")
    return int(y) * 12 + int(m)


def _migrate_settings(entries: dict, kid: str) -> None:
    """예전 settings.json 의 `kipris_calls`·`kipris_reported` → 지금 키의 항목. 두 칸은 지우지 않고 더는 쓰지 않는다."""
    try:
        with open(_chekpat_dir() / "settings.json", encoding="utf-8") as f:
            settings = json.load(f)
    except (OSError, ValueError):
        return
    calls = settings.get("kipris_calls") if isinstance(settings, dict) else None
    reported = settings.get("kipris_reported") if isinstance(settings, dict) else None
    if not isinstance(calls, dict):
        return
    reported = reported if isinstance(reported, dict) else {}
    for ym, n in calls.items():
        try:
            c, r = int(n), int(reported.get(ym, 0))
        except (TypeError, ValueError):
            continue
        entries[f"{kid}|{ym}"] = {"calls": c, "reported": min(r, c), "full": False}


class _Ledger:
    """잠금 안에서만 쓴다 — `with _ledger() as led:` 안에서 `led.entries` 를 고치고 `led.save()`."""

    def __init__(self):
        self.data = {"version": 1, "src": "", "entries": {}}

    @property
    def src(self) -> str:
        return self.data["src"]

    @property
    def entries(self) -> dict:
        return self.data["entries"]

    def load(self, kid: str) -> None:
        path = _ledger_path()
        data = None
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict) or not isinstance(data.get("entries"), dict):
                raise ValueError("형식")
        except FileNotFoundError:
            data = None
        except (OSError, ValueError):
            try:
                os.replace(path, path.with_name(path.name + ".bad"))      # 못 읽는 파일은 치워 두고 새로
            except OSError:
                pass
            data = None
        fresh = data is None
        if fresh:
            data = {"version": 1, "src": "", "entries": {}}
        if not re.fullmatch(r"[0-9a-f]{16}", str(data.get("src") or "")):
            data["src"] = secrets.token_hex(8)
        self.data = data
        if fresh:
            if kid:
                _migrate_settings(self.entries, kid)
            self.save()

    def save(self) -> None:
        path = _ledger_path()
        text = json.dumps(self.data, ensure_ascii=False, indent=1)
        tmp = path.with_name(f"{path.name}.{os.getpid()}.tmp")
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            with open(tmp, "w", encoding="utf-8") as f:
                f.write(text)
            for attempt in range(5):
                try:
                    os.replace(tmp, path)
                    return
                except OSError:
                    time.sleep(0.05)
            with open(path, "w", encoding="utf-8") as f:       # 윈도에서 끝내 막히면 그 자리에
                f.write(text)
            os.unlink(tmp)
        except OSError as e:
            logging.error(f"KIPRIS 호출 수 기록 실패: {e}")

    def entry(self, kid: str, ym: str) -> dict:
        e = self.entries.setdefault(f"{kid}|{ym}", {})
        e["calls"] = int(e.get("calls", 0))
        e["reported"] = min(int(e.get("reported", 0)), e["calls"])
        e["full"] = bool(e.get("full", False))
        return e

    def prune(self, this_month: str) -> None:
        """미보고분이 없고 가득 차지도 않은, 석 달 넘게 지난 항목만 지운다."""
        now = _month_index(this_month)
        for k in [k for k, e in self.entries.items()
                  if int(e.get("calls", 0)) <= int(e.get("reported", 0)) and not e.get("full")
                  and now - _month_index(k.split("|", 1)[1]) >= 3]:
            del self.entries[k]


@contextlib.contextmanager
def _ledger(kid: str):
    with _count_lock:                                  # 같은 프로세스의 스레드 — 파일 잠금보다 먼저
        with _locked_file(_chekpat_dir() / "kipris_ledger.lock"):
            led = _Ledger()
            led.load(kid)
            yield led


# NAS 정산 (규약 4절)

class _Full(Exception):
    def __init__(self, total: int):
        self.total = total
        super().__init__(str(total))


class _Offline(Exception):
    """NAS 에 닿지 않았거나 반영 여부를 모른다."""


_known_total: Dict[tuple, int] = {}      # (key_id, 달) → 지금까지 본 가장 큰 공유 값 — 줄지 않는다


def _remember(kid: str, ym: str, total: int) -> None:
    _known_total[(kid, ym)] = max(int(total), _known_total.get((kid, ym), 0))


def _one(cur) -> int:
    return int(list(cur.fetchone().values())[0])


def _nas_settle(kid: str, ym: str, src: str, entry: dict, want_new: bool, cap: int):
    """한 항목을 NAS 에 정산한다. 돌려주는 값: (새 호출 허가 여부, 정산 뒤 공유 값). 닿지 않으면 `_Offline`.

    미보고분은 이미 쓴 호출이라 조건 없이 더하고, 새 1회만 정산 뒤의 값이 cap 미만일 때 허가한다(x1 4).
    발신자 줄이 누적값이라 커밋 응답을 잃고 다시 보내도 두 번 더해지지 않는다(x1 5)."""
    import db                                       # db 가 kipris 를 들이므로 여기서
    calls, reported, full = entry["calls"], entry["reported"], entry["full"]
    try:
        conn = db.connect()
    except db.DbError:
        raise _Offline()
    try:
        with conn.cursor() as cur:
            cur.execute("INSERT IGNORE INTO kipris_usage (ym, key_id, calls) VALUES (%s, %s, 0)", (ym, kid))
            try:
                cur.execute("INSERT IGNORE INTO kipris_usage_src (ym, key_id, src, calls) VALUES (%s, %s, %s, %s)",
                            (ym, kid, src, reported))
                has_src = True
            except Exception as e:
                if not (e.args and e.args[0] == 1146):        # 발신자 표가 없는 예전 NAS 만 예전 표로 잇는다
                    raise
                has_src = False
            if has_src:
                cur.execute("SELECT calls FROM kipris_usage_src WHERE ym = %s AND key_id = %s AND src = %s FOR UPDATE",
                            (ym, kid, src))
                have = _one(cur)
                cur.execute("SELECT calls FROM kipris_usage WHERE ym = %s AND key_id = %s FOR UPDATE", (ym, kid))
                total = _one(cur)
                pending = max(0, calls - have)
                after = total + pending
                grant = bool(want_new and not full and after < cap)
                add = pending + (1 if grant else 0)
                if add > 0:
                    cur.execute("UPDATE kipris_usage SET calls = calls + %s WHERE ym = %s AND key_id = %s", (add, ym, kid))
                cur.execute("UPDATE kipris_usage_src SET calls = GREATEST(calls, %s) WHERE ym = %s AND key_id = %s AND src = %s",
                            (have + add, ym, kid, src))
                total = after + (1 if grant else 0)
            else:
                pending = max(0, calls - reported)
                grant = False
                if pending > 0:
                    cur.execute("UPDATE kipris_usage SET calls = calls + %s WHERE ym = %s AND key_id = %s", (pending, ym, kid))
                if want_new and not full:
                    cur.execute("UPDATE kipris_usage SET calls = calls + 1 WHERE ym = %s AND key_id = %s AND calls < %s",
                                (ym, kid, int(cap)))
                    grant = cur.rowcount > 0
                cur.execute("SELECT calls FROM kipris_usage WHERE ym = %s AND key_id = %s", (ym, kid))
                total = _one(cur)
            if full:
                cur.execute("UPDATE kipris_usage SET calls = GREATEST(calls, %s) WHERE ym = %s AND key_id = %s",
                            (int(cap), ym, kid))
                total = max(total, int(cap))
        conn.commit()
    except Exception as e:                          # 표 없음·쓰기 거부·네트워크·커밋 응답 불명 — 이 PC 셈으로 잇는다
        logging.warning(f"KIPRIS 공유 셈 쓰기 실패: {type(e).__name__}: {e}")
        db._last_fail = time.monotonic()            # 한동안 다시 붙어 보지 않는다
        raise _Offline()
    finally:
        try:
            conn.close()
        except Exception:
            pass
    _remember(kid, ym, total)
    return grant, total


def _count_call(today: Optional[date] = None) -> None:
    """한 번 보낸다고 적는다. 상한이면 보내기 전에 막는다 (규약 3·4절).

    잠금을 쥔 채 다른 키·지난 달의 미보고분을 먼저 NAS 에 정산하고(소유권 = 잠금을 쥔 쪽만 `reported` 를 올린다),
    끝으로 지금 키·이번 달 항목에 새 1회를 청구한다. NAS 에 닿지 않으면 이 PC 셈으로 판단한다."""
    key = api_key()
    if not key:
        return                                      # 키가 없으면 보내지 않는다 — _send 가 먼저 막는다
    kid, month = key_id(key), (today or date.today()).strftime("%Y-%m")
    cur_key = f"{kid}|{month}"
    with _ledger(kid) as led:
        offline = False
        stale = sorted((k for k, e in led.entries.items()
                        if k != cur_key and (int(e.get("calls", 0)) > int(e.get("reported", 0)) or e.get("full"))),
                       key=lambda k: k.split("|", 1)[1])
        for k in stale:
            k_id, ym = k.split("|", 1)
            e = led.entry(k_id, ym)
            try:
                _nas_settle(k_id, ym, led.src, e, False, MONTHLY_CAP)
            except _Offline:
                offline = True
                break
            e["reported"] = e["calls"]
            led.save()
        ent = led.entry(kid, month)
        if not offline:
            try:
                grant, total = _nas_settle(kid, month, led.src, ent, True, MONTHLY_CAP)
            except _Offline:
                offline = True
            else:
                ent["reported"] = ent["calls"]
                if grant:
                    ent["calls"] += 1
                    ent["reported"] = ent["calls"]
                led.prune(month)
                led.save()
                if not grant:
                    raise KiprisError("이번 달 KIPRIS 호출이 상한에 닿았습니다 (%d/%d)." % (total, MONTHLY_CAP))
                return
        base = _known_total.get((kid, month))
        used = ent["calls"] if base is None else base + (ent["calls"] - ent["reported"])
        if ent["full"] or used >= MONTHLY_CAP:
            raise KiprisError("이번 달 KIPRIS 호출이 상한에 닿았습니다 (%d/%d)." % (used, MONTHLY_CAP))
        ent["calls"] += 1
        led.save()


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
    if code == "20":
        raise KiprisError("KIPRIS 가 접근을 거부했습니다 (결과 코드 20) — 이 키가 막혔을 수 있어 호출을 멈춥니다.")
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
_proc_lock = threading.Lock()
_last_ts = 0.0


def _acquire():
    """이 프로세스의 간격(MIN_INTERVAL) 뒤에 이 컴퓨터 전체의 속도 문턱(규약 5절)을 거친다."""
    global _last_ts
    while True:
        with _proc_lock:
            now = time.monotonic()
            wait = _last_ts + MIN_INTERVAL - now
            if wait <= 0:
                _last_ts = now
                break
        time.sleep(wait)
    _gate_wait()


def _gate_reserve(clock=time.time) -> float:
    """gate 파일에 다음 자리 시각을 받아 적는다(잠금 안). 기다리지는 않는다."""
    interval = 1.0 / SAFE_MAX_RATE
    path = _guard_dir() / "gate"
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "a+b") as fh:
        _flock(fh)
        try:
            fh.seek(0)
            parts = (fh.read(80) or b"").split()
            now = clock()
            try:
                last = float(parts[0]) if parts else 0.0
                written = float(parts[1]) if len(parts) > 1 else None
            except ValueError:
                last, written = 0.0, now
            if written is None:
                base = min(last, now + 5.0)                     # 예전 판이 적은 값
            elif written > now + REWIND_SLACK:
                base = now - interval                           # 시계가 되감겼다 — 옛 예약을 버린다
            else:
                base = last                                     # 정상 대기열 — 자르지 않는다
            at = max(now, base + interval)
            fh.seek(0)
            fh.truncate()
            fh.write(b"%.6f %.6f" % (at, now))
            fh.flush()
        finally:
            _funlock(fh)
    return at


def _gate_wait(clock=time.time, sleep=time.sleep) -> None:
    at = _gate_reserve(clock)
    while True:
        left = at - clock()
        if left <= 0:
            return
        sleep(left)


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
        _check_not_blocked()
        _count_call()
        _acquire()
        _check_not_blocked()                            # 문턱을 기다리는 사이 누가 막힘을 남겼을 수 있다
        try:
            with _opener.open(url, timeout=TIMEOUT) as res:
                text = res.read().decode("utf-8", errors="replace")
            if re.search(r"<resultCode>\s*20\s*</resultCode>", text, re.I):
                _mark_blocked(f"접근 거부(결과 코드 20) — TMProject {path}")
            return text
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
