"""받은 지시 — 받은편지함의 고객 지시 메일을 골라 .msg(Graph 는 .eml)로 두고 NAS `instructions` 에 적는다. Qt·네이버를 모른다.

CPProject `docs/INSTRUCTION_MAIL_PLAN.md` 6절. 확정(사건·서신으로 넣기·할 일)은 ChekPat 「받은 지시」 가 한다.

- 고객·해외 대리인 메일은 회사 대표 메일함(「Lee International IP & Law」)이 `FW:` 로 전달하고, 본문 첫 줄에
  `LEEDM:PA-2T3228` 꼬리표를 붙인다. 우리가 보낸 서신의 BCC 사본도 같은 보낸 사람·같은 꼴 제목으로 오지만
  `FW:` 가 없고 꼬리표가 `LEEDM:…#F` 다 → 고르지 않는다(10-24 윈도 PC 실측).
- 받은편지함만 본다(사용자 10-24). 메일을 보내지·옮기지·고치지 않는다(범주·플래그·읽음 표시도).
- 보낸 사람은 main 이 Restrict 필터로 가린다 — 주소 속성을 읽으면 Outlook 「전자 메일 주소 정보에 액세스」 경고가 뜬다.
  같은 까닭으로 Message-ID(`PR_INTERNET_MESSAGE_ID`)도 아직 읽지 않는다(경고가 안 뜨는지 본 뒤에).
- 대응기한은 다루지 않는다 — 공식 대응기한은 KIPRIS 로 정한 DB 값이다(사용자 기한 원칙 10-24).
- NAS 에 못 쓰면 `instruction_state.json` 에 쌓았다가 다음에 쓴다(`db.record`, 받은 위임과 같음).
- 트레이 알림은 고객 글에 연장 낱말이 있을 때만(사용자 10-24) — 「기연 지시일 수 있음」.
"""

import logging
import os
import re
import tempfile
from typing import List, Optional, Tuple

import db

DEFAULT_SENDERS = ["Lee International IP & Law"]      # 고객 메일을 전달하는 대표 메일함(settings.json instruction_senders)
STATE_NAME = "instruction_state.json"
MSG_DIR = "instructions"                               # 자료 폴더 아래 — 두 PC 의 Drive(맥 ChekPat 은 to_local)
BODY_LIMIT = 60000
OL_MSG_UNICODE = 9
OL_EMBEDDED_ITEM = 5
_PROP = "http://schemas.microsoft.com/mapi/proptag/"
PR_ATTACH_CONTENT_ID = _PROP + "0x3712001F"
PR_ATTACHMENT_HIDDEN = _PROP + "0x7FFE000B"

# main.ref_num_key 와 같은 Ref 규칙 — 앞뒤에 영문·숫자가 붙은 것은 아니다
_REF = (r"[a-zA-Z]{2,3}-[tT]?2[a-zA-Z][0-9]{4}[a-zA-Z0-9-]{0,5}|[a-zA-Z]{2,3}-[0-9]{5}")
REF = re.compile(r"(?<![A-Za-z0-9])(%s)(?![A-Za-z0-9])" % _REF)
LEEDM = re.compile(r"LEEDM\s*:\s*(%s)(#\w+)?" % _REF, re.I)
OUR_LABEL = re.compile(r"(?:Lee|Our)\s*Ref(?:erence)?\.?\s*[:.]?\s*(%s)" % _REF, re.I)
FORWARD_SUBJECT = re.compile(r"^\s*(fw|fwd|전달)\s*:", re.I)
SKIP_SUBJECT = re.compile(r"\[위임\]|사건 담당자 지정|automatic reply|auto(matic)?[- ]?reply|out of office|자동 회신|부재",
                          re.I)
EXTENSION = re.compile(r"extension|extend|more\s+time|additional\s+time|연장|기연", re.I)   # ChekPat extension 과 같음
_HEAD = re.compile(r"^\s*(From|보낸\s*사람)\s*:", re.I)


def _first_lines(body: str, n: int = 3) -> List[str]:
    return [ln.strip() for ln in (body or "").splitlines() if ln.strip()][:n]


def leedm_ref(body: str) -> Tuple[str, str]:
    """본문 첫머리의 `LEEDM:<Ref>` → (Ref, 꼬리). 꼬리 `#F` 는 우리가 보낸 서신. 없으면 ("", "")."""
    for line in _first_lines(body):
        m = LEEDM.search(line)
        if m:
            return m.group(1).upper(), (m.group(2) or "")
    return "", ""


def ref_hint(subject: str, body: str) -> str:
    """우리 Ref — LEEDM 꼬리표, 제목의 Lee/Our Ref 라벨 뒤, 제목의 첫 Ref 순. 확정은 ChekPat."""
    ref, tail = leedm_ref(body)
    if ref and not tail:
        return ref
    m = OUR_LABEL.search(subject or "")
    if m:
        return m.group(1).upper()
    m = REF.search(subject or "")
    return m.group(1).upper() if m else ""


def choose(subject: str, body: str) -> str:
    """대표 메일함이 보낸 메일(보낸 사람·받은편지함은 main 이 본다) 가운데 받은 지시면 Ref, 아니면 "".

    전달이어야 한다 — 제목 `FW:`·`Fwd:`·`전달:`, 또는 첫머리 `LEEDM:<Ref>`(꼬리 없음). 위임·예고·자동 회신은 뺀다.
    """
    if SKIP_SUBJECT.search(subject or ""):
        return ""
    ref, tail = leedm_ref(body)
    if not (FORWARD_SUBJECT.match(subject or "") or (ref and not tail)):
        return ""                                       # 우리 서신 BCC 사본(FW: 없음 · LEEDM:…#F) 등
    return ref_hint(subject, body)


def client_part(body: str) -> str:
    """고객 글 — 첫 전달 머리부터 다음 전달 머리 앞까지(그 아래는 우리가 보냈던 서신). 머리가 없으면 본문 전체."""
    lines = (body or "").splitlines()
    heads = [i for i, ln in enumerate(lines) if _HEAD.match(ln)]
    if not heads:
        return body or ""
    end = heads[1] if len(heads) > 1 else len(lines)
    return "\n".join(lines[heads[0]:end])


def wants_alert(subject: str, body: str) -> bool:
    """트레이 알림 — 고객 글에 연장 낱말이 있을 때만(인용된 우리 서신의 낱말은 세지 않는다)."""
    return bool(EXTENSION.search(client_part(body)))


def msg_file(root: str, received, ref: str, entry_id: str, ext: str = ".msg") -> str:
    """`<자료 폴더>\\instructions\\YYYY-MM\\YYYYMMDD-HHMM_<Ref>_<EntryID 끝 8자>.msg` — Graph 로 읽은 메일은 `.eml`."""
    safe = re.sub(r"[^A-Za-z0-9-]", "", ref or "") or "NOREF"
    return os.path.join(root, MSG_DIR, received.strftime("%Y-%m"),
                        "%s_%s_%s%s" % (received.strftime("%Y%m%d-%H%M"), safe, (entry_id or "")[-8:], ext))


def _prop(att, tag):
    try:
        return att.PropertyAccessor.GetProperty(tag)
    except Exception:                                   # noqa: BLE001 — 그 속성이 없는 첨부
        return None


def attachments(message) -> Tuple[List[str], list]:
    """(첨부 이름 — 본문에 박힌 그림·숨김 빼고, 항목으로 붙은 원 메일들)."""
    try:
        html = (message.HTMLBody or "").casefold()
    except Exception:                                   # noqa: BLE001
        html = ""
    names, embedded = [], []
    for att in message.Attachments:
        try:
            if att.Type == OL_EMBEDDED_ITEM:
                embedded.append(att)
            if _prop(att, PR_ATTACHMENT_HIDDEN):
                continue
            cid = str(_prop(att, PR_ATTACH_CONTENT_ID) or "").strip("<>")
            if cid and ("cid:" + cid).casefold() in html:
                continue                                # 서명·로고 그림
            names.append(str(att.FileName))
        except Exception as e:                          # noqa: BLE001 — 첨부 하나 못 읽어도 메일은 적는다
            logging.warning(f"받은 지시 첨부 읽기 실패: {e}")
    return names, embedded


def read_embedded(att, session) -> Tuple[str, str]:
    """항목으로 붙은 원 메일 → (본문, 머리 「제목: …」·「보낸 날짜: …」). 보낸 사람은 읽지 않는다(보안 경고).
    ChekPat 은 저장한 .msg 에서 원 메일 첨부를 다시 열 수 있다."""
    tmp = os.path.join(tempfile.mkdtemp(prefix="tm_instr_"), "orig.msg")
    try:
        att.SaveAsFile(tmp)
        item = session.OpenSharedItem(tmp)
        head = ["제목: %s" % (item.Subject or "")]
        try:
            head.append("보낸 날짜: %s" % item.SentOn.strftime("%Y-%m-%d %H:%M"))
        except Exception:                               # noqa: BLE001
            pass
        body = str(item.Body or "")[:BODY_LIMIT]
        try:
            item.Close(1)                               # olDiscard
        except Exception:                               # noqa: BLE001
            pass
        return body, "\n".join(head)
    except Exception as e:                              # noqa: BLE001
        logging.warning(f"받은 지시 — 붙은 원 메일을 열지 못했습니다: {e}")
        return "", ""
    finally:
        try:
            os.remove(tmp)
            os.rmdir(os.path.dirname(tmp))
        except OSError:
            pass


def received_of(message):
    try:
        return message.ReceivedTime
    except Exception:                                   # noqa: BLE001
        return message.CreationTime


def build_row(message, ref: str, msg_path: str, from_name: str, session=None, source: str = "outlook") -> dict:
    """`instructions` 한 줄(db.INSTRUCTION_COLUMNS). 원 발신인·고객 날짜는 ChekPat 이 본문에서 푼다."""
    names, embedded = attachments(message)
    orig_text, orig_head = read_embedded(embedded[0], session) if embedded and session is not None else ("", "")
    return {
        "mail_entry_id": message.EntryID,
        "message_id": None,
        "received_at": received_of(message).strftime("%Y-%m-%d %H:%M:%S"),
        "from_name": from_name,
        "subject": message.Subject,
        "ref_hint": ref or None,
        "attachments": "; ".join(names) or None,
        "body_text": str(message.Body or "")[:BODY_LIMIT],
        "orig_text": orig_text or None,
        "orig_head": orig_head or None,
        "msg_path": msg_path,
        "source": source,
    }


def handle(message, ref: str, data_root: str, from_name: str, session=None, connect_fn=None,
           source: str = "outlook") -> Optional[dict]:
    """고른 메일 하나 — .msg 저장 + NAS 한 줄(못 쓰면 쌓기). 이미 적은 메일이면 None.
    source "graph" 는 새 Outlook(graph_mail) — 원본을 .eml 로 둔다(ChekPat 은 둘 다 읽는다).

    반환값: {"ref", "alert"(연장 낱말), "result"(db.record)}
    """
    state = os.path.join(data_root, STATE_NAME)
    if db.is_recorded(state, message.EntryID):
        return None
    received = received_of(message)
    path = msg_file(data_root, received, ref, message.EntryID, ".eml" if source == "graph" else ".msg")
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        if not os.path.exists(path):
            message.SaveAs(path, OL_MSG_UNICODE)
    except Exception as e:                              # noqa: BLE001 — .msg 를 못 둬도 줄은 적는다(ChekPat 은 본문으로)
        logging.error(f"받은 지시 .msg 저장 실패: {path} — {e}")
        path = ""
    row = build_row(message, ref, path or None, from_name, session, source)
    result = db.record([row], state, connect_fn=connect_fn, table=db.INSTRUCTIONS)
    alert = wants_alert(message.Subject, row["body_text"])
    logging.info(f"받은 지시: {ref or '-'} · {message.Subject}" + (" · 연장 낱말" if alert else ""))
    return {"ref": ref, "alert": alert, "result": result}
