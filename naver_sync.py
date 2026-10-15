"""Outlook 캘린더 → 네이버 캘린더 일괄 복사, 그리고 옮긴 Outlook 일정 정리. 제목·날짜만 다룬다(본문은 보내지 않음).

    python naver_sync.py [--since 2026-10-16]                            미리보기 — 「신규」/「이미 있음」
    python naver_sync.py [--since 2026-10-16] --apply                    네이버에 등록
    python naver_sync.py [--since 2026-10-16] --delete-outlook           미리보기 — 「삭제 예정」/「유지」
    python naver_sync.py [--since 2026-10-16] --delete-outlook --apply   네이버에 있는 일정만 Outlook 에서 삭제

- Outlook 기본 캘린더의 하위 캘린더(FOLDERS)를 이름이 같은 네이버 캘린더로 복사한다.
- 중복: UID 가 (캘린더·제목·날짜)로 정해져 앱 자동 등록 분과도 겹치지 않는다. 여러 번 돌려도 안전하다.
- 삭제는 네이버에 같은 일정(이 앱이 등록한 것)이 있는 경우에만 한다. 지운 일정은 Outlook 「지운 항목」 폴더로 간다.
"""

import argparse
import sys
from datetime import date

import naver_cal

FOLDERS = ["Mail", "Response", "No_Instructions", "Complete"]   # Outlook 하위 캘린더 = 네이버 캘린더 이름


def outlook_events(calendar, since, folders=FOLDERS):
    """[(폴더, 제목, 날짜, 종일 여부, item), ...] — since 이후(포함) 일정만, 폴더별 날짜순.

    Items.Restrict 의 날짜 형식은 로캘마다 달라서 쓰지 않고 여기서 거른다.
    """
    events = []
    for name in folders:
        items = calendar.Folders[name].Items
        items.Sort("[Start]")
        for item in items:
            day = item.Start.date()
            if day >= since:
                events.append((name, item.Subject, day, bool(item.AllDayEvent), item))
    return events


def copy_status(session, folder, subject, day, apply):
    if apply:
        return "추가됨" if session.add_all_day(folder, subject, day, "") else "건너뜀(중복)"
    return "이미 있음" if session.exists(folder, subject, day) else "신규"


def delete_status(session, folder, subject, day, item, apply):
    """네이버에 같은 일정이 있을 때만 Outlook 일정을 지운다(지운 항목 폴더로 이동)."""
    if item.IsRecurring:
        return "반복 일정 — 유지"
    if not session.exists(folder, subject, day):
        return "네이버에 없음 — 유지"
    if not apply:
        return "삭제 예정"
    item.Delete()
    return "삭제됨"


def main(argv=None):
    parser = argparse.ArgumentParser(description="Outlook 캘린더 → 네이버 캘린더 복사·정리")
    parser.add_argument("--since", type=date.fromisoformat, default=date.today(), help="이 날짜 이후 일정 (YYYY-MM-DD, 기본: 오늘)")
    parser.add_argument("--delete-outlook", action="store_true", help="복사 대신, 네이버로 옮긴 Outlook 일정 삭제")
    parser.add_argument("--apply", action="store_true", help="실제로 실행 (없으면 미리보기)")
    args = parser.parse_args(argv)

    import win32com.client
    calendar = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(9)
    events = outlook_events(calendar, args.since)
    session = naver_cal.Session()

    job = "Outlook 일정 삭제" if args.delete_outlook else "네이버 등록"
    print(f"{job} {'실행' if args.apply else '미리보기'}: {args.since} 이후 Outlook 일정 {len(events)}건\n")
    counts = {}
    current = None
    for folder, subject, day, all_day, item in events:
        if folder != current:
            current = folder
            print(f"## {folder}")
        try:
            if args.delete_outlook:
                status = delete_status(session, folder, subject, day, item, args.apply)
            else:
                status = copy_status(session, folder, subject, day, args.apply)
        except Exception as e:
            status = f"오류: {e}"
        key = status.split(":")[0]
        counts[key] = counts.get(key, 0) + 1
        print(f"  {day}  {subject}{'' if all_day else '  [시간→종일]'}  — {status}")

    print("\n합계: " + (", ".join(f"{k} {v}건" for k, v in counts.items()) or "0건"))
    if not args.apply and (counts.get("신규") or counts.get("삭제 예정")):
        print("실행하려면 --apply 를 붙여 다시 실행하세요.")
    if args.delete_outlook:
        print("※ 지운 일정은 Outlook 「지운 항목」 폴더에서 되살릴 수 있습니다. Exchange 등과 동기화된 계정이면 다른 기기에서도 지워집니다.")
    return 1 if counts.get("오류") else 0


if __name__ == "__main__":
    sys.exit(main())
