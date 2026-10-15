"""Outlook 캘린더 → 네이버 캘린더 일괄 복사. 제목·날짜만 옮긴다(본문은 보내지 않음).

    python naver_sync.py [--since 2026-10-16]           미리보기 — 목록과 「신규」/「이미 있음」만 보여 줌
    python naver_sync.py [--since 2026-10-16] --apply   네이버에 등록

- Outlook 기본 캘린더의 하위 캘린더(FOLDERS)를 이름이 같은 네이버 캘린더로 복사한다.
- 중복: UID 가 (캘린더·제목·날짜)로 정해져 앱 자동 등록 분과도 겹치지 않는다. 여러 번 돌려도 안전하다.
- 복사만 한다. Outlook 에서 옮기거나·날짜를 바꾸거나·지운 일정은 네이버에 그대로 남는다.
"""

import argparse
import sys
from datetime import date

import naver_cal

FOLDERS = ["Mail", "Response", "No_Instructions", "Complete"]   # Outlook 하위 캘린더 = 네이버 캘린더 이름


def outlook_events(calendar, since, folders=FOLDERS):
    """[(폴더, 제목, 날짜, 종일 여부), ...] — since 이후(포함) 일정만, 폴더별 날짜순.

    Items.Restrict 의 날짜 형식은 로캘마다 달라서 쓰지 않고 여기서 거른다.
    """
    events = []
    for name in folders:
        items = calendar.Folders[name].Items
        items.Sort("[Start]")
        for item in items:
            day = item.Start.date()
            if day >= since:
                events.append((name, item.Subject, day, bool(item.AllDayEvent)))
    return events


def main(argv=None):
    parser = argparse.ArgumentParser(description="Outlook 캘린더 → 네이버 캘린더 복사")
    parser.add_argument("--since", type=date.fromisoformat, default=date.today(), help="이 날짜 이후 일정 (YYYY-MM-DD, 기본: 오늘)")
    parser.add_argument("--apply", action="store_true", help="실제로 등록 (없으면 미리보기)")
    args = parser.parse_args(argv)

    import win32com.client
    calendar = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(9)
    events = outlook_events(calendar, args.since)
    session = naver_cal.Session()

    print(f"{'등록' if args.apply else '미리보기'}: {args.since} 이후 Outlook 일정 {len(events)}건\n")
    counts = {}
    current = None
    for folder, subject, day, all_day in events:
        if folder != current:
            current = folder
            print(f"## {folder}")
        try:
            if args.apply:
                status = "추가됨" if session.add_all_day(folder, subject, day, "") else "건너뜀(중복)"
            else:
                status = "이미 있음" if session.exists(folder, subject, day) else "신규"
        except naver_cal.NaverError as e:
            status = f"오류: {e}"
        counts[status.split(":")[0]] = counts.get(status.split(":")[0], 0) + 1
        print(f"  {day}  {subject}{'' if all_day else '  [시간→종일]'}  — {status}")

    print("\n합계: " + ", ".join(f"{k} {v}건" for k, v in counts.items()))
    if not args.apply and counts.get("신규"):
        print("등록하려면 --apply 를 붙여 다시 실행하세요.")
    return 1 if counts.get("오류") else 0


if __name__ == "__main__":
    sys.exit(main())
