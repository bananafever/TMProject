from datetime import datetime

from dateutil.relativedelta import relativedelta

# 두 날짜 문자열
filing_date = "2020-05-26"
req_of_exam_date = "2021-10-19"
receipt_of_allow_date = "2024-12-17"
OA_issued_date = "2024-04-25"
OA_deadline = "2024-06-25"

# 문자열을 datetime 객체로 변환
date_format = "%Y-%m-%d"
parsed_date1 = datetime.strptime(filing_date, date_format)
parsed_date2 = datetime.strptime(req_of_exam_date, date_format)

# date1에 4년 추가, date2에 3년 추가
date1_plus_4yrs = parsed_date1 + relativedelta(years=4)
date2_plus_3yrs = parsed_date2 + relativedelta(years=3)

# 두 날짜 비교하여 늦은 날짜를 선택
reference_date = max(date1_plus_4yrs, date2_plus_3yrs)

print(f"기준일: {reference_date}")

allowed_date_parsed = datetime.strptime(receipt_of_allow_date, date_format)
delayed_period = allowed_date_parsed - reference_date

print(f"지연 기간: {delayed_period}")

parsed_OA_issued_date = datetime.strptime(OA_issued_date, date_format)
parsed_OA_deadline = datetime.strptime(OA_deadline, date_format)

responsible_period = parsed_OA_deadline - parsed_OA_issued_date
print(f"출원인 책임 지연기간: {responsible_period}")
