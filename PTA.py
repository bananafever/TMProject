from datetime import datetime

from dateutil.relativedelta import relativedelta


def calculate_pta(
    filing_date,
    req_of_exam_date,
    receipt_of_allow_date,
    OA_issued_date,
    OA_deadline,
):
    """특허 기간 조정(PTA) 관련 날짜를 계산합니다.

    Args:
        filing_date: 출원일 (YYYY-MM-DD)
        req_of_exam_date: 심사청구일 (YYYY-MM-DD)
        receipt_of_allow_date: 허가 수령일 (YYYY-MM-DD)
        OA_issued_date: OA 발행일 (YYYY-MM-DD)
        OA_deadline: OA 마감일 (YYYY-MM-DD)

    Returns:
        dict: 기준일, 지연 기간, 출원인 책임 지연기간
    """
    date_format = "%Y-%m-%d"

    parsed_filing = datetime.strptime(filing_date, date_format)
    parsed_exam = datetime.strptime(req_of_exam_date, date_format)

    # 출원일 + 4년, 심사청구일 + 3년 중 늦은 날짜를 기준일로 선택
    date1_plus_4yrs = parsed_filing + relativedelta(years=4)
    date2_plus_3yrs = parsed_exam + relativedelta(years=3)
    reference_date = max(date1_plus_4yrs, date2_plus_3yrs)

    print(f"기준일: {reference_date.strftime(date_format)}")

    allowed_date_parsed = datetime.strptime(receipt_of_allow_date, date_format)
    delayed_period = allowed_date_parsed - reference_date
    print(f"지연 기간: {delayed_period.days}일")

    parsed_OA_issued = datetime.strptime(OA_issued_date, date_format)
    parsed_OA_deadline = datetime.strptime(OA_deadline, date_format)
    responsible_period = parsed_OA_deadline - parsed_OA_issued
    print(f"출원인 책임 지연기간: {responsible_period.days}일")

    return {
        "reference_date": reference_date,
        "delayed_period": delayed_period,
        "responsible_period": responsible_period,
    }


if __name__ == "__main__":
    calculate_pta(
        filing_date="2020-05-26",
        req_of_exam_date="2021-10-19",
        receipt_of_allow_date="2024-12-17",
        OA_issued_date="2024-04-25",
        OA_deadline="2024-06-25",
    )
