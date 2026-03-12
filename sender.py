import win32com.client as win32


def send_test_email(to, subject, body):
    """Outlook을 통해 테스트 이메일을 발송합니다.

    Args:
        to: 수신자 이메일 주소
        subject: 이메일 제목
        body: 이메일 본문
    """
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0은 메일 항목을 나타냅니다.

    mail.To = to
    mail.Subject = subject
    mail.Body = body

    try:
        mail.Send()
        print("이메일이 성공적으로 발송되었습니다.")
    except Exception as e:
        print(f"이메일 발송 중 오류가 발생했습니다: {e}")


if __name__ == "__main__":
    send_test_email(
        to="ysshin@leeinternational.com",
        subject="[위임] PJA-2R8603; 한송희(고객지원팀-책임); 거절결정서 [서 2024-12-23]; 2412130993",
        body="This is a test email sent using Python and Outlook.",
    )
