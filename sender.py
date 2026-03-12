import win32com.client as win32

# Outlook 애플리케이션 시작
outlook = win32.Dispatch("Outlook.Application")

# 새 이메일 생성
mail = outlook.CreateItem(0)  # 0은 메일 항목을 나타냅니다.

# 이메일 정보 설정
mail.To = "ysshin@leeinternational.com"  # 수신자 이메일
mail.Subject = "[위임] PJA-2R8603; 한송희(고객지원팀-책임); 거절결정서 [서 2024-12-23]; 2412130993"  # 제목
mail.Body = "This is a test email sent using Python and Outlook."  # 본문

try:
    mail.Send()
    print("이메일이 성공적으로 발송되었습니다.")
except Exception as e:
    print(f"이메일 발송 중 오류가 발생했습니다: {e}")
