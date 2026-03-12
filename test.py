import win32com.client

# Outlook 애플리케이션 초기화
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# 기본 Calendar 폴더 가져오기
calendar_folder = namespace.GetDefaultFolder(9)

# Calendar 폴더의 하위 폴더 출력
for subfolder in calendar_folder.Folders:
    print(f"Folder Name: {subfolder.Name}")
