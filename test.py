import win32com.client


def list_calendar_subfolders():
    """Outlook Calendar 폴더의 하위 폴더 목록을 출력합니다."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    calendar_folder = namespace.GetDefaultFolder(9)

    for subfolder in calendar_folder.Folders:
        print(f"Folder Name: {subfolder.Name}")


if __name__ == "__main__":
    list_calendar_subfolders()
