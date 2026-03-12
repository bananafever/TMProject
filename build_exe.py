import os
import shutil
import PyInstaller.__main__

def build_exe():
    # 1. 빌드 설정
    entry_point = "main.py"  # 실행 파일의 시작점
    exe_name = "TaskManager"   # 생성될 EXE 파일 이름
    icon_file = "attach_down.ico" # EXE 아이콘
    
    # 추가 데이터 파일 (이미지 등)
    # 형식: "소스파일;대상폴더" (Windows 기준)
    datas = [
        ("icon.png", "."),
    ]
    
    # 숨겨진 임포트 (발견되지 않는 모듈 강제 포함)
    hidden_imports = [
        "win32timezone", # Outlook 연동시 필요할 수 있음
    ]

    # 2. PyInstaller 인자 구성
    args = [
        entry_point,
        "--name=" + exe_name,
        "--onefile",       # 단일 파일로 빌드
        "--noconsole",     # GUI 앱인 경우 콘솔 창 숨김
        "--clean",         # 빌드 전 캐시 삭제
    ]

    # 아이콘 추가
    if os.path.exists(icon_file):
        args.append(f"--icon={icon_file}")

    # 데이터 파일 추가
    for src, dst in datas:
        if os.path.exists(src):
            args.append(f"--add-data={src};{dst}")

    # 숨겨진 임포트 추가
    for imp in hidden_imports:
        args.append(f"--hidden-import={imp}")

    # 3. 빌드 실행
    print(f"Building {exe_name}...")
    PyInstaller.__main__.run(args)
    
    # 4. 마무리 (선택 사항: build 폴더 등 정리)
    print("\nBuild complete!")
    print(f"The executable can be found in the 'dist' folder.")

if __name__ == "__main__":
    # 필요한 라이브러리 설치 확인 (선택 사항)
    # os.system("pip install pyinstaller")
    
    build_exe()
