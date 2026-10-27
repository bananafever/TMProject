#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TaskManager Mac Application Build Script
macOS용 TaskManager.app 번들 및 TaskManager.dmg 설치 파일 빌더
"""

import argparse
import os
import shutil
import subprocess
import sys


def ensure_icns(icon_png: str = "icon.png", out_icns: str = "app_icon.icns") -> str:
    """macOS 내장 sips 와 iconutil을 사용해 PNG로부터 .icns 생성"""
    if os.path.exists(out_icns):
        return out_icns

    if not os.path.exists(icon_png):
        print(f"[!] 아이콘 원본 파일이 없습니다: {icon_png}")
        return ""

    print(f"[*] macOS 아이콘 파일 생성 중: {icon_png} -> {out_icns}")
    iconset_dir = "tmp_app_icon.iconset"
    try:
        os.makedirs(iconset_dir, exist_ok=True)
        sizes = [
            (16, "icon_16x16.png"),
            (32, "icon_16x16@2x.png"),
            (32, "icon_32x32.png"),
            (64, "icon_32x32@2x.png"),
            (128, "icon_128x128.png"),
            (256, "icon_128x128@2x.png"),
            (256, "icon_256x256.png"),
            (512, "icon_256x256@2x.png"),
            (512, "icon_512x512.png"),
        ]
        for sz, name in sizes:
            subprocess.run(
                ["sips", "-z", str(sz), str(sz), icon_png, "--out", os.path.join(iconset_dir, name)],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        subprocess.run(["iconutil", "-c", "icns", iconset_dir, "-o", out_icns], check=True)
        print(f"[✓] 아이콘 생성 완료: {out_icns}")
        return out_icns
    except Exception as e:
        print(f"[!] .icns 생성 실패 ({e}). 기본 아이콘으로 진행합니다.")
        return ""
    finally:
        if os.path.exists(iconset_dir):
            shutil.rmtree(iconset_dir, ignore_errors=True)


def ad_hoc_sign(app_path: str) -> bool:
    """macOS 보안 정책(Gatekeeper)을 통과하기 위한 ad-hoc 코드 서명"""
    print(f"\n[*] ad-hoc 코드 서명 적용 중: {app_path}")
    try:
        cmd = ["codesign", "--force", "--deep", "-s", "-", app_path]
        res = subprocess.run(cmd, capture_output=True, text=True)
        if res.returncode == 0:
            print("[✓] 코드 서명 완료")
            return True
        else:
            print(f"[!] 코드 서명 경고: {res.stderr.strip()}")
            return False
    except Exception as e:
        print(f"[!] 코드 서명 중 오류 발생: {e}")
        return False


def create_dmg(app_path: str, dmg_path: str, volname: str = "TaskManager") -> bool:
    """macOS hdiutil을 사용해 드래그 앤 드롭 설치 가능한 .dmg 생성"""
    print(f"\n[*] DMG 설치 파일 생성 중: {dmg_path}")
    staging_dir = os.path.abspath("dist/_dmg_staging")
    if os.path.exists(staging_dir):
        shutil.rmtree(staging_dir, ignore_errors=True)
    os.makedirs(staging_dir, exist_ok=True)

    try:
        dest_app = os.path.join(staging_dir, os.path.basename(app_path))
        shutil.copytree(app_path, dest_app, symlinks=True)

        # /Applications 바로가기 심볼릭 링크 생성 (드래그 앤 드롭 설치 지원)
        apps_link = os.path.join(staging_dir, "Applications")
        os.symlink("/Applications", apps_link)

        if os.path.exists(dmg_path):
            os.remove(dmg_path)

        cmd = [
            "hdiutil", "create",
            "-volname", volname,
            "-srcfolder", staging_dir,
            "-ov",
            "-format", "UDZO",
            dmg_path,
        ]
        res = subprocess.run(cmd, capture_output=True, text=True)
        if res.returncode != 0:
            print(f"[!] DMG 생성 실패:\n{res.stderr.strip()}")
            return False

        size_mb = os.path.getsize(dmg_path) / (1024 * 1024)
        print(f"[✓] DMG 생성 완료: {dmg_path} ({size_mb:.1f} MB)")
        return True
    except Exception as e:
        print(f"[!] DMG 생성 중 오류 발생: {e}")
        return False
    finally:
        if os.path.exists(staging_dir):
            shutil.rmtree(staging_dir, ignore_errors=True)


def build_mac(onefile: bool = False, make_dmg: bool = True, clean: bool = True, skip_sign: bool = False):
    import PyInstaller.__main__

    # 1. 빌드 기본 설정
    entry_point = "main.py"
    app_name = "TaskManager"
    bundle_id = "com.tmproject.taskmanager"

    # 아이콘 준비 (.icns 없으면 자동 생성)
    icon_file = ensure_icns("icon.png", "app_icon.icns")

    # 추가 데이터 파일 (macOS/POSIX 구분자는 콜론 ':')
    datas = [
        ("icon.png", "."),
        ("assets/fonts", "assets/fonts"),
        ("assets/icons", "assets/icons"),
    ]

    # macOS에 필요한 숨겨진 임포트 (win32 제외)
    hidden_imports = [
        "pymysql",
        "pymysql.cursors",
        "msal",
        "msal_extensions",
        "keyring",
    ]

    # 2. PyInstaller 인자 구성
    args = [
        entry_point,
        f"--name={app_name}",
        "--windowed",       # macOS .app 번들 생성 (GUI 앱, 콘솔 숨김)
        "--noconfirm",
        f"--osx-bundle-identifier={bundle_id}",
    ]

    if onefile:
        args.append("--onefile")
    else:
        args.append("--onedir")

    if clean:
        args.append("--clean")

    # 아이콘 설정
    if icon_file and os.path.exists(icon_file):
        args.append(f"--icon={icon_file}")

    # 데이터 파일 추가 (macOS에서는 ':' 구분자 사용)
    for src, dst in datas:
        if os.path.exists(src):
            args.append(f"--add-data={src}{os.pathsep}{dst}")

    # 숨겨진 임포트 추가
    for imp in hidden_imports:
        args.append(f"--hidden-import={imp}")

    # 3. PyInstaller 빌드 실행
    print(f"[*] Building {app_name}.app for macOS...")
    print(f"[*] PyInstaller command args: {' '.join(args)}")
    PyInstaller.__main__.run(args)

    app_path = os.path.abspath(os.path.join("dist", f"{app_name}.app"))
    if not os.path.exists(app_path):
        print(f"\n[!] 빌드 실패: {app_path} 가 생성되지 않았습니다.")
        sys.exit(1)

    # 4. ad-hoc 서명
    if not skip_sign:
        ad_hoc_sign(app_path)

    # 5. DMG 생성
    dmg_path = os.path.abspath(os.path.join("dist", f"{app_name}.dmg"))
    if make_dmg:
        create_dmg(app_path, dmg_path, volname=app_name)

    # 6. 마무리 안내
    print("\n" + "=" * 60)
    print("🎉 macOS 빌드가 성공적으로 완료되었습니다!")
    print("=" * 60)
    print(f"• 앱 번들 (.app) : {app_path}")
    if make_dmg and os.path.exists(dmg_path):
        print(f"• DMG 파일 (.dmg) : {dmg_path}")
    print("\n[실행 및 배포 팁]")
    print("1. 로컬에서 바로 실행:")
    print(f"   open '{app_path}'")
    print("2. 다른 Mac에서 '손상되었거나 확인할 수 없는 개발자' 경고 발생 시 터미널에서 격리 해제:")
    print(f"   xattr -cr /Applications/{app_name}.app")
    print("=" * 60)


def main():
    parser = argparse.ArgumentParser(description="TaskManager macOS Application Builder")
    parser.add_argument("--onefile", action="store_true", help="단일 번들(onefile) 모드로 빌드 (기본: onedir)")
    parser.add_argument("--no-dmg", action="store_true", help="DMG 생성 건너뛰기")
    parser.add_argument("--no-clean", action="store_true", help="빌드 전 캐시 정리 건너뛰기")
    parser.add_argument("--no-sign", action="store_true", help="ad-hoc 코드 서명 건너뛰기")

    args = parser.parse_args()
    build_mac(
        onefile=args.onefile,
        make_dmg=not args.no_dmg,
        clean=not args.no_clean,
        skip_sign=args.no_sign,
    )


if __name__ == "__main__":
    main()
