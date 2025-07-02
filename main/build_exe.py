# build_exe.py - 실행파일 생성 스크립트
import PyInstaller.__main__
import sys
import os

def build_executable():
    """실행파일 생성"""
    
    # PyInstaller 옵션
    pyinstaller_args = [
        '--onefile',  # 단일 실행파일 생성
        '--windowed',  # 콘솔 창 숨기기 (GUI 앱)
        '--name=한글PDF변환기',  # 실행파일 이름
        '--icon=icon.ico',  # 아이콘 파일 (선택사항)
        '--add-data=icon.ico;.',
        '--distpath=dist',  # 출력 폴더
        '--workpath=build',  # 임시 작업 폴더
        '--clean',  # 이전 빌드 파일 정리
        '--noconfirm',  # 확인 없이 진행
        'hwp_to_pdf_gui.py'  # 메인 파이썬 파일
    ]
    
    print("실행파일을 생성합니다...")
    print("이 과정은 몇 분 정도 소요될 수 있습니다.")
    
    try:
        PyInstaller.__main__.run(pyinstaller_args)
        print("\n✅ 실행파일 생성 완료!")
        print("생성된 파일: dist/한글PDF변환기.exe")
        print("\n이 파일을 다른 PC에 복사하여 사용할 수 있습니다.")
        print("(파이썬이 설치되지 않은 PC에서도 실행 가능합니다)")
        
    except Exception as e:
        print(f"❌ 실행파일 생성 실패: {e}")
        return False
    
    return True

if __name__ == "__main__":
    build_executable()