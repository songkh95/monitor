import PyInstaller.__main__
import customtkinter
import os

# 1. CustomTkinter 라이브러리 경로 찾기 (이게 제일 중요!)
ctk_path = os.path.dirname(customtkinter.__file__)

# 2. PyInstaller 실행 설정
print("=== EXE 변환을 시작합니다 ===")
print(f" * CustomTkinter 경로 감지됨: {ctk_path}")

PyInstaller.__main__.run([
    'main.py',                       # 1) 변환할 메인 파일 이름
    '--name=PrintManager',           # 2) 생성될 EXE 파일 이름
    '--noconsole',                   # 3) 검은색 CMD 창 안 뜨게 하기 (중요!)
    '--onefile',                     # 4) 파일 하나로 합치기
    f'--add-data={ctk_path};customtkinter', # 5) 디자인 파일 강제 포함 (Windows용 문법)
    '--clean',                       # 6) 캐시 삭제 후 빌드
])

print("\n=== 변환 완료! ===")
print("dist 폴더 안에 생긴 PrintManager.exe 파일을 확인하세요.")