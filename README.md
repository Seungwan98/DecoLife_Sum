
# ExcelNamePicker

엑셀에 있는 파일명 목록을 기준으로, 소스 폴더에서 같은 이름(옵션에 따라 확장자 무시 가능)의 파일을 찾아 출력 폴더로 복사합니다. 결과는 CSV 로그로 남습니다.

## 빠른 시작

```bash
# (선택) 가상환경 권장
python -m venv .venv
. .venv/Scripts/activate   # Windows PowerShell
# 또는
. .venv/bin/activate       # macOS/Linux

pip install -r requirements.txt
python main.py
```

## 옵션 설명
- **하위폴더까지 재귀적으로 검색**: 소스 폴더의 모든 하위 디렉토리까지 조회합니다.
- **대소문자 무시**: 이름 비교 시 대소문자를 무시합니다.
- **확장자 무시(이름만 매칭)**: 엑셀의 `report`와 소스의 `report.pdf`를 동일하게 간주합니다.
- **덮어쓰기**: 출력 폴더에 같은 파일명이 있을 때 덮어쓸지 여부(기본은 `(1)`, `(2)` 번호를 붙여 충돌 방지).

## 빌드(.exe) 방법 - Windows
PyInstaller 사용:

```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --name "ExcelNamePicker" main.py
```

빌드 결과(`dist/ExcelNamePicker.exe`)를 배포하면 됩니다.

> 참고: pandas/openpyxl 포함으로 인해 exe 용량이 비교적 큽니다. 용량을 더 줄이고 싶다면 `openpyxl`만으로 엑셀을 읽도록 커스텀하거나 `polars`로 대체하는 등 경량화가 가능합니다.
