# HWP2PDF - HWP/HWPX → PDF 대량 변환기

HWP 및 HWPX 파일을 PDF로 대량 변환하는 도구입니다.
GUI 앱과 CLI 스크립트 두 가지 방식을 제공합니다.

## 요구사항

- macOS
- [LibreOffice](https://www.libreoffice.org/) 설치 필요

```bash
brew install --cask libreoffice
```

## GUI 앱 사용법

```bash
python3 hwp2pdf_gui.py
```

### 기능
- 파일/폴더 단위로 HWP/HWPX 파일 추가
- 출력 폴더 지정 (미지정 시 원본과 같은 위치)
- 진행률 바 및 실시간 로그
- 변환 결과 요약 (성공/실패 목록)

### 실행 파일 빌드

```bash
pip install pyinstaller
python3 -m PyInstaller --onefile --windowed --name "HWP2PDF" hwp2pdf_gui.py
```

빌드 결과물은 `dist/HWP2PDF.app`에 생성됩니다.

## CLI 스크립트 사용법

```bash
# 현재 폴더의 모든 HWP/HWPX 파일 변환
./hwp2pdf.sh

# 입력 폴더 지정
./hwp2pdf.sh /path/to/hwp_files

# 입력 폴더 + 출력 폴더 지정
./hwp2pdf.sh /path/to/hwp_files /path/to/pdf_output
```

## 라이선스

MIT
