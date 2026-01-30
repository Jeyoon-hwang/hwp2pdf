#!/bin/bash
#
# HWP/HWPX → PDF 대량 변환 스크립트
# 사용법:
#   ./hwp2pdf.sh [입력폴더] [출력폴더]
#
# 입력폴더를 지정하지 않으면 현재 디렉토리의 HWP/HWPX 파일을 변환합니다.
# 출력폴더를 지정하지 않으면 입력폴더와 같은 위치에 PDF를 생성합니다.

set -euo pipefail

SOFFICE="/Applications/LibreOffice.app/Contents/MacOS/soffice"

INPUT_DIR="${1:-.}"
OUTPUT_DIR="${2:-}"

# 입력 폴더 확인
if [ ! -d "$INPUT_DIR" ]; then
    echo "오류: 입력 폴더를 찾을 수 없습니다: $INPUT_DIR"
    exit 1
fi

# 출력 폴더 설정
if [ -z "$OUTPUT_DIR" ]; then
    OUTPUT_DIR="$INPUT_DIR"
fi
mkdir -p "$OUTPUT_DIR"

# 절대 경로로 변환
INPUT_DIR="$(cd "$INPUT_DIR" && pwd)"
OUTPUT_DIR="$(cd "$OUTPUT_DIR" && pwd)"

# LibreOffice 확인
if [ ! -x "$SOFFICE" ]; then
    echo "오류: LibreOffice가 설치되어 있지 않습니다."
    echo "  brew install --cask libreoffice"
    exit 1
fi

# HWP/HWPX 파일 수집
FILES=()
while IFS= read -r -d '' f; do
    FILES+=("$f")
done < <(find "$INPUT_DIR" -maxdepth 1 \( -iname "*.hwp" -o -iname "*.hwpx" \) -print0 | sort -z)

TOTAL=${#FILES[@]}

if [ "$TOTAL" -eq 0 ]; then
    echo "변환할 HWP/HWPX 파일이 없습니다: $INPUT_DIR"
    exit 0
fi

echo "========================================="
echo " HWP → PDF 대량 변환"
echo "========================================="
echo " 입력: $INPUT_DIR"
echo " 출력: $OUTPUT_DIR"
echo " 파일 수: $TOTAL"
echo "========================================="
echo ""

SUCCESS=0
FAIL=0
FAILED_FILES=()

for i in "${!FILES[@]}"; do
    FILE="${FILES[$i]}"
    BASENAME="$(basename "$FILE")"
    NUM=$((i + 1))

    printf "[%d/%d] %s ... " "$NUM" "$TOTAL" "$BASENAME"

    if "$SOFFICE" --headless --convert-to pdf --outdir "$OUTPUT_DIR" "$FILE" > /dev/null 2>&1; then
        echo "완료"
        SUCCESS=$((SUCCESS + 1))
    else
        echo "실패"
        FAIL=$((FAIL + 1))
        FAILED_FILES+=("$BASENAME")
    fi
done

echo ""
echo "========================================="
echo " 변환 결과"
echo "========================================="
echo " 성공: $SUCCESS / $TOTAL"
echo " 실패: $FAIL / $TOTAL"

if [ "$FAIL" -gt 0 ]; then
    echo ""
    echo " 실패한 파일:"
    for f in "${FAILED_FILES[@]}"; do
        echo "   - $f"
    done
fi

echo "========================================="
