#!/bin/bash
export GEMINI_API_KEY="AIzaSyBdU6fo7jzlzCdVneJuSGb_z4CKmm5d4t8"
PYTHON="/Users/2dev/youtubelive-capture/venv/bin/python"
SCRIPT="/Users/2dev/youtubelive-capture/analyze.py"
URL="https://www.youtube.com/watch?v=Vq3jz0c7oMs"
COMMON="--segment-duration 10 --screenshot-interval 60 --skip-analysis --whisper-model base"

echo "$(date): === 前半 0-180分 開始 ==="
$PYTHON $SCRIPT --url "$URL" --start 0 --duration 180 $COMMON \
    --output /Users/2dev/youtubelive-capture/output/6h_first
echo "$(date): === 前半 完了 ==="

echo "$(date): === 後半 180-360分 開始 ==="
$PYTHON $SCRIPT --url "$URL" --start 180 --duration 180 $COMMON \
    --output /Users/2dev/youtubelive-capture/output/6h_latter
echo "$(date): === 後半 完了 ==="

echo "$(date): === 全6時間処理完了 ==="
