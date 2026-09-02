#!/data/data/com.termux/files/usr/bin/bash
set -e

# 默认目录是当前目录
ROOT_DIR="${1:-.}"

# 判断源目录是否存在
if [[ ! -d "$ROOT_DIR" ]]; then
  echo "目录不存在: $ROOT_DIR"
  exit 1
fi

# 查找所有最里层目录（没有子目录的目录）
LEAF_DIRS=$(find "$ROOT_DIR" -type d ! -exec sh -c 'ls -1 "$0"/*/ >/dev/null 2>&1' {} \; -print)

# 遍历每个最里层目录生成 media.json
for DIR in $LEAF_DIRS; do
  OUT="$DIR/media.json"
  echo "[" > "$OUT"

  first=1

  find "$DIR" -maxdepth 1 -type f \
    \( \
      -iname "*.jpg" -o -iname "*.jpeg" -o -iname "*.png" -o \
      -iname "*.gif" -o -iname "*.webp" -o -iname "*.bmp" -o \
      -iname "*.mp4" -o -iname "*.mkv" -o -iname "*.webm" -o \
      -iname "*.mov" -o -iname "*.avi" -o \
      -iname "*.mp3" -o -iname "*.flac" -o -iname "*.wav" -o \
      -iname "*.ogg" -o -iname "*.aac" -o -iname "*.m4a" -o \
      -iname "*.wma" -o -iname "*.aiff" -o -iname "*.aif" -o \
      -iname "*.opus" \
    \) \
    -printf "%f\n" \
  | sort -V \
  | while IFS= read -r file; do
      if [ $first -eq 1 ]; then
        first=0
      else
        echo "," >> "$OUT"
      fi
      printf '  "%s"' "$file" >> "$OUT"
    done

  echo "" >> "$OUT"
  echo "]" >> "$OUT"
  echo "✔ 已生成 $OUT"
done
