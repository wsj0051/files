#!/data/data/com.termux/files/usr/bin/bash
set -e

# ================== 配置区 ==================

# 源文件 / 目录
SOURCES=(
  "/data/data/com.termux/files/home/html/video_op.html"
  "/data/data/com.termux/files/home/html/index_op.html"
)

# 需要忽略的目录名
IGNORE_DIRS=(
  ".git"
  "node_modules"
  ".svn"
  ".ps"
)

# 目标根目录（第一个参数，不传则当前目录）
TARGET_ROOT="${1:-.}"

# ================== 构建 prune 参数（数组方式） ==================

FIND_PRUNE=()
for IG in "${IGNORE_DIRS[@]}"; do
  FIND_PRUNE+=( -path "*/$IG" -o )
done

# ================== 查找最里层目录 ==================
# 定义：没有子目录的目录

LEAF_DIRS=$(
  find "$TARGET_ROOT" \
    \( "${FIND_PRUNE[@]}" -false \) -prune -o \
    -type d \
    ! -exec sh -c 'ls -1 "$1"/*/ >/dev/null 2>&1' _ {} \; \
    -print
)

# ================== 拷贝逻辑 ==================

for DIR in $LEAF_DIRS; do
  for SRC in "${SOURCES[@]}"; do
    if [[ ! -e "$SRC" ]]; then
      echo "⚠ 源不存在: $SRC"
      continue
    fi

    BASENAME=$(basename "$SRC")
    DEST="$DIR/$BASENAME"

    cp -r "$SRC" "$DEST"

    echo "✔ 拷贝 $SRC -> $DEST"
  done
done
