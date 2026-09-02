#!/data/data/com.termux/files/usr/bin/bash
set -e

# ================== 配置区 ==================

# 目标根目录（第一个参数，不传默认当前目录）
TARGET_ROOT="${1:-.}"

# 忽略的目录名（可多个）
IGNORE_DIRS=(
  ".git"
  "node_modules"
  ".svn"
)

# 是否真的删除（false = 预览，true = 真删）
DO_DELETE=true

# ================== 构建 find 的 prune 参数 ==================

FIND_PRUNE=()
for IG in "${IGNORE_DIRS[@]}"; do
  FIND_PRUNE+=( -path "*/$IG" -o )
done

# ================== 查找 html 文件 ==================

HTML_FILES=$(
  find "$TARGET_ROOT" \
    \( "${FIND_PRUNE[@]}" -false \) -prune -o \
    -type f -name "*.html" -print
)

# ================== 执行清理 ==================

if [[ -z "$HTML_FILES" ]]; then
  echo "✅ 未发现需要清理的 *.html 文件"
  exit 0
fi

echo "🧹 将处理以下 *.html 文件："
echo "$HTML_FILES"
echo "----------------------------------"

if [[ "$DO_DELETE" == true ]]; then
  echo "🚨 正在删除..."
  rm -f $HTML_FILES
  echo "✅ 清理完成"
else
  echo "👀 预览模式（未删除）"
fi
