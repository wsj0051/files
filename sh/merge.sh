#!/data/data/com.termux/files/usr/bin/bash

# 颜色
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
BLUE='\033[0;34m'
NC='\033[0m'

DIR="${1:-.}"

echo -e "${BLUE}扫描目录: $DIR${NC}"

# 检查 ffmpeg
if ! command -v ffmpeg &> /dev/null; then
    echo -e "${RED}错误: 未安装 ffmpeg${NC}"
    echo "安装命令: pkg install ffmpeg"
    exit 1
fi

shopt -s nullglob
found=0

for mp4 in "$DIR"/*.mp4; do
    base="${mp4%.mp4}"
    mp3="$base.mp3"
    
    if [ -f "$mp3" ]; then
        found=1
        output="${base}_merged.mp4"
        
        echo -e "${YELLOW}发现匹配文件:${NC}"
        echo "  视频: $mp4"
        echo "  音频: $mp3"
        
        ffmpeg -y \
            -i "$mp4" \
            -i "$mp3" \
            -c:v copy \
            -c:a aac \
            -map 0:v:0 \
            -map 1:a:0 \
            "$output"
        
        if [ $? -eq 0 ]; then
            echo -e "${GREEN}✓ 合并成功: $output${NC}"
        else
            echo -e "${RED}✗ 合并失败: $base${NC}"
        fi
        
        echo ""
    fi
done

shopt -u nullglob

if [ $found -eq 0 ]; then
    echo -e "${YELLOW}未找到可合并的 mp3 + mp4 文件${NC}"
else
    echo -e "${GREEN}处理完成！${NC}"
fi
