#!/data/data/com.termux/files/usr/bin/bash

# 配置
#DOWNLOAD_DIR=~/storage/shared/Downloads
DOWNLOAD_DIR=~/storage/shared/DCIM
#DOWNLOAD_DIR=~/storage/shared/Android/media/com.tencent.mm/WAuxiliary/Cache
# cookies.txt 需要使用浏览器插件导出，路径自行调整
#COOKIES=~/storage/shared/Documents/cookies.txt
COOKIES=~/storage/shared/DCIM/all_cookies.txt
# 代理
PROXY="http://127.0.0.1:7890"
#PROXY=""
LOG_FILE="$DOWNLOAD_DIR/yt-dlp.log"
TEMP_DIR="$DOWNLOAD_DIR/temp"

# yt-dlp 下载路径记录文件（临时工作文件，供最终结果精准定位刚下载的文件）
YT_PATH_FILE="$TEMP_DIR/yt_downloaded_path.txt"

# ==================== 统一日志 ====================
log_debug() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] $$ $*" >> "$LOG_FILE"
}

# 去除变量中的 \r (Windows CRLF 兼容)
strip_cr() {
    echo "${1//$'\r'/}"
}

# ==================== URL 清洗 ====================
# 自动去除网页分享链接自带的追踪参数（spm_id_from/vd_source/utm_*/share_* 等），
# 保留关键参数（v/list/p/t/index 等）。这样即使粘贴了带一长串参数的链接也能干净下载
clean_url() {
    local url="$1"
    local base query part key clean_query=""
    local -a parts
    case "$url" in
        *\?*) ;;
        *) echo "$url"; return 0 ;;
    esac
    base="${url%%\?*}"
    query="${url#*\?}"
    IFS='&' read -r -a parts <<< "$query"
    for part in "${parts[@]}"; do
        key="${part%%=*}"
        case "$key" in
            spm_id_from|vd_source|utm_source|utm_medium|utm_campaign|utm_term|utm_content|\
            from|share_source|share_medium|share_plat|share_tag|bbid|ts|timestamp|unique_k|\
            search_id|click_from|seid|rdid|source|hotword|qr_login)
                continue ;;
        esac
        [ -n "$part" ] && clean_query="${clean_query:+${clean_query}&}${part}"
    done
    if [ -n "$clean_query" ]; then
        echo "${base}?${clean_query}"
    else
        echo "$base"
    fi
}

# ==================== 文件名格式化 ====================
format_video_name() {
    local raw="$1"
    raw=$(echo "$raw" | sed 's/^[[:space:]]*//;s/[[:space:]]*$//')
    # 清理所有空白：普通空白 + Unicode 特殊空白（不换行空格 U+00A0、全角空格 U+3000、
    # 零宽字符 U+200B~U+200F、U+2060、BOM U+FEFF）。Termux 默认 locale 下 [[:space:]] 匹配不到
    raw=$(echo "$raw" | sed -e 's/[[:space:]]/_/g' \
        -e 's/\xC2\xA0/_/g' \
        -e 's/\xE3\x80\x80/_/g' \
        -e 's/\xE2\x80\x8B//g' -e 's/\xE2\x80\x8C//g' -e 's/\xE2\x80\x8D//g' \
        -e 's/\xE2\x80\x8E//g' -e 's/\xE2\x80\x8F//g' -e 's/\xE2\x81\xA0//g' \
        -e 's/\xEF\xBB\xBF//g')
    raw=$(echo "$raw" | sed 's|[\/\\:*?"<>|]||g')
    raw=$(echo "$raw" | sed -e 's/!//g' -e 's/@//g' -e 's/#//g' -e 's/\$//g' \
        -e 's/%//g' -e 's/\^//g' -e 's/&//g' -e 's/\*//g' \
        -e 's/(//g' -e 's/)//g' -e 's/+//g' -e 's/=//g' \
        -e 's/{//g' -e 's/}//g' -e 's/\[//g' -e 's/\]//g' \
        -e "s/'//g" -e 's/;//g' -e 's/,//g' -e 's/~//g' -e 's/`//g')
    raw=$(echo "$raw" | sed 's/__*/_/g')
    raw=$(echo "$raw" | cut -c1-80)
    raw=$(echo "$raw" | sed 's/[._-]$//')
    [ -z "$raw" ] && raw="video_$(date +%s)"
    echo "$raw"
}

# 启动时记录环境
mkdir -p "$DOWNLOAD_DIR" 2>/dev/null
log_debug "========== yt.sh 启动 =========="
log_debug "参数: ARGV=($*)"
log_debug "DOWNLOAD_DIR=$DOWNLOAD_DIR (exists=$(test -d "$DOWNLOAD_DIR" && echo yes || echo no))"
log_debug "yt-dlp=$(command -v yt-dlp 2>/dev/null || echo 'NOT FOUND')"
log_debug "ffmpeg=$(command -v ffmpeg 2>/dev/null || echo 'NOT FOUND')"
log_debug "aria2c=$(command -v aria2c 2>/dev/null || echo 'NOT FOUND')"

# 颜色输出
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

# 创建临时目录
mkdir -p "$TEMP_DIR"

# 清理函数
cleanup() {
    # 终止残留子进程（yt-dlp/ffmpeg/aria2c），避免 Ctrl+C 后仍在后台偷偷下载
    jobs -p 2>/dev/null | xargs -r kill 2>/dev/null
    pkill -P $$ 2>/dev/null
    rm -rf "$TEMP_DIR"/* 2>/dev/null
    disown -a 2>/dev/null || true
}
# 错误处理
set -e
# 关键: Ctrl+C 必须真正退出。旧版 trap 只清理不 exit，导致"停不下来"
trap 'echo ""; echo -e "${RED}已中断 (Ctrl+C)${NC}"; exit 130' INT
trap 'cleanup; exit 143' TERM
trap 'cleanup' EXIT

# 显示帮助
show_help() {
    echo "用法: $0 [选项] <URL>"
    echo "选项:"
    echo "  -m, --mp3       下载并转换为MP3格式"
    echo "  -a, --audio     下载最佳音频（保持原格式，失败时自动转MP3）"
    echo "  -v, --video     下载视频（默认）"
    echo "  -w, --wechat    微信兼容模式（下载后转码 H.264 yuv420p+aac）"
    echo "  -f, --forum     从论坛页面提取并下载所有音视频"
    echo "  -l, --list      仅列出论坛页面中的音视频链接，不下载"
    echo "  -p, --playlist  下载整个播放列表/频道"
    echo "  -h, --help      显示此帮助"
    echo ""
    echo "示例:"
    echo "  $0 -v https://youtube.com/watch?v=xxx        # 下载视频"
    echo "  $0 -w https://youtube.com/watch?v=xxx        # 下载视频（微信兼容）"
    echo "  $0 -m https://youtube.com/watch?v=xxx        # 下载MP3"
    echo "  $0 -a https://youtube.com/watch?v=xxx        # 下载最佳音频"
    echo "  $0 -f https://hifiti.com/thread-2570.htm     # 爬取论坛页面所有音视频"
    echo "  $0 -l https://hifiti.com/thread-2570.htm     # 列出论坛页面音视频链接"
    echo "  $0 -p 'https://youtube.com/playlist?list=...' # 下载播放列表"
    echo "  $0                                          # 交互模式：选模式后粘贴分享内容自动下载"
    echo "  # 可直接粘贴分享内容（标题+链接），自动提取链接并清洗追踪参数:"
    echo "  $0 老子四章 老子〔先秦〕_哔哩哔哩_bilibili https://www.bilibili.com/video/BVxxx/"
    exit 0
}

# 检查必需工具
check_dependencies() {
    local missing=()
    local tools=("yt-dlp" "aria2c" "curl" "grep" "sed" "awk")
    for tool in "${tools[@]}"; do
        if ! command -v "$tool" &> /dev/null; then
            missing+=("$tool")
        fi
    done
    if [ ${#missing[@]} -ne 0 ]; then
        echo -e "${RED}错误: 缺少必需工具: ${missing[*]}${NC}"
        echo "请运行: pkg install ${missing[*]}"
        cleanup 1
    fi
}

# 从HTML中提取音视频链接（用于论坛模式）
extract_media_urls() {
    local html_file=$1
    local urls=()
    echo -e "${BLUE}正在从HTML中提取音视频链接...${NC}"
    while IFS= read -r line; do
        if [[ $line =~ (https?://(www\.)?youtube\.com/watch\?v=[a-zA-Z0-9_-]+) ]] || \
           [[ $line =~ (https?://youtu\.be/[a-zA-Z0-9_-]+) ]] || \
           [[ $line =~ (https?://(www\.)?bilibili\.com/video/[a-zA-Z0-9]+) ]] || \
           [[ $line =~ (https?://(www\.)?vimeo\.com/[0-9]+) ]] || \
           [[ $line =~ (https?://(www\.)?dailymotion\.com/video/[a-zA-Z0-9]+) ]] || \
           [[ $line =~ (https?://.*\.(mp4|webm|mkv|avi|mov|flv|mp3|m4a|wav|flac|ogg)(\?[^\"\']*)?) ]]; then
            url="${BASH_REMATCH[1]}"
            url=$(echo "$url" | sed 's/["'\''<>].*$//')
            urls+=("$url")
        fi
    done < <(grep -Eo 'https?://[^"'\''<> ]+(youtube\.com|youtu\.be|bilibili\.com|vimeo\.com|dailymotion\.com|\.mp4|\.webm|\.mkv|\.avi|\.mov|\.flv|\.mp3|\.m4a|\.wav|\.flac|\.ogg)[^"'\''<> ]*' "$html_file" 2>/dev/null || true)
    if [ ${#urls[@]} -gt 0 ]; then
        printf '%s\n' "${urls[@]}" | sort -u
    fi
}

# 论坛模式
forum_mode() {
    local url=$1
    local action=$2
    local html_file="$TEMP_DIR/forum_page.html"
    echo -e "${YELLOW}正在获取论坛页面: $url${NC}"
    if [ -n "$PROXY" ]; then
        curl -s -L -x "$PROXY" -A "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36" "$url" -o "$html_file"
    else
        curl -s -L -A "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36" "$url" -o "$html_file"
    fi
    if [ ! -f "$html_file" ] || [ ! -s "$html_file" ]; then
        echo -e "${RED}错误: 无法获取论坛页面${NC}"
        return 1
    fi
    mapfile -t media_urls < <(extract_media_urls "$html_file")
    if [ ${#media_urls[@]} -eq 0 ]; then
        echo -e "${YELLOW}未找到任何音视频链接${NC}"
        return 0
    fi
    echo -e "${GREEN}找到 ${#media_urls[@]} 个音视频链接:${NC}"
    local index=1
    for media_url in "${media_urls[@]}"; do
        echo -e "${BLUE}$index.${NC} $media_url"
        ((index++))
    done
    if [ "$action" = "list" ]; then
        return 0
    fi
    echo ""
    echo -e "${YELLOW}请选择:${NC}"
    echo "1) 下载所有"
    echo "2) 选择单个下载"
    echo "3) 取消"
    read -p "选择 (1-3): " choice
    case $choice in
        1)
            echo -e "${GREEN}开始下载所有 ${#media_urls[@]} 个文件...${NC}"
            for media_url in "${media_urls[@]}"; do
                echo -e "${YELLOW}正在下载: $media_url${NC}"
                download_with_fallback "bestaudio" "$media_url" || echo -e "${RED}下载失败: $media_url${NC}"
            done
            ;;
        2)
            read -p "请输入序号 (1-${#media_urls[@]}): " num
            if [[ "$num" =~ ^[0-9]+$ ]] && [ "$num" -ge 1 ] && [ "$num" -le "${#media_urls[@]}" ]; then
                selected_url="${media_urls[$((num-1))]}"
                echo -e "${GREEN}下载: $selected_url${NC}"
                echo "选择格式:"
                echo "1) 最佳音频"
                echo "2) MP3"
                echo "3) 视频"
                read -p "选择 (1-3): " fmt_choice
                case $fmt_choice in
                    1) download_with_fallback "bestaudio" "$selected_url" ;;
                    2) download_with_fallback "mp3" "$selected_url" ;;
                    3) download_with_fallback "video" "$selected_url" ;;
                    *) echo -e "${RED}无效选择${NC}" ;;
                esac
            else
                echo -e "${RED}无效序号${NC}"
            fi
            ;;
        3)
            echo -e "${YELLOW}取消下载${NC}"
            return 0
            ;;
        *)
            echo -e "${RED}无效选择${NC}"
            ;;
    esac
}

# 解析参数
MODE="video"
URL=""
URL_PARTS=()
FORUM_ACTION=""
WECHAT_MODE="false"

log_debug "解析参数前: MODE=$MODE URL=$URL WECHAT_MODE=$WECHAT_MODE"

while [[ $# -gt 0 ]]; do
    case $1 in
        -m|--mp3)
            MODE="mp3"
            shift
            ;;
        -a|--audio)
            MODE="bestaudio"
            shift
            ;;
        -v|--video)
            MODE="video"
            WECHAT_MODE="false"
            shift
            ;;
        -w|--wechat)
            MODE="video"
            WECHAT_MODE="true"
            shift
            ;;
        -f|--forum)
            MODE="forum"
            FORUM_ACTION="download"
            shift
            ;;
        -l|--list)
            MODE="forum"
            FORUM_ACTION="list"
            shift
            ;;
        -p|--playlist)
            MODE="playlist"
            shift
            ;;
        -h|--help)
            show_help
            ;;
        *)
            # 收集所有非选项参数（兼容整段粘贴：标题 + 链接会拆成多个参数）
            URL_PARTS+=("$1")
            shift
            ;;
    esac
done

MODE=$(strip_cr "$MODE")
WECHAT_MODE=$(strip_cr "$WECHAT_MODE")
FORUM_ACTION=$(strip_cr "$FORUM_ACTION")

# ==================== 从参数中提取 URL ====================
# 兼容直接粘贴分享内容（如："老子四章 老子〔先秦〕_哔哩哔哩_bilibili\nhttps://..."），
# 自动从参数中找出第一个 http(s):// 链接
if [ ${#URL_PARTS[@]} -gt 0 ]; then
    for part in "${URL_PARTS[@]}"; do
        part=$(strip_cr "$part")
        case "$part" in
            http://*|https://*)
                URL="$part"
                break
                ;;
        esac
        # 参数本身不是 URL，但内部可能嵌了 URL（标题和链接被粘在同一个参数里）
        inner=$(echo "$part" | grep -oE 'https?://[^[:space:]]+' | head -1)
        if [ -n "$inner" ]; then
            URL="$inner"
            break
        fi
    done
    # 回退：没找到带协议的链接，用第一个参数（兼容旧用法直接传 URL）
    if [ -z "$URL" ]; then
        URL=$(strip_cr "${URL_PARTS[0]}")
    fi
    if [ ${#URL_PARTS[@]} -gt 1 ]; then
        echo -e "${BLUE}已从粘贴内容中提取链接:${NC} $URL"
        log_debug "提取 URL: ${URL_PARTS[*]} → $URL"
    fi
fi

# 自动去除 URL 中的网页追踪参数（spm_id_from/vd_source/utm_* 等）
if [ -n "$URL" ]; then
    URL_CLEAN=$(clean_url "$URL")
    if [ "$URL_CLEAN" != "$URL" ]; then
        echo -e "${BLUE}URL 自动清洗: 已去除追踪参数${NC}"
        echo -e "${BLUE}  原始: $URL${NC}"
        echo -e "${BLUE}  清洗: $URL_CLEAN${NC}"
        log_debug "URL 清洗: $URL → $URL_CLEAN"
        URL="$URL_CLEAN"
    fi
fi

log_debug "解析参数后: MODE=$MODE URL=$URL WECHAT_MODE=$WECHAT_MODE FORUM_ACTION=$FORUM_ACTION"

if [ -z "$URL" ]; then
    echo -e "${YELLOW}未提供 URL，将进入交互下载模式${NC}"
    echo -e "${YELLOW}（也可以在命令行直接带链接运行: $0 https://...）${NC}"
fi

echo -e "${GREEN}=============================="
echo " Ultimate yt-dlp Downloader "
echo -e "==============================${NC}"

check_dependencies
log_debug "依赖检查通过"

# 构建基础参数
BASE_ARGS=()
if [ "${NO_COOKIES:-0}" != "1" ] && [ -f "$COOKIES" ] && [ -s "$COOKIES" ]; then
    BASE_ARGS+=(--cookies "$COOKIES")
fi
if [ -n "$PROXY" ]; then
    BASE_ARGS+=(--proxy "$PROXY")
fi

YT_EXTRACTOR_ARGS=(--extractor-args "youtube:player_client=ios,web_embedded")

COMMON_ARGS=(
    -N 32
    --downloader aria2c
    --downloader-args "aria2c:-x 16 -k 1M"
    --retries 10
    --fragment-retries 10
    --file-access-retries 10
    --extractor-retries 10
    --throttled-rate 100K
    -o "$DOWNLOAD_DIR/%(title)s.%(ext)s"
    --user-agent "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    --no-quiet
    --progress
    --console-title
    --no-warnings
    --compat-options no-live-chat
)

PLAYLIST_ARGS=(
    --yes-playlist
    --playlist-start 1
    --playlist-end 999
    --download-archive "$DOWNLOAD_DIR/.downloaded.txt"
)

# ==================== 微信兼容转码 ====================
transcode_to_h264() {
    local input="$1"
    local output="${input%.*}_compatible.mp4"
    echo -e "${YELLOW}[转码] 转换为微信兼容格式 (H.264 baseline + AAC)...${NC}" >&2
    log_debug "[转码] 输入: $input"
    ffprobe -v error -show_entries stream=codec_name,width,height,profile \
        -of default=noprint_wrappers=1 "$input" 2>/dev/null | paste -sd ' ' - >> "$LOG_FILE" || true
    ffmpeg -i "$input" -y \
        -c:v libx264 -profile:v baseline -level 3.1 \
        -preset fast -crf 24 \
        -pix_fmt yuv420p \
        -vf 'scale=min(1920,iw):-2' \
        -c:a aac -b:a 128k -ac 2 \
        -movflags +faststart \
        "$output" 2>/dev/null
    if [ -f "$output" ] && [ -s "$output" ]; then
        local in_size=$(stat -c%s "$input" 2>/dev/null || echo 0)
        local out_size=$(stat -c%s "$output" 2>/dev/null || echo 0)
        log_debug "[转码] 成功: $(basename "$input") (${in_size}B) → $(basename "$output") (${out_size}B)"
        rm -f "$input" 2>/dev/null
        echo "$output"
    else
        log_debug "[转码] 失败: 输出文件为空或不存在，保留原文件"
        echo "$input"
    fi
}

# ==================== 下载核心函数 ====================
download_with_fallback() {
    local mode=$1
    local url=$2
    local extra_args=()
    
    mode=$(strip_cr "$mode")
    url=$(strip_cr "$url")
    
    log_debug "download_with_fallback 入口: mode=$mode url=$url"
    
    if [[ "$url" == *"bilibili"* ]] || [[ "$url" == *"b23.tv"* ]]; then
        extra_args+=(--add-header "Referer:https://www.bilibili.com/")
    fi

    case $mode in
        "playlist")
            echo -e "${YELLOW}播放列表模式 - 下载整个列表/频道（直接存成品）${NC}"
            log_debug "执行: yt-dlp playlist mode url=$url"
            yt-dlp "${BASE_ARGS[@]}" "${COMMON_ARGS[@]}" "${PLAYLIST_ARGS[@]}" "${extra_args[@]}" \
                --print after_move:filepath \
                --print-to-file after_move:filepath "$YT_PATH_FILE" \
                "$url" || { log_debug "yt-dlp 播放列表下载失败, ret=$?"; return 1; }
            log_debug "yt-dlp playlist 成功"
            ;;
        "mp3")
            echo -e "${YELLOW}开始下载并转换为MP3...${NC}"
            log_debug "执行: yt-dlp mp3 mode url=$url"
            yt-dlp "${BASE_ARGS[@]}" "${COMMON_ARGS[@]}" "${extra_args[@]}" \
                -x \
                --audio-format mp3 \
                --audio-quality 0 \
                --embed-metadata \
                --embed-thumbnail \
                --ppa "ffmpeg:-id3v2_version 3" \
                --print after_move:filepath \
                --print-to-file after_move:filepath "$YT_PATH_FILE" \
                "$url" || { log_debug "yt-dlp mp3 失败, ret=$?"; return 1; }
            log_debug "yt-dlp mp3 成功"
            ;;
            
        "bestaudio")
            echo -e "${YELLOW}尝试下载最佳原始音频...${NC}"
            log_debug "执行: yt-dlp bestaudio mode url=$url"
            if ! yt-dlp "${BASE_ARGS[@]}" "${COMMON_ARGS[@]}" "${extra_args[@]}" \
                -x \
                -f bestaudio \
                --embed-metadata \
                --print after_move:filepath \
                --print-to-file after_move:filepath "$YT_PATH_FILE" \
                "$url"; then
                log_debug "yt-dlp bestaudio 失败, 回退 mp3"
                echo -e "${YELLOW}下载最佳音频失败，尝试转换为MP3...${NC}"
                yt-dlp "${BASE_ARGS[@]}" "${COMMON_ARGS[@]}" "${extra_args[@]}" \
                    -x \
                    --audio-format mp3 \
                    --audio-quality 0 \
                    --embed-metadata \
                    --embed-thumbnail \
                    --ppa "ffmpeg:-id3v2_version 3" \
                    --print after_move:filepath \
                    --print-to-file after_move:filepath "$YT_PATH_FILE" \
                    "$url" || { log_debug "yt-dlp bestaudio回退mp3 失败, ret=$?"; return 1; }
            fi
            log_debug "yt-dlp bestaudio 成功"
            ;;
            
        "video")
            echo -e "${YELLOW}解析视频标题...${NC}"
            local raw_title
            raw_title=$(yt-dlp --print "%(title)s" --no-download \
                "${BASE_ARGS[@]}" "${YT_EXTRACTOR_ARGS[@]}" "$url" 2>/dev/null | tail -1)
            local safe_name=$(format_video_name "$raw_title")
            local video_dir="$DOWNLOAD_DIR/$safe_name"
            local merged_file="$video_dir/$safe_name.mp4"
            mkdir -p "$video_dir"
            echo -e "${YELLOW}下载视频: ${GREEN}$safe_name${NC}"
            echo -e "${BLUE}保存到: $video_dir${NC}"
            log_debug "video 模式: title=$raw_title safe=$safe_name dir=$video_dir"

            # 去重：检查最终合并文件是否存在
            if [ -f "$merged_file" ] && [ -s "$merged_file" ]; then
                echo -e "${GREEN}已存在完整文件，跳过: $merged_file${NC}" >&2
                echo "$merged_file" > "$YT_PATH_FILE"
                log_debug "去重命中: $merged_file"
                return 0
            fi

            # 2. 下载视频流
            echo -e "${YELLOW}  → 下载视频流...${NC}"
            yt-dlp "${BASE_ARGS[@]}" "${YT_EXTRACTOR_ARGS[@]}" \
                -N 16 \
                --retries 10 --fragment-retries 10 --file-access-retries 10 --extractor-retries 10 \
                --throttled-rate 100K \
                --user-agent "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36" \
                --no-warnings --no-quiet --progress \
                "${extra_args[@]}" \
                -f "bestvideo[ext=mp4][vcodec^=avc1]/bestvideo[ext=mp4]/bestvideo" \
                -o "$video_dir/${safe_name}_video.%(ext)s" \
                "$url" || { log_debug "视频流下载失败, ret=$?"; return 1; }
            log_debug "视频流下载完成"

            # 3. 下载音频流
            echo -e "${YELLOW}  → 下载音频流...${NC}"
            yt-dlp "${BASE_ARGS[@]}" "${YT_EXTRACTOR_ARGS[@]}" \
                -N 16 \
                --retries 10 --fragment-retries 10 --file-access-retries 10 --extractor-retries 10 \
                --throttled-rate 100K \
                --user-agent "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36" \
                --no-warnings --no-quiet --progress \
                "${extra_args[@]}" \
                -f "bestaudio[ext=m4a]/bestaudio" \
                -o "$video_dir/${safe_name}_audio.%(ext)s" \
                "$url" || { log_debug "音频流下载失败, ret=$?"; return 1; }
            log_debug "音频流下载完成"

            # 4. 尝试下载字幕（失败不影响）
            echo -e "${YELLOW}  → 检查字幕...${NC}"
            yt-dlp "${BASE_ARGS[@]}" "${YT_EXTRACTOR_ARGS[@]}" \
                --no-warnings --no-quiet \
                "${extra_args[@]}" \
                --skip-download \
                --write-sub --sub-lang zh,en --convert-subs srt \
                -o "$video_dir/${safe_name}.%(ext)s" \
                "$url" 2>/dev/null || true
            log_debug "字幕检查完成"

            # 5. 定位下载的流文件（使用 find 避免通配符问题）
            echo -e "${YELLOW}  → 查找视频/音频流文件...${NC}"
            local video_file=$(find "$video_dir" -maxdepth 1 -type f -name "*_video.*" -printf "%T@ %p\n" 2>/dev/null | sort -rn | head -1 | cut -d' ' -f2-)
            local audio_file=$(find "$video_dir" -maxdepth 1 -type f -name "*_audio.*" -printf "%T@ %p\n" 2>/dev/null | sort -rn | head -1 | cut -d' ' -f2-)

            # 调试信息（可保留，方便排查）
            echo "========== DEBUG INFO =========="
            echo "video_dir = '$video_dir'"
            echo "merged_file = '$merged_file'"
            echo "video_file = '$video_file'"
            echo "audio_file = '$audio_file'"
            if [ -n "$video_file" ]; then
                ls -l "$video_file" 2>&1
                #file "$video_file" 2>&1
            else
                echo "video_file is EMPTY"
            fi
            if [ -n "$audio_file" ]; then
                ls -l "$audio_file" 2>&1
                #file "$audio_file" 2>&1
            else
                echo "audio_file is EMPTY"
            fi
            echo "================================="

            log_debug "流文件: video=$video_file audio=$audio_file"

            # 6. 合并（彻底隔离子 shell，解决卡顿）
            if [ -f "$video_file" ] && [ -f "$audio_file" ]; then
                # 合并：先 -c copy 直拷（最快）；失败/超时则回退为音频重编码 AAC（错误写入日志便于排查）。
                # 注：不再加 -movflags +faststart —— 本地播放无需 moov 前置，
                #     可避免 Android FUSE(/sdcard) 二次写盘导致合并极慢甚至卡死；
                #     微信兼容模式 transcode_to_h264 输出时仍保留 faststart。
                # 用 timeout 兜底，避免 FUSE 写入卡死时脚本无限等待。
                echo -e "${YELLOW}  → 合并视频+音频 (copy 模式)...${NC}"
                timeout 600 ffmpeg -v error -i "$video_file" -i "$audio_file" -c copy -y "$merged_file" 2>>"$LOG_FILE"
                local ret=$?
                if [ $ret -eq 0 ]; then
                    echo -e "${GREEN}  → 合并完成: $merged_file${NC}"
                    log_debug "合并完成: $merged_file"
                else
                    if [ $ret -eq 124 ]; then
                        log_debug "copy 合并超时(>600s)，尝试回退：音频重编码 AAC"
                        echo -e "${YELLOW}  → copy 合并超时(>10分钟)，改用音频重编码重试...${NC}"
                    else
                        log_debug "copy 合并失败(ret=$ret)，尝试回退：音频重编码 AAC"
                        echo -e "${YELLOW}  → copy 合并失败(ret=$ret)，改用音频重编码重试...${NC}"
                    fi
                    timeout 900 ffmpeg -v error -i "$video_file" -i "$audio_file" \
                        -c:v copy -c:a aac -b:a 192k -y "$merged_file" 2>>"$LOG_FILE"
                    ret=$?
                    if [ $ret -eq 0 ]; then
                        echo -e "${GREEN}  → 回退重编码合并完成: $merged_file${NC}"
                        log_debug "回退重编码合并完成: $merged_file"
                    else
                        echo -e "${RED}  → 合并失败(ret=$ret)，详情见: $LOG_FILE${NC}"
                        log_debug "回退重编码也失败，ret=$ret"
                        return 1
                    fi
                fi
            elif [ -f "$video_file" ]; then
                echo -e "${YELLOW}  → 单流模式，直接复制...${NC}"
                cp "$video_file" "$merged_file"
                echo -e "${GREEN}  → 复制完成: $merged_file${NC}"
                log_debug "单流复制: $merged_file"
            else
                echo -e "${RED}  → 错误: 未找到视频流文件${NC}"
                log_debug "错误: 未找到视频流文件"
                return 1
            fi

            # 7. 记录合并后路径 + 自动清理中间流文件
            if [ -f "$merged_file" ]; then
                echo "$merged_file" > "$YT_PATH_FILE"
                echo -e "${GREEN}最终文件: $merged_file${NC}" >&2
                local msize=$(stat -c%s "$merged_file" 2>/dev/null || echo '?')
                log_debug "最终文件大小: ${msize}B"
                # 自动删除中间流文件，DCIM 只保留最终成品
                if [ -n "$video_file" ] && [ -f "$video_file" ] && [ "$video_file" != "$merged_file" ]; then
                    rm -f "$video_file" 2>/dev/null && log_debug "已清理中间视频流: $video_file"
                fi
                if [ -n "$audio_file" ] && [ -f "$audio_file" ]; then
                    rm -f "$audio_file" 2>/dev/null && log_debug "已清理中间音频流: $audio_file"
                fi
                echo -e "${GREEN}  → 已自动清理中间流文件 (_video/_audio)${NC}"
            fi
            log_debug "yt-dlp video 完成"
            ;;
            
        *)
            log_debug "未知下载模式: $mode"
            echo -e "${RED}未知下载模式: $mode${NC}"
            return 1
            ;;
    esac
    
    return 0
}

# ==================== 交互模式 ====================
# 无参数运行时进入：选择下载模式 → 粘贴分享内容（标题+链接，可多行）→ 自动解析下载。
# 交互模式下直接粘贴文本，& 等字符不会被 shell 拆分，无需加引号
interactive_mode() {
    echo -e "${GREEN}====================================${NC}"
    echo -e "${GREEN}  交互下载模式${NC}"
    echo -e "${GREEN}====================================${NC}"
    echo -e "${BLUE}选择下载模式，然后粘贴分享内容（标题+链接均可），自动解析下载。${NC}"
    while true; do
        echo ""
        echo -e "${YELLOW}请选择下载模式:${NC}"
        echo "  v)    视频 (默认)"
        echo "  a)    最佳音频"
        echo "  m)    MP3"
        echo "  w)    微信兼容 (转码 H.264)"
        echo "  p)    播放列表/频道"
        echo "  f)    论坛页面提取"
        echo "  q)    退出"
        read -r -p "选项 [v]: " choice || { echo ""; echo -e "${GREEN}再见${NC}"; return 0; }
        WECHAT_MODE="false"   # 每次选择都重置，避免 w 状态残留
        case $choice in
            q|Q|exit|quit) echo -e "${GREEN}再见${NC}"; return 0 ;;
            a|A) imode="bestaudio" ;;
            m|M) imode="mp3" ;;
            w|W) imode="video"; WECHAT_MODE="true" ;;
            p|P) imode="playlist" ;;
            f|F) imode="forum" ;;
            v|V|"") imode="video" ;;
            *) echo -e "${RED}无效选项，使用默认视频模式${NC}"; imode="video" ;;
        esac
        echo ""
        echo -e "${YELLOW}粘贴分享内容（标题+链接可一起粘）:${NC}"
        echo -e "${YELLOW}粘贴后检测到链接即自动开始；若末尾无换行，按一次回车确认。${NC}"
        local text=""
        local line
        while IFS= read -r line; do
            line="${line%$'\r'}"   # 去除可能残留的 CR
            case "$line" in
                *http://*|*https://*)   # 读到含链接的行：立即开始（无需空行/回车）
                    text="$text $line"
                    break
                    ;;
                "")                     # 空行：已有内容则视为粘贴结束（旧行为兜底）
                    [ -n "$text" ] && break
                    ;;
                *)                      # 其他行（标题等）：先收集
                    text="$text $line"
                    ;;
            esac
        done
        text=$(echo "$text" | sed 's/^[[:space:]]*//;s/[[:space:]]*$//')
        if [ -z "$text" ]; then
            echo -e "${RED}未输入内容，跳过${NC}"
            continue
        fi
        local url
        url=$(echo "$text" | grep -oE 'https?://[^[:space:]]+' | head -1)
        if [ -z "$url" ]; then
            echo -e "${RED}未从内容中解析到链接，跳过${NC}"
            continue
        fi
        url=$(clean_url "$url")
        echo -e "${GREEN}解析到链接: $url${NC}"
        log_debug "交互模式: 模式=$imode 链接=$url"
        case $imode in
            "forum")
                forum_mode "$url" "download"
                ;;
            "playlist")
                echo -e "${YELLOW}播放列表模式 - 下载整个列表/频道${NC}"
                download_with_fallback "playlist" "$url"
                ;;
            *)
                download_with_fallback "$imode" "$url"
                # 交互模式下 w 不经过脚本尾部的转码逻辑，这里手动补上
                if [ "$WECHAT_MODE" = "true" ]; then
                    local latest
                    latest=$(tail -1 "$YT_PATH_FILE" 2>/dev/null)
                    if [ -n "$latest" ] && [ -f "$latest" ]; then
                        latest=$(transcode_to_h264 "$latest")
                        echo -e "${GREEN}微信兼容转码完成: $latest${NC}"
                    fi
                fi
                ;;
        esac
        echo ""
        echo -e "${GREEN}本次下载完成。继续输入下一个，或按 q 退出。${NC}"
    done
}

# ==================== 执行下载 ====================
# 无 URL 时进入交互模式（原"错误: 请输入URL"改为直接交互）
if [ -z "$URL" ]; then
    log_debug "未提供 URL，进入交互模式"
    interactive_mode
    exit 0
fi

log_debug "进入下载分发: MODE=$MODE URL=$URL WECHAT_MODE=$WECHAT_MODE"
case $MODE in
    "forum")
        log_debug "forum 模式"
        forum_mode "$URL" "$FORUM_ACTION"
        ;;
    "playlist")
        log_debug "playlist 模式"
        echo -e "${YELLOW}播放列表模式 - 下载所有视频（默认最高画质）${NC}"
        download_with_fallback "playlist" "$URL"
        ;;
    *)
        log_debug "默认模式: 调用 download_with_fallback '$MODE' '$URL'"
        download_with_fallback "$MODE" "$URL"
        ;;
esac

# ==================== 最终结果处理 ====================
RET=$?
log_debug "download_with_fallback 返回: RET=$RET"
if [ $RET -eq 0 ]; then
    LATEST=$(tail -1 "$YT_PATH_FILE" 2>/dev/null)
    if [ -z "$LATEST" ] || [ ! -f "$LATEST" ]; then
        LATEST=$(ls -t "$DOWNLOAD_DIR"/*.mp4 "$DOWNLOAD_DIR"/*.m4a "$DOWNLOAD_DIR"/*.mp3 "$DOWNLOAD_DIR"/*.webm "$DOWNLOAD_DIR"/*.mkv "$DOWNLOAD_DIR"/*/*.mp4 "$DOWNLOAD_DIR"/*/*.m4a "$DOWNLOAD_DIR"/*/*.mp3 2>/dev/null | head -1)
        log_debug "yt_path_file 不可用，回退 ls -t: LATEST=$LATEST"
    fi
    log_debug "最新文件检测: LATEST=$LATEST WECHAT_MODE=$WECHAT_MODE"
    if [ -n "$LATEST" ]; then
        if [ "$WECHAT_MODE" = "true" ]; then
            log_debug "微信兼容: 开始转码 input=$LATEST"
            LATEST=$(transcode_to_h264 "$LATEST")
            log_debug "微信兼容: 转码完成 output=$LATEST"
        fi
        echo "FILEPATH=$LATEST"
        log_debug "输出: FILEPATH=$LATEST"
    else
        log_debug "警告: 未找到下载文件 (DOWNLOAD_DIR=$DOWNLOAD_DIR)"
        echo "FILEPATH="
    fi
    echo "[$(date)] 下载成功: $URL → $(basename "$LATEST" 2>/dev/null || echo '?')" >> "$LOG_FILE"
    log_debug "========== yt.sh 成功结束 =========="
else
    echo "[$(date)] 下载失败: $URL, 模式: $MODE, 错误代码: $RET" >> "$LOG_FILE"
    log_debug "========== yt.sh 失败结束 RET=$RET =========="
fi

echo
exit $RET