#!/data/data/com.termux/files/usr/bin/bash

# 颜色输出
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'

# No Color
NC='\033[0m'

# 配置
BACKUP_DIR="/data/data/com.termux/files/home/tf/Private/bak/termux"
TIMESTAMP=$(date +"%Y%m%d_%H%M%S")
BACKUP_NAME="termux-backup-${TIMESTAMP}.tar.gz"
LATEST_LINK="${BACKUP_DIR}/termux-backup-latest.tar.gz"

# 要备份的目录
BACKUP_SRC=(
    "/data/data/com.termux/files/home"
    "/data/data/com.termux/files/usr"
)

# 排除的文件/目录（支持通配符）
EXCLUDE=(
    "--exclude=home/storage"
    "--exclude=home/.cache"
    "--exclude=home/.npm"
    "--exclude=home/.cargo"
    "--exclude=home/.gradle"
    "--exclude=home/.m2"
    "--exclude=usr/tmp"
    "--exclude=usr/cache"
    "--exclude=*.log"
    "--exclude=*.tmp"

    # 排除备份目标目录自身
    "--exclude=home/tf"  
)

# 创建备份目录
create_backup_dir() {
    if [ ! -d "$BACKUP_DIR" ]; then
        echo -e "${YELLOW}创建备份目录: $BACKUP_DIR${NC}"
        mkdir -p "$BACKUP_DIR"
        if [ $? -ne 0 ]; then
            echo -e "${RED}错误: 无法创建备份目录${NC}"
            exit 1
        fi
    fi
}

# 检查存储空间（兼容 Android）
check_space() {
    echo -e "${BLUE}检查存储空间...${NC}"
    
    # 计算需要备份的大小（排除排除项）
    local required=0
    
    # 估算 home 目录大小（排除 cache 等）
    local home_size=$(du -sm /data/data/com.termux/files/home 2>/dev/null | awk '{print $1}')
    local usr_size=$(du -sm /data/data/com.termux/files/usr 2>/dev/null | awk '{print $1}')
    
    # 减去一些估计的排除项大小
    local exclude_size=200  # 估计排除项约 200MB
    required=$((home_size + usr_size - exclude_size))
    
    # 确保至少为 1MB
    if [ $required -lt 1 ]; then
        required=1
    fi
    
    # 检查目标目录的可用空间
    local available=""
    
    # Android 的 df 输出格式示例：
    # Filesystem              1K-blocks    Used   Available   Use%   Mounted on
    if command -v df &> /dev/null; then
        # 获取目标目录所在分区的可用空间（以KB为单位）
        local df_output=$(df -k "$BACKUP_DIR" 2>/dev/null | tail -1)
        available=$(echo "$df_output" | awk '{print $4}')
        
        # 如果没获取到，尝试另一种格式
        if [ -z "$available" ] || [ "$available" = "0" ]; then
            # 有些 Android 的 df 输出在第四列是 Available
            available=$(echo "$df_output" | awk '{print $4}')
        fi
        
        # 如果还是没获取到，尝试使用 /data 分区
        if [ -z "$available" ] || [ "$available" = "0" ]; then
            df_output=$(df -k /data 2>/dev/null | tail -1)
            available=$(echo "$df_output" | awk '{print $4}')
        fi
        
        # 转换为 MB
        if [ -n "$available" ] && [ "$available" -gt 0 ]; then
            available=$((available / 1024))
        else
            available="unknown"
        fi
    fi
    
    # 如果无法获取可用空间，跳过检查
    if [ "$available" = "unknown" ] || [ -z "$available" ]; then
        echo -e "${YELLOW}警告: 无法获取可用空间信息，跳过空间检查${NC}"
        return 0
    fi
    
    echo -e "${BLUE}需要空间: ~${required}MB${NC}"
    echo -e "${BLUE}可用空间: ${available}MB${NC}"
    
    # 预留 100MB 缓冲区
    if [ $available -lt $((required + 100)) ]; then
        echo -e "${RED}错误: 可用空间不足${NC}"
        echo -e "${RED}需要至少 $((required + 100))MB，可用 ${available}MB${NC}"
        echo -e "${YELLOW}是否继续？(y/n)${NC}"
        read -p "选择: " force_continue
        if [[ ! "$force_continue" =~ ^[Yy]$ ]]; then
            exit 1
        fi
    else
        echo -e "${GREEN}空间充足，继续备份...${NC}"
    fi
}

# 执行备份
do_backup() {
    echo -e "${GREEN}开始备份 Termux..."
    echo -e "================================${NC}"
    
    # 切换到 Termux 根目录
    cd /data/data/com.termux/files || {
        echo -e "${RED}错误: 无法进入 /data/data/com.termux/files${NC}"
        exit 1
    }
    
    # 构建排除参数
    local EXCLUDE_ARGS=""
    for excl in "${EXCLUDE[@]}"; do
        EXCLUDE_ARGS="$EXCLUDE_ARGS $excl"
    done
    
    echo -e "${YELLOW}备份文件: ${BACKUP_DIR}/${BACKUP_NAME}${NC}"
    echo -e "${YELLOW}备份内容: home usr${NC}"
    echo -e "${YELLOW}排除项: ${EXCLUDE[*]}${NC}"
    echo ""
    
    # 使用 pv 显示进度（如果安装）
    if command -v pv &> /dev/null; then
        echo -e "${GREEN}使用 pv 显示进度...${NC}"
        tar -cf - $EXCLUDE_ARGS home usr 2>/dev/null | pv -s $(du -sb home usr 2>/dev/null | awk '{total+=$1} END {print total}') | gzip > "${BACKUP_DIR}/${BACKUP_NAME}"
    else
        # 使用 tar 的内置进度
        echo -e "${YELLOW}提示: 安装 pv 可以显示更详细的进度 (pkg install pv)${NC}"
        tar -czf "${BACKUP_DIR}/${BACKUP_NAME}" \
            $EXCLUDE_ARGS \
            home usr \
            --warning=no-file-changed \
            --checkpoint=1000 \
            --checkpoint-action="ttyout=已打包 %u 个文件\r"
    fi
    
    local TAR_EXIT=$?
    
    if [ $TAR_EXIT -eq 0 ] || [ $TAR_EXIT -eq 1 ]; then  # tar 返回 1 有时只是文件更改警告
        # 创建最新版本的软链接
        ln -sf "${BACKUP_DIR}/${BACKUP_NAME}" "$LATEST_LINK" 2>/dev/null
        
        # 计算备份大小
        local BACKUP_SIZE=$(du -h "${BACKUP_DIR}/${BACKUP_NAME}" | cut -f1)
        
        echo -e "\n${GREEN}✓ 备份完成！${NC}"
        echo -e "${GREEN}  文件: ${BACKUP_DIR}/${BACKUP_NAME}${NC}"
        echo -e "${GREEN}  大小: ${BACKUP_SIZE}${NC}"
        
        # 显示备份文件列表
        echo ""
        echo -e "${YELLOW}是否查看备份内容列表？${NC}"
        echo "1) 是"
        echo "2) 否"
        read -p "选择 (1-2): " list_choice
        
        if [ "$list_choice" = "1" ]; then
            echo -e "${BLUE}备份文件内容:${NC}"
            tar -tzf "${BACKUP_DIR}/${BACKUP_NAME}" 2>/dev/null | head -20
            echo "..."
        fi
        
        # 可选：删除旧备份（保留最近5个）
        echo ""
        echo -e "${YELLOW}是否删除旧备份？（保留最近5个）${NC}"
        echo "1) 是"
        echo "2) 否"
        read -p "选择 (1-2): " clean_choice
        
        if [ "$clean_choice" = "1" ]; then
            clean_old_backups
        fi
    else
        echo -e "${RED}✗ 备份失败！错误代码: $TAR_EXIT${NC}"
        exit 1
    fi
}

# 清理旧备份
clean_old_backups() {
    echo -e "${YELLOW}清理旧备份...${NC}"
    
    cd "$BACKUP_DIR" || return
    
    # 保留最近5个备份
    local count=0
    ls -t termux-backup-*.tar.gz 2>/dev/null | while read backup; do
        count=$((count + 1))
        if [ $count -gt 5 ]; then
            echo "删除: $backup"
            rm -f "$backup"
        fi
    done
    
    echo -e "${GREEN}清理完成${NC}"
}

# 主函数
main() {
    echo -e "${GREEN}=============================="
    echo "     Termux 备份工具"
    echo -e "==============================${NC}"
    
    # 创建备份目录
    create_backup_dir
    
    # 检查空间（带错误处理）
    check_space
    
    # 执行备份
    do_backup
}

# 运行主函数
main "$@"
