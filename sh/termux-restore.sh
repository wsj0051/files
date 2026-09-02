#!/data/data/com.termux/files/usr/bin/bash

# 颜色输出
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 配置
BACKUP_DIR="/data/data/com.termux/files/home/tf/Private/bak/termux"
LATEST_LINK="${BACKUP_DIR}/termux-backup-latest.tar.gz"

# 显示警告
show_warning() {
    echo -e "${RED}╔════════════════════════════════════════════════════════╗${NC}"
    echo -e "${RED}║                     警告！                             ║${NC}"
    echo -e "${RED}║  恢复操作将覆盖现有文件，建议先备份当前环境              ║${NC}"
    echo -e "${RED}║  恢复过程中请勿关闭Termux或中断操作                     ║${NC}"
    echo -e "${RED}╚════════════════════════════════════════════════════════╝${NC}"
    echo ""
}

# 列出可用备份
list_backups() {
    echo -e "${GREEN}可用备份文件:${NC}"
    echo "================================"
    
    local backups=()
    local index=1
    
    while IFS= read -r file; do
        if [ -f "$file" ]; then
            size=$(du -h "$file" | cut -f1)
            date=$(stat -c %y "$file" | cut -d. -f1)
            echo -e "${BLUE}$index)${NC} $(basename "$file")"
            echo "   大小: $size, 日期: $date"
            backups+=("$file")
            ((index++))
        fi
    done < <(find "$BACKUP_DIR" -name "termux-backup-*.tar.gz" | sort -r)
    
    if [ ${#backups[@]} -eq 0 ]; then
        echo -e "${RED}没有找到备份文件${NC}"
        exit 1
    fi
    
    # 返回备份列表
    BACKUP_FILES=("${backups[@]}")
}

# 选择备份文件
select_backup() {
    echo ""
    echo -e "${YELLOW}请选择要恢复的备份:${NC}"
    echo "0) 退出"
    
    local i
    for i in "${!BACKUP_FILES[@]}"; do
        echo "$((i+1))) $(basename "${BACKUP_FILES[$i]}")"
    done
    
    read -p "选择 (0-${#BACKUP_FILES[@]}): " choice
    
    if [ "$choice" = "0" ]; then
        echo -e "${YELLOW}已取消${NC}"
        exit 0
    fi
    
    if [[ "$choice" =~ ^[0-9]+$ ]] && [ "$choice" -ge 1 ] && [ "$choice" -le "${#BACKUP_FILES[@]}" ]; then
        SELECTED_BACKUP="${BACKUP_FILES[$((choice-1))]}"
    else
        echo -e "${RED}无效选择${NC}"
        select_backup
    fi
}

# 确认恢复
confirm_restore() {
    echo ""
    echo -e "${RED}即将恢复: $(basename "$SELECTED_BACKUP")${NC}"
    echo -e "${RED}这将覆盖以下目录:${NC}"
    echo "  - /data/data/com.termux/files/home"
    echo "  - /data/data/com.termux/files/usr"
    echo ""
    echo -e "${YELLOW}是否继续？${NC}"
    echo "1) 是，开始恢复"
    echo "2) 否，退出"
    
    read -p "选择 (1-2): " confirm
    
    if [ "$confirm" != "1" ]; then
        echo -e "${YELLOW}已取消${NC}"
        exit 0
    fi
}

# 执行恢复
do_restore() {
    echo -e "${GREEN}开始恢复 Termux...${NC}"
    echo "================================"
    
    # 检查备份文件
    if [ ! -f "$SELECTED_BACKUP" ]; then
        echo -e "${RED}错误: 备份文件不存在${NC}"
        exit 1
    fi
    
    # 检查备份文件完整性
    echo -e "${YELLOW}验证备份文件...${NC}"
    if ! gzip -t "$SELECTED_BACKUP" 2>/dev/null; then
        echo -e "${RED}错误: 备份文件损坏${NC}"
        exit 1
    fi
    
    # 切换到 Termux 根目录
    cd /data/data/com.termux/files || {
        echo -e "${RED}错误: 无法进入 /data/data/com.termux/files${NC}"
        exit 1
    }
    
    # 创建临时备份（可选）
    echo -e "${YELLOW}创建当前环境临时备份...${NC}"
    local temp_backup="/data/local/tmp/termux-pre-restore-$(date +%Y%m%d_%H%M%S).tar.gz"
    tar -czf "$temp_backup" home usr --warning=no-file-changed 2>/dev/null
    
    if [ $? -eq 0 ]; then
        echo -e "${GREEN}临时备份已创建: $temp_backup${NC}"
    fi
    
    echo ""
    echo -e "${YELLOW}开始解压备份文件...${NC}"
    
    # 执行恢复
    tar -xzf "$SELECTED_BACKUP" \
        --checkpoint=1000 \
        --checkpoint-action="ttyout=已完成 %u 个文件\r"
    
    local TAR_EXIT=$?
    
    if [ $TAR_EXIT -eq 0 ]; then
        echo -e "\n${GREEN}✓ 恢复完成！${NC}"
        
        # 修复权限
        echo -e "${YELLOW}修复文件权限...${NC}"
        chmod -R 755 /data/data/com.termux/files/usr/bin 2>/dev/null
        chmod 700 /data/data/com.termux/files/home/.ssh 2>/dev/null
        chmod 600 /data/data/com.termux/files/home/.ssh/* 2>/dev/null
        
        echo -e "${GREEN}所有操作完成${NC}"
        echo ""
        echo -e "${YELLOW}建议重启 Termux 或执行:${NC}"
        echo "  exit"
        echo "  然后重新打开 Termux"
    else
        echo -e "\n${RED}✗ 恢复失败！错误代码: $TAR_EXIT${NC}"
        echo -e "${YELLOW}临时备份保存在: $temp_backup${NC}"
        exit 1
    fi
}

# 主函数
main() {
    echo -e "${GREEN}=============================="
    echo "     Termux 恢复工具"
    echo "==============================${NC}"
    
    # 显示警告
    show_warning
    
    # 检查备份目录
    if [ ! -d "$BACKUP_DIR" ]; then
        echo -e "${RED}错误: 备份目录不存在${NC}"
        exit 1
    fi
    
    # 列出备份
    list_backups
    
    # 选择备份
    select_backup
    
    # 确认
    confirm_restore
    
    # 执行恢复
    do_restore
}

# 运行主函数
main "$@"
