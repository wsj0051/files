#!/data/data/com.termux/files/usr/bin/bash
# miniserve 统一控制脚本（Android / Termux 最终版）

TERMUX_HOME="/data/data/com.termux/files/home"
MINISERVE="/data/data/com.termux/files/usr/bin/miniserve"
PORT=8088
ROOT="$TERMUX_HOME/disk/Public/media"
SERVER_DIR="$TERMUX_HOME/server"
PIDFILE="$SERVER_DIR/miniserve.pid"
LOGFILE="$SERVER_DIR/miniserve.log"

URL="http://127.0.0.1:$PORT"

mkdir -p "$SERVER_DIR" "$ROOT"

# -----------------------------
# 工具函数
# -----------------------------

check_port() {
    if command -v ss >/dev/null 2>&1; then
        ss -ltn 2>/dev/null | grep -q ":$PORT "
    elif command -v netstat >/dev/null 2>&1; then
        netstat -ltn 2>/dev/null | grep -q ":$PORT "
    else
        return 1
    fi
}

http_alive() {
    command -v curl >/dev/null 2>&1 || return 1
    curl -fs --max-time 2 "$URL" >/dev/null 2>&1
}

wait_http_up() {
    for _ in 1 2 3 4 5; do
        http_alive && return 0
        sleep 1
    done
    return 1
}

# -----------------------------
# stop
# -----------------------------

stop() {
    echo "停止 miniserve 服务..."

    IS_ROOT=0
    [ "$(id -u)" = "0" ] && IS_ROOT=1

    if [ "$IS_ROOT" = "0" ]; then
        echo "⚠️  当前为普通用户，无法停止 root 启动的服务"
        echo "👉 请使用：su -c \"$0 stop\""
        return 1
    fi

    if [ -f "$PIDFILE" ]; then
        PID=$(cat "$PIDFILE" 2>/dev/null)
        [ -n "$PID" ] && kill "$PID" 2>/dev/null || true
    fi

    pkill -f "miniserve.*$PORT" 2>/dev/null || true
    sleep 1

    rm -f "$PIDFILE"

    if check_port; then
        echo "❌ 停止失败：端口 $PORT 仍在监听"
        return 1
    else
        echo "✅ 服务已停止"
        return 0
    fi
}


# -----------------------------
# start
# -----------------------------

start() {
    echo "启动 miniserve 服务..."

    if http_alive; then
        echo "⚠️  服务已在运行"
        return 0
    fi

    if check_port; then
        echo "端口 $PORT 被占用，尝试清理..."
        stop || true
    fi

    > "$LOGFILE"
    chmod 644 "$LOGFILE" 2>/dev/null || true

    nohup env MINISERVE_CTL=1 \
        "$MINISERVE" "$ROOT" \
        --port "$PORT" \
        --hidden \
        --verbose \
        >>"$LOGFILE" 2>&1 &

    echo $! >"$PIDFILE" 2>/dev/null || true
    chmod 644 "$PIDFILE" 2>/dev/null || true

    if wait_http_up; then
        echo "✅ miniserve 启动成功"
        echo "访问地址: $URL"
        return 0
    else
        echo "❌ 启动失败（HTTP 未响应）"
        echo "最近日志："
        tail -10 "$LOGFILE" 2>/dev/null
        return 1
    fi
}

# -----------------------------
# restart
# -----------------------------

restart() {
    stop || true
    sleep 1
    start
}

# -----------------------------
# status（最终权威版）
# -----------------------------

status() {
    echo "=== miniserve 服务状态 ==="
    echo "端口: $PORT"
    echo "目录: $ROOT"
    echo "地址: $URL"
    echo ""

    if http_alive; then
        echo "✅ 状态: 运行中（HTTP 可访问）"
        return 0
    fi

    if check_port; then
        echo "⚠️  状态: 端口监听中，但 HTTP 不可访问"
        echo "可能原因："
        echo " - 服务启动中"
        echo " - 服务异常"
        echo " - 非 miniserve 占用端口"
        return 1
    fi

    echo "❌ 状态: 未运行"
    return 1
}

# -----------------------------
# log
# -----------------------------

log() {
    [ -f "$LOGFILE" ] && tail -20 "$LOGFILE" || echo "日志文件不存在"
}

# -----------------------------
# fix
# -----------------------------

fix() {
    echo "修复文件权限..."
    [ -f "$LOGFILE" ] && chmod 644 "$LOGFILE"
    [ -f "$PIDFILE" ] && chmod 644 "$PIDFILE"
    echo "✅ 完成"
}

# -----------------------------
# clean
# -----------------------------

clean() {
    [ -f "$LOGFILE" ] && : >"$LOGFILE"
    echo "✅ 日志已清空"
}

# -----------------------------
# main
# -----------------------------

case "$1" in
    start)   start ;;
    stop)    stop ;;
    restart) restart ;;
    status)  status ;;
    log)     log ;;
    fix)     fix ;;
    clean)   clean ;;
    *)
        echo "用法: $0 {start|stop|restart|status|log|fix|clean}"
        exit 1
        ;;
esac
