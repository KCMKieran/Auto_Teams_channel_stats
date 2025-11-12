#!/bin/bash

# =============================================================================
# Teams 客服群消息统计脚本
# 功能：自动执行数据统计并发送邮件报告
# 作者：Kieran
# =============================================================================

# 设置脚本在出错时退出
set -e

# 定义项目目录（请根据实际路径调整）
PROJECT_DIR="/opt/myproject/channel_stats_for_rebecca/channel_stats"

# 定义日志函数
log() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] $1"
}

log "==================== 开始执行统计任务 ===================="

# 1. 检查并切换到项目目录
if [ ! -d "$PROJECT_DIR" ]; then
    log "错误：项目目录不存在: $PROJECT_DIR"
    exit 1
fi

cd "$PROJECT_DIR" || {
    log "错误：无法切换到项目目录: $PROJECT_DIR"
    exit 1
}

log "当前工作目录: $(pwd)"

# 2. 检查虚拟环境是否存在
if [ ! -f "venv/bin/activate" ]; then
    log "错误：虚拟环境不存在，请先创建虚拟环境"
    exit 1
fi

# 3. 激活虚拟环境
log "激活Python虚拟环境..."
source venv/bin/activate

# 验证Python环境
python_version=$(python --version 2>&1)
log "使用Python版本: $python_version"

# 4. 检查主程序文件是否存在
if [ ! -f "channels_stats.py" ]; then
    log "错误：主程序文件 channels_stats.py 不存在"
    exit 1
fi

# 5. 执行主程序
log "开始执行数据统计程序..."
if python channels_stats.py; then
    log "✅ 数据统计程序执行成功"
else
    log "❌ 数据统计程序执行失败，请检查 run_stats.log 文件"
    exit 1
fi

# 6. 显示当前CSV文件列表
csv_count=$(find . -name "channel_message_stats_*.csv" -type f | wc -l)
if [ "$csv_count" -gt 0 ]; then
    log "当前生成的CSV文件:"
    find . -name "channel_message_stats_*.csv" -type f -exec basename {} \; | sort
else
    log "未找到生成的CSV文件"
fi

# 7. 任务完成
log "==================== 统计任务执行完毕 ===================="

# 退出虚拟环境（可选，脚本结束时会自动退出）
deactivate