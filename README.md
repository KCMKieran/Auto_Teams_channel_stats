# Teams 客服群消息统计项目

## 项目简介
本项目用于自动统计 Microsoft Teams 各客服群上一周的消息数量，生成 CSV 报表，并自动通过邮件发送给指定收件人。

---

## 环境准备

### 1. 部署目录
建议将项目部署在 Ubuntu 系统的 `/opt/channel_stats_for_rebecca/channel_stats/` 目录下。

### 2. Python 环境
- 推荐使用 Python 3.7 及以上版本。
- 使用虚拟环境以隔离项目依赖。

#### 创建与激活虚拟环境
```bash
# 确保系统已安装 python3-venv
sudo apt update
sudo apt install python3-pip python3-venv -y

# 进入项目目录
cd /opt/channel_stats_for_rebecca/channel_stats

# 创建虚拟环境
python3 -m venv venv

# 激活虚拟环境
source venv/bin/activate
```
> **注意**：后续所有命令（如 `pip install`, `python`）都应在激活的虚拟环境下执行。

---

## 项目配置与安装

### 1. 配置文件说明
项目包含两个主要的配置文件：

- **`config.py`**: 用于存放需要统计的 Teams 客服群组列表 `TARGET_TEAMS`。
- **`.env`**: 用于存放所有敏感信息和环境相关配置，**此文件不应提交到版本控制中**。

### 2. 设置环境变量 (`.env`)
在项目根目录下创建 `.env` 文件，并填入以下内容：

```env
# --- Microsoft Graph API 配置 ---
TENANT_ID=你的Azure应用租户ID
CLIENT_ID=你的Azure应用客户端ID
CLIENT_SECRET=你的Azure应用客户端密钥

# --- 邮件发送配置 ---
SMTP_SERVER=你的SMTP服务器地址
SMTP_PORT=587
USERNAME_MAIL=你的发件邮箱账号
PASSWORD_MAIL=你的邮箱密码或应用专用密码
MAIL_SEND_TOO=主送人邮箱地址 (多个用英文逗号,分隔)
MAIL_CCC=抄送人邮箱地址 (多个用英文逗号,分隔)
```

### 3. 安装依赖
```bash
pip install -r requirements.txt
```

---

## 运行与部署

### 1. 手动运行
你可以通过直接运行主脚本来手动触发一次统计任务，这对于测试非常有用：
```bash
python channels_stats.py
```
> 任务执行日志会同时显示在控制台并记录在 `run_stats.log` 文件中。

### 2. 使用Shell脚本运行
为了方便 `cron` 调用，项目提供了一个 `run_stats.sh` 脚本。

#### 脚本内容
```bash
#!/bin/bash
# 跳转到项目目录
cd /opt/channel_stats_for_rebecca/channel_stats || exit

# 激活虚拟环境
source venv/bin/activate

# 运行主程序
python channels_stats.py

# 清理30天前的CSV文件
find . -name "channel_message_stats_*.csv" -type f -mtime +30 -exec rm -f {} \;
```

#### 赋予执行权限
```bash
chmod +x run_stats.sh
```

### 3. 设置定时任务 (crontab)
编辑 `crontab` 以设置定时执行。

```bash
crontab -e
```
在文件末尾添加以下行，设置在每周一早上 9:00 执行脚本：
```
0 9 * * 1 /opt/channel_stats_for_rebecca/channel_stats/run_stats.sh >> /opt/channel_stats_for_rebecca/channel_stats/cron.log 2>&1
```
> `cron` 的所有输出（包括错误）都会被重定向到 `cron.log` 文件，便于排查问题。

---

## 文件与日志管理

- **CSV 报告**: 生成的统计报告会保存在项目根目录，文件名格式为 `channel_message_stats_YYYYMMDD-YYYYMMDD.csv`。
- **运行日志**:
    - `run_stats.log`: 记录了 `channels_stats.py` 脚本每次运行的详细日志，包括 API 请求、数据处理、邮件发送等步骤。
    - `cron.log`: 记录了 `crontab` 定时任务执行的输出信息，主要用于排查调度层面的问题。
- **自动清理**: `run_stats.sh` 脚本会自动删除30天前的 CSV 文件，防止占用过多磁盘空间。天数可在脚本中通过修改 `-mtime +30` 参数来调整。

---

## 常见问题排查

1.  **脚本没执行？**
    - 首先检查 `cron.log` 是否有内容，以及 `run_stats.log` 的时间戳是否更新。
    - 确认 `crontab` 任务的路径是否完全正确。
2.  **依赖包找不到 (ModuleNotFoundError)？**
    - 确认已通过 `source venv/bin/activate` 激活了虚拟环境。
    - 确认已运行 `pip install -r requirements.txt` 安装了所有依赖。
3.  **获取Token失败 或 API请求401/403错误？**
    - 检查 `.env` 文件中的 `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET` 是否正确无误。
4.  **邮件发送失败？**
    - 检查 `.env` 文件中的 `SMTP_*`, `USERNAME_MAIL`, `PASSWORD_MAIL` 等配置是否正确。
    - 确认服务器的网络策略是否允许访问外部SMTP服务。

---

## 附录：项目目录结构
```
/opt/channel_stats_for_rebecca/channel_stats/
  ├── channels_stats.py         # 主统计脚本
  ├── email_utils.py           # 邮件发送工具
  ├── config.py                # 存放Teams群组列表
  ├── requirements.txt         # Python依赖包列表
  ├── .env                     # 环境变量配置文件 (本地创建，不提交)
  ├── venv/                    # Python虚拟环境目录
  ├── run_stats.sh             # 运行脚本 (用于cron)
  ├── run_stats.log            # 主程序运行日志
  ├── cron.log                 # 定时任务日志
  └── ...                      # 其他生成的文件
``` 