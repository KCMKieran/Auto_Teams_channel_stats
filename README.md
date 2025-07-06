# Teams 客服群消息统计项目

## 项目简介
本项目用于自动统计 Microsoft Teams 各客服群上一周的消息数量，并生成 CSV 报表。支持定时任务自动运行，并通过 bash 脚本实现历史数据自动清理，适合公司内网服务器部署。

---

## 环境准备

### 1. 推荐部署目录
建议将项目部署在 Ubuntu 系统的 `/opt/teams-stats/` 目录下，便于管理和权限控制。

### 2. Python 环境与虚拟环境
- 推荐使用 Python 3.7 及以上版本。
- 使用虚拟环境隔离依赖，避免与系统环境冲突。

#### 创建虚拟环境
```bash
sudo apt update
sudo apt install python3 python3-pip python3-venv -y
cd /opt/teams-stats
python3 -m venv venv
source venv/bin/activate
```

---

## 目录结构建议
```
/opt/teams-stats/
  ├── channels_stats.py         # 主统计脚本
  ├── requirements.txt         # 依赖包列表
  ├── venv/                    # Python 虚拟环境
  ├── run_stats.sh             # 运行脚本
  ├── run_stats.log            # 运行日志
  ├── cron.log                 # 定时任务日志
  ├── output/                  # 生成的 CSV 文件
  └── ...                      # 其他配置文件
```

---

## 依赖安装

激活虚拟环境后，安装依赖：
```bash
source /opt/teams-stats/venv/bin/activate
pip install -r requirements.txt
```

---

## Bash 脚本编写与说明

新建 `run_stats.sh`，内容如下：
```bash
#!/bin/bash
cd /opt/teams-stats
source venv/bin/activate
python channels_stats.py
# 清理30天前的csv文件
find ./output -name "channel_message_stats_*.csv" -type f -mtime +30 -exec rm -f {} \;
echo "$(date '+%Y-%m-%d %H:%M:%S') 统计任务已完成" >> run_stats.log
```
> 注意：
> - `cd /opt/teams-stats` 跳转到项目目录
> - `source venv/bin/activate` 激活虚拟环境
> - `find` 命令自动清理30天前的csv文件，可根据需要调整天数
> - 日志追加到 `run_stats.log`

赋予脚本可执行权限：
```bash
chmod +x /opt/teams-stats/run_stats.sh
```

---

## crontab 定时任务设置

编辑定时任务：
```bash
crontab -e
```
添加如下行（每周一早上9点执行）：
```
0 9 * * 1 /opt/teams-stats/run_stats.sh >> /opt/teams-stats/cron.log 2>&1
```
> 这样每次执行的输出和错误信息都会记录到 `cron.log` 文件中，方便排查问题。

---

## CSV 文件自动清理方案

- 脚本中已集成自动清理命令：
  ```bash
  find ./output -name "channel_message_stats_*.csv" -type f -mtime +30 -exec rm -f {} \;
  ```
- 只保留最近30天的统计文件，防止磁盘空间被长期占用。
- 可根据实际需求调整 `-mtime +30` 的天数。

---

## 常见问题与维护建议

1. **脚本没执行怎么办？**
   - 检查 `cron.log` 和 `run_stats.log`，查看是否有报错信息。
   - 确认 crontab 任务是否正确添加。
2. **依赖包找不到？**
   - 确认已激活虚拟环境，并在虚拟环境中安装依赖。
3. **csv文件会不会被误删？**
   - 只会删除30天前的csv文件，最近的都保留。
4. **如何修改保留天数？**
   - 修改脚本中 `-mtime +30` 的数字即可。
5. **如何手动测试？**
   - 直接运行 `./run_stats.sh`，看是否能正常生成csv和清理旧文件。
6. **如何查看csv文件？**
   - `ls /opt/teams-stats/output/channel_message_stats_*.csv`

---

## 结语

本项目适合公司内网定时统计任务，部署简单，维护方便。如有更多需求（如邮件发送、日志告警等），可在此基础上扩展。遇到问题欢迎随时咨询！ 