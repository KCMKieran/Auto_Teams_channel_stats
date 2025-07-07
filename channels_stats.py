from msal import ConfidentialClientApplication
from dotenv import load_dotenv
import os
import requests
import datetime
import time
from collections import defaultdict
import pandas as pd
from typing import List, Dict, Any, Optional

# 从新模块导入
from email_utils import send_email_with_attachments
from config import TARGET_TEAMS

# --- 日志配置 ---
import logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("run_stats.log", mode='w', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# --- 统一的API请求函数 ---
def robust_request(url: str, headers: Dict[str, str], retries: int = 3, timeout: int = 30) -> Dict:
    """封装请求，包含重试、超时和错误处理"""
    for attempt in range(retries):
        try:
            response = requests.get(url, headers=headers, timeout=timeout)
            if response.status_code == 429:
                retry_after = int(response.headers.get("Retry-After", 5))
                logging.warning(f"触发API限流，将在 {retry_after} 秒后重试...")
                time.sleep(retry_after)
                continue
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            logging.error(f"请求失败 (尝试 {attempt + 1}/{retries}): {e}")
            if attempt < retries - 1:
                time.sleep(2 ** attempt)  # 指数退避策略
    return {} # 多次失败后返回空字典

def get_access_token() -> str:
    """获取Microsoft Graph API访问令牌"""
    load_dotenv(override=True) 
    tenant_id = os.getenv('TENANT_ID')
    client_id = os.getenv('CLIENT_ID')
    client_secret = os.getenv('CLIENT_SECRET')
    if not all([tenant_id, client_id, client_secret]):
        raise ValueError("环境变量 TENANT_ID, CLIENT_ID, 或 CLIENT_SECRET 未设置!")

    authority = f'https://login.microsoftonline.com/{tenant_id}'
    scope = ['https://graph.microsoft.com/.default']
    app = ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)
    token_result = app.acquire_token_for_client(scopes=scope)
    
    if not token_result or "access_token" not in token_result:
        logging.error(f"获取Token失败: {token_result.get('error_description') if token_result else '无响应'}")
        raise Exception("无法获取访问令牌")
        
    logging.info("成功获取访问令牌")
    return token_result['access_token']

def get_teams(headers: Dict[str, str]) -> List[Dict[str, Any]]:
    """获取所有Teams"""
    teams = []
    url = "https://graph.microsoft.com/v1.0/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"
    while url:
        resp = robust_request(url, headers=headers)
        teams.extend(resp.get('value', []))
        url = resp.get('@odata.nextLink')
    return teams

def get_target_channels(teams: List[Dict[str, Any]], headers: Dict[str, str], target_names: List[str]) -> List[Dict[str, Any]]:
    """获取目标团队的频道"""
    target_names_normalized = [name.replace(' ', '') for name in target_names]
    channels_to_check = []
    for team in teams:
        team_name = team.get('displayName', '')
        if team_name.replace(' ', '') in target_names_normalized:
            team_id = team['id']
            ch_url = f'https://graph.microsoft.com/v1.0/teams/{team_id}/channels'
            resp = robust_request(ch_url, headers=headers)
            for ch in resp.get('value', []):
                channels_to_check.append({
                    'team_id': team_id,
                    'channel_id': ch['id'],
                    'channel_name': ch['displayName'],
                    'team_name': team_name
                })
    return channels_to_check

def parse_datetime(datetime_str: str) -> Optional[datetime.datetime]:
    """将API返回的日期字符串解析为datetime对象"""
    for fmt in ('%Y-%m-%dT%H:%M:%S.%fZ', '%Y-%m-%dT%H:%M:%SZ'):
        try:
            return datetime.datetime.strptime(datetime_str, fmt)
        except (ValueError, TypeError):
            continue
    logging.warning(f"无法解析日期格式: {datetime_str}")
    return None

def extract_sender_name(message: Dict[str, Any]) -> str:
    """从消息中提取发件人姓名"""
    sender_info = message.get('from')
    if sender_info and 'user' in sender_info:
        user = sender_info.get('user', {})
        return user.get('displayName') or user.get('id') or 'Unknown'
    return 'Unknown'

def get_channel_messages(channel: Dict[str, Any], headers: Dict[str, str], cutoff_time: datetime.datetime) -> Dict[str, Dict[str, int]]:
    """获取指定频道的消息统计数据，包括回复"""
    sender_stats = defaultdict(lambda: defaultdict(int))
    team_id = channel['team_id']
    channel_id = channel['channel_id']
    channel_key = f"{channel['team_name']} - {channel['channel_name']}"
    
    def process_messages(messages: List[Dict[str, Any]], is_reply: bool = False) -> bool:
        """处理消息列表（主消息或回复）"""
        for msg in messages:
            msg_time = parse_datetime(msg.get('createdDateTime', ''))
            if not msg_time or msg_time < cutoff_time:
                return True # 已超出时间范围，停止后续分页
            
            sender = extract_sender_name(msg)
            sender_stats[sender][channel_key] += 1
            
            if not is_reply and msg.get('id'):
                fetch_replies(msg['id'])
        return False

    def fetch_replies(message_id: str) -> None:
        """获取并统计单条消息的所有回复"""
        replies_url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies"
        while replies_url:
            reply_resp = robust_request(replies_url, headers)
            if process_messages(reply_resp.get('value', []), is_reply=True):
                break
            replies_url = reply_resp.get('@odata.nextLink')

    url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages"
    while url:
        resp = robust_request(url, headers)
        if process_messages(resp.get('value', [])):
            break
        url = resp.get('@odata.nextLink')
    
    return dict(sender_stats) # 转换为普通字典以匹配类型提示

def generate_message_stats(output_file: Optional[str] = None) -> None:
    """主函数，生成消息统计数据"""
    try:
        if output_file is None:
            now = datetime.datetime.utcnow()
            last_sunday = now - datetime.timedelta(days=now.weekday() + 1)
            last_monday = last_sunday - datetime.timedelta(days=6)
            start_date = last_monday.strftime("%Y%m%d")
            end_date = last_sunday.strftime("%Y%m%d")
            output_file = f"channel_message_stats_{start_date}-{end_date}.csv"

        # 时间范围设定
        last_sunday = datetime.datetime.utcnow() - datetime.timedelta(days=datetime.datetime.utcnow().weekday() + 1)
        end_time = last_sunday.replace(hour=23, minute=59, second=59, microsecond=999999)
        start_time = (last_sunday - datetime.timedelta(days=6)).replace(hour=0, minute=0, second=0, microsecond=0)
        logging.info(f"搜索时间范围: {start_time} 到 {end_time}")

        access_token = get_access_token()
        headers = {'Authorization': f'Bearer {access_token}'}

        logging.info("正在获取所有团队列表...")
        teams = get_teams(headers)
        logging.info(f"正在从 {len(TARGET_TEAMS)} 个目标群组中筛选频道...")
        channels_to_check = get_target_channels(teams, headers, TARGET_TEAMS)
        logging.info(f"共找到 {len(channels_to_check)} 个频道需要处理。")

        all_sender_stats = defaultdict(lambda: defaultdict(int))
        for i, channel in enumerate(channels_to_check, 1):
            logging.info(f"处理进度: {i}/{len(channels_to_check)} - {channel['team_name']} - {channel['channel_name']}")
            channel_stats = get_channel_messages(channel, headers, start_time)
            for sender, stats in channel_stats.items():
                for channel_key, count in stats.items():
                    all_sender_stats[sender][channel_key] += count
        
        if not all_sender_stats:
            logging.warning("未收集到任何消息数据，不生成CSV文件，也不发送邮件。")
            return

        rows = []
        for sender, channel_counts in all_sender_stats.items():
            row = {'Sender': sender, 'Total Messages': sum(channel_counts.values())}
            row.update(channel_counts)
            rows.append(row)

        df = pd.DataFrame(rows).sort_values(by='Total Messages', ascending=False)
        df.to_csv(output_file, index=False, encoding='utf-8-sig')
        logging.info(f"✅ 统计完成，结果已保存到 {output_file}")
        logging.info(f"\n{df.to_string()}")

        # --- 集成邮件发送 ---
        logging.info("开始发送统计报告邮件...")
        subject = f"客服群消息统计周报 ({start_date} - {end_date})"
        body = f"""
        <p>Hi Rebecca，</p>
        <p>附件为上周({start_date} - {end_date})的客服群消息统计报告。</p>
        <br>
        <p>Best </p>
        <hr>
        <p><i>此邮件为Kieran自动发送</i></p>
        """
        try:
            send_email_with_attachments(subject=subject, body=body, attachments=[output_file])
        except Exception as e:
            logging.error(f"邮件发送失败: {e}")

    except Exception as e:
        logging.error(f"任务执行失败: {e}", exc_info=True)
        # exc_info=True 会记录完整的traceback

if __name__ == "__main__":
    start_time_main = time.time()
    logging.info("="*20 + " 开始执行统计任务 " + "="*20)
    generate_message_stats()
    end_time_main = time.time()
    logging.info(f"总执行耗时: {end_time_main - start_time_main:.2f} 秒")
    logging.info("="*20 + " 任务执行结束 " + "="*20)
