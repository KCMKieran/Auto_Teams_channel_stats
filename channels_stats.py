from msal import ConfidentialClientApplication
from dotenv import load_dotenv
import os
import requests
import datetime
from collections import defaultdict
import pandas as pd
from typing import List, Dict, Any

def get_access_token() -> str:
    """Get Microsoft Graph API access token"""
    load_dotenv()
    tenant_id = os.getenv('TENANT_ID')
    client_id = os.getenv('CLIENT_ID')
    client_secret = os.getenv('CLIENT_SECRET')

    authority = f'https://login.microsoftonline.com/{tenant_id}'
    scope = ['https://graph.microsoft.com/.default']

    app = ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )
    token_result = app.acquire_token_for_client(scopes=scope)
    return token_result['access_token']

def get_teams(headers: Dict[str, str]) -> List[Dict[str, Any]]:
    """Get all Teams from Microsoft Graph API"""
    teams = []
    url = "https://graph.microsoft.com/v1.0/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"

    while url:
        resp = requests.get(url, headers=headers).json()
        teams.extend(resp.get('value', []))
        url = resp.get('@odata.nextLink')
    
    return teams

def get_target_channels(teams: List[Dict[str, Any]], headers: Dict[str, str], target_names: List[str]) -> List[Dict[str, Any]]:
    """Get channels for target teams"""
    target_names_normalized = [name.replace(' ', '') for name in target_names]
    channels_to_check = []

    for team in teams:
        team_name = team.get('displayName', '')
        if team_name.replace(' ', '') in target_names_normalized:
            team_id = team['id']
            ch_url = f'https://graph.microsoft.com/v1.0/teams/{team_id}/channels'
            resp = requests.get(ch_url, headers=headers).json()
            for ch in resp.get('value', []):
                channels_to_check.append({
                    'team_id': team_id,
                    'channel_id': ch['id'],
                    'channel_name': ch['displayName'],
                    'team_name': team_name
                })
    
    return channels_to_check

def get_channel_messages(channel: Dict[str, Any], headers: Dict[str, str], cutoff_time: datetime.datetime) -> Dict[str, Dict[str, int]]:
    """Get message statistics for a specific channel including replies"""
    sender_stats = defaultdict(lambda: defaultdict(int))
    team_id = channel['team_id']
    channel_id = channel['channel_id']
    channel_key = f"{channel['team_name']} - {channel['channel_name']}"
    
    print(f"▶ Processing: {channel_key}")
    
    def fetch_replies(message_id: str) -> None:
        """Fetch and count all replies for a specific message"""
        replies_url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies"
        
        while replies_url:
            reply_resp = requests.get(replies_url, headers=headers)
            if reply_resp.status_code == 429:  # Rate limit hit
                retry_after = int(reply_resp.headers.get("Retry-After", 1))
                print(f"Rate limit hit in replies, waiting {retry_after} seconds...")
                time.sleep(retry_after)
                continue
                
            reply_data = reply_resp.json()
            
            for reply in reply_data.get('value', []):
                reply_time_str = reply.get('createdDateTime')
                try:
                    reply_time = datetime.datetime.strptime(reply_time_str, '%Y-%m-%dT%H:%M:%S.%fZ')
                except ValueError:
                    reply_time = datetime.datetime.strptime(reply_time_str, '%Y-%m-%dT%H:%M:%SZ')
                
                if reply_time < cutoff_time:
                    return
                
                reply_sender_info = reply.get('from')
                if reply_sender_info and 'user' in reply_sender_info:
                    reply_sender = reply_sender_info.get('user', {}).get('displayName') or reply_sender_info.get('user', {}).get('id') or 'Unknown'
                else:
                    reply_sender = 'Unknown'
                
                sender_stats[reply_sender][channel_key] += 1
            
            replies_url = reply_data.get('@odata.nextLink')
    
    url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages"
    
    while url:
        resp = requests.get(url, headers=headers)
        if resp.status_code == 429:  # Rate limit hit
            retry_after = int(resp.headers.get("Retry-After", 1))
            print(f"Rate limit hit in messages, waiting {retry_after} seconds...")
            time.sleep(retry_after)
            continue

        data = resp.json()
        
        for msg in data.get('value', []):
            msg_time_str = msg.get('createdDateTime')
            try:
                msg_time = datetime.datetime.strptime(msg_time_str, '%Y-%m-%dT%H:%M:%S.%fZ')
            except ValueError:
                msg_time = datetime.datetime.strptime(msg_time_str, '%Y-%m-%dT%H:%M:%SZ')

            if msg_time < cutoff_time:
                return sender_stats

            # Count main message
            sender_info = msg.get('from')
            if sender_info and 'user' in sender_info:
                sender = sender_info.get('user', {}).get('displayName') or sender_info.get('user', {}).get('id') or 'Unknown'
            else:
                sender = 'Unknown'

            sender_stats[sender][channel_key] += 1
            
            # Get and count replies for this message
            message_id = msg.get('id')
            if message_id:
                fetch_replies(message_id)

        url = data.get('@odata.nextLink')
    
    return sender_stats

def generate_message_stats(output_file: str = None) -> None:
    """Main function to generate message statistics"""    # Generate output filename with date range if not provided
    if output_file is None:
        now = datetime.datetime.utcnow()
        today_weekday = now.weekday()
        # 计算上周日的日期
        last_sunday = now - datetime.timedelta(days=today_weekday + 1)
        # 计算上周一的日期
        last_monday = last_sunday - datetime.timedelta(days=6)
        
        # 格式化日期为YYYYMMDD格式
        start_date = last_monday.strftime("%Y%m%d")
        end_date = last_sunday.strftime("%Y%m%d")
        
        # 使用 os.path.join 确保路径正确
        output_dir = os.path.join(os.path.dirname(__file__))  # 获取当前脚本所在目录
        output_file = os.path.join(output_dir, f"channel_message_stats_{start_date}-{end_date}.csv")

    # Target Teams names
    target_names = [
        "HZL013客服群", "SHF001 客服群", "SZU000 客服群",
        "HZL009 客服群", "SHY03 客服群", "SHT042 客服群",
        "HZL008 客服群", "HZL0111客服群", "JSA000 客服群",
        "SHT000 客服群", "HNE000 客服群", "SZS000 客服群",
        "SHS000 客服群", "SHY01 客服群", "SHY02 客服群",
        "HZL 005 客服群", "SHP000 客服群", "SHT049 客服群",
        "CCX000 客服群", "HZL014客服群", "CCX003 客服群",
        "VSH000 客服群", "CCX004 客服群", "SZS003 客服群",
        "HZL012客服群"
    ]

    # Get access token and set up headers
    access_token = get_access_token()
    headers = {'Authorization': f'Bearer {access_token}'}

    # Get teams and channels
    teams = get_teams(headers)
    channels_to_check = get_target_channels(teams, headers, target_names)    # 设置时间范围（当前时间往前推7天）
    now = datetime.datetime.utcnow()
    # 获取今天是周几（0是周一，6是周日）
    today_weekday = now.weekday()
    # 计算上周日的日期（当前日期 - 今天是周几 - 1）
    last_sunday = now - datetime.timedelta(days=today_weekday + 1)
    # 将上周日设置为23:59:59
    end_time = last_sunday.replace(hour=23, minute=59, second=59, microsecond=999999)
    # 计算上周一的日期（上周日 - 6天）
    last_monday = last_sunday - datetime.timedelta(days=6)
    # 将上周一设置为00:00:00
    seven_days_ago = last_monday.replace(hour=0, minute=0, second=0, microsecond=0)
    print(f"搜索时间范围: {seven_days_ago} 到 {end_time}")

    # Process all channels and collect statistics
    all_sender_stats = defaultdict(lambda: defaultdict(int))
    
    for channel in channels_to_check:
        channel_stats = get_channel_messages(channel, headers, seven_days_ago)
        for sender, stats in channel_stats.items():
            for channel_key, count in stats.items():
                all_sender_stats[sender][channel_key] += count

    # Convert to DataFrame and save to CSV
    rows = []
    for sender, channel_counts in all_sender_stats.items():
        total = sum(channel_counts.values())
        row = {'Sender': sender, 'Total Messages': total}
        row.update(channel_counts)
        rows.append(row)

    df = pd.DataFrame(rows)
    df = df.sort_values(by='Total Messages', ascending=False)
    df.to_csv(output_file, index=False, encoding='utf-8-sig')
    print(f"✅ Statistics completed, results saved to {output_file}")
    print(df)

if __name__ == "__main__":
    import time
    start_time = time.time()
    generate_message_stats()
    end_time = time.time()
    print(f"Total execution time: {end_time - start_time:.2f} seconds")
