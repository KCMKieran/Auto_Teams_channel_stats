from msal import ConfidentialClientApplication
from dotenv import load_dotenv
import os
import requests
import datetime
import time
from collections import defaultdict
import pandas as pd
from typing import List, Dict, Any, Optional

# Import from new module
from email_utils import send_email_with_attachments
from config import TARGET_TEAMS

# --- Logging Configuration ---
import logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("run_stats.log", mode='a', encoding='utf-8'),
        # mode = 'a' means append, 'w' means overwrite
        logging.StreamHandler()
    ]
)

# --- Unified API Request Function ---
def robust_request(url: str, headers: Dict[str, str], retries: int = 3, timeout: int = 30) -> Dict:
    """Encapsulated request with retry, timeout, and error handling"""
    for attempt in range(retries):
        try:
            response = requests.get(url, headers=headers, timeout=timeout)
            if response.status_code == 429:
                retry_after = int(response.headers.get("Retry-After", 5))
                logging.warning(f"API rate limit triggered, retrying in {retry_after} seconds...")
                time.sleep(retry_after)
                continue
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            logging.error(f"Request failed (attempt {attempt + 1}/{retries}): {e}")
            if attempt < retries - 1:
                time.sleep(2 ** attempt)  # Exponential backoff
    return {} # Return empty dict after multiple failures

def get_access_token() -> str:
    """Get Microsoft Graph API access token"""
    load_dotenv(override=True) 
    tenant_id = os.getenv('TENANT_ID')
    client_id = os.getenv('CLIENT_ID')
    client_secret = os.getenv('CLIENT_SECRET')
    if not all([tenant_id, client_id, client_secret]):
        raise ValueError("Environment variables TENANT_ID, CLIENT_ID, or CLIENT_SECRET are not set!")

    authority = f'https://login.microsoftonline.com/{tenant_id}'
    scope = ['https://graph.microsoft.com/.default']
    app = ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)
    token_result = app.acquire_token_for_client(scopes=scope)
    
    if not token_result or "access_token" not in token_result:
        logging.error(f"Failed to get token: {token_result.get('error_description') if token_result else 'No response'}")
        raise Exception("Unable to obtain access token")
        
    logging.info("Successfully obtained access token")
    return token_result['access_token']

def get_teams(headers: Dict[str, str]) -> List[Dict[str, Any]]:
    """Get all Teams"""
    teams = []
    url = "https://graph.microsoft.com/v1.0/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"
    while url:
        resp = robust_request(url, headers=headers)
        teams.extend(resp.get('value', []))
        url = resp.get('@odata.nextLink')
    return teams

def get_target_channels(teams: List[Dict[str, Any]], headers: Dict[str, str], target_names: List[str]) -> List[Dict[str, Any]]:
    """Get channels of target teams"""
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
    """Parse API returned date string to datetime object"""
    for fmt in ('%Y-%m-%dT%H:%M:%S.%fZ', '%Y-%m-%dT%H:%M:%SZ'):
        try:
            return datetime.datetime.strptime(datetime_str, fmt)
        except (ValueError, TypeError):
            continue
    logging.warning(f"Unable to parse date format: {datetime_str}")
    return None

def extract_sender_name(message: Dict[str, Any]) -> str:
    """Extract sender name from message"""
    sender_info = message.get('from')
    if sender_info and 'user' in sender_info:
        user = sender_info.get('user', {})
        return user.get('displayName') or user.get('id') or 'Unknown'
    return 'Unknown'

def get_channel_messages(channel: Dict[str, Any], headers: Dict[str, str], cutoff_time: datetime.datetime) -> Dict[str, Dict[str, int]]:
    """Get message statistics for the specified channel, including replies"""
    sender_stats = defaultdict(lambda: defaultdict(int))
    team_id = channel['team_id']
    channel_id = channel['channel_id']
    channel_key = f"{channel['team_name']} - {channel['channel_name']}"
    
    def process_messages(messages: List[Dict[str, Any]], is_reply: bool = False) -> bool:
        """Process message list (main messages or replies)"""
        for msg in messages:
            msg_time = parse_datetime(msg.get('createdDateTime', ''))
            if not msg_time or msg_time < cutoff_time:
                return True # Out of time range, stop further paging
            
            sender = extract_sender_name(msg)
            sender_stats[sender][channel_key] += 1
            
            if not is_reply and msg.get('id'):
                fetch_replies(msg['id'])
        return False

    def fetch_replies(message_id: str) -> None:
        """Get and count all replies to a single message"""
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
    
    return dict(sender_stats) # Convert to regular dict to match type hint

def generate_message_stats(output_file: Optional[str] = None) -> None:
    """Main function, generate message statistics"""
    try:
        if output_file is None:
            now = datetime.datetime.utcnow()
            last_sunday = now - datetime.timedelta(days=now.weekday() + 1)
            last_monday = last_sunday - datetime.timedelta(days=6)
            start_date = last_monday.strftime("%Y%m%d")
            end_date = last_sunday.strftime("%Y%m%d")
            output_file = f"channel_message_stats_{start_date}-{end_date}.csv"

        # Time range setup
        last_sunday = datetime.datetime.utcnow() - datetime.timedelta(days=datetime.datetime.utcnow().weekday() + 1)
        end_time = last_sunday.replace(hour=23, minute=59, second=59, microsecond=999999)
        start_time = (last_sunday - datetime.timedelta(days=6)).replace(hour=0, minute=0, second=0, microsecond=0)
        logging.info(f"Search time range: {start_time} to {end_time}")

        access_token = get_access_token()
        headers = {'Authorization': f'Bearer {access_token}'}

        logging.info("Fetching all team list...")
        teams = get_teams(headers)
        logging.info(f"Filtering channels from {len(TARGET_TEAMS)} target groups...")
        channels_to_check = get_target_channels(teams, headers, TARGET_TEAMS)
        logging.info(f"Found {len(channels_to_check)} channels to process.")

        all_sender_stats = defaultdict(lambda: defaultdict(int))
        for i, channel in enumerate(channels_to_check, 1):
            logging.info(f"Progress: {i}/{len(channels_to_check)} - {channel['team_name']} - {channel['channel_name']}")
            channel_stats = get_channel_messages(channel, headers, start_time)
            for sender, stats in channel_stats.items():
                for channel_key, count in stats.items():
                    all_sender_stats[sender][channel_key] += count
        
        if not all_sender_stats:
            logging.warning("No message data collected, no CSV file will be generated, and no email will be sent.")
            return

        rows = []
        for sender, channel_counts in all_sender_stats.items():
            row = {'Sender': sender, 'Total Messages': sum(channel_counts.values())}
            row.update(channel_counts)
            rows.append(row)

        df = pd.DataFrame(rows).sort_values(by='Total Messages', ascending=False)
        df.to_csv(output_file, index=False, encoding='utf-8-sig')
        logging.info(f"✅ Statistics completed, result saved to {output_file}")
        logging.info(f"\n{df.to_string()}")

        # --- Email sending integration ---
        logging.info("Sending statistics report email...")
        subject = f"Teams Channel Message Weekly Report ({start_date} - {end_date})"
        body = f"""
        <p>Hi Rebecca,</p>
        <p>The attachment is the Teams channel message statistics report for last week ({start_date} - {end_date}).</p>
        <br>
        <p>Best regards,</p>
        <hr>
        <p><i>This email was sent automatically by Kieran</i></p>
        """
        try:
            send_email_with_attachments(subject=subject, body=body, attachments=[output_file])
        except Exception as e:
            logging.error(f"Failed to send email: {e}")

    except Exception as e:
        logging.error(f"Task execution failed: {e}", exc_info=True)
        # exc_info=True will record the full traceback

if __name__ == "__main__":
    start_time_main = time.time()
    logging.info("="*20 + " Start statistics task " + "="*20)
    generate_message_stats()
    end_time_main = time.time()
    logging.info(f"Total execution time: {end_time_main - start_time_main:.2f} seconds")
    logging.info("="*20 + " Task finished " + "="*20)
