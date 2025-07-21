# Teams Channel Statistics Project 

**Author**: Kieran  
**Last Updated**: 2025-07-21    
**Email**: Kieran.xiang@kohleservices.com


## Project Overview
This project automatically counts the number of messages in each Microsoft Teams channels for the past week, generates a CSV report, and automatically sends it via email to the designated recipient (rebecca).

---

## Environment Preparation

### 1. Deployment Directory
The project has deployed in the `/opt/channel_stats_for_rebecca/channel_stats/` directory on Ubuntu. 

### 2. Python Environment
- Python 3.7 or above is recommended.
- Use a virtual environment to isolate project dependencies.

#### Create and Activate Virtual Environment
```bash
# Ensure python3-venv is installed
sudo apt update
sudo apt install python3-pip python3-venv -y

# Enter the project directory
cd /opt/channel_stats_for_rebecca/channel_stats

# Create a virtual environment
python3 -m venv venv

# Activate the virtual environment
source venv/bin/activate
```
> **Note:** All subsequent commands (such as `pip install`, `python`) should be run within the activated virtual environment.

---

## Project Configuration & Installation

### 1. Configuration Files
The project contains two main configuration files:

- **`config.py`**: Stores the list of Teams customer service groups to be counted as `TARGET_TEAMS`.
- **`.env`**: Stores all sensitive information and environment-related configurations. **This file should NOT be committed to version control.**

### 2. Set Environment Variables (`.env`)
Create a `.env` file in the project root directory and fill in the following content:

```env
# --- Microsoft Graph API Configuration ---
TENANT_ID=Your Azure Tenant ID
CLIENT_ID=Your Azure App Client ID
CLIENT_SECRET=Your Azure App Client Secret

# --- Email Sending Configuration ---
SMTP_SERVER=Your SMTP server address
SMTP_PORT=587
USERNAME_MAIL=Your sender email account
PASSWORD_MAIL=Your email password or app-specific password
MAIL_SEND_TOO=Primary recipient email address (comma-separated for multiple)
MAIL_CCC=CC recipient email address (comma-separated for multiple)
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

---

## Running & Deployment

### 1. Manual Run
You can manually trigger a statistics task by running the main script, which is useful for testing:
```bash
python channels_stats.py
```
> Execution logs will be displayed in the console and recorded in the `run_stats.log` file.

### 2. Run with Shell Script
To facilitate `cron` scheduling, the project provides a `run_stats.sh` script.

#### Script Content
```bash
#!/bin/bash
# Change to project directory
cd /opt/channel_stats_for_rebecca/channel_stats || exit

# Activate virtual environment
source venv/bin/activate

# Run the main program
python channels_stats.py

# Clean up CSV files older than 30 days
find . -name "channel_message_stats_*.csv" -type f -mtime +30 -exec rm -f {} \;
```

#### Grant Execute Permission
```bash
chmod +x run_stats.sh
```

### 3. Set Up Scheduled Task (crontab)
Edit `crontab` to schedule automatic execution.

```bash
# Edit scheduled tasks:
crontab -e  # Open editor to modify cron job configuration
crontab -l  # View all current cron jobs
crontab -r  # Delete all cron job configurations

# Add the following content using crontab -e to run script every Monday at 9am:
```

```
0 9 * * 1 /opt/channel_stats_for_rebecca/channel_stats/run_stats.sh >> /opt/channel_stats_for_rebecca/channel_stats/cron.log 2>&1
```
> All output (including errors) from `cron` will be redirected to the `cron.log` file for troubleshooting.

---

## File & Log Management

- **CSV Reports**: Generated reports are saved in the project root directory, named as `channel_message_stats_YYYYMMDD-YYYYMMDD.csv`.
- **Run Logs**:
    - `run_stats.log`: Detailed log of each run of `channels_stats.py`, including API requests, data processing, email sending, etc.
    - `cron.log`: Output from scheduled tasks, mainly for troubleshooting scheduling issues.
- **Automatic Cleanup**: The `run_stats.sh` script automatically deletes CSV files older than 30 days to prevent excessive disk usage. You can adjust the number of days by modifying the `-mtime +30` parameter in the script.

---

## Troubleshooting

1.  **Script not running?**
    - First, check if `cron.log` has content and whether the timestamp in `run_stats.log` is updated.
    - Make sure the `crontab` task path is absolutely correct.
2.  **ModuleNotFoundError for dependencies?**
    - Ensure the virtual environment is activated with `source venv/bin/activate`.
    - Ensure all dependencies are installed with `pip install -r requirements.txt`.
3.  **Token acquisition failed or API returns 401/403?**
    - Check if `TENANT_ID`, `CLIENT_ID`, and `CLIENT_SECRET` in `.env` are correct.
4.  **Email sending failed?**
    - Check if `SMTP_*`, `USERNAME_MAIL`, `PASSWORD_MAIL` in `.env` are correct.
    - Ensure the server's network policy allows access to external SMTP services.

---

## Appendix: Project Directory Structure
```
/opt/channel_stats_for_rebecca/channel_stats/
  ├── channels_stats.py         # Main statistics script
  ├── email_utils.py           # Email sending utilities
  ├── config.py                # Teams group list
  ├── requirements.txt         # Python dependencies
  ├── .env.example             # Environment variable config (local only, not committed)
  ├── venv/                    # Python virtual environment
  ├── run_stats.sh             # Run script (for cron)
  ├── run_stats.log            # Main program log
  ├── cron.log                 # Scheduled task log
  └── ...                      # Other generated files
``` 