# Quick Start Guide - Python Implementation

## Overview
This guide will help you quickly set up and run the Python-based Copilot audit data collector.

## Prerequisites
- Python 3.7 or higher
- PowerShell 7 (for setup script)
- Azure AD Global Administrator or Application Administrator access

## Step-by-Step Setup

### 1. Install Python Dependencies
```bash
pip install -r requirements.txt
```

### 2. Create Azure AD App Registration

**Option A: Using PowerShell Script (Recommended)**
```powershell
# Open PowerShell 7 as Administrator
.\Setup-AzureAppRegistration.ps1
```

This script will:
- Create the Azure AD App Registration
- Configure required permissions
- Generate a client secret
- Create the .env file with your credentials
- Provide instructions for granting admin consent

**Important**: After the script completes, grant admin consent in the Azure Portal using the URL provided by the script.

**Option B: Manual Setup**
1. Go to Azure Portal > Azure Active Directory > App registrations
2. Click "New registration"
3. Name: "Copilot Audit Data Collector"
4. Supported account types: "Accounts in this organizational directory only"
5. Click "Register"
6. Note the Application (client) ID and Directory (tenant) ID
7. Go to "Certificates & secrets" > "New client secret"
8. Create a secret and note the value
9. Go to "API permissions"
10. Add the following permissions:
    - Microsoft Graph API:
      - User.Read.All (Application)
      - Directory.Read.All (Application)
      - AuditLog.Read.All (Application)
    - Office 365 Management APIs:
      - ActivityFeed.Read (Application)
      - ServiceHealth.Read (Application)
11. Click "Grant admin consent for [Your Organization]"
12. Copy .env.example to .env and fill in your values

### 3. Verify Configuration

Check that your .env file contains:
```bash
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
```

### 4. Run the Script

**First Run - Collect All Data**
```bash
python copilot_audit.py
```

**Subsequent Runs - Incremental Updates**
```bash
python copilot_audit.py
```
The script automatically detects the last event timestamp and only retrieves new data.

**Collect Only User Data**
```bash
python copilot_audit.py --users-only
```

**Collect Only Event Data**
```bash
python copilot_audit.py --events-only
```

**Custom Output Directory**
```bash
python copilot_audit.py --output-dir /path/to/output
```

## Expected Output

After running successfully, you'll find these files in the output directory:

1. **Copilot_Users.csv** - User information with Copilot license status
2. **Copilot_Events.csv** - Copilot interaction events
3. **copilot_audit.log** - Detailed execution log
4. **AuditScriptLog.txt** - Audit-specific log

## Scheduling Automated Runs

### Windows Task Scheduler

1. Open Task Scheduler
2. Create Basic Task
3. Name: "Copilot Audit Data Collection"
4. Trigger: Daily (or your preferred schedule)
5. Action: Start a program
   - Program: `python.exe` (or find your Python path: `where python`)
   - Arguments: `copilot_audit.py`
   - Start in: `C:\path\to\GonvarriDashboards`
6. Finish

### Linux/macOS Cron

Edit crontab:
```bash
crontab -e
```

Add entry (runs daily at 2 AM):
```bash
0 2 * * * cd /path/to/GonvarriDashboards && /usr/bin/python3 copilot_audit.py >> /var/log/copilot_audit.log 2>&1
```

## Troubleshooting

### Issue: "Missing required environment variables"
**Solution**: Ensure .env file exists and contains all three required variables

### Issue: Authentication fails
**Solution**: 
- Verify credentials in .env are correct
- Ensure admin consent was granted
- Check the app registration still has valid permissions

### Issue: No events retrieved
**Solution**:
- Verify audit log is enabled in your tenant
- Check that users have actually used Copilot
- Ensure Office 365 Management API subscription is active
- Note: Audit logs may have up to 24-hour delay

### Issue: Permission errors
**Solution**:
- Verify all API permissions are added to the app registration
- Ensure admin consent was granted
- Wait a few minutes after granting consent for changes to propagate

## Performance Notes

- **First run**: May take longer (up to 365 days of audit data)
- **Subsequent runs**: Only retrieves new events since last run
- **Large tenants**: Consider using --users-only and --events-only separately
- **API throttling**: Script includes basic retry logic

## Security Best Practices

1. **Protect the .env file**: Never commit it to version control
2. **Rotate secrets**: Update client secrets every 12-24 months
3. **Monitor usage**: Review app registration activity logs
4. **Least privilege**: Only grant required permissions
5. **Secure output**: Store CSV files in secure location

## Next Steps

After collecting data:
1. Open the Power BI template (.pbit file)
2. Point it to your CSV files
3. Configure the report parameters
4. Save as .pbix file

For detailed information, see [PYTHON_SETUP.md](PYTHON_SETUP.md)
