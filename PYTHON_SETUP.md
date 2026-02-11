# Microsoft 365 Copilot Audit Data Collector - Python Implementation

This Python script consolidates the functionality of the original PowerShell scripts (`Audit-Get-Users.ps1` and `Audit-Get-Events.ps1`) into a single, automated solution that can run unattended.

## Features

- **User Data Collection**: Retrieves user information from Microsoft Graph including:
  - User details (name, UPN, job title, department, location)
  - Manager information
  - Copilot license status
  
- **Event Data Collection**: Retrieves Copilot interaction events from Office 365 audit logs including:
  - Interaction timestamps
  - Application usage (Word, Excel, PowerPoint, Teams, etc.)
  - Accessed resources
  - Context and location information
  - Copilot Studio Agent usage

- **Unattended Execution**: Designed to run without user interaction, suitable for scheduled tasks
- **Comprehensive Logging**: Detailed logging to both console and file
- **Error Handling**: Robust error handling with graceful failures
- **Incremental Updates**: Events retrieval continues from the last collected event

## Prerequisites

1. **Python 3.7 or higher**
2. **Azure AD App Registration** with the following permissions:
   - Microsoft Graph API:
     - `User.Read.All` (Application)
     - `Directory.Read.All` (Application)
     - `AuditLog.Read.All` (Application)
   - Office 365 Management APIs:
     - `ActivityFeed.Read` (Application)
     - `ServiceHealth.Read` (Application)
3. **Admin Consent** granted for the above permissions
4. **Python packages** (see requirements.txt)

## Quick Start

### 1. Setup Azure AD App Registration

Run the PowerShell setup script to create the Azure AD app registration and generate the `.env` file:

```powershell
# Run PowerShell as Administrator
.\Setup-AzureAppRegistration.ps1
```

This script will:
- Create an Azure AD App Registration with required permissions
- Generate a client secret
- Create a `.env` file with your credentials
- Provide instructions for granting admin consent

**Important**: After running the script, make sure to grant admin consent in the Azure Portal using the provided URL.

### 2. Install Python Dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the Script

```bash
# Run both user and event collection
python copilot_audit.py

# Run only user collection
python copilot_audit.py --users-only

# Run only event collection
python copilot_audit.py --events-only

# Specify custom output directory
python copilot_audit.py --output-dir /path/to/output
```

## Configuration

### Environment Variables

The script requires the following environment variables in a `.env` file:

```env
AZURE_TENANT_ID=your-tenant-id-here
AZURE_CLIENT_ID=your-client-id-here
AZURE_CLIENT_SECRET=your-client-secret-here
```

These are automatically created by the `Setup-AzureAppRegistration.ps1` script.

### Copilot SKU IDs

The script includes the default Copilot SKU ID for commercial tenants:
- `639dec6b-bb19-468b-871c-c5c441c4b0cb` (Microsoft 365 Copilot)

For GCC environments, modify the script or add to your `.env`:
```env
COPILOT_SKU_IDS=a920a45e-67da-4a1a-b408-460d7a2453ce
```

## Output Files

The script generates the following files in the output directory (default: `./output`):

1. **Copilot_Users.csv**: User information including:
   - EntraID, DisplayName, UserPrincipalName
   - JobTitle, Department, City, Country, UsageLocation
   - ManagerName, ManagerUPN
   - HasCopilotLicense

2. **Copilot_Events.csv**: Copilot interaction events including:
   - TimeStamp, User, App, Location
   - App context, Accessed Resources
   - Accessed Resource Locations, Action
   - AgentName (for Copilot Studio)

3. **copilot_audit.log**: Detailed execution log
4. **AuditScriptLog.txt**: Audit-specific log messages

## Usage Examples

### Scheduled Execution (Windows Task Scheduler)

Create a scheduled task to run the script automatically:

1. Open Task Scheduler
2. Create a new task with the following settings:
   - Trigger: Daily (or as needed)
   - Action: Start a program
     - Program: `python.exe`
     - Arguments: `C:\path\to\copilot_audit.py`
     - Start in: `C:\path\to\`

### Scheduled Execution (Linux/macOS cron)

Add to crontab:

```bash
# Run daily at 2 AM
0 2 * * * /usr/bin/python3 /path/to/copilot_audit.py >> /path/to/copilot_audit_cron.log 2>&1
```

### Docker Container

Create a `Dockerfile`:

```dockerfile
FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt

COPY copilot_audit.py .
COPY .env .

CMD ["python", "copilot_audit.py"]
```

Build and run:

```bash
docker build -t copilot-audit .
docker run -v $(pwd)/output:/app/output copilot-audit
```

## Comparison with PowerShell Scripts

| Feature | PowerShell Scripts | Python Script |
|---------|-------------------|---------------|
| User Authentication | Interactive | Unattended (Service Principal) |
| Scheduling | Manual/Task Scheduler | Fully automated |
| Dependencies | ExchangeOnlineManagement, Microsoft.Graph | msal, requests, python-dotenv |
| Event Retrieval | Search-UnifiedAuditLog | Office 365 Management API |
| Logging | Basic | Comprehensive (file + console) |
| Error Handling | Basic | Robust with retries |
| Cross-platform | Windows only | Windows, Linux, macOS |

## Troubleshooting

### Authentication Errors

- Verify the `.env` file contains correct credentials
- Ensure admin consent has been granted
- Check that the app registration has the required permissions

### No Events Retrieved

- Verify the Office 365 audit log subscription is active
- Check that users have generated Copilot interactions
- Ensure the service principal has `ActivityFeed.Read` permission
- Note: Audit logs may have a delay of up to 24 hours

### API Rate Limiting

If you encounter rate limiting errors:
- The script includes basic retry logic
- For large tenants, consider running user and event collection separately
- Implement exponential backoff for production use

### Missing Manager Information

- Ensure managers are populated in Entra ID (Azure AD)
- Verify the service principal has `Directory.Read.All` permission

## Security Considerations

1. **Protect Credentials**: Never commit the `.env` file to version control
2. **Secret Rotation**: Rotate client secrets regularly (maximum 2 years)
3. **Least Privilege**: Only grant the minimum required permissions
4. **Audit Logs**: Monitor the service principal's activity
5. **Secure Storage**: Store output files in a secure location

## Migration from PowerShell Scripts

To migrate from the PowerShell scripts:

1. **Back up existing CSV files**: Archive `Copilot_Users.csv` and `Copilot_Events.csv`
2. **Run setup script**: Execute `Setup-AzureAppRegistration.ps1`
3. **Grant admin consent**: Complete the admin consent in Azure Portal
4. **Install Python**: Ensure Python 3.7+ is installed
5. **Install dependencies**: Run `pip install -r requirements.txt`
6. **Test execution**: Run `python copilot_audit.py` to verify
7. **Schedule automation**: Set up scheduled task or cron job

The Python script will automatically detect existing event files and continue from the last event timestamp.

## Support and Contributions

For issues or questions:
- Check the logs in `copilot_audit.log`
- Review Azure AD app permissions
- Verify network connectivity to Microsoft APIs

## License

This script is provided as-is for use with Microsoft 365 Copilot audit reporting.

## Acknowledgments

- Original PowerShell scripts concept and audit log parsing logic
- Microsoft Graph and Office 365 Management API documentation
- Special thanks to the Office365itpros team (https://github.com/12Knocksinna/Office365itpros)
