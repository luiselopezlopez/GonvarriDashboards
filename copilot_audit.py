#!/usr/bin/env python3
"""
Microsoft 365 Copilot Audit Data Retrieval Script

This script consolidates the functionality of two PowerShell scripts:
- Audit-Get-Users.ps1: Retrieves user information from Microsoft Graph
- Audit-Get-Events.ps1: Retrieves Copilot interaction events from audit logs

The script can run unattended and includes appropriate logging and error handling.

Requirements:
- Python 3.7+
- See requirements.txt for required packages
- Azure AD App Registration with appropriate permissions
- Environment variables configured in .env file

Usage:
    python copilot_audit.py [--users-only] [--events-only] [--output-dir OUTPUT_DIR]
"""

import os
import sys
import csv
import logging
import argparse
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict, Optional
import json

# Third-party imports
try:
    from msal import ConfidentialClientApplication
    import requests
    from dotenv import load_dotenv
except ImportError as e:
    print(f"Missing required package: {e}")
    print("Please install requirements: pip install -r requirements.txt")
    sys.exit(1)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('copilot_audit.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class CopilotAuditClient:
    """Client for retrieving Microsoft 365 Copilot audit data"""
    
    def __init__(self, tenant_id: str, client_id: str, client_secret: str, output_dir: str = "./output"):
        """
        Initialize the Copilot Audit Client
        
        Args:
            tenant_id: Azure AD tenant ID
            client_id: Azure AD application (client) ID
            client_secret: Azure AD application client secret
            output_dir: Directory for output files
        """
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        
        # File paths
        self.users_csv_path = self.output_dir / "Copilot_Users.csv"
        self.events_csv_path = self.output_dir / "Copilot_Events.csv"
        self.log_file_path = self.output_dir / "AuditScriptLog.txt"
        
        # Copilot SKU IDs - load from environment or use default (commercial)
        sku_ids_env = os.getenv("COPILOT_SKU_IDS", "")
        if sku_ids_env:
            self.copilot_sku_ids = [sid.strip() for sid in sku_ids_env.split(",") if sid.strip()]
        else:
            self.copilot_sku_ids = ["639dec6b-bb19-468b-871c-c5c441c4b0cb"]
        
        # Initialize MSAL app
        self.app = ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}"
        )
        
        self.graph_token = None
        self.management_token = None
    
    def _get_graph_token(self) -> Optional[str]:
        """Get access token for Microsoft Graph API"""
        if self.graph_token:
            return self.graph_token
            
        try:
            result = self.app.acquire_token_for_client(
                scopes=["https://graph.microsoft.com/.default"]
            )
            
            if "access_token" in result:
                self.graph_token = result["access_token"]
                logger.info("Successfully acquired Graph API token")
                return self.graph_token
            else:
                logger.error(f"Failed to acquire Graph token: {result.get('error_description', 'Unknown error')}")
                return None
        except Exception as e:
            logger.error(f"Error acquiring Graph token: {e}")
            return None
    
    def _get_management_token(self) -> Optional[str]:
        """Get access token for Office 365 Management API"""
        if self.management_token:
            return self.management_token
            
        try:
            result = self.app.acquire_token_for_client(
                scopes=["https://manage.office.com/.default"]
            )
            
            if "access_token" in result:
                self.management_token = result["access_token"]
                logger.info("Successfully acquired Management API token")
                return self.management_token
            else:
                logger.error(f"Failed to acquire Management token: {result.get('error_description', 'Unknown error')}")
                return None
        except Exception as e:
            logger.error(f"Error acquiring Management token: {e}")
            return None
    
    def _make_graph_request(self, endpoint: str, params: Optional[Dict] = None) -> Optional[Dict]:
        """
        Make a request to Microsoft Graph API
        
        Args:
            endpoint: API endpoint (e.g., '/users')
            params: Query parameters
            
        Returns:
            JSON response or None on error
        """
        token = self._get_graph_token()
        if not token:
            return None
        
        url = f"https://graph.microsoft.com/v1.0{endpoint}"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        try:
            response = requests.get(url, headers=headers, params=params)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            logger.error(f"Graph API request failed for {endpoint}: {e}")
            return None
    
    def _get_all_pages(self, endpoint: str, params: Optional[Dict] = None) -> List[Dict]:
        """
        Get all pages from a paginated Microsoft Graph endpoint
        
        Args:
            endpoint: API endpoint
            params: Query parameters
            
        Returns:
            List of all items from all pages
        """
        all_items = []
        next_link = None
        
        while True:
            if next_link:
                # Use the next link directly
                token = self._get_graph_token()
                if not token:
                    break
                    
                headers = {
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json"
                }
                
                try:
                    response = requests.get(next_link, headers=headers)
                    response.raise_for_status()
                    data = response.json()
                except requests.exceptions.RequestException as e:
                    logger.error(f"Error fetching next page: {e}")
                    break
            else:
                data = self._make_graph_request(endpoint, params)
                if not data:
                    break
            
            items = data.get("value", [])
            all_items.extend(items)
            
            next_link = data.get("@odata.nextLink")
            if not next_link:
                break
        
        return all_items
    
    def get_users_with_copilot_license(self) -> bool:
        """
        Retrieve users with job titles and check for Copilot licenses
        
        Returns:
            True if successful, False otherwise
        """
        logger.info("Starting user data retrieval...")
        
        try:
            # Get all users with job titles
            logger.info("Fetching users from Microsoft Graph...")
            params = {
                "$filter": "jobTitle ne null",
                "$select": "id,displayName,userPrincipalName,jobTitle,department,city,country,usageLocation,assignedLicenses",
                "$top": 999
            }
            
            users = self._get_all_pages("/users", params)
            logger.info(f"Retrieved {len(users)} users with job titles")
            
            # Process each user
            results = []
            for i, user in enumerate(users):
                if (i + 1) % 50 == 0:
                    logger.info(f"Processing user {i + 1} of {len(users)}...")
                
                user_data = {
                    "EntraID": user.get("id", ""),
                    "DisplayName": user.get("displayName", ""),
                    "UserPrincipalName": user.get("userPrincipalName", ""),
                    "JobTitle": user.get("jobTitle", ""),
                    "Department": user.get("department", ""),
                    "City": user.get("city", ""),
                    "Country": user.get("country", ""),
                    "UsageLocation": user.get("usageLocation", ""),
                    "ManagerName": "",
                    "ManagerUPN": "",
                    "HasCopilotLicense": False
                }
                
                # Get manager information
                try:
                    manager_data = self._make_graph_request(f"/users/{user['id']}/manager")
                    if manager_data:
                        user_data["ManagerName"] = manager_data.get("displayName", "")
                        user_data["ManagerUPN"] = manager_data.get("userPrincipalName", "")
                except Exception as e:
                    logger.debug(f"Could not retrieve manager for user {user.get('userPrincipalName', '')}: {e}")
                
                # Check for Copilot license
                assigned_licenses = user.get("assignedLicenses", [])
                for license in assigned_licenses:
                    sku_id = str(license.get("skuId", "")).lower()
                    if sku_id in [sid.lower() for sid in self.copilot_sku_ids]:
                        user_data["HasCopilotLicense"] = True
                        break
                
                results.append(user_data)
            
            # Write to CSV
            logger.info(f"Writing {len(results)} users to {self.users_csv_path}")
            with open(self.users_csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                if results:
                    fieldnames = results[0].keys()
                    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                    writer.writeheader()
                    writer.writerows(results)
            
            logger.info(f"Successfully exported users to {self.users_csv_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error retrieving users: {e}", exc_info=True)
            return False
    
    def _write_log_file(self, message: str):
        """Write a timestamped message to the audit log file"""
        timestamp = datetime.utcnow().isoformat()
        with open(self.log_file_path, 'a') as f:
            f.write(f"{timestamp}:{message}\n")
    
    def _get_last_event_timestamp(self) -> datetime:
        """
        Get the timestamp of the last event from the CSV file
        
        Returns:
            Last event timestamp or configured days ago if file doesn't exist
        """
        if not self.events_csv_path.exists():
            # Default lookback period (configurable via environment)
            lookback_days = int(os.getenv("AUDIT_LOOKBACK_DAYS", "90"))
            return datetime.utcnow() - timedelta(days=lookback_days)
        
        try:
            with open(self.events_csv_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                lines = list(reader)
                if len(lines) > 1:  # More than just header
                    last_row = lines[-1]
                    # Extract timestamp from first field (format: dd-MMM-yyyy HH:mm:ss)
                    timestamp_str = last_row[0]
                    return datetime.strptime(timestamp_str, '%d-%b-%Y %H:%M:%S')
        except Exception as e:
            logger.warning(f"Could not parse last event timestamp: {e}")
        
        # Default lookback period (configurable via environment)
        lookback_days = int(os.getenv("AUDIT_LOOKBACK_DAYS", "90"))
        return datetime.utcnow() - timedelta(days=lookback_days)
    
    def _parse_copilot_event(self, record: Dict) -> Dict:
        """
        Parse a Copilot interaction audit record
        
        Args:
            record: Raw audit record
            
        Returns:
            Parsed event data
        """
        try:
            audit_data = record.get("AuditData", {})
            if isinstance(audit_data, str):
                audit_data = json.loads(audit_data)
            
            copilot_event_data = audit_data.get("CopilotEventData", {})
            contexts = copilot_event_data.get("Contexts", [])
            context_type = contexts[0].get("Type", "") if contexts else ""
            context_id = contexts[0].get("Id", "") if contexts else ""
            
            # Determine Copilot app
            copilot_app = "Copilot for M365"
            copilot_location = None
            
            # Map context type to app
            app_mapping = {
                "xlsx": "Excel",
                "docx": "Word",
                "pptx": "PowerPoint",
                "TeamsMeeting": "Teams",
                "whiteboard": "Whiteboard",
                "loop": "Loop",
                "StreamVideo": "Stream"
            }
            
            copilot_app = app_mapping.get(context_type, copilot_app)
            
            if context_type == "TeamsMeeting":
                copilot_location = "Teams meeting"
            elif context_type == "StreamVideo":
                copilot_location = "Stream video player"
            
            # Additional app host checks
            app_host = copilot_event_data.get("AppHost", "")
            # Note: context_id is from Microsoft audit logs, not user input
            teams_url_pattern = "https://teams.microsoft.com/"
            if context_id and context_id.startswith(teams_url_pattern):
                copilot_app = "Teams"
            elif app_host == "bizchat":
                copilot_app = "Copilot for M365 Chat"
            elif app_host == "Outlook":
                copilot_app = "Outlook"
            elif app_host == "Copilot Studio":
                copilot_app = "Copilot Studio Agent"
            
            # Determine context
            context = context_id or copilot_event_data.get("ThreadId", "")
            
            # Extract agent name for Copilot Studio
            agent_name = ""
            if copilot_app == "Copilot Studio Agent":
                app_identity = audit_data.get("AppIdentity", "")
                # Try to extract agent name from AppIdentity
                if "_" in app_identity:
                    agent_name = app_identity.split("_")[-1]
                elif "-" in app_identity:
                    agent_name = app_identity.split("-")[-1]
            
            # Determine location
            # Note: context_id is from Microsoft audit logs, not user input
            if "/sites/" in context_id:
                copilot_location = "SharePoint Online"
            elif context_id and context_id.startswith("https://teams.microsoft.com/"):
                if "ctx=channel" in context_id:
                    copilot_location = "Teams Channel"
                else:
                    copilot_location = "Teams Chat"
            elif "/personal/" in context_id:
                copilot_location = "OneDrive for Business"
            
            # Extract accessed resources
            accessed_resources = copilot_event_data.get("AccessedResources", [])
            resource_names = sorted(set([r.get("Name", "") for r in accessed_resources if r.get("Name")]))
            resource_ids = sorted(set([r.get("Id", "") for r in accessed_resources if r.get("Id")]))
            resource_actions = sorted(set([r.get("Action", "") for r in accessed_resources if r.get("Action")]))
            
            # Format timestamp
            creation_date = record.get("CreationTime", record.get("CreationDate", ""))
            if creation_date:
                dt = datetime.fromisoformat(creation_date.replace("Z", "+00:00"))
                timestamp = dt.strftime("%d-%b-%Y %H:%M:%S")
            else:
                timestamp = ""
            
            return {
                "TimeStamp": timestamp,
                "User": record.get("UserId", ""),
                "App": copilot_app,
                "Location": copilot_location or "",
                "App context": context,
                "Accessed Resources": ", ".join(resource_names),
                "Accessed Resource Locations": ", ".join(resource_ids),
                "Action": ", ".join(resource_actions),
                "AgentName": agent_name
            }
            
        except Exception as e:
            logger.error(f"Error parsing event: {e}")
            return None
    
    def get_copilot_events(self) -> bool:
        """
        Retrieve Copilot interaction events from audit logs using Office 365 Management API
        
        Returns:
            True if successful, False otherwise
        """
        logger.info("Starting Copilot events retrieval...")
        
        try:
            # Determine start date
            start_date = self._get_last_event_timestamp()
            end_date = datetime.utcnow()
            
            logger.info(f"Retrieving audit records between {start_date} and {end_date}")
            self._write_log_file(f"BEGIN: Retrieving audit records between {start_date} and {end_date}")
            
            token = self._get_management_token()
            if not token:
                logger.error("Failed to acquire Management API token")
                return False
            
            # Office 365 Management Activity API
            base_url = f"https://manage.office.com/api/v1.0/{self.tenant_id}/activity/feed"
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            }
            
            # Start subscription (if not already started)
            try:
                subscription_url = f"{base_url}/subscriptions/start?contentType=Audit.General&PublisherIdentifier={self.tenant_id}"
                response = requests.post(subscription_url, headers=headers)
                if response.status_code in [200, 400]:  # 400 if already subscribed
                    logger.info("Audit subscription active")
            except Exception as e:
                logger.warning(f"Could not start subscription: {e}")
            
            # List available content
            # Note: The Management API works differently than Search-UnifiedAuditLog
            # We need to list content blobs and download them
            total_count = 0
            # Configurable interval minutes (default 24 hours)
            interval_minutes = int(os.getenv("AUDIT_INTERVAL_MINUTES", "1440"))
            current_start = start_date
            
            results = []
            
            while current_start < end_date:
                current_end = min(current_start + timedelta(minutes=interval_minutes), end_date)
                
                logger.info(f"Retrieving audit records for activities between {current_start} and {current_end}")
                self._write_log_file(f"INFO: Retrieving audit records for activities between {current_start} and {current_end}")
                
                # List available content
                start_time_str = current_start.strftime("%Y-%m-%dT%H:%M:%S")
                end_time_str = current_end.strftime("%Y-%m-%dT%H:%M:%S")
                
                list_url = f"{base_url}/subscriptions/content?contentType=Audit.General&startTime={start_time_str}&endTime={end_time_str}&PublisherIdentifier={self.tenant_id}"
                
                try:
                    response = requests.get(list_url, headers=headers)
                    
                    if response.status_code == 200:
                        content_blobs = response.json()
                        logger.info(f"Found {len(content_blobs)} content blobs")
                        
                        # Download each content blob
                        for blob in content_blobs:
                            content_uri = blob.get("contentUri")
                            if not content_uri:
                                continue
                            
                            try:
                                content_response = requests.get(content_uri, headers=headers)
                                if content_response.status_code == 200:
                                    records = content_response.json()
                                    
                                    # Filter for CopilotInteraction records
                                    copilot_records = [r for r in records if r.get("RecordType") == 91]  # CopilotInteraction = 91
                                    
                                    logger.info(f"Found {len(copilot_records)} Copilot records in blob")
                                    
                                    for record in copilot_records:
                                        parsed_event = self._parse_copilot_event(record)
                                        if parsed_event:
                                            results.append(parsed_event)
                                            total_count += 1
                                    
                            except Exception as e:
                                logger.error(f"Error downloading content blob: {e}")
                    
                    elif response.status_code == 204:
                        logger.info("No content available for this time range")
                    else:
                        logger.warning(f"List content request returned status {response.status_code}")
                        
                except Exception as e:
                    logger.error(f"Error listing content: {e}")
                
                current_start = current_end
            
            # Write results to CSV
            if results:
                logger.info(f"Writing {len(results)} events to {self.events_csv_path}")
                
                # Check if file exists to determine if we should append
                file_exists = self.events_csv_path.exists()
                mode = 'a' if file_exists else 'w'
                
                with open(self.events_csv_path, mode, newline='', encoding='utf-8') as csvfile:
                    fieldnames = ["TimeStamp", "User", "App", "Location", "App context", 
                                  "Accessed Resources", "Accessed Resource Locations", 
                                  "Action", "AgentName"]
                    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                    
                    if not file_exists:
                        writer.writeheader()
                    
                    writer.writerows(results)
            
            self._write_log_file(f"END: Retrieved {total_count} audit records")
            logger.info(f"Successfully retrieved {total_count} Copilot events")
            
            return True
            
        except Exception as e:
            logger.error(f"Error retrieving events: {e}", exc_info=True)
            return False


def main():
    """Main entry point for the script"""
    parser = argparse.ArgumentParser(
        description="Microsoft 365 Copilot Audit Data Retrieval Script"
    )
    parser.add_argument(
        "--users-only",
        action="store_true",
        help="Only retrieve user data"
    )
    parser.add_argument(
        "--events-only",
        action="store_true",
        help="Only retrieve event data"
    )
    parser.add_argument(
        "--output-dir",
        default="./output",
        help="Output directory for CSV files (default: ./output)"
    )
    
    args = parser.parse_args()
    
    # Load environment variables
    load_dotenv()
    
    # Get credentials from environment
    tenant_id = os.getenv("AZURE_TENANT_ID")
    client_id = os.getenv("AZURE_CLIENT_ID")
    client_secret = os.getenv("AZURE_CLIENT_SECRET")
    
    if not all([tenant_id, client_id, client_secret]):
        logger.error("Missing required environment variables. Please ensure AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET are set in .env file")
        sys.exit(1)
    
    # Initialize client
    client = CopilotAuditClient(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret,
        output_dir=args.output_dir
    )
    
    success = True
    
    # Retrieve users (unless events-only is specified)
    if not args.events_only:
        logger.info("=" * 60)
        logger.info("Retrieving user data...")
        logger.info("=" * 60)
        if not client.get_users_with_copilot_license():
            logger.error("Failed to retrieve user data")
            success = False
    
    # Retrieve events (unless users-only is specified)
    if not args.users_only:
        logger.info("=" * 60)
        logger.info("Retrieving Copilot events...")
        logger.info("=" * 60)
        if not client.get_copilot_events():
            logger.error("Failed to retrieve event data")
            success = False
    
    if success:
        logger.info("=" * 60)
        logger.info("Script completed successfully!")
        logger.info("=" * 60)
        sys.exit(0)
    else:
        logger.error("=" * 60)
        logger.error("Script completed with errors")
        logger.error("=" * 60)
        sys.exit(1)


if __name__ == "__main__":
    main()
