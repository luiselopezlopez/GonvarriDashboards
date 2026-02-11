# Azure AD App Registration Setup Script
# This script creates an Azure AD Service Principal (App Registration) with the necessary permissions
# for the Copilot Audit Python script and generates a .env file with the credentials

Write-Host "Microsoft 365 Copilot Audit - Azure AD App Setup" -ForegroundColor Cyan
Write-Host "=" * 60

# Check if Microsoft Graph PowerShell module is installed
$module = Get-Module -ListAvailable | Where-Object { $_.Name -eq 'Microsoft.Graph' }

if ($null -eq $module) {
    try {
        Write-Host "Installing Microsoft.Graph module..." -ForegroundColor Yellow
        Install-Module -Name Microsoft.Graph -Force -AllowClobber -Scope CurrentUser
    } 
    catch {
        Write-Host "Failed to install Microsoft.Graph module: $_" -ForegroundColor Red
        exit
    }
}

# Import required modules
Import-Module Microsoft.Graph.Applications
Import-Module Microsoft.Graph.Authentication

# Connect to Microsoft Graph with required permissions
try {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Write-Host "You will need Global Administrator or Application Administrator permissions" -ForegroundColor Yellow
    Connect-MgGraph -Scopes "Application.ReadWrite.All", "Directory.Read.All" -NoWelcome
    Write-Host "Connected successfully!" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
    exit
}

# Get tenant information
$context = Get-MgContext
$tenantId = $context.TenantId
Write-Host "Tenant ID: $tenantId" -ForegroundColor Green

# App registration details
$appName = "Copilot Audit Data Collector"
$appDescription = "Service Principal for automated collection of Microsoft 365 Copilot audit data"

Write-Host ""
Write-Host "Creating Azure AD App Registration..." -ForegroundColor Yellow

# Check if app already exists
$existingApp = Get-MgApplication -Filter "displayName eq '$appName'" -ErrorAction SilentlyContinue

if ($existingApp) {
    Write-Host "App registration '$appName' already exists." -ForegroundColor Yellow
    $choice = Read-Host "Do you want to (R)euse existing, (D)elete and recreate, or (C)ancel? [R/D/C]"
    
    switch ($choice.ToUpper()) {
        "R" {
            $app = $existingApp
            Write-Host "Reusing existing app registration..." -ForegroundColor Green
        }
        "D" {
            Write-Host "Deleting existing app registration..." -ForegroundColor Yellow
            Remove-MgApplication -ApplicationId $existingApp.Id
            Start-Sleep -Seconds 5
            $existingApp = $null
        }
        "C" {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            Disconnect-MgGraph
            exit
        }
        default {
            Write-Host "Invalid choice. Operation cancelled." -ForegroundColor Red
            Disconnect-MgGraph
            exit
        }
    }
}

if (-not $existingApp) {
    # Define required API permissions
    # Microsoft Graph API permissions
    $graphResourceId = "00000003-0000-0000-c000-000000000000" # Microsoft Graph
    
    $requiredResourceAccess = @(
        @{
            ResourceAppId = $graphResourceId
            ResourceAccess = @(
                @{
                    # User.Read.All - Application permission
                    Id = "df021288-bdef-4463-88db-98f22de89214"
                    Type = "Role"
                },
                @{
                    # Directory.Read.All - Application permission
                    Id = "7ab1d382-f21e-4acd-a863-ba3e13f7da61"
                    Type = "Role"
                },
                @{
                    # AuditLog.Read.All - Application permission
                    Id = "b0afded3-3588-46d8-8b3d-9842eff778da"
                    Type = "Role"
                }
            )
        },
        @{
            # Office 365 Management APIs
            ResourceAppId = "c5393580-f805-4401-95e8-94b7a6ef2fc2"
            ResourceAccess = @(
                @{
                    # ActivityFeed.Read - Application permission
                    Id = "594c1fb6-4f81-4475-ae41-0c394909246c"
                    Type = "Role"
                },
                @{
                    # ServiceHealth.Read - Application permission  
                    Id = "e2cea78f-e743-4d8f-a16a-75b629a038ae"
                    Type = "Role"
                }
            )
        }
    )

    # Create the app registration
    $appParams = @{
        DisplayName = $appName
        Description = $appDescription
        RequiredResourceAccess = $requiredResourceAccess
        SignInAudience = "AzureADMyOrg"
    }

    try {
        $app = New-MgApplication @appParams
        Write-Host "App registration created successfully!" -ForegroundColor Green
        Write-Host "Application (client) ID: $($app.AppId)" -ForegroundColor Green
    } catch {
        Write-Host "Failed to create app registration: $_" -ForegroundColor Red
        Disconnect-MgGraph
        exit
    }

    # Create a client secret
    Write-Host ""
    Write-Host "Creating client secret..." -ForegroundColor Yellow
    
    $passwordCredential = @{
        DisplayName = "Copilot Audit Script Secret"
        EndDateTime = (Get-Date).AddYears(2)
    }

    try {
        $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential $passwordCredential
        $clientSecret = $secret.SecretText
        Write-Host "Client secret created successfully!" -ForegroundColor Green
        Write-Host "Secret will expire on: $($secret.EndDateTime)" -ForegroundColor Yellow
    } catch {
        Write-Host "Failed to create client secret: $_" -ForegroundColor Red
        Disconnect-MgGraph
        exit
    }

    # Grant admin consent
    Write-Host ""
    Write-Host "IMPORTANT: Admin consent is required for the application permissions." -ForegroundColor Yellow
    Write-Host "Please grant admin consent in the Azure Portal:" -ForegroundColor Yellow
    Write-Host "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($app.AppId)" -ForegroundColor Cyan
    Write-Host ""
    
    $grantConsent = Read-Host "Press Enter after granting admin consent (or type 'skip' to skip this step)"
    
    if ($grantConsent.ToLower() -ne "skip") {
        Write-Host "Waiting 10 seconds for consent propagation..." -ForegroundColor Yellow
        Start-Sleep -Seconds 10
    }
}
else {
    # For existing app, we need to create a new secret
    $app = $existingApp
    
    Write-Host ""
    Write-Host "Creating new client secret for existing app..." -ForegroundColor Yellow
    
    $passwordCredential = @{
        DisplayName = "Copilot Audit Script Secret - $(Get-Date -Format 'yyyy-MM-dd')"
        EndDateTime = (Get-Date).AddYears(2)
    }

    try {
        $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential $passwordCredential
        $clientSecret = $secret.SecretText
        Write-Host "Client secret created successfully!" -ForegroundColor Green
        Write-Host "Secret will expire on: $($secret.EndDateTime)" -ForegroundColor Yellow
    } catch {
        Write-Host "Failed to create client secret: $_" -ForegroundColor Red
        Disconnect-MgGraph
        exit
    }
}

# Create .env file
Write-Host ""
Write-Host "Creating .env file..." -ForegroundColor Yellow

$envContent = @"
# Azure AD Authentication Configuration
# Generated on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

# Azure AD Tenant ID
AZURE_TENANT_ID=$tenantId

# Azure AD Application (Client) ID
AZURE_CLIENT_ID=$($app.AppId)

# Azure AD Application Client Secret
AZURE_CLIENT_SECRET=$clientSecret

# Optional: Copilot SKU IDs (comma-separated)
# Default: 639dec6b-bb19-468b-871c-c5c441c4b0cb (Microsoft 365 Copilot)
# GCC environments: a920a45e-67da-4a1a-b408-460d7a2453ce
# COPILOT_SKU_IDS=639dec6b-bb19-468b-871c-c5c441c4b0cb

# Optional: Audit log lookback period in days (default: 90)
# Used when no previous events file exists
# AUDIT_LOOKBACK_DAYS=90

# Optional: Audit interval in minutes (default: 1440 = 24 hours)
# Smaller intervals for high-volume tenants, larger for low-volume
# AUDIT_INTERVAL_MINUTES=1440
"@

try {
    $envContent | Out-File -FilePath ".env" -Encoding UTF8 -Force
    Write-Host ".env file created successfully!" -ForegroundColor Green
} catch {
    Write-Host "Failed to create .env file: $_" -ForegroundColor Red
}

# Create .env.example file (without secrets)
$envExampleContent = @"
# Azure AD Authentication Configuration Template
# Copy this file to .env and fill in your actual values

# Azure AD Tenant ID
AZURE_TENANT_ID=your-tenant-id-here

# Azure AD Application (Client) ID  
AZURE_CLIENT_ID=your-client-id-here

# Azure AD Application Client Secret
AZURE_CLIENT_SECRET=your-client-secret-here

# Optional: Copilot SKU IDs (comma-separated)
# Default: 639dec6b-bb19-468b-871c-c5c441c4b0cb (Microsoft 365 Copilot)
# GCC environments: a920a45e-67da-4a1a-b408-460d7a2453ce
# COPILOT_SKU_IDS=639dec6b-bb19-468b-871c-c5c441c4b0cb

# Optional: Audit log lookback period in days (default: 90)
# Used when no previous events file exists
# AUDIT_LOOKBACK_DAYS=90

# Optional: Audit interval in minutes (default: 1440 = 24 hours)
# Smaller intervals for high-volume tenants, larger for low-volume
# AUDIT_INTERVAL_MINUTES=1440
"@

try {
    $envExampleContent | Out-File -FilePath ".env.example" -Encoding UTF8 -Force
    Write-Host ".env.example file created successfully!" -ForegroundColor Green
} catch {
    Write-Host "Failed to create .env.example file: $_" -ForegroundColor Red
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null

# Summary
Write-Host ""
Write-Host "=" * 60
Write-Host "Setup Complete!" -ForegroundColor Green
Write-Host "=" * 60
Write-Host ""
Write-Host "Application Details:" -ForegroundColor Cyan
Write-Host "  Name: $appName"
Write-Host "  Application (client) ID: $($app.AppId)"
Write-Host "  Tenant ID: $tenantId"
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Cyan
Write-Host "  1. Verify admin consent has been granted in Azure Portal"
Write-Host "  2. Install Python dependencies: pip install -r requirements.txt"
Write-Host "  3. Run the script: python copilot_audit.py"
Write-Host ""
Write-Host "IMPORTANT SECURITY NOTES:" -ForegroundColor Yellow
Write-Host "  - The .env file contains sensitive credentials"
Write-Host "  - Never commit .env file to version control"
Write-Host "  - Keep the client secret secure"
Write-Host "  - The secret will expire on: $($secret.EndDateTime)" -ForegroundColor Yellow
Write-Host ""
