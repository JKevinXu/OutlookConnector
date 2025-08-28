# Deploy Office Add-in to M365 Beta Environment
# Run this script with Global Admin privileges

# Install required modules
Write-Host "Installing O365 Centralized Deployment module..." -ForegroundColor Green
Install-Module -Name O365CentralizedAddInDeployment -Force -AllowClobber

# Import the module
Import-Module -Name O365CentralizedAddInDeployment

# Connect to M365
Write-Host "Connecting to Microsoft 365..." -ForegroundColor Green
Connect-OrganizationAddInService

# Deploy the add-in from manifest
Write-Host "Deploying AM Personal Assistant add-in..." -ForegroundColor Green

# Option 1: Deploy from URL (GitHub Pages - Production)
New-OrganizationAddIn -ManifestPath 'https://jkevinxu.github.io/OutlookConnector/manifest.json' -Locale 'en-US'

# Option 2: Deploy from local file (if testing locally built version)
# New-OrganizationAddIn -ManifestPath 'dist/manifest.json' -Locale 'en-US'

# Option 3: Deploy to specific beta users
# New-OrganizationAddIn -ManifestPath 'https://jkevinxu.github.io/OutlookConnector/manifest.json' -Locale 'en-US' -Members 'beta-user1@domain.com', 'beta-user2@domain.com'

Write-Host "Deployment initiated. Check Microsoft 365 Admin Center for status." -ForegroundColor Green
Write-Host "Add-in will be available to users within 24 hours." -ForegroundColor Yellow

# Get deployment details
Write-Host "Current add-ins:" -ForegroundColor Green
Get-OrganizationAddIn 