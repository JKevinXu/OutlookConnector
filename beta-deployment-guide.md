# 🧪 M365 Beta Environment Deployment Guide

## Overview
This guide explains how to deploy the AM Personal Assistant Office Add-in to your M365 beta environment for testing and validation.

## 🚀 Quick Start (Recommended)

### Method 1: Admin Center Deployment

1. **Build the Add-in**:
   ```bash
   npm run build
   ```

2. **Access Microsoft 365 Admin Center**:
   - Go to [admin.microsoft.com](https://admin.microsoft.com)
   - Sign in with Global Admin credentials

3. **Deploy via Integrated Apps**:
   - Navigate to **Settings** → **Integrated apps**
   - Click **Upload custom apps**
   - Select **App package file**
   - Upload `dist/manifest.json` (production) or `manifest.beta.json` (beta-specific)

4. **Configure Beta Deployment**:
   - **Who has access**: Select "Specific users/groups"
   - Add your beta testing group or individual users
   - **App settings**: Choose "Optional, enabled by default"
   - Click **Deploy**

## 🛠️ Advanced Deployment Methods

### Method 2: PowerShell Deployment

1. **Install PowerShell Module**:
   ```powershell
   Install-Module -Name O365CentralizedAddInDeployment -Force
   ```

2. **Run Deployment Script**:
   ```powershell
   # Use the provided deploy-to-m365-beta.ps1 script
   ./deploy-to-m365-beta.ps1
   ```

3. **Verify Deployment**:
   ```powershell
   Get-OrganizationAddIn
   ```

### Method 3: Exchange Admin Center

1. **Access Exchange Admin**:
   - Go to [admin.exchange.microsoft.com](https://admin.exchange.microsoft.com)
   - Navigate to **Organization** → **Add-ins**

2. **Add Custom Add-in**:
   - Click **+ Add add-in** → **Add from URL**
   - Enter: `https://jkevinxu.github.io/OutlookConnector/manifest.json`
   - Select installation options for beta users

## 🎯 Beta Environment Considerations

### User Targeting for Beta Testing

1. **Create Beta User Group**:
   - In M365 Admin Center → **Groups** → **Active groups**
   - Create new security group: "Office Add-in Beta Testers"
   - Add beta users to this group

2. **Deployment Configuration**:
   ```json
   {
     "targetAudience": "SpecificGroups",
     "groups": ["Office Add-in Beta Testers"],
     "deploymentType": "Optional",
     "defaultEnabled": true
   }
   ```

### Beta-Specific Features

1. **Beta API Endpoints**:
   - Your add-in already uses: `https://bwzo9wnhy3.execute-api.us-west-2.amazonaws.com/beta`
   - This is configured in webpack.config.js proxy settings

2. **Beta Manifest Differences**:
   - Version: `1.0.0-beta`
   - Name includes "(Beta)" identifier
   - Additional beta API domain in validDomains

## 📋 Pre-Deployment Checklist

- [ ] **Requirements Met**:
  - [ ] Users have M365 Enterprise licenses
  - [ ] Exchange Online mailboxes active
  - [ ] Modern authentication enabled
  - [ ] Beta channel access configured

- [ ] **Manifest Validation**:
  ```bash
  npm run validate  # Validates manifest.json
  ```

- [ ] **Build Verification**:
  - [ ] Production build successful
  - [ ] All assets accessible via GitHub Pages
  - [ ] Manifest URLs point to correct endpoints

- [ ] **Beta Environment Setup**:
  - [ ] Beta user group created
  - [ ] Test users identified
  - [ ] Rollback plan prepared

## 🔧 Deployment Commands

### Validate Manifest
```bash
npm run validate
```

### Build for Production
```bash
npm run build
```

### Deploy to Beta Users (PowerShell)
```powershell
# Connect to M365
Connect-OrganizationAddInService

# Deploy to specific beta users
New-OrganizationAddIn -ManifestPath 'https://jkevinxu.github.io/OutlookConnector/manifest.json' -Locale 'en-US' -Members 'beta-user1@domain.com', 'beta-user2@domain.com'

# Or deploy beta-specific manifest
New-OrganizationAddIn -ManifestPath 'manifest.beta.json' -Locale 'en-US' -Members 'beta-testing-group@domain.com'
```

### Monitor Deployment
```powershell
# Check deployment status
Get-OrganizationAddIn

# Get specific add-in details
Get-OrganizationAddIn -ProductId <product-id>
```

## 🧪 Testing & Validation

### Post-Deployment Testing

1. **Verify Add-in Appearance**:
   - Beta users open Outlook (web or desktop)
   - Check for "AM Personal Assistant (Beta)" in ribbon
   - Verify add-in loads correctly

2. **Functional Testing**:
   - Test email analysis features
   - Verify API connectivity to beta endpoints
   - Check authentication flow
   - Validate seller metrics functionality

3. **User Feedback Collection**:
   - Set up feedback mechanism
   - Monitor usage analytics
   - Track performance metrics

### Troubleshooting

1. **Add-in Not Appearing**:
   - Wait 24 hours for propagation
   - Clear Office cache: File → Account → Office Updates → Update Options → Clear Cache
   - Verify user is in beta group

2. **Authentication Issues**:
   - Check validDomains in manifest
   - Verify redirect URIs
   - Clear browser cache

3. **API Connection Problems**:
   - Verify beta endpoint accessibility
   - Check CORS configuration
   - Review proxy settings in webpack.config.js

## 📊 Monitoring & Analytics

### Key Metrics to Track

1. **Deployment Metrics**:
   - Number of successful installations
   - Time to propagate to users
   - Installation failure rate

2. **Usage Metrics**:
   - Daily active users
   - Feature usage patterns
   - Error rates and types

3. **Performance Metrics**:
   - Add-in load times
   - API response times
   - User satisfaction scores

### Monitoring Commands

```powershell
# Get usage statistics
Get-OrganizationAddIn | Format-Table -Property DisplayName, EnabledState, AssignedUsers

# Check specific user assignments
Get-OrganizationAddIn -ProductId <id> | Select-Object -ExpandProperty Assignments
```

## 🔄 Update & Rollback Procedures

### Updating Beta Deployment

1. **Update Manifest**:
   ```powershell
   Set-OrganizationAddIn -ProductId <id> -ManifestPath 'new-manifest.json' -Locale 'en-US'
   ```

2. **Add/Remove Beta Users**:
   ```powershell
   # Add users
   Set-OrganizationAddInAssignments -ProductId <id> -Add -Members 'new-beta-user@domain.com'
   
   # Remove users
   Set-OrganizationAddInAssignments -ProductId <id> -Remove -Members 'user@domain.com'
   ```

### Rollback Plan

1. **Disable Add-in**:
   ```powershell
   Set-OrganizationAddIn -ProductId <id> -Enabled $false
   ```

2. **Remove Deployment**:
   ```powershell
   Remove-OrganizationAddIn -ProductId <id>
   ```

## 🎯 Production Promotion

Once beta testing is complete:

1. **Update Manifest Version**:
   - Change version from "1.0.0-beta" to "1.0.0"
   - Remove "(Beta)" from names and descriptions

2. **Expand Deployment**:
   ```powershell
   Set-OrganizationAddInAssignments -ProductId <id> -AssignToEveryone $true
   ```

3. **Monitor Production Deployment**:
   - Track adoption rates
   - Monitor support tickets
   - Gather user feedback

## 📞 Support & Resources

- **Microsoft Documentation**: [Office Add-ins Deployment](https://docs.microsoft.com/en-us/office/dev/add-ins/publish/centralized-deployment)
- **PowerShell Reference**: [Centralized Deployment Cmdlets](https://docs.microsoft.com/en-us/microsoft-365/enterprise/use-the-centralized-deployment-powershell-cmdlets-to-manage-add-ins)
- **Troubleshooting Guide**: [Office Add-ins Troubleshooting](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest) 