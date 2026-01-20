# PIMMeUpScotty

A PowerShell script for automating daily Azure Privileged Identity Management (PIM) role activations.

## Overview

PIMMeUpScotty simplifies the daily workflow of activating your eligible Azure AD (Entra ID) and Azure Resource PIM roles. Instead of manually activating each role through the Azure Portal, this script handles all configured roles with a single command.

## Features

- **Automated Role Activation**: Activate multiple Azure AD and Azure Resource roles in one go
- **Smart State Management**: Tracks previously activated roles to avoid unnecessary re-activations
- **Flexible Configuration**: JSON-based configuration for easy customization
- **Diagnostic Tools**: Built-in scanning to discover all eligible roles across your Azure environment
- **Optimal Duration Learning**: Automatically learns and applies optimal activation durations based on policy constraints
- **Failure Resilience**: Intelligent retry logic with automatic duration fallback
- **Activity Logging**: Maintains detailed logs of all activation attempts

## Prerequisites

### Required Permissions

- Azure AD roles that are PIM-eligible
- Azure Resource roles that are PIM-eligible (Owner, Contributor, etc.)

### Required PowerShell Modules

The following PowerShell modules are required:

- `Az.Accounts`
- `Az.Resources`
- `Microsoft.Graph.Authentication`
- `Microsoft.Graph.Users`
- `Microsoft.Graph.Identity.Governance`

**The setup wizard automatically checks and installs any missing modules.**

## Installation

1. Download or clone the script to your local machine
2. Open PowerShell (no need to run as Administrator)
3. Navigate to the script directory
4. Run the setup wizard:

```powershell
.\PIMMeUpScotty.ps1 -Setup
```

The setup wizard will:
- Check for required PowerShell modules and install any that are missing
- Display version information for modules already installed
- Prompt you to configure Tenant ID, Subscription ID, and other settings
- Create the `pim-config.json` configuration file

**Alternatively**, just run the script without parameters on first use:

```powershell
.\PIMMeUpScotty.ps1
```

If no configuration exists, the setup wizard will run automatically.

## Configuration

### Setup Wizard

The setup wizard (`-Setup` parameter or automatic first run) will prompt you for:

- **Tenant ID**: Your Azure AD tenant identifier
- **Subscription ID**: Your default Azure subscription
- **Justification**: Default text shown when activating roles (e.g., "Daily operational work")
- **Duration**: Maximum activation duration in hours (default: 8)

These values are saved to `pim-config.json`.

### Configuration File Structure

The script creates a `pim-config.json` file with the following structure:

```json
{
  "Scopes": [
    {
      "TenantId": "your-tenant-id",
      "TenantName": "Your Company",
      "SubscriptionId": "your-subscription-id",
      "SubscriptionName": "Production",
      "ScopeId": null,
      "LandingZoneName": null,
      "ActivationMode": "AllEligible",
      "MaxDurationHours": 8,
      "Roles": null,
      "AADRoles": []
    }
  ],
  "DefaultJustification": "Daily operational work"
}
```

### Advanced Configuration

#### Activate Specific Azure AD Roles

Add role display names to the `AADRoles` array:

```json
"AADRoles": [
  "Global Administrator",
  "User Administrator",
  "Security Administrator"
]
```

Leave empty (`[]`) to activate **all** eligible Azure AD roles.

#### Configure Multiple Scopes

Add multiple scope objects to activate roles at different levels:

```json
{
  "Scopes": [
    {
      "TenantId": "your-tenant-id",
      "SubscriptionId": "subscription-1",
      "MaxDurationHours": 8,
      "AADRoles": []
    },
    {
      "TenantId": "your-tenant-id",
      "ScopeId": "/subscriptions/subscription-2",
      "MaxDurationHours": 4,
      "Roles": ["Owner", "Contributor"]
    }
  ],
  "DefaultJustification": "Daily operational work"
}
```

## Usage

### Run Setup

Run the setup wizard to install modules and configure the script:

```powershell
.\PIMMeUpScotty.ps1 -Setup
```

Setup output:
```
╔════════════════════════════════════════════╗
║        PIMMeUpScotty Setup Wizard          ║
╚════════════════════════════════════════════╝

Step 1: Checking Required PowerShell Modules
═══════════════════════════════════════════════

  Checking: Az.Accounts... ✓ Found (v3.0.0)
  Checking: Az.Resources... ✓ Found (v7.0.0)
  Checking: Microsoft.Graph.Authentication... ✗ Not found
    Installing Microsoft.Graph.Authentication...
    ✓ Installed successfully (v2.10.0)
  ...

Module Check Summary:
  Already present: 4
  Newly installed: 1

Step 2: Configuration Setup
═══════════════════════════════════════════════

Required Information:
...
```

### Basic Usage

Activate all configured roles:

```powershell
.\PIMMeUpScotty.ps1
```

### List Eligible Roles

View eligible roles without activating them:

```powershell
.\PIMMeUpScotty.ps1 -ListOnly
```

### Scan All Resources

Discover all eligible roles across your entire Azure environment:

```powershell
.\PIMMeUpScotty.ps1 -ScanAll
```

This diagnostic mode scans:
- All subscriptions
- All resource groups
- All individual resources
- Management groups (if accessible)

### Check Specific Resource

Scan a specific Azure resource for eligible roles:

```powershell
.\PIMMeUpScotty.ps1 -ResourceScope "/subscriptions/{sub-id}/resourceGroups/{rg-name}"
```

### Custom Duration

Override default activation duration:

```powershell
.\PIMMeUpScotty.ps1 -AADDurationHours 4 -AzureDurationHours 6
```

## Authentication

The script requires interactive authentication:

1. **Azure Authentication**: Browser-based sign-in for Azure PowerShell
2. **Microsoft Graph Authentication**: Browser-based consent for Graph API permissions

Required Microsoft Graph permissions:
- `User.Read`
- `RoleManagement.ReadWrite.Directory`

## State Management

The script maintains state in `pim-state.json` to track:

- Last activation times
- Expiration timestamps
- Optimal activation durations
- Failure counts

This enables smart features like:
- Skipping roles that are still active
- Avoiding repeated failures
- Learning the best activation durations for your policies

## Logging

All activity is logged to `pim-activation.log` including:

- Activation successes and failures
- Connection attempts
- Configuration loading
- State changes

## Troubleshooting

### Authentication Issues

If you encounter authentication errors:

```powershell
# Disconnect existing sessions
Disconnect-AzAccount
Disconnect-MgGraph

# Re-run the script
.\PIMMeUpScotty.ps1
```

### No Eligible Roles Found

Use diagnostic mode to verify your eligible roles:

```powershell
.\PIMMeUpScotty.ps1 -ScanAll
```

Compare results with the Azure Portal:
- Azure AD roles: https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadmigratedroles
- Azure Resource roles: https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/azurerbac

### Activation Failures

Common causes:
- **Duration too long**: The script automatically retries with shorter durations (4h, then 2h)
- **Missing prerequisites**: Ensure PIM policies allow self-activation
- **Approval required**: Some roles may require additional approval
- **Already active**: The role is already activated (shown as "Already Active")

## Automation

### Windows Task Scheduler

Create a scheduled task to run daily:

```powershell
$action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument '-ExecutionPolicy Bypass -File "C:\Path\To\PIMMeUpScotty.ps1"'
$trigger = New-ScheduledTaskTrigger -Daily -At 8:00AM
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive
Register-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -TaskName "PIM Daily Activation" -Description "Activate PIM roles daily"
```

**Note**: Scheduled tasks require the user to be logged in for interactive authentication.

## Files Created

- **pim-config.json**: Your role configuration
- **pim-state.json**: Activation state and history
- **pim-activation.log**: Activity log

## Security Considerations

- Configuration files are stored locally in the script directory
- No credentials are stored; authentication uses interactive browser-based flows
- Justification text is included in Azure audit logs
- The script requires appropriate PIM role eligibility

## Examples

### Example 1: Initial Setup

```powershell
# Run setup wizard
.\PIMMeUpScotty.ps1 -Setup
```

The wizard will check modules, install missing ones, and guide you through configuration.

### Example 2: Daily Morning Activation

```powershell
# Run at the start of your workday
.\PIMMeUpScotty.ps1
```

Output:
```
  /\     PIM Me Up, Scotty!
 /  \    -------------------------
/____\   Azure PIM Role Assignment

✓ Configuration loaded
=== Connecting to Azure and Microsoft Graph ===
Connected to Azure as: user@company.com
Connected to Microsoft Graph as: user@company.com
=== Retrieving Eligible Azure AD Roles ===
Found 3 eligible Azure AD role(s)
=== Activating Azure AD Roles ===
  → Activating: Global Administrator... ✓ Activated
  → Activating: Security Administrator... ✓ Already Active
  → Activating: User Administrator... ✓ Activated
=== Activating Azure Resource Roles ===
Scope: Subscription - /subscriptions/...
  → Activating: Owner... ✓ Activated
╔════════════════════════════════════════════╗
║              Summary                       ║
╚════════════════════════════════════════════╝
Successful activations: 4
Skipped (still active): 0
Failed activations: 0
Execution time: 12.34 seconds
```

### Example 3: Discovery Mode

```powershell
# Find all eligible roles
.\PIMMeUpScotty.ps1 -ScanAll
```

### Example 4: Quick Check

```powershell
# See what would be activated without activating
.\PIMMeUpScotty.ps1 -ListOnly
```

## Author

**Claus Lehmann Munch**  
January 2026

## License

This script is provided as-is without warranty. Use at your own discretion.

## Contributing

Feel free to submit issues or enhancement suggestions.

---

**Beam me up, Scotty! ✨**
