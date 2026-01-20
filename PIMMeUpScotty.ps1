<#
.SYNOPSIS
    Daily PIM activation script for Azure AD and Azure Resource roles.

.DESCRIPTION
    This script activates your eligible PIM roles in Azure AD (Entra ID) and Azure Resources.
    Configure the roles you want to activate in the configuration section below.

.NOTES
    Author: Claus Lehmann Munch
    Date: January 2026
    
    Required Modules:
    - Az.Accounts
    - Az.Resources
    - Microsoft.Graph.Authentication
    - Microsoft.Graph.Users
    - Microsoft.Graph.Identity.Governance

.PARAMETER Setup
    Run the setup wizard to check and install required PowerShell modules and configure the script.

.PARAMETER ListOnly
    Only list eligible roles without activating them.

.PARAMETER ScanAll
    Scan ALL Azure resources to find eligible roles (diagnostic mode).

.PARAMETER ResourceScope
    Scan a specific resource scope (e.g., resource ID or path).

.PARAMETER AADDurationHours
    Duration in hours for Azure AD role activations (default: 8).

.PARAMETER AzureDurationHours
    Duration in hours for Azure Resource role activations (default: 8).

.EXAMPLE
    .\PIMMeUpScotty.ps1 -Setup
    Runs the setup wizard to install modules and configure the script.

.EXAMPLE
    .\PIMMeUpScotty.ps1
    Activates all configured PIM roles for daily work.

.EXAMPLE
    .\PIMMeUpScotty.ps1 -ListOnly
    Lists all eligible roles without activating them.

.EXAMPLE
    .\PIMMeUpScotty.ps1 -ScanAll
    Scans all Azure resources to discover eligible roles.
#>

[CmdletBinding()]
param(
    [switch]$Setup,     # Run setup wizard to check modules and configure
    [switch]$ListOnly,  # Only list eligible roles without activating
    [switch]$ScanAll,   # Scan ALL Azure resources to find eligible roles
    [string]$ResourceScope,  # Scan a specific resource scope (e.g., resource ID or path)
    [int]$AADDurationHours = 8,  # Duration for Azure AD roles
    [int]$AzureDurationHours = 8  # Duration for Azure Resource roles
)

#region Configuration
# File paths
$ConfigFilePath = Join-Path $PSScriptRoot "pim-config.json"
$StateFilePath = Join-Path $PSScriptRoot "pim-state.json"
$LogFilePath = Join-Path $PSScriptRoot "pim-activation.log"

# Configuration variables (loaded from pim-config.json)
$TenantID = $null
$SubscriptionID = $null
$DefaultJustification = $null
$AADRolesToActivate = @()
$AzureResourcesToActivate = @()
#endregion Configuration

#region Setup Functions
function Invoke-Setup {
    Write-Host "`n╔════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║        PIMMeUpScotty Setup Wizard          ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    
    # Step 1: Check and install required modules
    Write-Host "Step 1: Checking Required PowerShell Modules" -ForegroundColor Yellow
    Write-Host "═══════════════════════════════════════════════" -ForegroundColor Yellow
    Write-Host ""
    
    $requiredModules = @(
        'Az.Accounts'
        'Az.Resources'
        'Microsoft.Graph.Authentication'
        'Microsoft.Graph.Users'
        'Microsoft.Graph.Identity.Governance'
    )
    
    $modulesInstalled = 0
    $modulesAlreadyPresent = 0
    
    foreach ($moduleName in $requiredModules) {
        Write-Host "  Checking: $moduleName..." -ForegroundColor Cyan -NoNewline
        
        $module = Get-Module -ListAvailable -Name $moduleName | Select-Object -First 1
        
        if ($module) {
            Write-Host " ✓ Found (v$($module.Version))" -ForegroundColor Green
            $modulesAlreadyPresent++
        }
        else {
            Write-Host " ✗ Not found" -ForegroundColor Red
            Write-Host "    Installing $moduleName..." -ForegroundColor Yellow
            
            try {
                Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                $installedModule = Get-Module -ListAvailable -Name $moduleName | Select-Object -First 1
                Write-Host "    ✓ Installed successfully (v$($installedModule.Version))" -ForegroundColor Green
                $modulesInstalled++
            }
            catch {
                Write-Host "    ✗ Failed to install: $_" -ForegroundColor Red
                Write-Error "Failed to install required module: $moduleName"
                return $false
            }
        }
    }
    
    Write-Host ""
    Write-Host "Module Check Summary:" -ForegroundColor Green
    Write-Host "  Already present: $modulesAlreadyPresent" -ForegroundColor Gray
    if ($modulesInstalled -gt 0) {
        Write-Host "  Newly installed: $modulesInstalled" -ForegroundColor Gray
    }
    Write-Host ""
    
    # Step 2: Configuration setup
    Write-Host "Step 2: Configuration Setup" -ForegroundColor Yellow
    Write-Host "═══════════════════════════════════════════════" -ForegroundColor Yellow
    Write-Host ""
    
    # Check if config already exists
    if (Test-Path $ConfigFilePath) {
        Write-Host "Existing configuration found at: $ConfigFilePath" -ForegroundColor Yellow
        $overwrite = Read-Host "Do you want to overwrite it? (y/N)"
        if ($overwrite -ne 'y' -and $overwrite -ne 'Y') {
            Write-Host "Setup cancelled. Existing configuration preserved." -ForegroundColor Cyan
            return $true
        }
        Write-Host ""
    }
    
    # Prompt for required information
    Write-Host "Required Information:" -ForegroundColor Green
    $tenantId = Read-Host "Enter your Azure Tenant ID"
    if ([string]::IsNullOrWhiteSpace($tenantId)) {
        Write-Error "Tenant ID is required"
        return $false
    }
    
    $subscriptionId = Read-Host "Enter your Azure Subscription ID"
    if ([string]::IsNullOrWhiteSpace($subscriptionId)) {
        Write-Error "Subscription ID is required"
        return $false
    }
    
    Write-Host "`nOptional Information (press Enter to use defaults):" -ForegroundColor Green
    $tenantName = Read-Host "Tenant Name (optional)"
    $subscriptionName = Read-Host "Subscription Name (optional)"
    
    Write-Host "`nJustification text will be shown when activating roles." -ForegroundColor Gray
    Write-Host "Default: 'Daily operational work'" -ForegroundColor Gray
    $justification = Read-Host "Custom justification text (press Enter for default)"
    if ([string]::IsNullOrWhiteSpace($justification)) {
        $justification = "Daily operational work"
    }
    
    $durationStr = Read-Host "Maximum activation duration in hours (press Enter for 8)"
    $maxDuration = 8
    if (-not [string]::IsNullOrWhiteSpace($durationStr) -and [int]::TryParse($durationStr, [ref]$maxDuration)) {
        # Use parsed value
    } else {
        $maxDuration = 8
    }
    
    # Create config with provided values
    $newConfig = @{
        Scopes = @(
            @{
                TenantId = $tenantId
                TenantName = if ($tenantName) { $tenantName } else { "" }
                SubscriptionId = $subscriptionId
                SubscriptionName = if ($subscriptionName) { $subscriptionName } else { "" }
                ScopeId = $null
                LandingZoneName = $null
                ActivationMode = "AllEligible"
                MaxDurationHours = $maxDuration
                Roles = $null
                AADRoles = @()
            }
        )
        DefaultJustification = $justification
    }
    
    # Save configuration
    try {
        $newConfig | ConvertTo-Json -Depth 10 | Set-Content -Path $ConfigFilePath -Force
        Write-Host "`n✓ Configuration saved to: $ConfigFilePath" -ForegroundColor Green
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "║          Setup Complete!                   ║" -ForegroundColor Green
        Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
        Write-Host "You can now run the script without -Setup to activate your PIM roles." -ForegroundColor Gray
        Write-Host "Edit $ConfigFilePath to customize your configuration." -ForegroundColor Gray
        Write-Host ""
        return $true
    }
    catch {
        Write-Error "Failed to save configuration: $_"
        return $false
    }
}
#endregion Setup Functions

#region Configuration Loading
function Load-Configuration {
    if (-not (Test-Path $ConfigFilePath)) {
        Write-Host ""
        Write-Host "No configuration file found. Running setup..." -ForegroundColor Yellow
        Write-Host ""
        
        if (-not (Invoke-Setup)) {
            Write-Error "Setup failed. Cannot continue."
            return $false
        }
        
        # After setup, the config file should exist, so reload
        if (-not (Test-Path $ConfigFilePath)) {
            Write-Error "Configuration file was not created. Setup may have failed."
            return $false
        }
    }
    
    try {
        $config = Get-Content $ConfigFilePath -Raw | ConvertFrom-Json
        Write-Verbose "Loaded configuration from $ConfigFilePath"
        
        # Validate and extract configuration
        if (-not $config.Scopes -or $config.Scopes.Count -eq 0) {
            throw "Configuration file must contain at least one scope"
        }
        
        $firstScope = $config.Scopes[0]
        
        # Extract required values
        $script:TenantID = $firstScope.TenantId
        $script:SubscriptionID = $firstScope.SubscriptionId
        $script:DefaultJustification = if ($config.DefaultJustification) { $config.DefaultJustification } else { "Daily operational work" }
        
        # Prompt for missing required values
        if ([string]::IsNullOrWhiteSpace($script:TenantID)) {
            $script:TenantID = Read-Host "Enter Tenant ID (required)"
            if ([string]::IsNullOrWhiteSpace($script:TenantID)) {
                throw "Tenant ID is required"
            }
        }
        
        if ([string]::IsNullOrWhiteSpace($script:SubscriptionID)) {
            $script:SubscriptionID = Read-Host "Enter Subscription ID (required)"
            if ([string]::IsNullOrWhiteSpace($script:SubscriptionID)) {
                throw "Subscription ID is required"
            }
        }
        
        # Extract AAD roles if specified
        if ($firstScope.AADRoles -and $firstScope.AADRoles.Count -gt 0) {
            $script:AADRolesToActivate = $firstScope.AADRoles
        } else {
            # Empty array means activate ALL eligible AAD roles
            $script:AADRolesToActivate = @()
        }
        
        # Build Azure resource scopes
        $script:AzureResourcesToActivate = @()
        
        foreach ($scope in $config.Scopes) {
            if ($scope.SubscriptionId) {
                # Subscription scope
                $script:AzureResourcesToActivate += @{
                    ScopeType = "Subscription"
                    ScopeId = "/subscriptions/$($scope.SubscriptionId)"
                    RolesToActivate = if ($scope.Roles) { $scope.Roles } else { @() }
                    MaxDurationHours = if ($scope.MaxDurationHours) { $scope.MaxDurationHours } else { 8 }
                }
            }
            
            if ($scope.ScopeId) {
                # Custom scope (could be management group, resource group, etc.)
                $scopeType = "Subscription"
                if ($scope.ScopeId -match '/managementGroups/') {
                    $scopeType = "ManagementGroup"
                } elseif ($scope.ScopeId -match '/resourceGroups/') {
                    $scopeType = "ResourceGroup"
                }
                
                $script:AzureResourcesToActivate += @{
                    ScopeType = $scopeType
                    ScopeId = $scope.ScopeId
                    RolesToActivate = if ($scope.Roles) { $scope.Roles } else { @() }
                    MaxDurationHours = if ($scope.MaxDurationHours) { $scope.MaxDurationHours } else { 8 }
                }
            }
        }
        
        # Use parameter values if provided, otherwise use config
        if ($AADDurationHours -eq 8 -and $config.Scopes[0].MaxDurationHours) {
            $script:AADDurationHours = $config.Scopes[0].MaxDurationHours
        }
        
        if ($AzureDurationHours -eq 8 -and $config.Scopes[0].MaxDurationHours) {
            $script:AzureDurationHours = $config.Scopes[0].MaxDurationHours
        }
        
        Write-Host "✓ Configuration loaded" -ForegroundColor Green
        Write-Verbose "  Tenant: $script:TenantID"
        Write-Verbose "  Subscription: $script:SubscriptionID"
        Write-Verbose "  AAD Roles to activate: $($script:AADRolesToActivate.Count) $(if ($script:AADRolesToActivate.Count -eq 0) { '(ALL eligible)' })"
        Write-Verbose "  Azure Resource scopes: $($script:AzureResourcesToActivate.Count)"
        
        return $true
    }
    catch {
        Write-Error "Failed to load configuration: $_"
        Write-Host "Please check your pim-config.json file format" -ForegroundColor Red
        return $false
    }
}
#endregion Configuration Loading

#region Logging and Persistence Functions
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Append to log file
    Add-Content -Path $LogFilePath -Value $logEntry -ErrorAction SilentlyContinue
    
    # Also write to verbose stream for debugging
    Write-Verbose $logEntry
}

function Load-State {
    if (Test-Path $StateFilePath) {
        try {
            $state = Get-Content $StateFilePath -Raw | ConvertFrom-Json
            Write-Log "Loaded state from $StateFilePath"
            return $state
        }
        catch {
            Write-Log "Failed to load state file: $_" -Level Warning
            return Initialize-State
        }
    }
    else {
        Write-Log "No state file found, initializing new state"
        return Initialize-State
    }
}

function Initialize-State {
    return [PSCustomObject]@{
        lastRun = $null
        userId = $null
        activationHistory = @{
            aad = @{}
            azure = @{}
        }
        preferences = @{
            defaultJustification = $DefaultJustification
            aadDurationHours = $AADDurationHours
            azureDurationHours = $AzureDurationHours
        }
    }
}

function Save-State {
    param([PSCustomObject]$State)
    
    try {
        $State | ConvertTo-Json -Depth 10 | Set-Content -Path $StateFilePath -Force
        Write-Log "Saved state to $StateFilePath"
    }
    catch {
        Write-Log "Failed to save state: $_" -Level Error
    }
}

function Test-RoleStillActive {
    param(
        [PSCustomObject]$RoleHistory
    )
    
    if (-not $RoleHistory.expiresAt) {
        return $false
    }
    
    try {
        # Handle both DateTime objects and string formats
        $expiresAt = $null
        if ($RoleHistory.expiresAt -is [DateTime]) {
            $expiresAt = $RoleHistory.expiresAt
        }
        else {
            # Try parsing with invariant culture for ISO 8601 format
            $expiresAt = [DateTime]::Parse($RoleHistory.expiresAt, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
        }
        
        $now = Get-Date
        
        # Consider active if expires more than 30 minutes from now
        if ($expiresAt -gt $now.AddMinutes(30)) {
            $remainingTime = $expiresAt - $now
            Write-Log "Role still active, expires in $([math]::Round($remainingTime.TotalHours, 1)) hours"
            return $true
        }
    }
    catch {
        Write-Log "Failed to parse expiration time: $_" -Level Warning
    }
    
    return $false
}

function Update-RoleHistory {
    param(
        [PSCustomObject]$State,
        [string]$RoleType,  # 'aad' or 'azure'
        [string]$RoleKey,
        [bool]$Success,
        [int]$Duration = 0
    )
    
    $history = $State.activationHistory.$RoleType
    
    # Handle both hashtables (new state) and PSObjects (loaded from JSON)
    $roleHistory = $null
    if ($history -is [hashtable]) {
        if (-not $history.ContainsKey($RoleKey)) {
            $history[$RoleKey] = @{
                lastActivated = $null
                expiresAt = $null
                optimalDuration = 0
                consecutiveFailures = 0
                totalActivations = 0
                totalFailures = 0
            }
        }
        $roleHistory = $history[$RoleKey]
    }
    else {
        # PSObject from JSON
        $roleHistory = $history.PSObject.Properties[$RoleKey]?.Value
        if (-not $roleHistory) {
            $roleHistory = [PSCustomObject]@{
                lastActivated = $null
                expiresAt = $null
                optimalDuration = 0
                consecutiveFailures = 0
                totalActivations = 0
                totalFailures = 0
            }
            $history | Add-Member -MemberType NoteProperty -Name $RoleKey -Value $roleHistory -Force
        }
    }
    
    if ($Success) {
        $now = Get-Date
        $roleHistory.lastActivated = $now.ToString("o")
        $roleHistory.expiresAt = $now.AddHours($Duration).ToString("o")
        $roleHistory.consecutiveFailures = 0
        $roleHistory.totalActivations++
        
        # Update optimal duration if this worked
        if ($Duration -gt 0 -and ($roleHistory.optimalDuration -eq 0 -or $Duration -gt $roleHistory.optimalDuration)) {
            $roleHistory.optimalDuration = $Duration
        }
        
        Write-Log "Updated history for $RoleKey - Success, expires at $($roleHistory.expiresAt)"
    }
    else {
        $roleHistory.consecutiveFailures++
        $roleHistory.totalFailures++
        Write-Log "Updated history for $RoleKey - Failure (consecutive: $($roleHistory.consecutiveFailures))" -Level Warning
    }
}

function Get-OptimalDuration {
    param(
        [PSCustomObject]$State,
        [string]$RoleType,
        [string]$RoleKey,
        [int]$DefaultDuration
    )
    
    $history = $State.activationHistory.$RoleType
    
    # Handle both hashtables (new state) and PSObjects (loaded from JSON)
    $roleHistory = $null
    if ($history -is [hashtable]) {
        $roleHistory = $history[$RoleKey]
    }
    else {
        # PSObject from JSON
        $roleHistory = $history.PSObject.Properties[$RoleKey]?.Value
    }
    
    if ($roleHistory -and $roleHistory.optimalDuration -gt 0) {
        $optimal = $roleHistory.optimalDuration
        Write-Log "Using optimal duration of ${optimal}h for $RoleKey (learned from previous runs)"
        return $optimal
    }
    
    return $DefaultDuration
}
#endregion Logging and Persistence Functions

#region Helper Functions
function Ensure-Module {
    param([string]$Name)
    
    # Check if module is already loaded
    $loadedModule = Get-Module -Name $Name
    if ($loadedModule) {
        Write-Verbose "Module $Name is already loaded (version $($loadedModule.Version))"
        return
    }
    
    # Check if module is available
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Installing module: $Name..." -ForegroundColor Yellow
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }
    
    # Import module if not already loaded
    try {
        Import-Module $Name -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Verbose "Imported module: $Name"
    }
    catch {
        # If import fails due to assembly conflict, check if it's actually usable
        if (Get-Module -Name $Name) {
            Write-Verbose "Module $Name is loaded despite import error"
        } else {
            throw
        }
    }
}

function Connect-AzureAndGraph {
    param(
        [string]$TenantId,
        [string]$SubscriptionId
    )
    
    Write-Host "`n=== Connecting to Azure and Microsoft Graph ===" -ForegroundColor Cyan
    
    # Ensure modules
    $modules = @('Az.Accounts', 'Az.Resources', 'Microsoft.Graph.Authentication', 
                 'Microsoft.Graph.Users', 'Microsoft.Graph.Identity.Governance')
    foreach ($module in $modules) {
        Ensure-Module -Name $module
    }
    
    # Connect to Azure
    Write-Host "Connecting to Azure..." -ForegroundColor Yellow
    $azContext = Get-AzContext -ErrorAction SilentlyContinue
    
    # For PIM operations, we need MicrosoftGraphEndpointResourceId scope
    # Always reconnect to ensure we have the proper auth scope
    $needsReconnect = $false
    if (-not $azContext -or $azContext.Tenant.Id -ne $TenantId) {
        $needsReconnect = $true
    } else {
        # Check if we can access role definitions (requires proper auth scope)
        try {
            $null = Get-AzRoleDefinition -Name "Reader" -ErrorAction Stop
            Write-Host "Using existing Azure connection: $($azContext.Account.Id)" -ForegroundColor Green
        } catch {
            # Existing connection lacks required scope, need to reconnect
            Write-Host "Existing Azure connection lacks required scope. Reconnecting..." -ForegroundColor Yellow
            $needsReconnect = $true
        }
    }
    
    if ($needsReconnect) {
        Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
        Connect-AzAccount -TenantId $TenantId -Subscription $SubscriptionId -AuthScope MicrosoftGraphEndpointResourceId | Out-Null
        Write-Host "Connected to Azure as: $((Get-AzContext).Account.Id)" -ForegroundColor Green
    }
    
    # Connect to Microsoft Graph
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    $mgContext = Get-MgContext -ErrorAction SilentlyContinue
    
    # Check if we have a valid session with required scopes
    $needsConnect = $true
    if ($mgContext -and $mgContext.TenantId -eq $TenantId) {
        $requiredScopes = @('User.Read', 'RoleManagement.ReadWrite.Directory')
        $hasAllScopes = $true
        
        foreach ($scope in $requiredScopes) {
            if ($mgContext.Scopes -notcontains $scope) {
                $hasAllScopes = $false
                break
            }
        }
        
        if ($hasAllScopes) {
            Write-Host "Reusing existing Microsoft Graph connection: $($mgContext.Account)" -ForegroundColor Green
            $needsConnect = $false
        } else {
            Write-Host "Existing session lacks required permissions. Reconnecting..." -ForegroundColor Yellow
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
    }
    
    if ($needsConnect) {
        $scopes = @('User.Read', 'RoleManagement.ReadWrite.Directory')
        
        Write-Host "A browser window will open for authentication." -ForegroundColor Yellow
        Write-Host "Please complete the sign-in and grant consent if prompted." -ForegroundColor Yellow
        Start-Sleep -Seconds 1
        
        $maxRetries = 2
        $retryCount = 0
        $connected = $false
        
        while (-not $connected -and $retryCount -lt $maxRetries) {
            try {
                if ($retryCount -gt 0) {
                    Write-Host "`nRetry attempt $retryCount of $($maxRetries - 1)..." -ForegroundColor Yellow
                    Start-Sleep -Seconds 2
                }
                
                Connect-MgGraph -Scopes $scopes -TenantId $TenantId -NoWelcome -ErrorAction Stop | Out-Null
                $mgContext = Get-MgContext
                
                if ($mgContext -and $mgContext.Account) {
                    $connected = $true
                    Write-Host "Connected to Microsoft Graph as: $($mgContext.Account)" -ForegroundColor Green
                }
            }
            catch {
                $retryCount++
                if ($retryCount -ge $maxRetries) {
                    Write-Host "`nTroubleshooting tips:" -ForegroundColor Red
                    Write-Host "1. Make sure you complete the authentication in the browser" -ForegroundColor White
                    Write-Host "2. Click 'Accept' on the consent page if prompted" -ForegroundColor White
                    Write-Host "3. Don't close the browser window until authentication completes" -ForegroundColor White
                    Write-Host "4. Try running: Disconnect-MgGraph" -ForegroundColor White
                    Write-Host "5. Then run this script again" -ForegroundColor White
                    throw "Failed to connect to Microsoft Graph after $maxRetries attempts: $_"
                } else {
                    Write-Warning "Authentication attempt failed: $_"
                }
            }
        }
    } else {
        $mgContext = Get-MgContext
    }
    
    if (-not $mgContext) {
        throw "Failed to connect to Microsoft Graph"
    }
    
    # Get current user's object ID
    $currentUser = (Get-MgUser -UserId $mgContext.Account).Id
    return $currentUser
}

function Get-EligibleAADRoles {
    param([string]$UserId)
    
    Write-Host "`n=== Retrieving Eligible Azure AD Roles ===" -ForegroundColor Cyan
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilitySchedules?`$filter=principalId eq '$UserId'&`$expand=roleDefinition"
        $result = Invoke-MgGraphRequest -Method GET -Uri $uri
        
        $roles = $result.value | ForEach-Object {
            [PSCustomObject]@{
                RoleId = $_.roleDefinitionId
                RoleName = $_.roleDefinition.displayName
                DirectoryScopeId = $_.directoryScopeId
            }
        }
        
        Write-Host "Found $($roles.Count) eligible Azure AD role(s)" -ForegroundColor Green
        return $roles
    }
    catch {
        Write-Warning "Failed to retrieve Azure AD roles: $_"
        return @()
    }
}

function Get-EligibleAzureResourceRoles {
    param(
        [string]$UserId,
        [string]$Scope
    )
    
    try {
        # Get all eligible assignments for this scope, then filter by user
        # The -Filter parameter doesn't always work reliably, so we filter after retrieval
        $eligibleAssignments = Get-AzRoleEligibilityScheduleInstance -Scope $Scope -ErrorAction Stop
        
        # Filter to current user
        $userAssignments = $eligibleAssignments | Where-Object { $_.PrincipalId -eq $UserId }
        
        return $userAssignments | ForEach-Object {
            [PSCustomObject]@{
                RoleDefinitionId = $_.RoleDefinitionId
                RoleName = $_.RoleDefinitionName
                ScopeId = $_.Scope
            }
        }
    }
    catch {
        Write-Warning "Failed to retrieve Azure roles for scope $Scope : $_"
        return @()
    }
}

function Get-AllEligibleAzureResourceRoles {
    param([string]$UserId)
    
    Write-Host "`n=== Scanning ALL Azure Resources for Eligible Roles ===" -ForegroundColor Cyan
    
    try {
        # Get all subscriptions
        $subscriptions = Get-AzSubscription
        $allEligibleRoles = @()
        
        foreach ($sub in $subscriptions) {
            Write-Host "Checking subscription: $($sub.Name)..." -ForegroundColor Yellow
            Set-AzContext -SubscriptionId $sub.Id | Out-Null
            
            # Check subscription level
            $subRoles = Get-AzRoleEligibilityScheduleInstance -Scope "/subscriptions/$($sub.Id)" -ErrorAction SilentlyContinue |
                Where-Object { $_.PrincipalId -eq $UserId }
            
            if ($subRoles) {
                $allEligibleRoles += $subRoles | ForEach-Object {
                    [PSCustomObject]@{
                        Scope = $_.Scope
                        ScopeType = "Subscription"
                        ScopeName = $sub.Name
                        RoleName = $_.RoleDefinitionName
                        RoleDefinitionId = $_.RoleDefinitionId
                    }
                }
            }
            
            # Check resource groups
            $resourceGroups = Get-AzResourceGroup
            foreach ($rg in $resourceGroups) {
                $rgRoles = Get-AzRoleEligibilityScheduleInstance -Scope $rg.ResourceId -ErrorAction SilentlyContinue |
                    Where-Object { $_.PrincipalId -eq $UserId }
                
                if ($rgRoles) {
                    $allEligibleRoles += $rgRoles | ForEach-Object {
                        [PSCustomObject]@{
                            Scope = $_.Scope
                            ScopeType = "ResourceGroup"
                            ScopeName = $rg.ResourceGroupName
                            RoleName = $_.RoleDefinitionName
                            RoleDefinitionId = $_.RoleDefinitionId
                        }
                    }
                }
                
                # Check individual resources within this resource group
                $resources = Get-AzResource -ResourceGroupName $rg.ResourceGroupName -ErrorAction SilentlyContinue
                foreach ($resource in $resources) {
                    $resourceRoles = Get-AzRoleEligibilityScheduleInstance -Scope $resource.ResourceId -ErrorAction SilentlyContinue |
                        Where-Object { $_.PrincipalId -eq $UserId }
                    
                    if ($resourceRoles) {
                        $allEligibleRoles += $resourceRoles | ForEach-Object {
                            [PSCustomObject]@{
                                Scope = $_.Scope
                                ScopeType = "Resource"
                                ScopeName = "$($resource.ResourceType)/$($resource.Name)"
                                RoleName = $_.RoleDefinitionName
                                RoleDefinitionId = $_.RoleDefinitionId
                            }
                        }
                    }
                }
            }
        }
        
        # Check management groups
        try {
            $mgs = Get-AzManagementGroup -ErrorAction SilentlyContinue
            foreach ($mg in $mgs) {
                $mgRoles = Get-AzRoleEligibilityScheduleInstance -Scope "/providers/Microsoft.Management/managementGroups/$($mg.Name)" -ErrorAction SilentlyContinue |
                    Where-Object { $_.PrincipalId -eq $UserId }
                
                if ($mgRoles) {
                    $allEligibleRoles += $mgRoles | ForEach-Object {
                        [PSCustomObject]@{
                            Scope = $_.Scope
                            ScopeType = "ManagementGroup"
                            ScopeName = $mg.DisplayName
                            RoleName = $_.RoleDefinitionName
                            RoleDefinitionId = $_.RoleDefinitionId
                        }
                    }
                }
            }
        }
        catch {
            Write-Verbose "Could not check management groups: $_"
        }
        
        return $allEligibleRoles
    }
    catch {
        Write-Warning "Failed to scan for Azure resource roles: $_"
        return @()
    }
}

function Activate-AADRole {
    param(
        [string]$UserId,
        [string]$RoleDefinitionId,
        [string]$DirectoryScopeId,
        [int]$DurationHours,
        [string]$Justification
    )
    
    $uri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests"
    
    # Try with requested duration first, then fallback to shorter durations if policy fails
    $durationsToTry = @($DurationHours)
    if ($DurationHours -gt 4) {
        $durationsToTry += 4  # Try 4 hours if original request fails
    }
    if ($DurationHours -gt 2) {
        $durationsToTry += 2  # Try 2 hours as last resort
    }
    
    foreach ($duration in $durationsToTry) {
        $body = @{
            action = "selfActivate"
            principalId = $UserId
            roleDefinitionId = $RoleDefinitionId
            directoryScopeId = $DirectoryScopeId
            justification = $Justification
            scheduleInfo = @{
                startDateTime = (Get-Date).ToUniversalTime().ToString("o")
                expiration = @{
                    type = "AfterDuration"
                    duration = "PT${duration}H"
                }
            }
        } | ConvertTo-Json -Depth 5
        
        try {
            Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType 'application/json' | Out-Null
            if ($duration -ne $DurationHours) {
                Write-Host " (${duration}h)" -ForegroundColor Yellow -NoNewline
            }
            return @{ Success = $true; Status = 'Activated' }
        }
        catch {
            $errorMessage = $_.ToString()
            
            # Check if role is already active or pending
            if ($errorMessage -like "*RoleAssignmentExists*") {
                return @{ Success = $true; Status = 'AlreadyActive' }
            }
            elseif ($errorMessage -like "*PendingRoleAssignmentRequest*") {
                return @{ Success = $true; Status = 'Pending' }
            }
            elseif ($errorMessage -like "*ExpirationRule*" -and $duration -ne $durationsToTry[-1]) {
                # Policy validation failed on duration, try shorter duration
                continue
            }
            
            Write-Warning "Failed to activate: $_"
            return @{ Success = $false; Status = 'Failed' }
        }
    }
    
    return @{ Success = $false; Status = 'Failed' }
}

function Activate-AzureResourceRole {
    param(
        [string]$UserId,
        [string]$RoleDefinitionId,
        [string]$ScopeId,
        [int]$DurationHours,
        [string]$Justification
    )
    
    # Use Az.Resources cmdlet for Azure resource role activation
    try {
        $params = @{
            Name = (New-Guid).Guid
            Scope = $ScopeId
            PrincipalId = $UserId
            RoleDefinitionId = $RoleDefinitionId
            RequestType = 'SelfActivate'
            Justification = $Justification
            ExpirationDuration = "PT${DurationHours}H"
            ExpirationType = 'AfterDuration'
            ScheduleInfoStartDateTime = (Get-Date).ToUniversalTime()
        }
        
        New-AzRoleAssignmentScheduleRequest @params -ErrorAction Stop | Out-Null
        return @{ Success = $true; Status = 'Activated' }
    }
    catch {
        $errorMessage = $_.Exception.Message
        
        # Check if role is already active
        if ($errorMessage -like "*already exists*" -or $errorMessage -like "*already active*") {
            return @{ Success = $true; Status = 'AlreadyActive' }
        }
        
        Write-Warning "Failed to activate: $_"
        return @{ Success = $false; Status = 'Failed' }
    }
}
#endregion Helper Functions

#region Main Execution
try {
    # Handle -Setup parameter
    if ($Setup) {
        exit (Invoke-Setup | ForEach-Object { if ($_) { 0 } else { 1 } })
    }
    
    $startTime = Get-Date
    Write-Host ""
    Write-Host "  /\     PIM Me Up, Scotty!" -ForegroundColor Cyan
    Write-Host " /  \    -------------------------" -ForegroundColor Cyan
    Write-Host "/____\   Azure PIM Role Assignment" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Log "=== PIM Activation Script Started ===" -Level Info
    
    # Load configuration
    if (-not (Load-Configuration)) {
        Write-Error "Failed to load configuration. Exiting."
        exit 1
    }
    
    # Load state
    $state = Load-State
    $state.lastRun = $startTime.ToString("o")
    
    # Connect and get user ID
    $currentUserId = Connect-AzureAndGraph -TenantId $TenantID -SubscriptionId $SubscriptionID
    Write-Host "Current User ID: $currentUserId" -ForegroundColor Gray
    Write-Log "Connected as user: $currentUserId"
    
    $state.userId = $currentUserId
    
    $activationCount = 0
    $failureCount = 0
    $skippedCount = 0
    
    #region Azure AD Roles
    $aadRoles = Get-EligibleAADRoles -UserId $currentUserId
    
    if ($aadRoles.Count -gt 0) {
        Write-Host "`nEligible Azure AD Roles:" -ForegroundColor Yellow
        Write-Log "Found $($aadRoles.Count) eligible Azure AD roles"
        $aadRoles | ForEach-Object { Write-Host "  - $($_.RoleName)" -ForegroundColor Gray }
        
        if (-not $ListOnly) {
            Write-Host "`nActivating Azure AD Roles..." -ForegroundColor Cyan
            
            foreach ($role in $aadRoles) {
                # Filter by configured roles if specified
                if ($AADRolesToActivate.Count -gt 0 -and $role.RoleName -notin $AADRolesToActivate) {
                    Write-Host "  ⊘ Skipping: $($role.RoleName) (not in activation list)" -ForegroundColor DarkGray
                    Write-Log "Skipped $($role.RoleName) - not in activation list"
                    continue
                }
                
                $roleKey = $role.RoleName
                
                # Check if role is still active from previous run
                $history = $state.activationHistory.aad
                $roleHistory = if ($history -is [hashtable]) { $history[$roleKey] } else { $history.PSObject.Properties[$roleKey]?.Value }
                if ($roleHistory -and (Test-RoleStillActive $roleHistory)) {
                    Write-Host "  ⊙ Skipping: $($role.RoleName) (still active)" -ForegroundColor Cyan
                    Write-Log "Skipped $($role.RoleName) - still active from previous run"
                    $skippedCount++
                    continue
                }
                
                # Skip if too many consecutive failures
                if ($roleHistory -and $roleHistory.consecutiveFailures -ge 3) {
                    Write-Host "  ⊘ Skipping: $($role.RoleName) (too many failures)" -ForegroundColor DarkGray
                    Write-Log "Skipped $($role.RoleName) - too many consecutive failures ($($roleHistory.consecutiveFailures))" -Level Warning
                    continue
                }
                
                # Use optimal duration if known
                $duration = Get-OptimalDuration -State $state -RoleType 'aad' -RoleKey $roleKey -DefaultDuration $AADDurationHours
                
                Write-Host "  → Activating: $($role.RoleName)..." -ForegroundColor Yellow -NoNewline
                Write-Log "Attempting to activate AAD role: $($role.RoleName) for ${duration}h"
                
                $result = Activate-AADRole -UserId $currentUserId `
                    -RoleDefinitionId $role.RoleId `
                    -DirectoryScopeId $role.DirectoryScopeId `
                    -DurationHours $duration `
                    -Justification $DefaultJustification
                
                if ($result.Success) {
                    switch ($result.Status) {
                        'Activated' {
                            Write-Host " ✓ Activated" -ForegroundColor Green
                            Write-Log "Successfully activated $($role.RoleName) for ${duration}h" -Level Success
                            Update-RoleHistory -State $state -RoleType 'aad' -RoleKey $roleKey -Success $true -Duration $duration
                            $activationCount++
                        }
                        'AlreadyActive' {
                            Write-Host " ✓ Already Active" -ForegroundColor Cyan
                            Write-Log "$($role.RoleName) is already active" -Level Info
                            Update-RoleHistory -State $state -RoleType 'aad' -RoleKey $roleKey -Success $true -Duration $duration
                            $activationCount++
                        }
                        'Pending' {
                            Write-Host " ⏳ Pending" -ForegroundColor Yellow
                            Write-Log "$($role.RoleName) activation is pending" -Level Info
                            Update-RoleHistory -State $state -RoleType 'aad' -RoleKey $roleKey -Success $true -Duration $duration
                            $activationCount++
                        }
                    }
                } else {
                    Write-Host " ✗ Failed" -ForegroundColor Red
                    Write-Log "Failed to activate $($role.RoleName)" -Level Error
                    Update-RoleHistory -State $state -RoleType 'aad' -RoleKey $roleKey -Success $false
                    $failureCount++
                }
            }
        }
    }
    #endregion Azure AD Roles
    
    #region Azure Resource Roles
    if ($ResourceScope) {
        # Quick scan of a specific resource
        Write-Host "`n🔍 Checking specific resource: $ResourceScope" -ForegroundColor Magenta
        
        try {
            $roles = Get-AzRoleEligibilityScheduleInstance -Scope $ResourceScope -ErrorAction Stop |
                Where-Object { $_.PrincipalId -eq $currentUserId }
            
            if ($roles.Count -eq 0) {
                Write-Host "❌ No eligible roles found for this resource" -ForegroundColor Red
            }
            else {
                Write-Host "✓ Found $($roles.Count) eligible role(s):" -ForegroundColor Green
                foreach ($role in $roles) {
                    Write-Host "  - $($role.RoleDefinitionName)" -ForegroundColor Yellow
                    Write-Host "    Scope: $($role.Scope)" -ForegroundColor Gray
                }
                
                Write-Host "`n📝 Configuration to add:" -ForegroundColor Cyan
                Write-Host "@{" -ForegroundColor White
                Write-Host "    ScopeType = `"Resource`"" -ForegroundColor White
                Write-Host "    ScopeId = `"$ResourceScope`"" -ForegroundColor White
                Write-Host "    RolesToActivate = @()  # Empty = activate all eligible roles" -ForegroundColor White
                Write-Host "}" -ForegroundColor White
            }
        }
        catch {
            Write-Error "Failed to check resource: $_"
        }
        
        return  # Exit after specific scan
    }
    
    if ($ScanAll) {
        # Diagnostic mode: scan everything
        Write-Host "`n🔍 DIAGNOSTIC MODE: Scanning all Azure resources..." -ForegroundColor Magenta
        $allAzureRoles = Get-AllEligibleAzureResourceRoles -UserId $currentUserId
        
        if ($allAzureRoles.Count -eq 0) {
            Write-Host "`n❌ No eligible Azure Resource roles found anywhere." -ForegroundColor Red
            Write-Host "This could mean:" -ForegroundColor Yellow
            Write-Host "  1. You don't have any PIM-eligible assignments for Azure resources" -ForegroundColor White
            Write-Host "  2. All eligible assignments might be at a different scope level" -ForegroundColor White
            Write-Host "  3. Check the Azure Portal: https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/azurerbac" -ForegroundColor White
        }
        else {
            Write-Host "`n✓ Found $($allAzureRoles.Count) eligible Azure Resource role(s):" -ForegroundColor Green
            Write-Host ""
            
            $grouped = $allAzureRoles | Group-Object ScopeType
            foreach ($group in $grouped) {
                Write-Host "  $($group.Name):" -ForegroundColor Cyan
                foreach ($role in $group.Group) {
                    Write-Host "    ├─ Scope: $($role.ScopeName)" -ForegroundColor Gray
                    Write-Host "    │  Role: $($role.RoleName)" -ForegroundColor Yellow
                    Write-Host "    │  Path: $($role.Scope)" -ForegroundColor DarkGray
                    Write-Host ""
                }
            }
            
            Write-Host "`n📝 To activate these roles daily, update your configuration:" -ForegroundColor Cyan
            Write-Host "`$AzureResourcesToActivate = @(" -ForegroundColor White
            
            $uniqueScopes = $allAzureRoles | Select-Object Scope, ScopeType, ScopeName -Unique
            foreach ($scope in $uniqueScopes) {
                Write-Host "    @{" -ForegroundColor White
                Write-Host "        ScopeType = `"$($scope.ScopeType)`"" -ForegroundColor White
                Write-Host "        ScopeId = `"$($scope.Scope)`"" -ForegroundColor White
                Write-Host "        # ScopeName: $($scope.ScopeName)" -ForegroundColor DarkGray
                Write-Host "        RolesToActivate = @()  # Empty = activate all eligible roles" -ForegroundColor White
                Write-Host "    }" -ForegroundColor White
            }
            Write-Host ")" -ForegroundColor White
        }
        
        return  # Exit after diagnostic scan
    }
    
    if ($AzureResourcesToActivate.Count -gt 0) {
        Write-Host "`n=== Activating Azure Resource Roles ===" -ForegroundColor Cyan
        Write-Log "Processing $($AzureResourcesToActivate.Count) Azure resource scope(s)"
        
        foreach ($resourceConfig in $AzureResourcesToActivate) {
            Write-Host "`nScope: $($resourceConfig.ScopeType) - $($resourceConfig.ScopeId)" -ForegroundColor Yellow
            
            if (-not $ListOnly) {
                # Get Owner role definition ID
                $ownerRoleDefId = (Get-AzRoleDefinition -Name "Owner").Id
                
                # Determine the RoleDefinitionId format based on scope type
                if ($resourceConfig.ScopeType -eq "Subscription") {
                    $roleDefId = "subscriptions/$SubscriptionID/providers/Microsoft.Authorization/roleDefinitions/$ownerRoleDefId"
                }
                elseif ($resourceConfig.ScopeType -eq "ManagementGroup") {
                    # Extract management group name from scope
                    $mgName = $resourceConfig.ScopeId -replace '.*/managementGroups/', ''
                    $roleDefId = "managementGroups/$mgName/providers/Microsoft.Authorization/roleDefinitions/$ownerRoleDefId"
                }
                else {
                    $roleDefId = $ownerRoleDefId
                }
                
                $roleKey = "$($resourceConfig.ScopeType)/$($resourceConfig.ScopeId)/Owner"
                
                # Check if role is still active from previous run
                $history = $state.activationHistory.azure
                $roleHistory = if ($history -is [hashtable]) { $history[$roleKey] } else { $history.PSObject.Properties[$roleKey]?.Value }
                if ($roleHistory -and (Test-RoleStillActive $roleHistory)) {
                    Write-Host "  ⊙ Skipping: Owner (still active)" -ForegroundColor Cyan
                    Write-Log "Skipped Owner at $($resourceConfig.ScopeId) - still active from previous run"
                    $skippedCount++
                    continue
                }
                
                # Skip if too many consecutive failures
                if ($roleHistory -and $roleHistory.consecutiveFailures -ge 3) {
                    Write-Host "  ⊘ Skipping: Owner (too many failures)" -ForegroundColor DarkGray
                    Write-Log "Skipped Owner at $($resourceConfig.ScopeId) - too many consecutive failures ($($roleHistory.consecutiveFailures))" -Level Warning
                    continue
                }
                
                Write-Host "  → Activating: Owner..." -ForegroundColor Yellow -NoNewline
                
                # Use config-specific duration if available, otherwise use script parameter
                $scopeMaxDuration = if ($resourceConfig.MaxDurationHours) { $resourceConfig.MaxDurationHours } else { $AzureDurationHours }
                
                # Use optimal duration if known, otherwise try fallback durations
                $optimalDuration = Get-OptimalDuration -State $state -RoleType 'azure' -RoleKey $roleKey -DefaultDuration $scopeMaxDuration
                $durationsToTry = @($optimalDuration)
                
                # Add fallback durations if not already optimal
                if ($optimalDuration -gt 4 -and $durationsToTry -notcontains 4) {
                    $durationsToTry += 4
                }
                if ($optimalDuration -gt 2 -and $durationsToTry -notcontains 2) {
                    $durationsToTry += 2
                }
                
                Write-Log "Attempting to activate Owner at $($resourceConfig.ScopeId) - trying durations: $($durationsToTry -join ', ')h"
                
                $activated = $false
                foreach ($duration in $durationsToTry) {
                    try {
                        $params = @{
                            Name                      = (New-Guid).Guid
                            Scope                     = $resourceConfig.ScopeId
                            PrincipalId               = $currentUserId
                            RoleDefinitionId          = $roleDefId
                            Justification             = $DefaultJustification
                            ScheduleInfoStartDateTime = Get-Date -Format o
                            ExpirationDuration        = "PT${duration}H"
                            ExpirationType            = 'AfterDuration'
                            RequestType               = 'SelfActivate'
                        }
                        
                        New-AzRoleAssignmentScheduleRequest @params -ErrorAction Stop | Out-Null
                        if ($duration -ne $scopeMaxDuration) {
                            Write-Host " ✓ Activated (${duration}h)" -ForegroundColor Green
                        } else {
                            Write-Host " ✓ Activated" -ForegroundColor Green
                        }
                        Write-Log "Successfully activated Owner at $($resourceConfig.ScopeId) for ${duration}h" -Level Success
                        Update-RoleHistory -State $state -RoleType 'azure' -RoleKey $roleKey -Success $true -Duration $duration
                        $activationCount++
                        $activated = $true
                        break
                    }
                    catch {
                        $errorMessage = $_.Exception.Message
                        
                        # Check if role is already active
                        if ($errorMessage -like "*already exists*" -or $errorMessage -like "*already active*" -or $errorMessage -like "*RoleAssignmentExists*") {
                            Write-Host " ✓ Already Active" -ForegroundColor Cyan
                            Write-Log "Owner at $($resourceConfig.ScopeId) is already active" -Level Info
                            Update-RoleHistory -State $state -RoleType 'azure' -RoleKey $roleKey -Success $true -Duration $duration
                            $activationCount++
                            $activated = $true
                            break
                        }
                        # If ExpirationRule failed and we have more durations to try, continue to next duration
                        elseif ($errorMessage -like "*ExpirationRule*" -and $duration -ne $durationsToTry[-1]) {
                            Write-Log "Duration ${duration}h failed policy, trying shorter duration" -Level Warning
                            continue
                        }
                        else {
                            # Final failure
                            if (-not $activated) {
                                Write-Host " ✗ Failed" -ForegroundColor Red
                                Write-Warning "  Error: $_"
                                Write-Log "Failed to activate Owner at $($resourceConfig.ScopeId): $errorMessage" -Level Error
                                Update-RoleHistory -State $state -RoleType 'azure' -RoleKey $roleKey -Success $false
                                $failureCount++
                            }
                            break
                        }
                    }
                }
            }
        }
    }
    #endregion Azure Resource Roles
    
    # Save state
    Save-State -State $state
    
    # Summary
    $duration = (Get-Date) - $startTime
    Write-Host "`n╔════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║              Summary                       ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Cyan
    
    if ($ListOnly) {
        Write-Host "List-only mode: No activations performed" -ForegroundColor Yellow
        Write-Log "List-only mode completed"
    } else {
        Write-Host "Successful activations: $activationCount" -ForegroundColor Green
        if ($skippedCount -gt 0) {
            Write-Host "Skipped (still active): $skippedCount" -ForegroundColor Cyan
        }
        if ($failureCount -gt 0) {
            Write-Host "Failed activations: $failureCount" -ForegroundColor Red
        }
        
        Write-Log "Activation summary - Success: $activationCount, Skipped: $skippedCount, Failed: $failureCount" -Level Info
    }
    Write-Host "Execution time: $([math]::Round($duration.TotalSeconds, 2)) seconds" -ForegroundColor Gray
    Write-Host ""
    Write-Log "Script completed in $([math]::Round($duration.TotalSeconds, 2)) seconds" -Level Info
    Write-Log "=== PIM Activation Script Completed ===" -Level Info
}
catch {
    Write-Error "Script execution failed: $_"
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    Write-Log "Script execution failed: $_" -Level Error
    Write-Log $_.ScriptStackTrace -Level Error
    exit 1
}
#endregion Main Execution