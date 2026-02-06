#Requires -Version 5.1
<#
.SYNOPSIS
    HYCU for Microsoft 365 - Environment Sizing Assessment Tool
    
.DESCRIPTION
    This script analyzes your Microsoft 365 environment to provide accurate sizing 
    information for HYCU backup and recovery planning. It generates a comprehensive
    HTML report with usage statistics for Exchange, OneDrive, and SharePoint.
    
.PARAMETER AzureAdGroupName
    Optional: Specify an Azure AD group to limit the analysis to specific users
    
.PARAMETER SkipArchiveMailbox
    Skip analysis of archive mailboxes (default: True)
    
.PARAMETER EnableDebug
    Enable detailed debug logging
    
.PARAMETER OutputObject
    Return the sizing object for programmatic use
    
.EXAMPLE
    .\Get-M365SizingInfo-HYCU.ps1
    Generates a full M365 sizing report
    
.EXAMPLE
    .\Get-M365SizingInfo-HYCU.ps1 -AzureAdGroupName "Marketing"
    Generates a report for users in the Marketing group only
    
.NOTES
    Version: 4.4-HYCU
    Requires: Microsoft.Graph.Reports, ExchangeOnlineManagement modules
    Author: HYCU Professional Services
    
.LINK
    https://hycu.com
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]$AzureAdGroupName,
    
    [Parameter()]
    [bool]$SkipArchiveMailbox = $true,
    
    [Parameter()]
    [bool]$EnableDebug = $false,
    
    [Parameter()]
    $OutputObject
)

#region Configuration
$Version = "v4.4-HYCU"
$Period = '180'
$systemTempFolder = [System.IO.Path]::GetTempPath()
$ProgressPreference = 'SilentlyContinue'

# Initialize counters
$ExchangeHTMLTitle = "User"
$ExchangeUserMailboxCount = 0
$ExchangeSharedMailboxCount = 0
#endregion

#region Helper Functions

function Write-Log {
    param([string]$Message)
    if ($EnableDebug) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Output "[$timestamp] $Message"
    }
}

function Get-MgReport {
    <#
    .SYNOPSIS
        Retrieves Microsoft Graph API reports
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$ReportName,

        [Parameter(Mandatory)]
        [ValidateSet("7", "30", "90", "180")]
        [string]$Period
    )
    
    process {
        try {
            $graphApiVersion = if ($ReportName -eq "getMailboxUsageDetail") { "beta" } else { "v1.0" }
            $uri = "https://graph.microsoft.com/$graphApiVersion/reports/$ReportName(period='D$Period')"
            $outputPath = Join-Path $systemTempFolder "$ReportName.csv"
            
            Write-Log "Fetching report: $ReportName"
            Invoke-MgGraphRequest -Uri $uri -OutputFilePath $outputPath
            
            return $outputPath
        }
        catch {
            $errorMessage = $_.Exception.Message
            
            if ($errorMessage -match 'Forbidden') {
                Disconnect-MgGraph
                throw "Authentication failed. The account requires 'Reports Reader' admin role permissions."
            }
            
            throw $_
        }
    }
}

function Measure-AverageGrowth {
    <#
    .SYNOPSIS
        Calculates annual storage growth rate from historical data
    #>
    param (
        [Parameter(Mandatory)]
        [string]$ReportCSV,
        
        [Parameter(Mandatory)]
        [string]$ReportName
    )
    
    try {
        $UsageReport = Import-Csv -Path $ReportCSV | 
            Where-Object { $_.'Is Deleted' -eq 'FALSE' } |
            Sort-Object -Property "Report Date"
        
        if ($ReportName -eq 'getOneDriveUsageStorage') {
            $UsageReport = $UsageReport | Where-Object { $_.'Site Type' -eq 'OneDrive' }
        }
        
        $Record = 1
        $StorageUsage = @()
        
        foreach ($item in $UsageReport) {
            if ($Record -eq 1) {
                $StorageUsed = [decimal]$item."Storage Used (Byte)"
            }
            else {
                $currentStorage = [decimal]$item."Storage Used (Byte)"
                
                if ($StorageUsed -gt 0) {
                    $growthPercent = [math]::Round(((($currentStorage / $StorageUsed) - 1) * 100), 2)
                    $StorageUsage += [PSCustomObject]@{ Growth = $growthPercent }
                }
                else {
                    $StorageUsage += [PSCustomObject]@{ Growth = 0 }
                }
                
                $StorageUsed = $currentStorage
            }
            $Record++
        }
        
        $AverageGrowth = ($StorageUsage | Measure-Object -Property Growth -Average).Average
        # Convert 180-day average to annual growth estimate
        $AnnualGrowth = [math]::Ceiling($AverageGrowth * 2)
        
        Write-Log "Calculated annual growth rate: $AnnualGrowth%"
        return $AnnualGrowth
    }
    catch {
        Write-Log "Error calculating growth rate: $_"
        return 10 # Default conservative estimate
    }
}

function ProcessUsageReport {
    <#
    .SYNOPSIS
        Processes usage reports and populates sizing data
    #>
    param (
        [Parameter(Mandatory)]
        [string]$ReportCSV,
        
        [Parameter(Mandatory)]
        [string]$ReportName,
        
        [Parameter(Mandatory)]
        [string]$Section
    )

    $ReportDetail = Import-Csv -Path $ReportCSV | Where-Object { $_.'Is Deleted' -eq 'FALSE' }
    
    # Apply Azure AD group filtering if specified
    if ($script:AzureAdRequired -and $Section -ne "SharePoint") {
        $FilterByField = switch ($Section) {
            "OneDrive" { "Owner Principal Name" }
            default { "User Principal Name" }
        }
        
        $ReportDetail = $ReportDetail | Where-Object { $_.$FilterByField -in $script:AzureAdGroupMembersByUserPrincipalName }
    }
    
    $SummarizedData = $ReportDetail | Measure-Object -Property 'Storage Used (Byte)' -Sum -Average
    
    # Update sizing object based on section
    switch ($Section) {
        'SharePoint' {
            $script:M365Sizing.$Section.NumberOfSites = $SummarizedData.Count
        }
        'Exchange' {
            $userMailboxes = $ReportDetail | Where-Object { $_.'Recipient Type' -eq 'User' }
            $sharedMailboxes = $ReportDetail | Where-Object { $_.'Recipient Type' -eq 'Shared' }
            
            if ($sharedMailboxes.Count -ge $userMailboxes.Count) {
                $script:M365Sizing.$Section.NumberOfUsers = $sharedMailboxes.Count
                $script:ExchangeHTMLTitle = "Mailboxes"
                $script:ExchangeSharedMailboxCount = $sharedMailboxes.Count
            }
            else {
                $script:M365Sizing.$Section.NumberOfUsers = $userMailboxes.Count
                $script:ExchangeUserMailboxCount = $userMailboxes.Count
            }
        }
        default {
            $script:M365Sizing.$Section.NumberOfUsers = $SummarizedData.Count
        }
    }

    $script:M365Sizing.$Section.TotalSizeGB = [math]::Round(($SummarizedData.Sum / 1GB), 2)
    $script:M365Sizing.$Section.SizePerUserGB = [math]::Round(($SummarizedData.Average / 1GB), 2)
}

function Start-SleepWithProgress {
    <#
    .SYNOPSIS
        Sleep with a progress bar for better UX
    #>
    param(
        [int]$SleepTime = 15,
        [int]$Milliseconds = 0
    )
    
    if ($Milliseconds -gt 0) {
        Start-Sleep -Milliseconds $Milliseconds
        return
    }
    
    for ($i = 0; $i -lt $SleepTime; $i++) {
        $percent = [math]::Round(($i / $SleepTime) * 100)
        Write-Progress -Activity "Processing" -Status "Please wait... $percent% complete" -PercentComplete $percent
        Start-Sleep -Seconds 1
    }
    
    Write-Progress -Completed -Activity "Processing"
}

function New-CleanO365Session {
    <#
    .SYNOPSIS
        Establishes a clean Exchange Online session
    #>
    Write-Log "Cleaning up existing PowerShell sessions"
    Get-PSSession | Remove-PSSession -Confirm:$false
    
    [System.GC]::Collect()
    
    Write-Log "Waiting for session cleanup (15s)"
    Start-SleepWithProgress -SleepTime 15
    
    $Error.Clear()
    
    Write-Log "Connecting to Exchange Online"
    try {
        Connect-ExchangeOnline -UserPrincipalName $script:UserPrincipalName -ShowBanner:$false -ErrorAction Stop
        $script:ErrorCount = 0
        $script:SessionStartTime = Get-Date
        Write-Log "Successfully connected to Exchange Online"
    }
    catch {
        Write-Log "ERROR: Failed to establish Exchange Online session - $_"
        $script:ErrorCount++
        
        if ($script:ErrorCount -gt 3) {
            throw "Failed to establish Exchange Online session after multiple attempts. Please check credentials and network connectivity."
        }
        
        Write-Log "Retrying connection in 60 seconds..."
        Start-SleepWithProgress -SleepTime 60
        New-CleanO365Session
    }
}

function Test-O365Session {
    <#
    .SYNOPSIS
        Validates Exchange Online session health
    #>
    $ObjectTime = Get-Date
    $SessionInfo = Get-PSSession
    
    if ($null -eq $SessionInfo) {
        Write-Log "ERROR: No active session found, reconnecting..."
        New-CleanO365Session
        return
    }
    
    if ($SessionInfo.State -ne "Opened") {
        Write-Log "ERROR: Session not in Open state, reconnecting..."
        New-CleanO365Session
        return
    }
    
    $ResetSeconds = 870 # 14.5 minutes
    
    if (($ObjectTime - $script:SessionStartTime).TotalSeconds -gt $ResetSeconds) {
        Write-Log "Session exceeded $ResetSeconds seconds, rebuilding connection"
        
        $DelaySeconds = [math]::Max(0, ((($ResetSeconds * 0.5) / 2) - 15))
        
        if ($DelaySeconds -gt 0) {
            Write-Log "Throttle recovery delay: $DelaySeconds seconds"
            Start-SleepWithProgress -SleepTime $DelaySeconds
        }
        
        New-CleanO365Session
    }
}

#endregion

#region Module Validation

Write-Output "`n==================================================================="
Write-Output "  HYCU for Microsoft 365 - Sizing Assessment Tool ($Version)"
Write-Output "===================================================================`n"

# Validate Microsoft.Graph.Reports module
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Reports)) {
    throw @"
Required module not found: Microsoft.Graph.Reports

Please install the module using:
    Install-Module Microsoft.Graph.Reports -Scope CurrentUser

For more information, visit: https://docs.microsoft.com/powershell/microsoftgraph/
"@
}

# Validate ExchangeOnlineManagement module
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw @"
Required module not found: ExchangeOnlineManagement

Please install the module using:
    Install-Module ExchangeOnlineManagement -Scope CurrentUser

For more information, visit: https://docs.microsoft.com/powershell/exchange/exchange-online-powershell
"@
}

Write-Output "[OK] All required PowerShell modules are installed`n"

#endregion

#region Initialize Data Structure

$M365Sizing = [ordered]@{
    Exchange = [ordered]@{
        NumberOfUsers = 0
        TotalSizeGB = 0
        SizePerUserGB = 0
        AverageGrowthPercent = 0
    }
    OneDrive = [ordered]@{
        NumberOfUsers = 0
        TotalSizeGB = 0
        SizePerUserGB = 0
        AverageGrowthPercent = 0
    }
    SharePoint = [ordered]@{
        NumberOfSites = 0
        TotalSizeGB = 0
        SizePerUserGB = 0
        AverageGrowthPercent = 0
    }
}

#endregion

#region Azure AD Group Processing

$AzureAdRequired = -not [string]::IsNullOrWhiteSpace($AzureAdGroupName)

if ($AzureAdRequired) {
    Write-Output "===================================================================="
    Write-Output "  Filtering by Azure AD Group: $AzureAdGroupName"
    Write-Output "====================================================================`n"
    
    try {
        Write-Output "[->] Connecting to Microsoft Graph..."
        Connect-MgGraph -Scopes "Group.Read.All", "GroupMember.Read.All", "User.Read.All", "Reports.Read.All" -NoWelcome
        
        Write-Output "[->] Retrieving group members..."
        $AzureAdGroup = Get-MgGroup -Filter "displayName eq '$AzureAdGroupName'" -ErrorAction Stop
        
        if ($null -eq $AzureAdGroup) {
            throw "Azure AD group '$AzureAdGroupName' not found. Please verify the group name."
        }
        
        $AzureAdGroupMembers = Get-MgGroupMember -GroupId $AzureAdGroup.Id -All
        $AzureAdGroupMembersByUserPrincipalName = @()
        
        foreach ($member in $AzureAdGroupMembers) {
            $user = Get-MgUser -UserId $member.Id -Property UserPrincipalName -ErrorAction SilentlyContinue
            if ($user.UserPrincipalName) {
                $AzureAdGroupMembersByUserPrincipalName += $user.UserPrincipalName
            }
        }
        
        Write-Output "[OK] Found $($AzureAdGroupMembersByUserPrincipalName.Count) users in group`n"
    }
    catch {
        throw "Failed to process Azure AD group: $_"
    }
}
else {
    Write-Output "===================================================================="
    Write-Output "  Analyzing Full M365 Environment"
    Write-Output "====================================================================`n"
    
    Write-Output "[->] Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "Reports.Read.All" -NoWelcome
    Write-Output "[OK] Connected to Microsoft Graph`n"
}

#endregion

#region Exchange Online Processing

Write-Output "===================================================================="
Write-Output "  Processing Exchange Online Data"
Write-Output "====================================================================`n"

try {
    Write-Output "[->] Retrieving Exchange mailbox usage report (last $Period days)..."
    $ExchangeReport = Get-MgReport -ReportName "getMailboxUsageDetail" -Period $Period
    Write-Output "[OK] Exchange usage report retrieved"
    
    Write-Output "[->] Retrieving Exchange storage report..."
    $ExchangeStorageReport = Get-MgReport -ReportName "getMailboxUsageStorage" -Period $Period
    Write-Output "[OK] Exchange storage report retrieved"
    
    Write-Output "[->] Processing Exchange data..."
    ProcessUsageReport -ReportCSV $ExchangeReport -ReportName "getMailboxUsageDetail" -Section "Exchange"
    
    Write-Output "[->] Calculating growth trends..."
    $M365Sizing.Exchange.AverageGrowthPercent = Measure-AverageGrowth -ReportCSV $ExchangeStorageReport -ReportName "getMailboxUsageStorage"
    
    Write-Output "[OK] Exchange analysis complete"
    Write-Output "    - Mailboxes: $($M365Sizing.Exchange.NumberOfUsers)"
    Write-Output "    - Total Storage: $($M365Sizing.Exchange.TotalSizeGB) GB"
    Write-Output "    - Annual Growth: $($M365Sizing.Exchange.AverageGrowthPercent) percent`n"
}
catch {
    Write-Output "[WARN] Warning: Could not retrieve Exchange data - $_`n"
}

#endregion

#region OneDrive Processing

Write-Output "===================================================================="
Write-Output "  Processing OneDrive for Business Data"
Write-Output "====================================================================`n"

try {
    Write-Output "[->] Retrieving OneDrive usage report (last $Period days)..."
    $OneDriveReport = Get-MgReport -ReportName "getOneDriveUsageAccountDetail" -Period $Period
    Write-Output "[OK] OneDrive usage report retrieved"
    
    Write-Output "[->] Retrieving OneDrive storage report..."
    $OneDriveStorageReport = Get-MgReport -ReportName "getOneDriveUsageStorage" -Period $Period
    Write-Output "[OK] OneDrive storage report retrieved"
    
    Write-Output "[->] Processing OneDrive data..."
    ProcessUsageReport -ReportCSV $OneDriveReport -ReportName "getOneDriveUsageAccountDetail" -Section "OneDrive"
    
    Write-Output "[->] Calculating growth trends..."
    $M365Sizing.OneDrive.AverageGrowthPercent = Measure-AverageGrowth -ReportCSV $OneDriveStorageReport -ReportName "getOneDriveUsageStorage"
    
    Write-Output "[OK] OneDrive analysis complete"
    Write-Output "    - Active Users: $($M365Sizing.OneDrive.NumberOfUsers)"
    Write-Output "    - Total Storage: $($M365Sizing.OneDrive.TotalSizeGB) GB"
    Write-Output "    - Annual Growth: $($M365Sizing.OneDrive.AverageGrowthPercent) percent`n"
}
catch {
    Write-Output "[WARN] Warning: Could not retrieve OneDrive data - $_`n"
}

#endregion

#region SharePoint Processing

Write-Output "===================================================================="
Write-Output "  Processing SharePoint Online Data"
Write-Output "====================================================================`n"

try {
    Write-Output "[->] Retrieving SharePoint site usage report (last $Period days)..."
    $SharePointReport = Get-MgReport -ReportName "getSharePointSiteUsageDetail" -Period $Period
    Write-Output "[OK] SharePoint usage report retrieved"
    
    Write-Output "[->] Retrieving SharePoint storage report..."
    $SharePointStorageReport = Get-MgReport -ReportName "getSharePointSiteUsageStorage" -Period $Period
    Write-Output "[OK] SharePoint storage report retrieved"
    
    Write-Output "[->] Processing SharePoint data..."
    ProcessUsageReport -ReportCSV $SharePointReport -ReportName "getSharePointSiteUsageDetail" -Section "SharePoint"
    
    Write-Output "[->] Calculating growth trends..."
    $M365Sizing.SharePoint.AverageGrowthPercent = Measure-AverageGrowth -ReportCSV $SharePointStorageReport -ReportName "getSharePointSiteUsageStorage"
    
    Write-Output "[OK] SharePoint analysis complete"
    Write-Output "    - Active Sites: $($M365Sizing.SharePoint.NumberOfSites)"
    Write-Output "    - Total Storage: $($M365Sizing.SharePoint.TotalSizeGB) GB"
    Write-Output "    - Annual Growth: $($M365Sizing.SharePoint.AverageGrowthPercent) percent`n"
}
catch {
    Write-Output "[WARN] Warning: Could not retrieve SharePoint data - $_`n"
}

#endregion

#region Archive Mailbox Processing (Optional)

if (-not $SkipArchiveMailbox) {
    Write-Output "===================================================================="
    Write-Output "  Processing Archive Mailboxes"
    Write-Output "====================================================================`n"
    
    try {
        Write-Output "[->] Connecting to Exchange Online..."
        
        $UserPrincipalName = (Get-MgContext).Account
        $script:UserPrincipalName = $UserPrincipalName
        $script:ErrorCount = 0
        $script:SessionStartTime = Get-Date
        
        New-CleanO365Session
        
        Write-Output "[->] Retrieving archive mailbox data..."
        
        if ($AzureAdRequired) {
            $Mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties ArchiveStatus, ArchiveDatabase |
                Where-Object { 
                    $_.UserPrincipalName -in $AzureAdGroupMembersByUserPrincipalName -and
                    $_.ArchiveStatus -eq "Active"
                }
        }
        else {
            $Mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties ArchiveStatus, ArchiveDatabase |
                Where-Object { $_.ArchiveStatus -eq "Active" }
        }
        
        Write-Output "[->] Found $($Mailboxes.Count) archive mailboxes to process"
        
        $ArchiveStats = @{
            TotalArchives = 0
            TotalSizeGB = 0
        }
        
        $processedCount = 0
        
        foreach ($mailbox in $Mailboxes) {
            Test-O365Session
            
            try {
                $stats = Get-EXOMailboxStatistics -Identity $mailbox.UserPrincipalName -Archive -ErrorAction Stop
                
                if ($stats.TotalItemSize) {
                    $sizeInBytes = [regex]::Match($stats.TotalItemSize.ToString(), "([0-9,]+) bytes").Groups[1].Value -replace ',', ''
                    $ArchiveStats.TotalSizeGB += [decimal]$sizeInBytes / 1GB
                    $ArchiveStats.TotalArchives++
                }
                
                $processedCount++
                
                if ($processedCount % 50 -eq 0) {
                    $percentComplete = [math]::Round(($processedCount / $Mailboxes.Count) * 100)
                    Write-Output "[->] Progress: $processedCount/$($Mailboxes.Count) ($percentComplete percent)"
                }
            }
            catch {
                Write-Log "Warning: Could not retrieve archive stats for $($mailbox.UserPrincipalName) - $_"
            }
        }
        
        $ArchiveStats.TotalSizeGB = [math]::Round($ArchiveStats.TotalSizeGB, 2)
        
        Write-Output "[OK] Archive mailbox analysis complete"
        Write-Output "    - Active Archives: $($ArchiveStats.TotalArchives)"
        Write-Output "    - Total Archive Storage: $($ArchiveStats.TotalSizeGB) GB`n"
        
        # Add to sizing object
        $M365Sizing.Exchange.ArchiveMailboxes = $ArchiveStats.TotalArchives
        $M365Sizing.Exchange.ArchiveSizeGB = $ArchiveStats.TotalSizeGB
        
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
    catch {
        Write-Output "[WARN] Warning: Archive mailbox processing failed - $_`n"
    }
}

#endregion

#region HTML Report Generation

Write-Output "===================================================================="
Write-Output "  Generating HYCU Sizing Report"
Write-Output "====================================================================`n"

$reportDate = Get-Date -Format "MMMM dd, yyyy 'at' HH:mm"
$totalUsers = [math]::Max($M365Sizing.Exchange.NumberOfUsers, $M365Sizing.OneDrive.NumberOfUsers)
$totalStorage = $M365Sizing.Exchange.TotalSizeGB + $M365Sizing.OneDrive.TotalSizeGB + $M365Sizing.SharePoint.TotalSizeGB
$totalStorage = [math]::Round($totalStorage, 2)

$HTML_CODE = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>HYCU for Microsoft 365 - Sizing Report</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            line-height: 1.6;
            color: #2d3748;
            background: linear-gradient(135deg, #6D28D9 0%, #4C1D95 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 16px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.5);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #6D28D9 0%, #4C1D95 100%);
            color: white;
            padding: 48px 48px 32px;
            position: relative;
            overflow: hidden;
        }
        
        .header::before {
            content: '';
            position: absolute;
            top: -50%;
            right: -10%;
            width: 600px;
            height: 600px;
            background: rgba(255, 255, 255, 0.05);
            border-radius: 50%;
        }
        
        .header-content {
            position: relative;
            z-index: 1;
        }
        
        .logo {
            font-size: 36px;
            font-weight: 700;
            letter-spacing: -0.5px;
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            gap: 12px;
        }
        
        .logo-icon {
            width: 48px;
            height: 48px;
            background: white;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 28px;
            font-weight: 900;
            color: #6D28D9;
        }
        
        .subtitle {
            font-size: 18px;
            font-weight: 400;
            opacity: 0.95;
            margin-bottom: 24px;
        }
        
        .report-meta {
            display: flex;
            gap: 32px;
            font-size: 14px;
            opacity: 0.9;
            flex-wrap: wrap;
        }
        
        .meta-item {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .meta-icon {
            font-size: 16px;
        }
        
        .content {
            padding: 48px;
        }
        
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 24px;
            margin-bottom: 48px;
        }
        
        .summary-card {
            background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
            padding: 28px;
            border-radius: 12px;
            border: 1px solid #e2e8f0;
            transition: transform 0.2s, box-shadow 0.2s;
        }
        
        .summary-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 12px 24px rgba(0, 0, 0, 0.1);
        }
        
        .summary-label {
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: #64748b;
            font-weight: 600;
            margin-bottom: 8px;
        }
        
        .summary-value {
            font-size: 36px;
            font-weight: 700;
            color: #6D28D9;
            line-height: 1;
        }
        
        .summary-unit {
            font-size: 16px;
            font-weight: 400;
            color: #64748b;
            margin-left: 4px;
        }
        
        .section {
            margin-bottom: 48px;
        }
        
        .section-header {
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 24px;
            padding-bottom: 16px;
            border-bottom: 2px solid #e2e8f0;
        }
        
        .section-icon {
            width: 40px;
            height: 40px;
            background: linear-gradient(135deg, #7C3AED 0%, #6D28D9 100%);
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-size: 20px;
        }
        
        .section-title {
            font-size: 24px;
            font-weight: 700;
            color: #1e293b;
        }
        
        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 24px;
        }
        
        .metric-card {
            background: white;
            border: 2px solid #e2e8f0;
            border-radius: 10px;
            padding: 20px;
            transition: border-color 0.2s;
        }
        
        .metric-card:hover {
            border-color: #7C3AED;
        }
        
        .metric-label {
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: #64748b;
            font-weight: 600;
            margin-bottom: 8px;
        }
        
        .metric-value {
            font-size: 28px;
            font-weight: 700;
            color: #6D28D9;
        }
        
        .metric-unit {
            font-size: 14px;
            font-weight: 400;
            color: #64748b;
            margin-left: 4px;
        }
        
        .growth-badge {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            background: #DDD6FE;
            color: #6D28D9;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            margin-top: 8px;
        }
        
        .footer {
            background: #f8fafc;
            padding: 32px 48px;
            border-top: 1px solid #e2e8f0;
            text-align: center;
        }
        
        .footer-content {
            max-width: 800px;
            margin: 0 auto;
        }
        
        .footer-title {
            font-size: 20px;
            font-weight: 700;
            color: #1e293b;
            margin-bottom: 12px;
        }
        
        .footer-text {
            font-size: 14px;
            color: #64748b;
            line-height: 1.8;
            margin-bottom: 20px;
        }
        
        .cta-button {
            display: inline-block;
            background: linear-gradient(135deg, #7C3AED 0%, #6D28D9 100%);
            color: white;
            padding: 14px 32px;
            border-radius: 8px;
            text-decoration: none;
            font-weight: 600;
            font-size: 15px;
            transition: transform 0.2s, box-shadow 0.2s;
            box-shadow: 0 4px 12px rgba(109, 40, 217, 0.5);
        }
        
        .cta-button:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(109, 40, 217, 0.5);
        }
        
        .divider {
            height: 1px;
            background: linear-gradient(to right, transparent, #e2e8f0, transparent);
            margin: 32px 0;
        }
        
        @media print {
            body {
                background: white;
                padding: 0;
            }
            
            .container {
                box-shadow: none;
                border-radius: 0;
            }
            
            .summary-card, .metric-card {
                break-inside: avoid;
            }
        }
        
        @media (max-width: 768px) {
            .content {
                padding: 32px 24px;
            }
            
            .header {
                padding: 32px 24px 24px;
            }
            
            .logo {
                font-size: 28px;
            }
            
            .summary-value {
                font-size: 28px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-content">
                <div class="logo">
                    <div class="logo-icon">H</div>
                    <span>HYCU for Microsoft 365</span>
                </div>
                <div class="subtitle">Environment Sizing Assessment Report</div>
                <div class="report-meta">
                    <div class="meta-item">
                        <span class="meta-icon">📅</span>
                        <span>Generated: $reportDate</span>
                    </div>
                    <div class="meta-item">
                        <span class="meta-icon">📊</span>
                        <span>Analysis Period: $Period days</span>
                    </div>
                    <div class="meta-item">
                        <span class="meta-icon">🔧</span>
                        <span>Version: $Version</span>
                    </div>
                </div>
            </div>
        </div>
        
        <div class="content">
            <div class="summary-grid">
                <div class="summary-card">
                    <div class="summary-label">Total Users</div>
                    <div class="summary-value">$totalUsers</div>
                </div>
                <div class="summary-card">
                    <div class="summary-label">Total Storage</div>
                    <div class="summary-value">$totalStorage<span class="summary-unit">GB</span></div>
                </div>
                <div class="summary-card">
                    <div class="summary-label">Workloads Protected</div>
                    <div class="summary-value">3</div>
                </div>
            </div>
            
            <div class="section">
                <div class="section-header">
                    <div class="section-icon">📧</div>
                    <h2 class="section-title">Exchange Online</h2>
                </div>
                <div class="metrics-grid">
                    <div class="metric-card">
                        <div class="metric-label">$ExchangeHTMLTitle Count</div>
                        <div class="metric-value">$($M365Sizing.Exchange.NumberOfUsers)</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Total Storage</div>
                        <div class="metric-value">$($M365Sizing.Exchange.TotalSizeGB)<span class="metric-unit">GB</span></div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Avg per User</div>
                        <div class="metric-value">$($M365Sizing.Exchange.SizePerUserGB)<span class="metric-unit">GB</span></div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Annual Growth</div>
                        <div class="metric-value">$($M365Sizing.Exchange.AverageGrowthPercent)<span class="metric-unit">%</span></div>
                        <div class="growth-badge">📈 Year-over-Year</div>
                    </div>
                </div>
            </div>
            
            <div class="divider"></div>
            
            <div class="section">
                <div class="section-header">
                    <div class="section-icon">☁️</div>
                    <h2 class="section-title">OneDrive for Business</h2>
                </div>
                <div class="metrics-grid">
                    <div class="metric-card">
                        <div class="metric-label">Active Users</div>
                        <div class="metric-value">$($M365Sizing.OneDrive.NumberOfUsers)</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Total Storage</div>
                        <div class="metric-value">$($M365Sizing.OneDrive.TotalSizeGB)<span class="metric-unit">GB</span></div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Avg per User</div>
                        <div class="metric-value">$($M365Sizing.OneDrive.SizePerUserGB)<span class="metric-unit">GB</span></div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Annual Growth</div>
                        <div class="metric-value">$($M365Sizing.OneDrive.AverageGrowthPercent)<span class="metric-unit">%</span></div>
                        <div class="growth-badge">📈 Year-over-Year</div>
                    </div>
                </div>
            </div>
            
            <div class="divider"></div>
            
            <div class="section">
                <div class="section-header">
                    <div class="section-icon">🌐</div>
                    <h2 class="section-title">SharePoint Online</h2>
                </div>
                <div class="metrics-grid">
                    <div class="metric-card">
                        <div class="metric-label">Active Sites</div>
                        <div class="metric-value">$($M365Sizing.SharePoint.NumberOfSites)</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Total Storage</div>
                        <div class="metric-value">$($M365Sizing.SharePoint.TotalSizeGB)<span class="metric-unit">GB</span></div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Avg per Site</div>
                        <div class="metric-value">$($M365Sizing.SharePoint.SizePerUserGB)<span class="metric-unit">GB</span></div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Annual Growth</div>
                        <div class="metric-value">$($M365Sizing.SharePoint.AverageGrowthPercent)<span class="metric-unit">%</span></div>
                        <div class="growth-badge">📈 Year-over-Year</div>
                    </div>
                </div>
            </div>
        </div>
        
        <div class="footer">
            <div class="footer-content">
                <div class="footer-title">Ready to Protect Your Microsoft 365 Environment?</div>
                <div class="footer-text">
                    HYCU for Microsoft 365 provides enterprise-grade backup and recovery with industry-leading RPOs and RTOs. 
                    Our SaaS-native solution ensures your data is always protected and instantly recoverable.
                </div>
                <a href="https://hycu.com" class="cta-button" target="_blank">Learn More at hycu.com</a>
            </div>
        </div>
    </div>
    
    <!-- Debug Information (Hidden, viewable in HTML source) -->
    <!--
    HYCU M365 Sizing Report - Debug Information
    ===========================================
    
    Exchange Details:
    - User Mailboxes: $ExchangeUserMailboxCount
    - Shared Mailboxes: $ExchangeSharedMailboxCount
    - Total: $($M365Sizing.Exchange.NumberOfUsers)
    
    Report Generated: $reportDate
    Script Version: $Version
    Analysis Period: $Period days
    
    For technical support, visit: https://support.hycu.com
    -->
</body>
</html>
"@

# Write the HTML report
$reportPath = Join-Path (Get-Location) "HYCU-M365-Sizing-Report.html"
$HTML_CODE | Out-File -FilePath $reportPath -Encoding UTF8 -Force

Write-Output "[OK] Report generated successfully!"
Write-Output "`n==================================================================="
Write-Output "  -- Report Location"
Write-Output "==================================================================="
Write-Output "`n    $reportPath`n"

#endregion

#region Cleanup and Return

if ($OutputObject) {
    return $M365Sizing
}

#endregion
