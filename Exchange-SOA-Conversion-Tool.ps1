#Requires -Version 5.1

<#
        .SYNOPSIS
        Exchange SOA Conversion tool (Cloud / On-Prem Source of Authority)

        .DESCRIPTION
        GUI tool to manage Exchange mailbox Source of Authority (SOA) conversion between cloud-managed and on-premises-managed by toggling the mailbox property `IsExchangeCloudManaged`.

        The tool connects to Exchange Online PowerShell, lists directory-synced mailboxes (`IsDirSynced = True`), and supports converting one or multiple selected users in batch.
        A CSV backup of mail-related attributes is taken before every conversion, and a full backup of all directory-synced mailboxes can be exported on demand.

        This tool is intended to support the approach described in:
        https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

        .EXAMPLE
        C:\PS> .\Exchange-SOA-Conversion-Tool.ps1

        .NOTES
        Version: 1.04

        .COPYRIGHT
        MIT License, feel free to distribute and use as you like, please leave author information.

       .LINK
        https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

        BLOG: http://www.hcandersen.dk
        Twitter: @dk_hcandersen
        LinkedIn: https://www.linkedin.com/in/hanschrandersen/

        .DISCLAIMER
        This script is provided AS-IS, with no warranty - Use at own risk.
    #>

$script:Version = "1.04"


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:ScriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:LogFile = Join-Path $script:ScriptPath "ExchangeSOAConversion_$(Get-Date -Format 'yyyyMMdd_HHmm').log"
$script:AllUsers = @()
$script:ViewUsers = @()
$script:UserIndex = @{}
$script:CurrentPage = 1
$script:PageSize = 100
$script:SortColumn = $null
$script:SortAscending = $true
$script:IsConnected = $false
$script:IsBusy = $false
$script:CancelRequested = $false
$script:CloseRequested = $false
$script:BackupFolder = $null

$script:SoaBackupAttributes = @(
    'WindowsEmailAddress'
    'EmailAddresses'
    'CustomAttribute1','CustomAttribute2','CustomAttribute3','CustomAttribute4','CustomAttribute5',
    'CustomAttribute6','CustomAttribute7','CustomAttribute8','CustomAttribute9','CustomAttribute10',
    'CustomAttribute11','CustomAttribute12','CustomAttribute13','CustomAttribute14','CustomAttribute15',
    'ExtensionCustomAttribute1','ExtensionCustomAttribute2','ExtensionCustomAttribute3','ExtensionCustomAttribute4','ExtensionCustomAttribute5',
    'HiddenFromAddressListsEnabled','ForwardingAddress','ForwardingSmtpAddress','DeliverToMailboxAndForward',
    'GrantSendOnBehalfTo','AcceptMessagesOnlyFrom','AcceptMessagesOnlyFromDLMembers','AcceptMessagesOnlyFromSendersOrMembers',
    'RejectMessagesFrom','RejectMessagesFromDLMembers','RejectMessagesFromSendersOrMembers',
    'RequireSenderAuthenticationEnabled','ModerationEnabled','ModeratedBy','SendModerationNotifications',
    'BypassModerationFromSendersOrMembers','LitigationHoldEnabled','LitigationHoldDate','LitigationHoldOwner',
    'RetentionHoldEnabled','StartDateForRetentionHold','EndDateForRetentionHold','RetentionComment','RetentionUrl',
    'SingleItemRecoveryEnabled','AuditEnabled','AuditLogAgeLimit','AuditAdmin','AuditDelegate','AuditOwner',
    'ArchiveName','ArchiveGuid','LegacyExchangeDN','MailTipTranslations','ResourceCapacity','ResourceCustom','ImmutableId'
)

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [ValidateSet('INFO','WARNING','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    Add-Content -Path $script:LogFile -Value $logEntry -Encoding UTF8

    Write-Host $logEntry
}

function Format-SoaValue {
    param(
        [Parameter(Mandatory=$false)]
        $Value
    )

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [datetime]) {
        return $Value.ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
    }

    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = @()
        foreach ($item in $Value) {
            if ($null -ne $item) {
                $items += $item.ToString()
            }
        }
        return ($items -join '|')
    }

    return $Value.ToString()
}

function ConvertTo-SoaBackupRecord {
    param(
        [Parameter(Mandatory=$true)]
        $Mailbox,

        [Parameter(Mandatory=$true)]
        [string]$Operation
    )

    $record = [ordered]@{
        BackupTimestamp          = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
        Operation                = $Operation
        DisplayName              = Format-SoaValue $Mailbox.DisplayName
        UserPrincipalName        = Format-SoaValue $Mailbox.UserPrincipalName
        Alias                    = Format-SoaValue $Mailbox.Alias
        PrimarySmtpAddress       = Format-SoaValue $Mailbox.PrimarySmtpAddress
        ExternalDirectoryObjectId = Format-SoaValue $Mailbox.ExternalDirectoryObjectId
        Guid                     = Format-SoaValue $Mailbox.Guid
        ExchangeGuid             = Format-SoaValue $Mailbox.ExchangeGuid
        IsDirSynced              = Format-SoaValue $Mailbox.IsDirSynced
        IsExchangeCloudManaged   = Format-SoaValue $Mailbox.IsExchangeCloudManaged
        RecipientType            = Format-SoaValue $Mailbox.RecipientType
        RecipientTypeDetails     = Format-SoaValue $Mailbox.RecipientTypeDetails
    }

    foreach ($attrName in $script:SoaBackupAttributes) {
        $prop = $Mailbox.PSObject.Properties[$attrName]
        $record[$attrName] = if ($prop) { Format-SoaValue $prop.Value } else { '' }
    }

    return [pscustomobject]$record
}

function Export-SoaBackupRecord {
    param(
        [Parameter(Mandatory=$true)]
        $Record,

        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    $Record | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Append
}

function Select-SoaUsers {
    param(
        [Parameter(Mandatory=$false)]
        $Users,

        [Parameter(Mandatory=$false)]
        [string]$SearchText
    )

    if ($null -eq $Users) {
        return @()
    }

    if ([string]::IsNullOrWhiteSpace($SearchText)) {
        return @($Users)
    }

    $term = $SearchText.ToLower()
    return @($Users | Where-Object { $_.SearchText -and $_.SearchText.Contains($term) })
}

function Get-SoaBackupFolder {
    if ($script:BackupFolder -and (Test-Path $script:BackupFolder)) {
        return $script:BackupFolder
    }

    $folder = Join-Path $script:ScriptPath 'Backups'

    if (-not (Test-Path $folder)) {
        try {
            New-Item -Path $folder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Log "Failed to create backup folder '$folder': $($_.Exception.Message). Falling back to Documents." -Level WARNING
            $folder = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'ExchangeSOABackups'
            New-Item -Path $folder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
    }

    $script:BackupFolder = $folder
    return $folder
}

function Set-BusyState {
    param(
        [Parameter(Mandatory=$true)]
        [bool]$Busy,

        [Parameter(Mandatory=$false)]
        [string]$Message = ''
    )

    $script:IsBusy = $Busy

    if ($Busy) {
        $script:CancelRequested = $false
        $buttonCancel.Enabled = $true
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $buttonConnect.Enabled = $false
        $buttonRefresh.Enabled = $false
        $buttonDisconnect.Enabled = $false
        $buttonBackupAll.Enabled = $false
        $buttonConvertToCloud.Enabled = $false
        $buttonConvertToOnPrem.Enabled = $false
        $buttonOpenLog.Enabled = $false
        $dataGridView.Enabled = $false
        $textBoxSearch.Enabled = $false
        $buttonPrevPage.Enabled = $false
        $buttonNextPage.Enabled = $false
        $buttonCancel.Visible = $true
        $progressBar.Visible = $true
        if ($Message) {
            $statusLabel.Text = $Message
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    else {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $buttonConnect.Enabled = (-not $script:IsConnected)
        $buttonRefresh.Enabled = $script:IsConnected
        $buttonDisconnect.Enabled = $script:IsConnected
        $buttonBackupAll.Enabled = $script:IsConnected
        $buttonConvertToCloud.Enabled = $script:IsConnected
        $buttonConvertToOnPrem.Enabled = $script:IsConnected
        $buttonOpenLog.Enabled = $true
        $dataGridView.Enabled = $true
        $textBoxSearch.Enabled = $true
        $buttonCancel.Visible = $false
        $progressBar.Visible = $false
        Update-UserGrid
    }
}

function Update-UserGrid {
    $totalUsers = $script:ViewUsers.Count
    $totalPages = [Math]::Ceiling($totalUsers / $script:PageSize)

    if ($script:CurrentPage -gt $totalPages -and $totalPages -gt 0) {
        $script:CurrentPage = $totalPages
    }

    if ($script:CurrentPage -lt 1) {
        $script:CurrentPage = 1
    }

    $startIndex = ($script:CurrentPage - 1) * $script:PageSize
    $endIndex = [Math]::Min($startIndex + $script:PageSize - 1, $totalUsers - 1)

    $dataGridView.Rows.Clear()

    if ($totalUsers -gt 0) {
        for ($i = $startIndex; $i -le $endIndex; $i++) {
            $user = $script:ViewUsers[$i]
            $cloudManagedStatus = if ($user.IsExchangeCloudManaged) { "True" } else { "False" }
            $dataGridView.Rows.Add($user.DisplayName, $user.PrimarySmtpAddress, $user.RecipientTypeDetails, $cloudManagedStatus, $user.Guid) | Out-Null
        }

        $showingCount = $endIndex - $startIndex + 1
        $filteredNote = ''
        if ($script:ViewUsers.Count -ne $script:AllUsers.Count) {
            $filteredNote = " (filtered from $($script:AllUsers.Count))"
        }
        $pageInfo.Text = "Page $($script:CurrentPage) of $totalPages - Showing $showingCount of $totalUsers users$filteredNote"
    } else {
        if ($script:AllUsers.Count -gt 0) {
            $pageInfo.Text = "No users match the search (filtered from $($script:AllUsers.Count))"
        } else {
            $pageInfo.Text = "No users to display"
        }
    }

    if (-not $script:IsBusy) {
        $buttonPrevPage.Enabled = ($script:CurrentPage -gt 1)
        $buttonNextPage.Enabled = ($script:CurrentPage -lt $totalPages)
    }
}

function Sort-Users {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ColumnName
    )

    if ($script:AllUsers.Count -eq 0) {
        return
    }

    if ($script:SortColumn -eq $ColumnName) {
        $script:SortAscending = -not $script:SortAscending
    } else {
        $script:SortColumn = $ColumnName
        $script:SortAscending = $true
    }

    $sortProperty = switch ($ColumnName) {
        "DisplayName"            { 'DisplayName' }
        "Email"                  { 'PrimarySmtpAddress' }
        "Type"                   { 'RecipientTypeDetails' }
        "IsExchangeCloudManaged" { 'IsExchangeCloudManaged' }
        default                  { $null }
    }

    if ($sortProperty) {
        if ($script:SortAscending) {
            $script:AllUsers = @($script:AllUsers | Sort-Object -Property $sortProperty)
        } else {
            $script:AllUsers = @($script:AllUsers | Sort-Object -Property $sortProperty -Descending)
        }
    }

    $script:ViewUsers = @(Select-SoaUsers -Users $script:AllUsers -SearchText $textBoxSearch.Text)
    $script:CurrentPage = 1
    Update-UserGrid

    $sortDirection = if ($script:SortAscending) { 'Ascending' } else { 'Descending' }
    Write-Log "Sorted users by $ColumnName ($sortDirection)"
}

function Initialize-ExchangeOnlineModule {
    Write-Log "Checking for Exchange Online Management module..."

    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement

    if (-not $module) {
        Write-Log "Exchange Online Management module not found. Attempting to install..." -Level WARNING

        try {
            [System.Windows.Forms.MessageBox]::Show(
                "Exchange Online Management module is not installed.`n`nThe tool will now attempt to install it. This may take a few minutes.",
                "Module Installation Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )

            if ($PSVersionTable.PSEdition -ne 'Core') {
                $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue
                if (-not $nuget) {
                    Write-Log "NuGet package provider not found. Installing..."
                    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
                }
            }

            Install-Module -Name ExchangeOnlineManagement -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Log "Exchange Online Management module installed successfully." -Level INFO

            [System.Windows.Forms.MessageBox]::Show(
                "Exchange Online Management module has been installed successfully.",
                "Installation Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )

            return $true
        }
        catch {
            Write-Log "Failed to install Exchange Online Management module: $($_.Exception.Message)" -Level ERROR

            [System.Windows.Forms.MessageBox]::Show(
                "Failed to install Exchange Online Management module.`n`nError: $($_.Exception.Message)`n`nPlease install manually using:`nInstall-Module -Name ExchangeOnlineManagement -Scope CurrentUser",
                "Installation Failed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )

            return $false
        }
    }
    else {
        Write-Log "Exchange Online Management module is already installed (Version: $($module.Version))."
        return $true
    }
}

function Connect-ExchangeOnlineSession {
    Write-Log "Attempting to connect to Exchange Online..."

    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop

        Connect-ExchangeOnline -ErrorAction Stop -ShowBanner:$false

        Write-Log "Successfully connected to Exchange Online."

        return $true
    }
    catch {
        Write-Log "Failed to connect to Exchange Online: $($_.Exception.Message)" -Level ERROR

        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to Exchange Online.`n`nError: $($_.Exception.Message)",
            "Connection Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )

        return $false
    }
}

function Get-ExchangeUsers {
    Write-Log "Retrieving directory-synced mailbox users from Exchange Online..."

    $list = New-Object System.Collections.Generic.List[object]
    $count = 0

    try {
        do {
            Get-Mailbox -Filter 'IsDirSynced -eq $true' -ResultSize Unlimited -ErrorAction Stop | ForEach-Object {
                if ($script:CancelRequested) {
                    break
                }

                $mbx = $_
                $searchParts = @($mbx.DisplayName, $mbx.UserPrincipalName, $mbx.PrimarySmtpAddress) + @($mbx.EmailAddresses)
                $searchText = (($searchParts | Where-Object { $_ }) -join ' ').ToLower()

                $list.Add([pscustomobject]@{
                    DisplayName            = [string]$mbx.DisplayName
                    PrimarySmtpAddress     = [string]$mbx.PrimarySmtpAddress
                    UserPrincipalName      = [string]$mbx.UserPrincipalName
                    IsExchangeCloudManaged = [bool]$mbx.IsExchangeCloudManaged
                    RecipientTypeDetails   = [string]$mbx.RecipientTypeDetails
                    Guid                   = [string]$mbx.Guid
                    SearchText             = $searchText
                })

                $count++
                if ($count % 50 -eq 0) {
                    $statusLabel.Text = "Loading users... $count loaded"
                    [System.Windows.Forms.Application]::DoEvents()
                }
            }
        } while ($false)
    }
    catch {
        Write-Log "Failed to retrieve mailbox users: $($_.Exception.Message)" -Level ERROR

        [System.Windows.Forms.MessageBox]::Show(
            "Failed to retrieve mailbox users.`n`nError: $($_.Exception.Message)",
            "Retrieval Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )

        return $null
    }

    if ($script:CancelRequested) {
        Write-Log "User load cancelled after $count users." -Level WARNING
        return $null
    }

    Write-Log "Retrieved $count directory-synced mailbox users."

    return ,$list
}

function Invoke-SoaConversion {
    param(
        [Parameter(Mandatory=$true)]
        [bool]$ToCloud
    )

    if ($dataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select at least one user from the list.",
            "No User Selected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $selected = @()
    foreach ($row in $dataGridView.SelectedRows) {
        $guid = [string]$row.Cells["Guid"].Value
        if ($script:UserIndex.ContainsKey($guid)) {
            $selected += [pscustomobject]@{ User = $script:UserIndex[$guid]; Row = $row }
        }
    }

    if ($selected.Count -eq 0) {
        return
    }

    $direction = if ($ToCloud) { 'Cloud Managed' } else { 'On-Premises Managed' }
    $target = $ToCloud

    $names = @($selected | ForEach-Object { $_.User.DisplayName } | Select-Object -First 20)
    $nameList = $names -join ", "
    if ($selected.Count -gt 20) {
        $nameList += "... and $($selected.Count - 20) more"
    }

    $confirmMessage = "Are you sure you want to convert $($selected.Count) user(s) to ${direction}?`n`nUsers: $nameList"

    if ($ToCloud) {
        $confirmMessage += "`n`nNote: Per Microsoft guidance, wait for a directory sync cycle plus up to 24 hours after Set-RemoteMailbox changes or a mailbox move before converting to cloud-managed."
    } else {
        $confirmMessage += "`n`nWARNING: Exchange attributes changed in the cloud will be overwritten by the on-premises values on the next sync cycle. A CSV backup is taken before conversion."
    }

    $result = [System.Windows.Forms.MessageBox]::Show(
        $confirmMessage,
        "Confirm Conversion",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    $operation = if ($ToCloud) { 'ToCloud' } else { 'ToOnPrem' }
    $backupFile = $null
    $successCount = 0
    $skippedCount = 0
    $failCount = 0
    $current = 0

    Set-BusyState -Busy $true -Message "Converting to $direction..."
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $progressBar.Minimum = 0
    $progressBar.Maximum = $selected.Count
    $progressBar.Value = 0

    try {
        $backupFolder = Get-SoaBackupFolder
        $backupFile = Join-Path $backupFolder ("SOA-Backup_{0}_{1}.csv" -f $operation, (Get-Date -Format 'yyyyMMdd_HHmmss'))

        foreach ($entry in $selected) {
            if ($script:CancelRequested) {
                Write-Log "Conversion cancelled by user." -Level WARNING
                break
            }

            $current++
            $user = $entry.User
            $statusLabel.Text = "Converting $current of $($selected.Count)..."
            $progressBar.Value = [Math]::Min($current, $progressBar.Maximum)
            [System.Windows.Forms.Application]::DoEvents()

            try {
                $mailbox = Get-Mailbox -Identity $user.Guid -ErrorAction Stop

                if ([bool]$mailbox.IsExchangeCloudManaged -eq $target) {
                    Write-Log "User '$($user.DisplayName)' ($($user.UserPrincipalName)) already has IsExchangeCloudManaged=$target. Skipped."
                    $skippedCount++
                    continue
                }

                try {
                    $record = ConvertTo-SoaBackupRecord -Mailbox $mailbox -Operation $operation
                    Export-SoaBackupRecord -Record $record -Path $backupFile
                }
                catch {
                    Write-Log "Backup failed for '$($user.DisplayName)' ($($user.UserPrincipalName)): $($_.Exception.Message). Skipping conversion." -Level ERROR
                    $failCount++
                    continue
                }

                Set-Mailbox -Identity $user.Guid -IsExchangeCloudManaged $target -ErrorAction Stop

                $user.IsExchangeCloudManaged = $target
                $entry.Row.Cells["IsExchangeCloudManaged"].Value = if ($target) { "True" } else { "False" }

                Write-Log "Converted user '$($user.DisplayName)' ($($user.UserPrincipalName)) to $direction."
                $successCount++
            }
            catch {
                Write-Log "Failed to convert user '$($user.DisplayName)' ($($user.UserPrincipalName)) to ${direction}: $($_.Exception.Message)" -Level ERROR
                $failCount++
            }
        }
    }
    finally {
        Set-BusyState -Busy $false
    }

    if ($successCount -gt 0 -and $backupFile -and (Test-Path $backupFile)) {
        Write-Log "Backup written to $backupFile"
    }

    Write-Log "Batch conversion to $direction completed. Successful: $successCount, Skipped: $skippedCount, Failed: $failCount"

    $summaryMessage = "Batch conversion completed.`n`nSuccessful: $successCount`nSkipped (already in target state): $skippedCount`nFailed: $failCount"
    if ($backupFile -and (Test-Path $backupFile)) {
        $summaryMessage += "`n`nBackup file:`n$backupFile"
    }
    if ($script:CancelRequested) {
        $summaryMessage += "`n`nOperation was cancelled before all users were processed."
    }

    [System.Windows.Forms.MessageBox]::Show(
        $summaryMessage,
        "Batch Conversion Summary",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )

    if ($script:CloseRequested) {
        $form.Close()
    }
}

function Backup-AllUsers {
    $operation = 'AllUsers'
    $count = 0
    $cancelled = $false
    $backupFile = $null

    Set-BusyState -Busy $true -Message "Backing up..."
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee

    try {
        $backupFolder = Get-SoaBackupFolder
        $backupFile = Join-Path $backupFolder ("SOA-Backup_{0}_{1}.csv" -f $operation, (Get-Date -Format 'yyyyMMdd_HHmmss'))

        $buffer = New-Object System.Collections.Generic.List[object]

        do {
            Get-Mailbox -Filter 'IsDirSynced -eq $true' -ResultSize Unlimited -ErrorAction Stop | ForEach-Object {
                if ($script:CancelRequested) {
                    break
                }

                $buffer.Add((ConvertTo-SoaBackupRecord -Mailbox $_ -Operation $operation))
                $count++

                if ($buffer.Count -ge 500) {
                    Export-SoaBackupRecord -Record $buffer -Path $backupFile
                    $buffer.Clear()
                }

                if ($count % 50 -eq 0) {
                    $statusLabel.Text = "Backing up... $count users"
                    [System.Windows.Forms.Application]::DoEvents()
                }
            }
        } while ($false)

        if ($buffer.Count -gt 0) {
            Export-SoaBackupRecord -Record $buffer -Path $backupFile
            $buffer.Clear()
        }

        $cancelled = $script:CancelRequested
    }
    catch {
        Write-Log "Backup All Users failed: $($_.Exception.Message)" -Level ERROR
        Set-BusyState -Busy $false

        [System.Windows.Forms.MessageBox]::Show(
            "Failed to back up users.`n`nError: $($_.Exception.Message)",
            "Backup Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )

        if ($script:CloseRequested) {
            $form.Close()
        }
        return
    }

    Set-BusyState -Busy $false

    if ($cancelled) {
        Write-Log "Backup All Users cancelled after $count users. Partial file: $backupFile" -Level WARNING

        [System.Windows.Forms.MessageBox]::Show(
            "Backup was cancelled. The file is partial and contains $count users.`n`nFile:`n$backupFile",
            "Backup Cancelled",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
    else {
        Write-Log "Backup All Users completed. $count users written to $backupFile"
        $statusLabel.Text = "Backup complete - $count users"

        $openResult = [System.Windows.Forms.MessageBox]::Show(
            "Backup completed successfully.`n`nUsers backed up: $count`nFile:`n$backupFile`n`nOpen the backup folder?",
            "Backup Complete",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        if ($openResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Start-Process explorer.exe -ArgumentList (Split-Path $backupFile -Parent)
            }
            catch {
                Write-Log "Failed to open backup folder: $($_.Exception.Message)" -Level WARNING
            }
        }
    }

    if ($script:CloseRequested) {
        $form.Close()
    }
}

Write-Log "========================================" -Level INFO
Write-Log "Exchange SOA Conversion tool Started" -Level INFO
Write-Log "========================================" -Level INFO

if (-not (Initialize-ExchangeOnlineModule)) {
    Write-Log "Cannot proceed without Exchange Online Management module. Exiting." -Level ERROR
    exit 1
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Exchange SOA Conversion"
$form.Size = New-Object System.Drawing.Size(1000, 700)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F3F3F3")
$form.ShowIcon = $false

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size = New-Object System.Drawing.Size(1000, 90)
$headerPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E8E8E8")
$headerPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($headerPanel)

$labelTitle = New-Object System.Windows.Forms.Label
$labelTitle.Location = New-Object System.Drawing.Point(20, 15)
$labelTitle.Size = New-Object System.Drawing.Size(850, 40)
$labelTitle.Text = "Exchange SOA Conversion Tool"
$labelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Regular)
$labelTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$labelTitle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$labelTitle.BackColor = [System.Drawing.Color]::Transparent
$labelTitle.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$headerPanel.Controls.Add($labelTitle)

$pictureBoxLogo = New-Object System.Windows.Forms.PictureBox
$pictureBoxLogo.Location = New-Object System.Drawing.Point(875, 8)
$pictureBoxLogo.Size = New-Object System.Drawing.Size(100, 75)
$pictureBoxLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$pictureBoxLogo.BackColor = [System.Drawing.Color]::Transparent
$pictureBoxLogo.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$logoPath = Join-Path $script:ScriptPath "logo.png"
if (Test-Path $logoPath) {
    try {
        $fileStream = [System.IO.File]::OpenRead($logoPath)
        $memoryStream = New-Object System.IO.MemoryStream
        $fileStream.CopyTo($memoryStream)
        $fileStream.Close()
        $fileStream.Dispose()
        $memoryStream.Position = 0
        $pictureBoxLogo.Image = [System.Drawing.Image]::FromStream($memoryStream)
    } catch {
        Write-Log "Failed to load logo image: $($_.Exception.Message)" -Level WARNING
    }
}
$headerPanel.Controls.Add($pictureBoxLogo)

$labelDescription = New-Object System.Windows.Forms.Label
$labelDescription.Location = New-Object System.Drawing.Point(20, 58)
$labelDescription.Size = New-Object System.Drawing.Size(850, 25)
$labelDescription.Text = "Convert Exchange mailbox Source of Authority"
$labelDescription.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$labelDescription.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#605E5C")
$labelDescription.BackColor = [System.Drawing.Color]::Transparent
$labelDescription.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$labelDescription.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$headerPanel.Controls.Add($labelDescription)

$buttonConnect = New-Object System.Windows.Forms.Button
$buttonConnect.Location = New-Object System.Drawing.Point(20, 105)
$buttonConnect.Size = New-Object System.Drawing.Size(160, 40)
$buttonConnect.Text = "Connect to EXO"
$buttonConnect.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
$buttonConnect.ForeColor = [System.Drawing.Color]::White
$buttonConnect.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$buttonConnect.FlatAppearance.BorderSize = 0
$buttonConnect.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonConnect.Add_Click({
    Set-BusyState -Busy $true -Message "Signing in to Exchange Online..."
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    [System.Windows.Forms.Application]::DoEvents()

    if (Connect-ExchangeOnlineSession) {
        $script:IsConnected = $true
        $buttonConnect.Text = "Connected"
        $buttonConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#107C10")

        $soaSuffix = ''
        try {
            $orgConfig = Get-OrganizationConfig -ErrorAction Stop
            Write-Log "Tenant BlockExchangeProvisioningFromOnPremEnabled: $($orgConfig.BlockExchangeProvisioningFromOnPremEnabled)"
            if ($orgConfig.BlockExchangeProvisioningFromOnPremEnabled -eq $true) {
                $soaSuffix = " | Tenant-wide SOA: Enabled"
            }
        }
        catch {
            Write-Log "Could not read organization config: $($_.Exception.Message)" -Level WARNING
        }

        $statusLabel.Text = "Loading users..."
        [System.Windows.Forms.Application]::DoEvents()

        $users = Get-ExchangeUsers

        Set-BusyState -Busy $false

        if ($null -ne $users) {
            $script:AllUsers = [object[]]$users.ToArray()
            $script:UserIndex = @{}
            foreach ($u in $script:AllUsers) {
                $script:UserIndex[$u.Guid] = $u
            }
            $script:CurrentPage = 1
            $script:ViewUsers = @(Select-SoaUsers -Users $script:AllUsers -SearchText $textBoxSearch.Text)
            Update-UserGrid
            $statusLabel.Text = "Connected - $($script:AllUsers.Count) users loaded$soaSuffix"
        }
        elseif ($script:CancelRequested) {
            $statusLabel.Text = "Load cancelled"
        }
        else {
            $statusLabel.Text = "Connected - failed to load users$soaSuffix"
        }
    }
    else {
        Set-BusyState -Busy $false
    }

    if ($script:CloseRequested) {
        $form.Close()
    }
})
$form.Controls.Add($buttonConnect)

$buttonRefresh = New-Object System.Windows.Forms.Button
$buttonRefresh.Location = New-Object System.Drawing.Point(195, 105)
$buttonRefresh.Size = New-Object System.Drawing.Size(140, 40)
$buttonRefresh.Text = "Refresh Users"
$buttonRefresh.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonRefresh.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonRefresh.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$buttonRefresh.FlatAppearance.BorderSize = 0
$buttonRefresh.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonRefresh.Enabled = $false
$buttonRefresh.Add_Click({
    Set-BusyState -Busy $true -Message "Loading users..."
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee

    $users = Get-ExchangeUsers

    Set-BusyState -Busy $false

    if ($null -ne $users) {
        $script:AllUsers = [object[]]$users.ToArray()
        $script:UserIndex = @{}
        foreach ($u in $script:AllUsers) {
            $script:UserIndex[$u.Guid] = $u
        }
        $script:CurrentPage = 1
        $script:ViewUsers = @(Select-SoaUsers -Users $script:AllUsers -SearchText $textBoxSearch.Text)
        Update-UserGrid
        $statusLabel.Text = "Refreshed - $($script:AllUsers.Count) users loaded"
    }
    elseif ($script:CancelRequested) {
        $statusLabel.Text = "Load cancelled"
    }

    if ($script:CloseRequested) {
        $form.Close()
    }
})
$form.Controls.Add($buttonRefresh)

$buttonDisconnect = New-Object System.Windows.Forms.Button
$buttonDisconnect.Location = New-Object System.Drawing.Point(350, 105)
$buttonDisconnect.Size = New-Object System.Drawing.Size(180, 40)
$buttonDisconnect.Text = "Disconnect from EXO"
$buttonDisconnect.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonDisconnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonDisconnect.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonDisconnect.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$buttonDisconnect.FlatAppearance.BorderSize = 0
$buttonDisconnect.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonDisconnect.Enabled = $false
$buttonDisconnect.Add_Click({
    Write-Log "Disconnecting from Exchange Online..."

    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-Log "Successfully disconnected from Exchange Online."
    }
    catch {
        Write-Log "Error during disconnect: $($_.Exception.Message)" -Level WARNING
    }

    $script:IsConnected = $false
    $script:AllUsers = @()
    $script:ViewUsers = @()
    $script:UserIndex = @{}
    $script:CurrentPage = 1
    $textBoxSearch.Text = ''

    $buttonConnect.Text = "Connect to EXO"
    $buttonConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
    $buttonConnect.Enabled = $true
    $buttonRefresh.Enabled = $false
    $buttonDisconnect.Enabled = $false
    $buttonBackupAll.Enabled = $false
    $buttonConvertToCloud.Enabled = $false
    $buttonConvertToOnPrem.Enabled = $false

    Update-UserGrid
    $pageInfo.Text = "No users to display"
    $buttonPrevPage.Enabled = $false
    $buttonNextPage.Enabled = $false

    $statusLabel.Text = "Disconnected"
})
$form.Controls.Add($buttonDisconnect)

$buttonBackupAll = New-Object System.Windows.Forms.Button
$buttonBackupAll.Location = New-Object System.Drawing.Point(545, 105)
$buttonBackupAll.Size = New-Object System.Drawing.Size(170, 40)
$buttonBackupAll.Text = "Backup All Users"
$buttonBackupAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonBackupAll.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonBackupAll.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonBackupAll.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$buttonBackupAll.FlatAppearance.BorderSize = 0
$buttonBackupAll.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonBackupAll.Enabled = $false
$buttonBackupAll.Add_Click({
    Backup-AllUsers
})
$form.Controls.Add($buttonBackupAll)

$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Location = New-Object System.Drawing.Point(725, 105)
$buttonCancel.Size = New-Object System.Drawing.Size(110, 40)
$buttonCancel.Text = "Cancel"
$buttonCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonCancel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#D83B01")
$buttonCancel.ForeColor = [System.Drawing.Color]::White
$buttonCancel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$buttonCancel.FlatAppearance.BorderSize = 0
$buttonCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonCancel.Visible = $false
$buttonCancel.Add_Click({
    $script:CancelRequested = $true
    $buttonCancel.Enabled = $false
    $statusLabel.Text = "Cancelling..."
})
$form.Controls.Add($buttonCancel)

$labelSearch = New-Object System.Windows.Forms.Label
$labelSearch.Location = New-Object System.Drawing.Point(20, 160)
$labelSearch.Size = New-Object System.Drawing.Size(60, 22)
$labelSearch.Text = "Search:"
$labelSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$labelSearch.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$labelSearch.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$form.Controls.Add($labelSearch)

$textBoxSearch = New-Object System.Windows.Forms.TextBox
$textBoxSearch.Location = New-Object System.Drawing.Point(80, 160)
$textBoxSearch.Size = New-Object System.Drawing.Size(400, 22)
$textBoxSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($textBoxSearch)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Size = New-Object System.Drawing.Size(315, 20)
$progressBar.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 335), 160)
$progressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

$searchTimer = New-Object System.Windows.Forms.Timer
$searchTimer.Interval = 300
$searchTimer.Add_Tick({
    $searchTimer.Stop()
    $script:ViewUsers = @(Select-SoaUsers -Users $script:AllUsers -SearchText $textBoxSearch.Text)
    $script:CurrentPage = 1
    Update-UserGrid
})
$textBoxSearch.Add_TextChanged({
    $searchTimer.Stop()
    $searchTimer.Start()
})

$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(20, 195)
$dataGridView.Size = New-Object System.Drawing.Size(960, 350)
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.ReadOnly = $true
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridView.MultiSelect = $true
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dataGridView.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$dataGridView.BackgroundColor = [System.Drawing.Color]::White
$dataGridView.GridColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$dataGridView.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
$dataGridView.DefaultCellStyle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$dataGridView.DefaultCellStyle.SelectionBackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
$dataGridView.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
$dataGridView.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$dataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#F9F9F9")
$dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$dataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$dataGridView.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
$dataGridView.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
$dataGridView.ColumnHeadersHeight = 40
$dataGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$dataGridView.EnableHeadersVisualStyles = $false
$dataGridView.RowHeadersVisible = $false
$dataGridView.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
$dataGridView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$dataGridView.Add_ColumnHeaderMouseClick({
    param($eventSender, $e)
    $columnName = $eventSender.Columns[$e.ColumnIndex].Name
    if ($columnName -ne "Guid") {
        Sort-Users -ColumnName $columnName
    }
})

$colDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDisplayName.Name = "DisplayName"
$colDisplayName.HeaderText = "Display Name"
$colDisplayName.FillWeight = 25
[void]$dataGridView.Columns.Add($colDisplayName)

$colEmail = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colEmail.Name = "Email"
$colEmail.HeaderText = "Email Address"
$colEmail.FillWeight = 30
[void]$dataGridView.Columns.Add($colEmail)

$colType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colType.Name = "Type"
$colType.HeaderText = "Type"
$colType.FillWeight = 20
[void]$dataGridView.Columns.Add($colType)

$colCloudManaged = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colCloudManaged.Name = "IsExchangeCloudManaged"
$colCloudManaged.HeaderText = "Cloud Managed"
$colCloudManaged.FillWeight = 25
[void]$dataGridView.Columns.Add($colCloudManaged)

$colGuid = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colGuid.Name = "Guid"
$colGuid.HeaderText = "Guid"
$colGuid.Visible = $false
[void]$dataGridView.Columns.Add($colGuid)

$form.Controls.Add($dataGridView)

$buttonPrevPage = New-Object System.Windows.Forms.Button
$buttonPrevPage.Location = New-Object System.Drawing.Point(20, 555)
$buttonPrevPage.Size = New-Object System.Drawing.Size(100, 30)
$buttonPrevPage.Text = "< Previous"
$buttonPrevPage.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonPrevPage.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonPrevPage.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonPrevPage.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$buttonPrevPage.FlatAppearance.BorderSize = 0
$buttonPrevPage.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonPrevPage.Enabled = $false
$buttonPrevPage.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$buttonPrevPage.Add_Click({
    if ($script:CurrentPage -gt 1) {
        $script:CurrentPage--
        Update-UserGrid
    }
})
$form.Controls.Add($buttonPrevPage)

$pageInfo = New-Object System.Windows.Forms.Label
$pageInfo.Location = New-Object System.Drawing.Point(130, 555)
$pageInfo.Size = New-Object System.Drawing.Size(550, 30)
$pageInfo.Text = "No users to display"
$pageInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$pageInfo.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#605E5C")
$pageInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$pageInfo.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($pageInfo)

$buttonNextPage = New-Object System.Windows.Forms.Button
$buttonNextPage.Location = New-Object System.Drawing.Point(690, 555)
$buttonNextPage.Size = New-Object System.Drawing.Size(100, 30)
$buttonNextPage.Text = "Next >"
$buttonNextPage.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonNextPage.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonNextPage.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonNextPage.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$buttonNextPage.FlatAppearance.BorderSize = 0
$buttonNextPage.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonNextPage.Enabled = $false
$buttonNextPage.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$buttonNextPage.Add_Click({
    $totalPages = [Math]::Ceiling($script:ViewUsers.Count / $script:PageSize)
    if ($script:CurrentPage -lt $totalPages) {
        $script:CurrentPage++
        Update-UserGrid
    }
})
$form.Controls.Add($buttonNextPage)

$buttonConvertToCloud = New-Object System.Windows.Forms.Button
$buttonConvertToCloud.Location = New-Object System.Drawing.Point(20, 595)
$buttonConvertToCloud.Size = New-Object System.Drawing.Size(220, 45)
$buttonConvertToCloud.Text = "Convert to Cloud Managed"
$buttonConvertToCloud.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonConvertToCloud.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonConvertToCloud.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonConvertToCloud.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$buttonConvertToCloud.FlatAppearance.BorderSize = 0
$buttonConvertToCloud.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonConvertToCloud.Enabled = $false
$buttonConvertToCloud.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$buttonConvertToCloud.Add_Click({
    Invoke-SoaConversion -ToCloud $true
})
$form.Controls.Add($buttonConvertToCloud)

$buttonConvertToOnPrem = New-Object System.Windows.Forms.Button
$buttonConvertToOnPrem.Location = New-Object System.Drawing.Point(255, 595)
$buttonConvertToOnPrem.Size = New-Object System.Drawing.Size(220, 45)
$buttonConvertToOnPrem.Text = "Convert to On-Prem Managed"
$buttonConvertToOnPrem.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonConvertToOnPrem.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonConvertToOnPrem.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonConvertToOnPrem.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$buttonConvertToOnPrem.FlatAppearance.BorderSize = 0
$buttonConvertToOnPrem.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonConvertToOnPrem.Enabled = $false
$buttonConvertToOnPrem.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$buttonConvertToOnPrem.Add_Click({
    Invoke-SoaConversion -ToCloud $false
})
$form.Controls.Add($buttonConvertToOnPrem)

$buttonOpenLog = New-Object System.Windows.Forms.Button
$buttonOpenLog.Location = New-Object System.Drawing.Point(490, 595)
$buttonOpenLog.Size = New-Object System.Drawing.Size(160, 45)
$buttonOpenLog.Text = "Open Log File"
$buttonOpenLog.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonOpenLog.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E1E1E1")
$buttonOpenLog.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1F1F1F")
$buttonOpenLog.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$buttonOpenLog.FlatAppearance.BorderSize = 0
$buttonOpenLog.Cursor = [System.Windows.Forms.Cursors]::Hand
$buttonOpenLog.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$buttonOpenLog.Add_Click({
    if (Test-Path $script:LogFile) {
        try {
            Start-Process notepad.exe -ArgumentList $script:LogFile
            Write-Log "Log file opened: $script:LogFile"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to open log file.`n`nError: $($_.Exception.Message)",
                "Error Opening Log",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show(
            "Log file does not exist yet.`n`nPath: $script:LogFile",
            "Log File Not Found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})
$form.Controls.Add($buttonOpenLog)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(665, 595)
$statusLabel.Size = New-Object System.Drawing.Size(315, 20)
$statusLabel.Text = "Not connected"
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$statusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#605E5C")
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$statusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($statusLabel)

$versionLabel = New-Object System.Windows.Forms.Label
$versionLabel.Location = New-Object System.Drawing.Point(665, 615)
$versionLabel.Size = New-Object System.Drawing.Size(315, 20)
$versionLabel.Text = "Version $script:Version"
$versionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$versionLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#808080")
$versionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$versionLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($versionLabel)

$form.Add_Resize({
    $buttonY = $form.ClientSize.Height - 105
    $buttonConvertToCloud.Location = New-Object System.Drawing.Point(20, $buttonY)
    $buttonConvertToOnPrem.Location = New-Object System.Drawing.Point(255, $buttonY)
    $buttonOpenLog.Location = New-Object System.Drawing.Point(490, $buttonY)

    $statusY = $form.ClientSize.Height - 105
    $statusLabel.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 335), $statusY)

    $versionY = $form.ClientSize.Height - 85
    $versionLabel.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 335), $versionY)

    $paginationY = $buttonY - 40
    $buttonPrevPage.Location = New-Object System.Drawing.Point(20, $paginationY)
    $pageInfo.Location = New-Object System.Drawing.Point(130, $paginationY)
    $pageInfo.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 270), 30)
    $buttonNextPage.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 120), $paginationY)

    $gridTop = 195
    $gridHeight = $paginationY - $gridTop - 10
    $dataGridView.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 40), $gridHeight)
})

$form.Add_FormClosing({
    param($eventSender, $e)

    if ($script:IsBusy) {
        $script:CancelRequested = $true
        $script:CloseRequested = $true
        $e.Cancel = $true
        Write-Log "Close requested while busy - will close after current operation." -Level INFO
        return
    }

    Write-Log "Exchange SOA Conversion tool Closing" -Level INFO

    if ($script:IsConnected) {
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Write-Log "Disconnected from Exchange Online." -Level INFO
        }
        catch {
            Write-Log "Error during disconnect: $($_.Exception.Message)" -Level WARNING
        }
        $script:IsConnected = $false
    }

    if ($pictureBoxLogo.Image) {
        $pictureBoxLogo.Image.Dispose()
    }
})

[void]$form.ShowDialog()

Write-Log "========================================" -Level INFO
Write-Log "Exchange SOA Conversion tool Ended" -Level INFO
Write-Log "========================================" -Level INFO
