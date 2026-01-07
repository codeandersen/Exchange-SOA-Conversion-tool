#Requires -Version 5.1

<#
        .SYNOPSIS
        Exchange SOA Conversion tool (Cloud / On-Prem Source of Authority)

        .DESCRIPTION
        GUI tool to manage Exchange mailbox Source of Authority (SOA) conversion between cloud-managed and on-premises-managed by toggling the mailbox property `IsExchangeCloudManaged`.
        
        The tool connects to Exchange Online PowerShell, lists directory-synced mailboxes (`IsDirSynced = True`), and supports converting one or multiple selected users in batch.
        
        This tool is intended to support the approach described in:
        https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

        .EXAMPLE
        C:\PS> .\Exchange-SOA-Conversion-Tool.ps1

        .NOTES
        Version: 1.0

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

$script:Version = "1.0"


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:LogFile = Join-Path $script:ScriptPath "ExchangeCloudManagement_$(Get-Date -Format 'yyyyMMdd_HHmm').log"
$script:AllUsers = @()
$script:CurrentPage = 1
$script:PageSize = 100

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

function Update-UserGrid {
    param(
        [Parameter(Mandatory=$false)]
        [array]$Users = $script:AllUsers
    )
    
    $script:AllUsers = $Users
    $totalUsers = $script:AllUsers.Count
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
            $user = $script:AllUsers[$i]
            $cloudManagedStatus = if ($user.IsExchangeCloudManaged) { "True" } else { "False" }
            $dataGridView.Rows.Add($user.DisplayName, $user.PrimarySmtpAddress, $cloudManagedStatus, $user.UserPrincipalName)
        }
        
        $showingCount = $endIndex - $startIndex + 1
        $pageInfo.Text = "Page $($script:CurrentPage) of $totalPages - Showing $showingCount of $totalUsers users"
    } else {
        $pageInfo.Text = "No users to display"
    }
    
    $buttonPrevPage.Enabled = ($script:CurrentPage -gt 1)
    $buttonNextPage.Enabled = ($script:CurrentPage -lt $totalPages)
}

function Search-ExchangeOnlineModule {
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
            
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
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
    Write-Log "Retrieving mailbox users from Exchange Online..."
    
    try {
        $allMailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop | Select-Object DisplayName, PrimarySmtpAddress, IsExchangeCloudManaged, UserPrincipalName, IsDirSynced, RecipientTypeDetails
        
        Write-Log "Retrieved $($allMailboxes.Count) total mailbox users."
        
        $hybridMailboxes = $allMailboxes | Where-Object { $_.IsDirSynced -eq $true }
        
        $cloudOnlyCount = $allMailboxes.Count - $hybridMailboxes.Count
        Write-Log "Filtered out $cloudOnlyCount cloud-only mailboxes. Displaying $($hybridMailboxes.Count) hybrid/on-premises synced mailboxes."
        
        return $hybridMailboxes
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
}

function Convert-ToCloudManaged {
    param(
        [Parameter(Mandatory=$true)]
        $SelectedUser
    )
    
    $identity = $SelectedUser.UserPrincipalName
    $displayName = $SelectedUser.DisplayName
    
    Write-Log "Converting user '$displayName' ($identity) to Cloud Managed..."
    
    try {
        Set-Mailbox -Identity $identity -IsExchangeCloudManaged $true -ErrorAction Stop
        
        Write-Log "Successfully converted user '$displayName' ($identity) to Cloud Managed." -Level INFO
        
        return $true
    }
    catch {
        Write-Log "Failed to convert user '$displayName' ($identity) to Cloud Managed. Error: $($_.Exception.Message)" -Level ERROR
        
        return $false
    }
}

function Convert-ToOnPremManaged {
    param(
        [Parameter(Mandatory=$true)]
        $SelectedUser
    )
    
    $identity = $SelectedUser.UserPrincipalName
    $displayName = $SelectedUser.DisplayName
    
    Write-Log "Converting user '$displayName' ($identity) to On-Premises Managed..."
    
    try {
        Set-Mailbox -Identity $identity -IsExchangeCloudManaged $false -ErrorAction Stop
        
        Write-Log "Successfully converted user '$displayName' ($identity) to On-Premises Managed." -Level INFO
        
        return $true
    }
    catch {
        Write-Log "Failed to convert user '$displayName' ($identity) to On-Premises Managed. Error: $($_.Exception.Message)" -Level ERROR
        
        return $false
    }
}

Write-Log "========================================" -Level INFO
Write-Log "Exchange SOA Conversion tool Started" -Level INFO
Write-Log "========================================" -Level INFO

if (-not (Search-ExchangeOnlineModule)) {
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
    $buttonConnect.Enabled = $false
    $buttonRefresh.Enabled = $false
    
    if (Connect-ExchangeOnlineSession) {
        $buttonConnect.Text = "Connected"
        $buttonConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#107C10")
        $buttonConnect.Enabled = $false
        $buttonRefresh.Enabled = $true
        $buttonDisconnect.Enabled = $true
        
        $users = Get-ExchangeUsers
        
        if ($users) {
            $script:CurrentPage = 1
            Update-UserGrid -Users $users
            $statusLabel.Text = "Connected - $($users.Count) users loaded"
        }
    }
    else {
        $buttonConnect.Enabled = $true
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
    $buttonRefresh.Enabled = $false
    
    $users = Get-ExchangeUsers
    
    if ($users) {
        $script:CurrentPage = 1
        Update-UserGrid -Users $users
        $statusLabel.Text = "Refreshed - $($users.Count) users loaded"
    }
    
    $buttonRefresh.Enabled = $true
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
        
        $buttonConnect.Text = "Connect to EXO"
        $buttonConnect.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
        $buttonConnect.Enabled = $true
        $buttonRefresh.Enabled = $false
        $buttonDisconnect.Enabled = $false
        
        $dataGridView.Rows.Clear()
        
        $statusLabel.Text = "Disconnected"
        
        [System.Windows.Forms.MessageBox]::Show(
            "Successfully disconnected from Exchange Online.",
            "Disconnected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        Write-Log "Error during disconnect: $($_.Exception.Message)" -Level WARNING
        
        [System.Windows.Forms.MessageBox]::Show(
            "An error occurred while disconnecting.`n`nError: $($_.Exception.Message)",
            "Disconnect Warning",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
})
$form.Controls.Add($buttonDisconnect)

$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(20, 165)
$dataGridView.Size = New-Object System.Drawing.Size(960, 380)
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

$colDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDisplayName.Name = "DisplayName"
$colDisplayName.HeaderText = "Display Name"
$colDisplayName.FillWeight = 30
[void]$dataGridView.Columns.Add($colDisplayName)

$colEmail = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colEmail.Name = "Email"
$colEmail.HeaderText = "Email Address"
$colEmail.FillWeight = 35
[void]$dataGridView.Columns.Add($colEmail)

$colCloudManaged = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colCloudManaged.Name = "IsExchangeCloudManaged"
$colCloudManaged.HeaderText = "Cloud Managed"
$colCloudManaged.FillWeight = 25
[void]$dataGridView.Columns.Add($colCloudManaged)

$colUserPrincipalName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colUserPrincipalName.Name = "UserPrincipalName"
$colUserPrincipalName.HeaderText = "User Principal Name"
$colUserPrincipalName.Visible = $false
[void]$dataGridView.Columns.Add($colUserPrincipalName)

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
    $totalPages = [Math]::Ceiling($script:AllUsers.Count / $script:PageSize)
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
$buttonConvertToCloud.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$buttonConvertToCloud.Add_Click({
    if ($dataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select at least one user from the list.",
            "No User Selected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $selectedCount = $dataGridView.SelectedRows.Count
    $userList = ($dataGridView.SelectedRows | ForEach-Object { $_.Cells["DisplayName"].Value }) -join ", "
    
    $confirmMessage = if ($selectedCount -eq 1) {
        "Are you sure you want to convert user '$userList' to Cloud Managed?"
    } else {
        "Are you sure you want to convert $selectedCount users to Cloud Managed?`n`nUsers: $userList"
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $confirmMessage,
        "Confirm Conversion",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $successCount = 0
        $failCount = 0
        
        foreach ($selectedRow in $dataGridView.SelectedRows) {
            $selectedUser = @{
                DisplayName = $selectedRow.Cells["DisplayName"].Value
                PrimarySmtpAddress = $selectedRow.Cells["Email"].Value
                UserPrincipalName = $selectedRow.Cells["UserPrincipalName"].Value
            }
            
            if (Convert-ToCloudManaged -SelectedUser $selectedUser) {
                $selectedRow.Cells["IsExchangeCloudManaged"].Value = "True"
                $successCount++
            } else {
                $failCount++
            }
        }
        
        $summaryMessage = "Batch conversion completed.`n`nSuccessful: $successCount`nFailed: $failCount"
        
        [System.Windows.Forms.MessageBox]::Show(
            $summaryMessage,
            "Batch Conversion Summary",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        Write-Log "Batch conversion to Cloud Managed completed. Success: $successCount, Failed: $failCount"
    }
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
$buttonConvertToOnPrem.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$buttonConvertToOnPrem.Add_Click({
    if ($dataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select at least one user from the list.",
            "No User Selected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $selectedCount = $dataGridView.SelectedRows.Count
    $userList = ($dataGridView.SelectedRows | ForEach-Object { $_.Cells["DisplayName"].Value }) -join ", "
    
    $confirmMessage = if ($selectedCount -eq 1) {
        "Are you sure you want to convert user '$userList' to On-Premises Managed?"
    } else {
        "Are you sure you want to convert $selectedCount users to On-Premises Managed?`n`nUsers: $userList"
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $confirmMessage,
        "Confirm Conversion",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $successCount = 0
        $failCount = 0
        
        foreach ($selectedRow in $dataGridView.SelectedRows) {
            $selectedUser = @{
                DisplayName = $selectedRow.Cells["DisplayName"].Value
                PrimarySmtpAddress = $selectedRow.Cells["Email"].Value
                UserPrincipalName = $selectedRow.Cells["UserPrincipalName"].Value
            }
            
            if (Convert-ToOnPremManaged -SelectedUser $selectedUser) {
                $selectedRow.Cells["IsExchangeCloudManaged"].Value = "False"
                $successCount++
            } else {
                $failCount++
            }
        }
        
        $summaryMessage = "Batch conversion completed.`n`nSuccessful: $successCount`nFailed: $failCount"
        
        [System.Windows.Forms.MessageBox]::Show(
            $summaryMessage,
            "Batch Conversion Summary",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        Write-Log "Batch conversion to On-Premises Managed completed. Success: $successCount, Failed: $failCount"
    }
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
    
    $gridHeight = $paginationY - 175
    $dataGridView.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 40), $gridHeight)
})

$form.Add_FormClosing({
    Write-Log "Exchange SOA Conversion tool Closing" -Level INFO
    
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Log "Disconnected from Exchange Online." -Level INFO
    }
    catch {
        Write-Log "Error during disconnect: $($_.Exception.Message)" -Level WARNING
    }
    
    if ($pictureBoxLogo.Image) {
        $pictureBoxLogo.Image.Dispose()
    }
})

[void]$form.ShowDialog()

Write-Log "========================================" -Level INFO
Write-Log "Exchange SOA Conversion tool Ended" -Level INFO
Write-Log "========================================" -Level INFO
