# Exchange Cloud Manage Conversion Tool

A PowerShell GUI tool for switching the Exchange mailbox *source of authority* between on-premises and Exchange Online in a hybrid environment by toggling `IsExchangeCloudManaged`.

This is intended to support the approach described in:
https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

## Features

- **GUI workflow**: Simple Windows Forms interface (connect, view users, convert, refresh, disconnect)
- **Automatic module installation**: Checks for `ExchangeOnlineManagement` and installs it in *CurrentUser* scope if needed
- **Exchange Online connectivity**: Uses modern auth via `Connect-ExchangeOnline`
- **Hybrid-focused user list**: Retrieves all EXO mailboxes and displays only directory-synced (`IsDirSynced = True`) objects
- **Batch conversion**: Multi-select users and convert in bulk
- **Cloud conversion**: Sets `IsExchangeCloudManaged = True`
- **On-prem conversion**: Sets `IsExchangeCloudManaged = False`
- **Logging + quick access**: Writes a timestamped log file and includes an **Open Log File** button

## Requirements

- Windows PowerShell 5.1 or later
- Exchange Online PowerShell module: `ExchangeOnlineManagement`
- Connectivity to Exchange Online endpoints
- Appropriate Exchange Online permissions to run `Get-Mailbox` and `Set-Mailbox`

### Permissions / roles

At minimum, the signed-in admin account must be able to:

- Run `Get-Mailbox` across the target scope
- Run `Set-Mailbox -IsExchangeCloudManaged ...` on the target recipients

Commonly used roles for this include (depending on your org model):

- Exchange Administrator
- Global Administrator

## Usage

1. **Run the Script**:
   ```powershell
   .\Exchange-SOA-Conversion.ps1
   ```

2. **Connect to Exchange Online**:
   - Click "Connect to EXO" button
   - Sign in with your Exchange Online admin credentials
   - The tool will automatically load directory-synced (`IsDirSynced = True`) mailbox users

3. **Convert Users**:
   - Select a user from the list
   - Click "Convert to Cloud Managed" to enable cloud management
   - Click "Convert to On-Prem Managed" to revert to on-premises management
   - You can multi-select users to run conversions in batch

4. **Refresh User List**:
   - Click "Refresh Users" to reload the mailbox list after conversions

5. **Disconnect**:
   - Click "Disconnect from EXO" when finished

## Log Files

Log files are created in the same directory as the script with the naming format:
```
ExchangeCloudManagement_YYYYMMDD_HHMM.log
```

Logged operations include:
- Exchange Online Management module installation attempts
- Connection to Exchange Online
- User conversions to cloud managed
- User conversions to on-premises managed
- Any errors or warnings

## PowerShell Commands

The tool executes the following commands:

**Convert to Cloud Managed**:
```powershell
Set-Mailbox -Identity <User> -IsExchangeCloudManaged $true
```

**Convert to On-Premises Managed**:
```powershell
Set-Mailbox -Identity <User> -IsExchangeCloudManaged $false
```

## How this maps to the Microsoft guidance

The Microsoft article describes enabling Exchange attribute management in the cloud for hybrid recipients. This tool focuses specifically on flipping the mailbox management flag (`IsExchangeCloudManaged`) for directory-synced mailboxes so you can move the recipient management “source of authority” between:

- Exchange on-premises (traditional hybrid management)
- Exchange Online (cloud-managed attributes)

## Reference

For more information about Exchange cloud attributes management, see:
https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

## Troubleshooting

- **Module Installation Fails**: Ensure you have internet connectivity and appropriate permissions. You can manually install the module using:
  ```powershell
  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
  ```

- **Connection Issues**: Verify your credentials have the necessary Exchange Online permissions

- **Conversion Failures**: Check the log file for detailed error messages

## Notes / limitations

- The tool intentionally filters out cloud-only mailboxes and shows only directory-synced mailboxes (`IsDirSynced = True`).
- Changes may take time to reflect depending on your environment and any directory sync / hybrid processes.
