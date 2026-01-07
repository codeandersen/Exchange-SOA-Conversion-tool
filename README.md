# Exchange SOA Conversion Tool

A PowerShell GUI tool for switching the Exchange mailbox *source of authority* between on-premises and Exchange Online in a hybrid environment by toggling `IsExchangeCloudManaged`.

This is intended to support the approach described in:
https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

## Features

- **Modern GUI interface**: Clean Windows Forms interface with responsive design and modern styling
- **Automatic module installation**: Checks for `ExchangeOnlineManagement` and installs it in *CurrentUser* scope if needed
- **Exchange Online connectivity**: Uses modern auth via `Connect-ExchangeOnline`
- **Hybrid-focused user list**: Retrieves all EXO mailboxes and displays only directory-synced (`IsDirSynced = True`) objects
- **Pagination support**: Displays users in pages of 100 with Previous/Next navigation
- **Batch conversion**: Multi-select users and convert in bulk with confirmation dialogs
- **Cloud conversion**: Sets `IsExchangeCloudManaged = True`
- **On-prem conversion**: Sets `IsExchangeCloudManaged = False`
- **Logging + quick access**: Writes a timestamped log file and includes an **Open Log File** button
- **Connection management**: Connect, refresh, and disconnect from Exchange Online with status indicators
- **Logo support**: Displays custom logo (logo.png) if present in script directory
- **Responsive layout**: Automatically adjusts to window resizing

## Screenshots

![Exchange SOA Conversion Tool](Images/Exchange-SOA-Conversion-tool-GUI.png)

*The Exchange SOA Conversion Tool showing a connected session with directory-synced mailboxes displayed in a paginated view.*

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
   .\Exchange-SOA-Conversion-Tool.ps1
   ```

2. **Connect to Exchange Online**:
   - Click "Connect to EXO" button
   - Sign in with your Exchange Online admin credentials
   - The tool will automatically load directory-synced (`IsDirSynced = True`) mailbox users
   - Button changes to "Connected" with green color upon successful connection

3. **Navigate Users**:
   - Use "Previous" and "Next" buttons to navigate through pages of users
   - Page information displays current page, total pages, and user count
   - Each page shows up to 100 users

4. **Convert Users**:
   - Select one or multiple users from the list (multi-select supported)
   - Click "Convert to Cloud Managed" to enable cloud management
   - Click "Convert to On-Prem Managed" to revert to on-premises management
   - Confirm the conversion when prompted
   - View batch conversion summary showing successful and failed conversions

5. **Refresh User List**:
   - Click "Refresh Users" to reload the mailbox list after conversions

6. **View Logs**:
   - Click "Open Log File" to view the session log in Notepad

7. **Disconnect**:
   - Click "Disconnect from EXO" when finished
   - Tool automatically disconnects when closing the window

## Log Files

Log files are created in the same directory as the script with the naming format:
```
ExchangeSOAConversion_YYYYMMDD_HHMM.log
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

### What happens after conversion to cloud-managed?

Once a user is converted to **cloud-managed** (`IsExchangeCloudManaged = True`), Exchange attributes for that mailbox can be managed directly in Exchange Online instead of on-premises Exchange. This means:

- **Exchange attributes** can be modified using Exchange Online PowerShell or the Exchange Admin Center
- Changes no longer need to be made in the on-premises Exchange Management Console/Shell
- The mailbox remains directory-synced from on-premises Active Directory, but Exchange-specific attributes are managed in the cloud
- This provides flexibility in hybrid environments where on-premises Exchange may be decommissioned or scaled down

#### Attributes that can be edited in Exchange Online after conversion:

**Custom and Extension Attributes:**
- ExtensionAttribute1 through ExtensionAttribute15

**Mailbox Settings:**
- mail
- altRecipient
- authoring
- msExchAssistantName
- msExchAuditAdmin
- etc. -> See more at https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management



**Important**: The user object itself is still synced from on-premises AD via Azure AD Connect. Only the Exchange recipient attributes are managed in the cloud.

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
- Users are displayed in pages of 100 for better performance with large mailbox counts.
- The tool automatically disconnects from Exchange Online when the window is closed.
- Optional: Place a `logo.png` file in the same directory as the script to display a custom logo in the header.

## Version

Current version: 1.0
