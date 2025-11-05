# Zendesk to Active Directory Automation

Automated PowerShell solution that creates Active Directory users from Zendesk tickets, manages group memberships, and sends HTML-formatted welcome emails with account credentials.

## Features

- 🔄 **Automated User Creation**: Creates AD users from Zendesk onboarding tickets
- 👥 **Group Management**: Automatically adds users to appropriate security groups
- 📧 **Email Notifications**: Sends HTML-formatted welcome emails with credentials
- 🔐 **Password Generation**: Secure random password generation (16 characters)
- 📊 **Comprehensive Logging**: Detailed logging for audit and troubleshooting
- 🔄 **Scheduled Execution**: Can be run on a schedule via Task Scheduler
- ☁️ **Google Workspace Integration**: Optional verification before sending emails

## Prerequisites

- Windows Server with Active Directory Domain Services
- PowerShell 5.1 or higher
- Active Directory PowerShell module
- SMTP server access (for email notifications)
- Zendesk API credentials
- (Optional) Google Workspace Admin SDK (GAM) for Google Workspace verification

## Installation

1. **Clone or download this repository**
   ```powershell
   git clone https://github.com/yourusername/zendesk-ad-integration.git
   cd zendesk-ad-integration
   ```

2. **Configure your settings**
   - Copy `config.example.ps1` to `config.ps1`
   - Edit `config.ps1` with your organization's settings
   - See [Configuration Guide](#configuration) below

3. **Set up SMTP credentials**
   ```powershell
   Read-Host 'Enter SMTP password' -AsSecureString | ConvertFrom-SecureString | Out-File "C:\Scripts\credentials\smtp-secret.txt"
   ```

## Scripts Overview

### `ADCreate.ps1` - Main Automation Script
Fetches tickets from Zendesk, exports to CSV, creates AD users, and sends welcome emails. This is the complete end-to-end automation script.

## Configuration

### Finding Zendesk Custom Field IDs

1. Log in to your Zendesk account as an administrator
2. Go to **Admin** → **Objects and rules** → **Tickets** → **Fields**
3. Click on the custom field you want to use
4. The Field ID is displayed in the URL: `https://yourcompany.zendesk.com/admin/objects-rules/tickets/ticket_fields/[FIELD_ID]`
5. Alternatively, use the Zendesk API:
   ```powershell
   $headers = @{
       Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("your-email@company.com/token:YOUR_API_TOKEN"))
   }
   $response = Invoke-RestMethod -Uri "https://yourcompany.zendesk.com/api/v2/ticket_fields.json" -Headers $headers
   $response.ticket_fields | Select-Object id, title, type
   ```

### Active Directory Groups

The script uses generic group names that should exist in most organizations:

- **MFA-Users** (mandatory): All users are added to this group for multi-factor authentication
- **IT-Equipment**: Users who need IT equipment (laptops, etc.)
- **Remote-Access**: Users who need VPN/remote access
- **Office-Users**: Users based in office locations

You can customize these group names in the configuration. The script will create CSV columns (`ITEquipment`, `RemoteAccess`, `OfficeUsers`) that default to `FALSE`. You can manually edit the CSV to set these to `TRUE` for specific users, or set up Zendesk workflows to populate these fields.

### CSV Structure

The scripts generate/expect a CSV with the following columns:

```
firstname,lastname,username,department,jobtitle,personalemail,employeetype,manager,ITEquipment,RemoteAccess,OfficeUsers
```

- `ITEquipment`, `RemoteAccess`, `OfficeUsers`: Must be `TRUE` or `FALSE` (case-sensitive)

## Usage

### Running the Main Automation Script

```powershell
# Load configuration
. .\config.ps1

# Run the main script
.\ADCreate.ps1 `
    -ZendeskUrl $ZendeskUrl `
    -ZendeskEmail $ZendeskEmail `
    -ZendeskApiToken $ZendeskApiToken `
    -ZendeskFormName $ZendeskFormName `
    -FieldID_FirstName $FieldID_FirstName `
    -FieldID_LastName $FieldID_LastName `
    -FieldID_PersonalEmail $FieldID_PersonalEmail `
    -FieldID_Department $FieldID_Department `
    -FieldID_Manager $FieldID_Manager `
    -FieldID_JobTitle $FieldID_JobTitle `
    -FieldID_EmployeeType $FieldID_EmployeeType `
    -AD_OU $AD_OU `
    -AD_DomainName $AD_DomainName `
    -EmailDomain $EmailDomain `
    -SmtpServer $SmtpServer `
    -SmtpUsername $SmtpUsername `
    -CredentialFile $CredentialFile `
    -FromEmail $FromEmail `
    -EmailSubject $EmailSubject `
    -CompanyName $CompanyName
```

### Scheduled Task Setup

1. Open Task Scheduler
2. Create a new task
3. Set trigger (e.g., daily at 9 AM)
4. Set action: Start a program
   - Program: `powershell.exe`
   - Arguments: `-File "C:\Scripts\ADCreate.ps1" -ZendeskUrl "..." [other parameters]`
   - Start in: `C:\Scripts`
5. Set "Run whether user is logged on or not"
6. Use an account with appropriate AD permissions

## Google Workspace Integration

If you use Google Workspace and want to verify users exist before sending emails:

1. Install [GAM (Google Workspace Admin SDK)](https://github.com/GAM-team/GAM)
2. Configure GAM with appropriate permissions
3. Set `GoogleWorkspaceDomain` in your configuration
4. The script will verify users exist in Google Workspace before sending emails

If Google Workspace is not configured, the script will skip verification and proceed with email sending.

## Email Template Customization

The email template includes placeholders for URLs. Update these in your configuration:

- `AccountPortalUrl`: Password change portal
- `EmailUrl`: Email access URL
- `CalendarUrl`: Calendar access URL
- `SupportUrl`: Support/ticketing system
- `HrSystemUrl`: HR system URL
- `TimesheetUrl`: Timesheet system URL
- `PerformanceUrl`: Performance management URL
- `TwoFactorUrl`: 2FA setup URL
- `AuthenticatorUrl`: Authenticator app setup guide
- `SupportEmail`: Support email address

## Security Considerations

- **API Tokens**: Store Zendesk API tokens securely. Consider using Windows Credential Manager or Azure Key Vault.
- **SMTP Credentials**: Use encrypted credential files (created with `ConvertFrom-SecureString`)
- **AD Permissions**: Scripts require appropriate AD permissions to create users and manage groups
- **Log Files**: Log files may contain sensitive information. Secure them appropriately.
- **CSV Files**: CSV files contain user data. Ensure they are stored securely and deleted after use.

## Troubleshooting

### "Active Directory module not available"
- Install RSAT-AD-PowerShell feature on your Windows machine
- On Windows Server, the module should already be available

### "Failed to decrypt credential file"
- The credential file must be encrypted by the same user account running the script
- Recreate the credential file as the user who will run the script

### "User already exists"
- The script checks for existing users and skips them
- This is expected behavior to prevent duplicate accounts

### "Google Workspace account not found"
- Ensure GAM is installed and configured correctly
- Verify the Google Workspace domain is correct
- Check that the user account sync has completed

### Zendesk API Errors
- Verify your API token is valid
- Check that the custom field IDs are correct
- Ensure the Zendesk form name matches exactly

## CSV Column Names

The scripts use these column names in the CSV:

- `firstname`, `lastname`, `username`: User name information
- `department`, `jobtitle`: Job information
- `personalemail`: Personal email for welcome email
- `employeetype`: Employee type (FT, PT, CT)
- `manager`: Manager username (format: firstname.lastname)
- `ITEquipment`, `RemoteAccess`, `OfficeUsers`: Group membership flags (TRUE/FALSE)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

For issues and questions:
- Open an issue on GitHub
- Check the troubleshooting section above
- Review the logs in the `logs/` directory

## Author

Abrham Wondimu

---

**Note**: This script has been generalized from a company-specific implementation. All hardcoded values have been replaced with configurable parameters. Please review and customize all settings before use in production.

