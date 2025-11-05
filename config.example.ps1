# Example Configuration File for Zendesk to AD Automation
# Copy this file and customize it with your organization's settings
# Then use this file to set PowerShell variables before running scripts

# ==============================================================================
# ZENDESK CONFIGURATION
# ==============================================================================
$ZendeskUrl = "yourcompany.zendesk.com"
$ZendeskEmail = "your-email@yourcompany.com"
$ZendeskApiToken = "YOUR_API_TOKEN_HERE"
$ZendeskFormName = "Onboarding"  # Name of the Zendesk form to search for

# Zendesk Custom Field IDs
# To find these IDs, see the README.md section "Finding Zendesk Custom Field IDs"
$FieldID_FirstName = 1234567890123
$FieldID_LastName = 1234567890124
$FieldID_PersonalEmail = 1234567890125
$FieldID_Department = 1234567890126
$FieldID_Manager = 1234567890127
$FieldID_JobTitle = 1234567890128
$FieldID_EmployeeType = 1234567890129
$FieldID_EquipmentIssuedFrom = 1234567890130  # Optional - set to 0 if not used

# ==============================================================================
# ACTIVE DIRECTORY CONFIGURATION
# ==============================================================================
$AD_OU = "OU=Users,OU=Company,DC=company,DC=local"
$AD_DomainName = "company.local"  # Used for UserPrincipalName
$EmailDomain = "company.com"  # Used for email addresses

# AD Group Names (use common organizational groups)
$MfaGroupName = "MFA-Users"  # Mandatory group for all users
$ItEquipmentGroupName = "IT-Equipment"  # For users who need IT equipment
$RemoteAccessGroupName = "Remote-Access"  # For VPN/remote access users
$OfficeLocationGroupName = "Office-Users"  # For office-based users

# ==============================================================================
# GOOGLE WORKSPACE CONFIGURATION (Optional)
# ==============================================================================
$GoogleWorkspaceDomain = "company.com"  # Leave empty "" to skip Google Workspace checks
$GamCommand = "gam"  # Path to GAM command (e.g., "gam" or "C:\path\to\gam.exe")

# ==============================================================================
# EMAIL CONFIGURATION
# ==============================================================================
$SmtpServer = "smtp.gmail.com"
$SmtpPort = 587
$SmtpUsername = "helpdesk@company.com"
$CredentialFile = "C:\Scripts\credentials\smtp-secret.txt"  # Path to encrypted credential file
$FromEmail = "IT Support <support@company.com>"
$CcEmail = "support@company.com"  # Leave empty "" if not needed
$EmailSubject = "Your Company Account"

# Company Information
$CompanyName = "Your Company Name"

# Email Template URLs (replace placeholders with actual URLs)
$AccountPortalUrl = "https://myaccount.company.com"
$EmailUrl = "https://mail.company.com"
$CalendarUrl = "https://calendar.company.com"
$SupportUrl = "https://support.company.com"
$HrSystemUrl = "https://hr.company.com"
$TimesheetUrl = "https://timesheet.company.com"
$PerformanceUrl = "https://performance.company.com"
$TwoFactorUrl = "https://myaccount.google.com/signinoptions/two-step-verification/enroll-welcome"
$AuthenticatorUrl = "https://docs.company.com/authenticator-setup"
$SupportEmail = "support@company.com"

# ==============================================================================
# SCRIPT CONFIGURATION
# ==============================================================================
$OutputCSV = "Newaccount.csv"
$LogDir = ""  # Leave empty to use script directory\logs
$BaseDirectory = ""  # Leave empty to use script location

# ==============================================================================
# USAGE EXAMPLE
# ==============================================================================
# After configuring this file, you can use it like this:
#
# . .\config.example.ps1
# .\ADCreate.ps1 `
#     -ZendeskUrl $ZendeskUrl `
#     -ZendeskEmail $ZendeskEmail `
#     -ZendeskApiToken $ZendeskApiToken `
#     -FieldID_FirstName $FieldID_FirstName `
#     -FieldID_LastName $FieldID_LastName `
#     -FieldID_PersonalEmail $FieldID_PersonalEmail `
#     -FieldID_Department $FieldID_Department `
#     -FieldID_Manager $FieldID_Manager `
#     -FieldID_JobTitle $FieldID_JobTitle `
#     -FieldID_EmployeeType $FieldID_EmployeeType `
#     -AD_OU $AD_OU `
#     -AD_DomainName $AD_DomainName `
#     -EmailDomain $EmailDomain `
#     -SmtpServer $SmtpServer `
#     -SmtpUsername $SmtpUsername `
#     -CredentialFile $CredentialFile `
#     -FromEmail $FromEmail `
#     -EmailSubject $EmailSubject `
#     -CompanyName $CompanyName

