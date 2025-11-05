# Zendesk to Active Directory Automation Script
# Fetches onboarding tickets from Zendesk, exports to CSV, creates AD users, and sends welcome emails
# Only fetches 'new' and 'open' tickets (duplicates are handled by AD existence check)
# Optimized for scheduled task execution
param(
    # Zendesk Configuration
    [Parameter(Mandatory=$true)]
    [string]$ZendeskUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$ZendeskEmail,
    
    [Parameter(Mandatory=$true)]
    [string]$ZendeskApiToken,
    
    [string]$ZendeskFormName = "Onboarding",
    
    # Zendesk Custom Field IDs (required - see README for how to find these)
    [Parameter(Mandatory=$true)]
    [long]$FieldID_FirstName,
    
    [Parameter(Mandatory=$true)]
    [long]$FieldID_LastName,
    
    [Parameter(Mandatory=$true)]
    [long]$FieldID_PersonalEmail,
    
    [Parameter(Mandatory=$true)]
    [long]$FieldID_Department,
    
    [Parameter(Mandatory=$true)]
    [long]$FieldID_Manager,
    
    [Parameter(Mandatory=$true)]
    [long]$FieldID_JobTitle,
    
    [Parameter(Mandatory=$true)]
    [long]$FieldID_EmployeeType,
    
    [long]$FieldID_EquipmentIssuedFrom = 0,  # Optional
    
    # AD Configuration
    [Parameter(Mandatory=$true)]
    [string]$AD_OU,
    
    [Parameter(Mandatory=$true)]
    [string]$AD_DomainName,  # e.g., "company.local"
    
    [Parameter(Mandatory=$true)]
    [string]$EmailDomain,  # e.g., "company.com"
    
    [string]$MfaGroupName = "MFA-Users",
    [string]$ItEquipmentGroupName = "IT-Equipment",
    [string]$RemoteAccessGroupName = "Remote-Access",
    [string]$OfficeLocationGroupName = "Office-Users",
    
    # Google Workspace Configuration (optional)
    [string]$GoogleWorkspaceDomain = "",  # If empty, Google Workspace checks are skipped
    [string]$GamCommand = "gam",  # GAM command path/name
    [int]$GSuiteCheckIntervalSeconds = 30,
    
    # Email Configuration
    [Parameter(Mandatory=$true)]
    [string]$SmtpServer,
    
    [int]$SmtpPort = 587,
    
    [Parameter(Mandatory=$true)]
    [string]$SmtpUsername,
    
    [Parameter(Mandatory=$true)]
    [string]$CredentialFile,  # Path to encrypted credential file
    
    [Parameter(Mandatory=$true)]
    [string]$FromEmail,
    
    [string]$CcEmail = "",
    
    [Parameter(Mandatory=$true)]
    [string]$EmailSubject,
    
    [Parameter(Mandatory=$true)]
    [string]$CompanyName,
    
    [string]$AccountPortalUrl = "{{ACCOUNT_PORTAL_URL}}",
    [string]$EmailUrl = "{{EMAIL_URL}}",
    [string]$CalendarUrl = "{{CALENDAR_URL}}",
    [string]$SupportUrl = "{{SUPPORT_URL}}",
    [string]$HrSystemUrl = "{{HR_SYSTEM_URL}}",
    [string]$TimesheetUrl = "{{TIMESHEET_URL}}",
    [string]$PerformanceUrl = "{{PERFORMANCE_URL}}",
    [string]$TwoFactorUrl = "{{TWO_FACTOR_URL}}",
    [string]$AuthenticatorUrl = "{{AUTHENTICATOR_URL}}",
    [string]$SupportEmail = "{{SUPPORT_EMAIL}}",
    
    # Script Configuration
    [string]$OutputCSV = "Newaccount.csv",
    [string]$LogDir = "",
    [string]$BaseDirectory = ""  # If empty, uses script location
)

#region Initialization
$ErrorActionPreference = "Continue"
if ([string]::IsNullOrWhiteSpace($BaseDirectory)) {
    $ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptPath = $BaseDirectory
}
$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Path)

# Setup logging directory
if ([string]::IsNullOrWhiteSpace($LogDir)) {
    $LogDir = Join-Path $ScriptPath "logs"
}

if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

$LogFile = Join-Path $LogDir "${ScriptName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Function to write to both console and log file
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    Write-Host $logMessage
    
    try {
        Add-Content -Path $LogFile -Value $logMessage -ErrorAction SilentlyContinue
    }
    catch {
        # If logging fails, continue execution
    }
}

# Setup output CSV path
$FullOutputPath = Join-Path $ScriptPath $OutputCSV

# Load required .NET assembly for URL encoding
Add-Type -AssemblyName System.Web

# Check for Active Directory module
$ADModuleAvailable = $false
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log "Active Directory module imported successfully" "INFO"
    $ADModuleAvailable = $true
}
catch {
    Write-Log "WARNING: Active Directory module not available. AD creation will be skipped." "WARN"
}

Write-Log "=== Script Started ===" "INFO"
Write-Log "Script: $ScriptName" "INFO"
Write-Log "Script Path: $ScriptPath" "INFO"
Write-Log "Output CSV: $FullOutputPath" "INFO"
Write-Log "Log File: $LogFile" "INFO"
Write-Log "AD OU: $AD_OU" "INFO"
Write-Log "AD Domain: $AD_DomainName" "INFO"
Write-Log "Email Domain: $EmailDomain" "INFO"
Write-Log "Fetching 'new' and 'open' tickets with form: $ZendeskFormName" "INFO"
#endregion

#region Helper Functions

# Function to generate random complex password (16 characters)
function Generate-ComplexPassword {
    $lowercase = "abcdefghijkmnpqrstuvwxyz"
    $uppercase = "ABCDEFGHJKLMNPQRSTUVWXYZ"
    $numbers = "23456789"
    $special = "!@#$%^&*()_+-=[]{}|;:,.<>?"
    
    $password = (
        (Get-Random -InputObject $lowercase.ToCharArray() -Count 2) +
        (Get-Random -InputObject $uppercase.ToCharArray() -Count 2) +
        (Get-Random -InputObject $numbers.ToCharArray() -Count 2) +
        (Get-Random -InputObject $special.ToCharArray() -Count 2) +
        (Get-Random -InputObject ($lowercase + $uppercase + $numbers + $special).ToCharArray() -Count 8)
    ) | Sort-Object {Get-Random}
    
    return (-join $password)
}

# Function to check if user exists in AD
function Test-ADUserExists {
    param([string]$Username)
    
    try {
        $user = Get-ADUser -Filter {SamAccountName -eq $Username} -ErrorAction SilentlyContinue
        return ($null -ne $user)
    }
    catch {
        Write-Log "Error checking if user $Username exists: $_" "WARN"
        return $false
    }
}

# Function to create AD user
function New-ADUserAccount {
    param(
        [PSCustomObject]$User,
        [string]$Password
    )
    
    $Username = ($User.username).ToLower()
    $Firstname = $User.firstname
    $Lastname = $User.lastname
    
    try {
        if (Test-ADUserExists -Username $Username) {
            Write-Log "User $Username already exists in AD. Skipping creation." "WARN"
            return @{ Success = $false; Message = "User already exists" }
        }
        
        # Get group membership flags from CSV (generic group names)
        $AddToItEquipment = $User.ITEquipment -eq "TRUE"
        $AddToRemoteAccess = $User.RemoteAccess -eq "TRUE"
        $AddToOfficeUsers = $User.OfficeUsers -eq "TRUE"
        
        Write-Log "Creating AD user: $Username" "INFO"
        Write-Log "  Name: $Firstname $Lastname" "INFO"
        Write-Log "  Groups: ITEquipment=$AddToItEquipment, RemoteAccess=$AddToRemoteAccess, OfficeUsers=$AddToOfficeUsers" "INFO"
        
        # Create new user
        New-ADUser `
            -SamAccountName $Username `
            -UserPrincipalName "$Username@$AD_DomainName" `
            -EmailAddress "$Username@$EmailDomain" `
            -Name "$Firstname $Lastname" `
            -GivenName $Firstname `
            -Surname $Lastname `
            -Enabled $True `
            -ChangePasswordAtLogon $True `
            -DisplayName "$Firstname $Lastname" `
            -Path $AD_OU `
            -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) `
            -Department $User.department `
            -Manager $($User.manager).ToLower() `
            -Title $User.jobtitle `
            -OtherAttributes @{'EmployeeType' = $User.employeetype; 'secondaryMail' = $User.personalemail} `
            -ErrorAction Stop
        
        Write-Log "Successfully created AD user: $Username" "INFO"
        
        # Add to mandatory MFA group
        try {
            Add-ADGroupMember -Identity $MfaGroupName -Members $Username -ErrorAction Stop
            Write-Log "Added $Username to mandatory $MfaGroupName group" "INFO"
        }
        catch {
            Write-Log "WARNING: Failed to add $Username to $MfaGroupName group: $_" "WARN"
        }
        
        # Add to optional groups
        if ($AddToItEquipment) {
            Add-ADGroupMember -Identity $ItEquipmentGroupName -Members $Username -ErrorAction SilentlyContinue
            Write-Log "Added $Username to $ItEquipmentGroupName group" "INFO"
        }
        if ($AddToRemoteAccess) {
            Add-ADGroupMember -Identity $RemoteAccessGroupName -Members $Username -ErrorAction SilentlyContinue
            Write-Log "Added $Username to $RemoteAccessGroupName group" "INFO"
        }
        if ($AddToOfficeUsers) {
            Add-ADGroupMember -Identity $OfficeLocationGroupName -Members $Username -ErrorAction SilentlyContinue
            Write-Log "Added $Username to $OfficeLocationGroupName group" "INFO"
        }
        
        return @{ Success = $true; Message = "User created successfully" }
    }
    catch {
        Write-Log "ERROR creating AD user $Username : $_" "ERROR"
        return @{ Success = $false; Message = "Error: $_" }
    }
}

# Function to check if user exists in Google Workspace
function Test-GoogleWorkspaceUser {
    param(
        [string]$Username,
        [int]$MaxWaitMinutes = 30,
        [int]$CheckIntervalSeconds = 60
    )
    
    if ([string]::IsNullOrWhiteSpace($GoogleWorkspaceDomain)) {
        Write-Log "Google Workspace domain not configured. Skipping check." "INFO"
        return $true  # Return true to proceed if GWS is not configured
    }
    
    $mail = "$Username@$GoogleWorkspaceDomain"
    $maxAttempts = [math]::Ceiling(($MaxWaitMinutes * 60) / $CheckIntervalSeconds)
    $attempt = 0
    
    Write-Log "Checking Google Workspace for user: $mail" "INFO"
    
    while ($attempt -lt $maxAttempts) {
        try {
            $GSuitInfo = cmd /c "$GamCommand info user $mail" 2>&1 | Out-String
            
            if ($GSuitInfo -match "User: $mail") {
                if ($GSuitInfo -match '(?i)creation[\s-]time[:\s]*([0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}\.[0-9]+Z)') {
                    $creationTimeStr = $matches[1].Trim()
                    try {
                        $creationTime = [datetime]::ParseExact($creationTimeStr, "yyyy-MM-ddTHH:mm:ss.fffZ", $null)
                        $timeDifference = (Get-Date) - $creationTime
                        
                        if ($timeDifference.TotalHours -le 1) {
                            Write-Log "Google Workspace account for $mail verified (created $([math]::Round($timeDifference.TotalMinutes, 1)) minutes ago)" "INFO"
                            return $true
                        }
                        else {
                            Write-Log "Google Workspace account for $mail exists but was created more than 1 hour ago" "WARN"
                            return $false
                        }
                    }
                    catch {
                        Write-Log "WARNING: Could not parse creation time for $mail : $_" "WARN"
                    }
                }
                else {
                    Write-Log "WARNING: No creationTime found for $mail" "WARN"
                    return $true
                }
            }
            else {
                $attempt++
                if ($attempt -lt $maxAttempts) {
                    Write-Log "Google Workspace account for $mail not found yet. Waiting $CheckIntervalSeconds seconds... (Attempt $attempt of $maxAttempts)" "INFO"
                    Start-Sleep -Seconds $CheckIntervalSeconds
                }
            }
        }
        catch {
            Write-Log "ERROR checking Google Workspace for $mail : $_" "WARN"
            $attempt++
            if ($attempt -lt $maxAttempts) {
                Start-Sleep -Seconds $CheckIntervalSeconds
            }
        }
    }
    
    Write-Log "Google Workspace account for $mail not found after $maxAttempts attempts" "WARN"
    return $false
}

# Function to wait until a specific minute past the hour
function Wait-UntilMinutePastHour {
    param([int]$TargetMinute)
    
    $now = Get-Date
    $currentMinute = $now.Minute
    $currentSecond = $now.Second
    
    if ($currentMinute -lt $TargetMinute) {
        $minutesToWait = $TargetMinute - $currentMinute
        $secondsToWait = ($minutesToWait * 60) - $currentSecond
    }
    elseif ($currentMinute -gt $TargetMinute) {
        $minutesToWait = (60 - $currentMinute) + $TargetMinute
        $secondsToWait = ($minutesToWait * 60) - $currentSecond
    }
    else {
        if ($currentSecond -lt 5) {
            $secondsToWait = 5 - $currentSecond
        }
        else {
            $minutesRemaining = 60 - $currentMinute
            $secondsToWait = ($minutesRemaining * 60) + ($TargetMinute * 60) - $currentSecond
        }
    }
    
    if ($secondsToWait -gt 0) {
        $targetTime = $now.AddSeconds($secondsToWait)
        $timeStr = $targetTime.ToString('HH:mm:ss')
        $targetStr = $TargetMinute.ToString('00')
        Write-Log "Waiting until $timeStr (target: :$targetStr past the hour)..." "INFO"
        Start-Sleep -Seconds $secondsToWait
        $currentTimeStr = (Get-Date).ToString('HH:mm:ss')
        Write-Log "Reached target time: $currentTimeStr" "INFO"
    }
    else {
        $targetStr = $TargetMinute.ToString('00')
        Write-Log "Already at or past target time :$targetStr" "INFO"
    }
}

# Function to send welcome email
function Send-WelcomeEmail {
    param(
        [PSCustomObject]$User,
        [string]$Password
    )
    
    $Username = ($User.username).ToLower()
    $Firstname = $User.firstname
    $EmailTo = $User.personalemail
    
    if ([string]::IsNullOrWhiteSpace($EmailTo)) {
        Write-Log "No personal email address for $Username. Skipping email." "WARN"
        return $false
    }
    
    if ([string]::IsNullOrWhiteSpace($Password)) {
        Write-Log "ERROR: Password is null or empty for user $Username. Cannot send email." "ERROR"
        return $false
    }
    
    try {
        try {
            $securePassword = Get-Content $CredentialFile -ErrorAction Stop | ConvertTo-SecureString
            $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SmtpUsername, $securePassword
        }
        catch {
            Write-Log "ERROR: Failed to decrypt credential file. This usually means:" "ERROR"
            Write-Log "  - The file was encrypted by a different user account" "ERROR"
            Write-Log "  - The script is running in a different execution context" "ERROR"
            Write-Log "  - Error details: $_" "ERROR"
            Write-Log "  - Solution: Run the script as the same user who created the credential file, or recreate the credential file" "ERROR"
            Write-Log "  - To recreate: Run this command as the user who will run the script:" "ERROR"
            Write-Log "    Read-Host 'Enter SMTP password' -AsSecureString | ConvertFrom-SecureString | Out-File '$CredentialFile'" "ERROR"
            return $false
        }
        
        # Generic email template with placeholders
        $Body = @"
<!DOCTYPE html>
<html>
<head>
<style>
body { font-family: Arial, sans-serif; line-height: 1.6; }
a { color: #0066cc; text-decoration: none; }
a:hover { text-decoration: underline; }
.info { margin-bottom: 15px; }
.important { color: #d63333; font-weight: bold; }
</style>
</head>
<body>
<p>Dear $Firstname,</p>

<div class="info">
<p>The following is your $CompanyName account information:</p>
<p class="important">Please use <a href="$AccountPortalUrl">$AccountPortalUrl</a> to change your temporary password.</p>
<p>You must change your temporary password within 7 days or your account will be locked.</p>
</div>

<div class="info">
<p><strong>Username:</strong> $Username<br>
<strong>Email Address:</strong> $Username@$EmailDomain<br>
<strong>Temp Password:</strong> $Password</p>
</div>

<div class="info">
<h3>Quick Links:</h3>
<ul>
<li><a href="$EmailUrl">Email</a></li>
<li><a href="$CalendarUrl">Calendar</a></li>
<li><a href="$SupportUrl">Ticketing System</a></li>
</ul>
</div>

<div class="info">
<h3>Additional Systems (may be available after start date):</h3>
<ul>
<li><a href="$HrSystemUrl">HR System</a></li>
<li><a href="$TimesheetUrl">Timesheet</a></li>
<li><a href="$PerformanceUrl">Performance Management System</a></li>
</ul>
</div>

<div class="info">
<h3>IT security requirements:</h3>
<ol>
<li>Password Complexity is Required (password must include a combination of uppercase letters, lowercase letters, numbers, and special characters, and be at least 16 characters long)</li>
<li>After updating your password, <a href="$TwoFactorUrl">enable 2-Step Verification</a>. This must be done within 7 days.</li>
<li>Set up 2FA for SSO using <a href="$AuthenticatorUrl">authenticator applications</a>.</li>
</ol>
</div>

<p>If you have any questions please email <a href="mailto:$SupportEmail">$SupportEmail</a>.</p>

<p>Welcome aboard!<br>
$CompanyName IT Team</p>
</body>
</html>
"@
        
        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = $FromEmail
        $mailMessage.To.Add($EmailTo)
        if (-not [string]::IsNullOrWhiteSpace($CcEmail)) {
            $mailMessage.CC.Add($CcEmail)
        }
        $mailMessage.Subject = $EmailSubject
        $mailMessage.Body = $Body
        $mailMessage.IsBodyHtml = $true
        
        $SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
        $SMTPClient.EnableSsl = $true
        $SMTPClient.Credentials = $cred
        $SMTPClient.Send($mailMessage)
        
        Write-Log "Sent welcome email to $EmailTo for user $Username" "INFO"
        return $true
    }
    catch {
        Write-Log "ERROR sending email to $EmailTo for user $Username : $_" "ERROR"
        return $false
    }
}

#endregion

# Create authentication header
$AuthString = "$ZendeskEmail/token:$ZendeskApiToken"
$EncodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($AuthString))
$Headers = @{
    Authorization = "Basic $EncodedAuth"
    "Content-Type" = "application/json"
}

# Custom field IDs mapping
$FieldIDs = @{
    FirstName = $FieldID_FirstName
    LastName = $FieldID_LastName
    PersonalEmail = $FieldID_PersonalEmail
    Department = $FieldID_Department
    Manager = $FieldID_Manager
    JobTitle = $FieldID_JobTitle
    EmployeeType = $FieldID_EmployeeType
    EquipmentIssuedFrom = $FieldID_EquipmentIssuedFrom
}

# Function to extract custom field value from ticket
function Get-CustomFieldValue {
    param(
        [object]$Ticket,
        [long]$FieldID
    )
    
    if (-not $Ticket -or $FieldID -eq 0) {
        return $null
    }
    
    if ($Ticket.custom_fields) {
        $field = $Ticket.custom_fields | Where-Object { $_.id -eq $FieldID }
        if ($field -and $field.value) {
            return $field.value.ToString().Trim()
        }
    }
    return $null
}

# Function to format manager name as firstname.lastname (lowercase, no apostrophes or hyphens)
function Format-ManagerName {
    param([string]$ManagerName)
    
    if ([string]::IsNullOrWhiteSpace($ManagerName) -or $ManagerName -eq "Unknown") {
        return "Unknown"
    }
    
    $cleanedName = $ManagerName -replace "['-]", ""
    $nameParts = $cleanedName.Trim().Split(' ', [StringSplitOptions]::RemoveEmptyEntries)
    
    if ($nameParts.Count -eq 0) {
        return "Unknown"
    }
    
    $firstName = $nameParts[0].ToLower().Trim()
    
    if ($nameParts.Count -gt 1) {
        $lastName = ($nameParts[1..($nameParts.Count-1)] -join '').ToLower().Trim()
    } else {
        $lastName = $firstName
    }
    
    $formattedManager = "$firstName.$lastName"
    $formattedManager = $formattedManager -replace '[^a-z0-9.]', ''
    
    return $formattedManager
}

# Get tickets filtered by Form field
Write-Log "=== FETCHING TICKETS WITH FORM: $ZendeskFormName ===" "INFO"
$allTickets = @()
try {
    $query = "type:ticket form:`"$ZendeskFormName`""
    $encodedQuery = [System.Web.HttpUtility]::UrlEncode($query)
    $url = "https://$ZendeskUrl/api/v2/search.json?query=$encodedQuery&sort_by=updated_at&sort_order=desc"
    
    Write-Log "Searching for tickets with Form: $ZendeskFormName..." "INFO"
    
    $response = Invoke-RestMethod -Uri $url -Method GET -Headers $Headers -TimeoutSec 30 -ErrorAction Stop
    $allTicketsRaw = $response.results
    
    $allTickets = $allTicketsRaw | Where-Object { 
        $formValue = $null
        if ($_.form) { $formValue = $_.form }
        elseif ($_.ticket_form) { $formValue = $_.ticket_form.name }
        
        if ($formValue) {
            $formValue -match $ZendeskFormName -or $formValue -eq $ZendeskFormName
        } else {
            $true
        }
    }
    
    Write-Log "Found: $($allTickets.Count) tickets with $ZendeskFormName form" "INFO"
}
catch {
    Write-Log "ERROR: Failed to fetch tickets: $($_.Exception.Message)" "ERROR"
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" "ERROR"
    exit 1
}

if ($allTickets.Count -eq 0) {
    Write-Log "No tickets found with $ZendeskFormName form." "WARN"
    Write-Log "=== Script Completed (No tickets found) ===" "INFO"
    exit 0
}

# Filter by status (new and open only)
$newTickets = $allTickets | Where-Object { $_.status -eq "new" }
$openTickets = $allTickets | Where-Object { $_.status -eq "open" }

Write-Log "" "INFO"
Write-Log "=== NEW TICKETS ===" "INFO"
$newCount = @($newTickets).Count
Write-Log "Found: $newCount New tickets" "INFO"

foreach ($ticket in $newTickets) {
    $createdDate = if ($ticket.created_at) { [DateTime]::Parse($ticket.created_at).ToString("yyyy-MM-dd") } else { "Unknown" }
    $updatedDate = if ($ticket.updated_at) { [DateTime]::Parse($ticket.updated_at).ToString("yyyy-MM-dd HH:mm") } else { "Unknown" }
    Write-Log "  #$($ticket.id): $($ticket.subject) - Status: $($ticket.status) | Created: $createdDate | Updated: $updatedDate" "INFO"
}

Write-Log "" "INFO"
Write-Log "=== OPEN TICKETS ===" "INFO"
$openCount = @($openTickets).Count
Write-Log "Found: $openCount Open tickets" "INFO"

foreach ($ticket in $openTickets) {
    $createdDate = if ($ticket.created_at) { [DateTime]::Parse($ticket.created_at).ToString("yyyy-MM-dd") } else { "Unknown" }
    $updatedDate = if ($ticket.updated_at) { [DateTime]::Parse($ticket.updated_at).ToString("yyyy-MM-dd HH:mm") } else { "Unknown" }
    Write-Log "  #$($ticket.id): $($ticket.subject) - Status: $($ticket.status) | Created: $createdDate | Updated: $updatedDate" "INFO"
}

Write-Log "" "INFO"

$allFilteredTickets = @($newTickets) + @($openTickets)
$totalTickets = $allFilteredTickets.Count

Write-Log "" "INFO"
Write-Log "Found $totalTickets new/open tickets to process" "INFO"
Write-Log "" "INFO"

if ($totalTickets -eq 0) {
    Write-Log "No new or open tickets found. Exiting." "WARN"
    Write-Log "=== Script Completed (No tickets to export) ===" "INFO"
    exit 0
}

$ticketsToExport = $allFilteredTickets

Write-Log "=== EXTRACTING CUSTOM FIELDS AND EXPORTING TO CSV ===" "INFO"
Write-Log "Processing $totalTickets tickets..." "INFO"

# Create CSV data
$csvData = @()
foreach ($ticket in $ticketsToExport) {
    Write-Log "Processing Ticket #$($ticket.id) - $($ticket.subject)" "INFO"
    
    $fullTicket = $null
    try {
        $ticketUrl = "https://$ZendeskUrl/api/v2/tickets/$($ticket.id).json"
        $ticketResponse = Invoke-RestMethod -Uri $ticketUrl -Method GET -Headers $Headers -TimeoutSec 30 -ErrorAction Stop
        $fullTicket = $ticketResponse.ticket
        Write-Log "  Fetched full ticket details for #$($ticket.id)" "INFO"
    }
    catch {
        Write-Log "  Error: Could not fetch full ticket details for #$($ticket.id): $($_.Exception.Message)" "WARN"
        $fullTicket = $ticket
    }
    
    # Extract custom field values
    $firstName = Get-CustomFieldValue -Ticket $fullTicket -FieldID $FieldIDs.FirstName
    $lastName = Get-CustomFieldValue -Ticket $fullTicket -FieldID $FieldIDs.LastName
    $personalEmail = Get-CustomFieldValue -Ticket $fullTicket -FieldID $FieldIDs.PersonalEmail
    $department = Get-CustomFieldValue -Ticket $fullTicket -FieldID $FieldIDs.Department
    $manager = Get-CustomFieldValue -Ticket $fullTicket -FieldID $FieldIDs.Manager
    $jobTitle = Get-CustomFieldValue -Ticket $fullTicket -FieldID $FieldIDs.JobTitle
    $employeeTypeRaw = Get-CustomFieldValue -Ticket $fullTicket -FieldID $FieldIDs.EmployeeType
    $equipmentIssuedFrom = Get-CustomFieldValue -Ticket $fullTicket -FieldID $FieldIDs.EquipmentIssuedFrom
    
    # Set default values if fields are empty
    if ([string]::IsNullOrWhiteSpace($firstName)) { $firstName = "Unknown" }
    if ([string]::IsNullOrWhiteSpace($lastName)) { $lastName = "Unknown" }
    if ([string]::IsNullOrWhiteSpace($personalEmail)) { $personalEmail = "" }
    if ([string]::IsNullOrWhiteSpace($department)) { $department = "Unknown" }
    if ([string]::IsNullOrWhiteSpace($manager)) { $manager = "Unknown" }
    if ([string]::IsNullOrWhiteSpace($jobTitle)) { $jobTitle = "Unknown" }
    if ([string]::IsNullOrWhiteSpace($employeeTypeRaw)) { $employeeTypeRaw = "Unknown" }
    if ([string]::IsNullOrWhiteSpace($equipmentIssuedFrom)) { $equipmentIssuedFrom = "-" }
    
    $manager = Format-ManagerName -ManagerName $manager
    
    # Format EmployeeType: Full time → FT, Part time → PT, else → CT
    $employeeType = $employeeTypeRaw
    if ($employeeType -and $employeeType -ne "Unknown") {
        $employeeTypeLower = $employeeType.ToLower().Trim()
        if ($employeeTypeLower -like "*full time*" -or $employeeTypeLower -like "*fulltime*" -or $employeeTypeLower -like "*full_time*" -or $employeeTypeLower -like "*full-time*") {
            $employeeType = "FT"
        }
        elseif ($employeeTypeLower -like "*part time*" -or $employeeTypeLower -like "*parttime*" -or $employeeTypeLower -like "*part_time*" -or $employeeTypeLower -like "*part-time*") {
            $employeeType = "PT"
        }
        else {
            $employeeType = "CT"
        }
    }
    else {
        $employeeType = "CT"
    }
    
    $originalDepartment = $department
    
    # Truncate Department to first 10 characters if longer than 10
    if ($department -and $department -ne "Unknown" -and $department.Length -gt 10) {
        $department = $department.Substring(0, 10).Trim()
    }
    
    # Generate username from firstname.lastname
    $username = "$($firstName.ToLower().Trim()).$($lastName.ToLower().Trim())"
    $username = $username -replace '[^a-z0-9.]', ''
    
    # If email is empty, generate one from username
    if ([string]::IsNullOrWhiteSpace($personalEmail)) {
        $personalEmail = "$username@$EmailDomain"
    }
    
    # Group memberships are now determined by CSV flags only (removed company-specific logic)
    # Default to FALSE - these can be set manually in CSV or via Zendesk workflow after export
    # Create CSV row with generic group names
    $csvRow = [PSCustomObject]@{
        firstname = $firstName
        lastname = $lastName
        username = $username
        department = $department
        jobtitle = $jobTitle
        personalemail = $personalEmail
        employeetype = $employeeType
        manager = $manager
        ITEquipment = "FALSE"  # Default - set manually in CSV or via Zendesk workflow
        RemoteAccess = "FALSE"  # Default - set manually in CSV or via Zendesk workflow
        OfficeUsers = "FALSE"   # Default - set manually in CSV or via Zendesk workflow
    }
    
    $csvData += $csvRow
    
    Write-Log "  Extracted: $firstName $lastName ($username)" "INFO"
}

# Export to CSV
try {
    $headers = @('firstname', 'lastname', 'username', 'department', 'jobtitle', 'personalemail', 'employeetype', 'manager', 'ITEquipment', 'RemoteAccess', 'OfficeUsers')
    
    $csvLines = @()
    $csvLines += ($headers -join ",")
    
    foreach ($row in $csvData) {
        function Escape-CsvField {
            param([string]$field)
            if ($null -eq $field) { return "" }
            if ($field -match '["\r\n]' -or $field.Contains(',')) {
                return '"' + ($field -replace '"', '""') + '"'
            }
            return $field
        }
        
        $csvLine = @(
            (Escape-CsvField $row.firstname),
            (Escape-CsvField $row.lastname),
            (Escape-CsvField $row.username),
            (Escape-CsvField $row.department),
            (Escape-CsvField $row.jobtitle),
            (Escape-CsvField $row.personalemail),
            (Escape-CsvField $row.employeetype),
            (Escape-CsvField $row.manager),
            (Escape-CsvField $row.ITEquipment),
            (Escape-CsvField $row.RemoteAccess),
            (Escape-CsvField $row.OfficeUsers)
        )
        $csvLines += ($csvLine -join ",")
    }
    
    $csvLines | Out-File -FilePath $FullOutputPath -Encoding UTF8 -NoNewline:$false
    Write-Log "Successfully exported $($csvData.Count) tickets to CSV: $FullOutputPath" "INFO"
}
catch {
    Write-Log "ERROR: Failed to export CSV: $($_.Exception.Message)" "ERROR"
    Write-Log "Make sure the CSV file is not open in another application." "WARN"
    exit 1
}

Write-Log "" "INFO"
Write-Log "=== ZENDESK FETCH SUMMARY ===" "INFO"
Write-Log "Total tickets processed: $totalTickets ($newCount New + $openCount Open)" "INFO"
Write-Log "CSV file created: $FullOutputPath" "INFO"
Write-Log "Custom fields extracted successfully!" "INFO"

if ($csvData.Count -eq 0) {
    Write-Log "=== Script Completed (No users to process) ===" "INFO"
    exit 0
}

#region AD User Creation Phase
Write-Log "" "INFO"
Write-Log "=== AD USER CREATION PHASE ===" "INFO"

if (-not $ADModuleAvailable) {
    Write-Log "Active Directory module not available. Skipping AD user creation." "WARN"
    exit 0
}

$ADUsers = @()
try {
    $ADUsers = Import-Csv $FullOutputPath
    Write-Log "Loaded $($ADUsers.Count) users from CSV" "INFO"
}
catch {
    Write-Log "ERROR: Failed to import CSV file: $_" "ERROR"
    exit 1
}

$stats = @{
    Total = $ADUsers.Count
    Created = 0
    AlreadyExists = 0
    Failed = 0
    EmailsSent = 0
    EmailsSkipped = 0
}

$userPasswords = @{}

foreach ($User in $ADUsers) {
    $Username = ($User.username).ToLower()
    Write-Log "Processing user: $Username" "INFO"
    
    if (Test-ADUserExists -Username $Username) {
        Write-Log "User $Username already exists. Skipping AD creation." "WARN"
        $stats.AlreadyExists++
        $User | Add-Member -NotePropertyName "ADStatus" -NotePropertyValue "AlreadyExists" -Force
        continue
    }
    
    $Password = Generate-ComplexPassword
    $result = New-ADUserAccount -User $User -Password $Password
    
    if ($result.Success) {
        $stats.Created++
        $User | Add-Member -NotePropertyName "ADStatus" -NotePropertyValue "Created" -Force
        $User | Add-Member -NotePropertyName "GeneratedPassword" -NotePropertyValue $Password -Force
        $userPasswords[$Username] = $Password
        Write-Log "Successfully created AD user: $Username" "INFO"
    }
    else {
        $stats.Failed++
        $User | Add-Member -NotePropertyName "ADStatus" -NotePropertyValue "Failed" -Force
        $User | Add-Member -NotePropertyName "ErrorMessage" -NotePropertyValue $result.Message -Force
        Write-Log "Failed to create AD user: $Username - $($result.Message)" "ERROR"
    }
}

Write-Log "" "INFO"
Write-Log "=== AD CREATION SUMMARY ===" "INFO"
Write-Log "Total users: $($stats.Total)" "INFO"
Write-Log "Created: $($stats.Created)" "INFO"
Write-Log "Already existed: $($stats.AlreadyExists)" "INFO"
Write-Log "Failed: $($stats.Failed)" "INFO"
#endregion

#region Google Workspace Sync and Email Phase
if (-not [string]::IsNullOrWhiteSpace($GoogleWorkspaceDomain)) {
    Write-Log "" "INFO"
    Write-Log "=== GOOGLE WORKSPACE SYNC AND EMAIL PHASE ===" "INFO"
    
    Write-Log "Waiting until :03 past the hour to check Google Workspace sync..." "INFO"
    Write-Log "AD and Google Workspace sync occurs at :00 every hour" "INFO"
    Wait-UntilMinutePastHour -TargetMinute 3
    
    Write-Log "" "INFO"
    Write-Log "=== PHASE 1: GOOGLE WORKSPACE VERIFICATION (at :03) ===" "INFO"
    
    $verifiedUsers = @()
    foreach ($User in $ADUsers) {
        $Username = ($User.username).ToLower()
        
        if ($User.ADStatus -ne "Created") {
            Write-Log "Skipping Google Workspace check for $Username (Status: $($User.ADStatus))" "INFO"
            continue
        }
        
        if ([string]::IsNullOrWhiteSpace($User.personalemail)) {
            Write-Log "No personal email for $Username. Skipping." "WARN"
            $stats.EmailsSkipped++
            continue
        }
        
        Write-Log "Verifying Google Workspace sync for $Username..." "INFO"
        $gwsExists = Test-GoogleWorkspaceUser -Username $Username -MaxWaitMinutes 2 -CheckIntervalSeconds $GSuiteCheckIntervalSeconds
        
        if ($gwsExists) {
            $storedPassword = $null
            if ($userPasswords.ContainsKey($Username)) {
                $storedPassword = $userPasswords[$Username]
            }
            elseif (-not [string]::IsNullOrWhiteSpace($User.GeneratedPassword)) {
                $storedPassword = $User.GeneratedPassword
            }
            
            if ([string]::IsNullOrWhiteSpace($storedPassword)) {
                Write-Log "WARNING: Password not found for $Username. Cannot send email." "ERROR"
                $stats.EmailsSkipped++
                continue
            }
            
            $User | Add-Member -NotePropertyName "GeneratedPassword" -NotePropertyValue $storedPassword -Force
            $userPasswords[$Username] = $storedPassword
            Write-Log "Google Workspace verified for $Username. Will send email at :05" "INFO"
            $User | Add-Member -NotePropertyName "GWSVerified" -NotePropertyValue $true -Force
            $verifiedUsers += $User
        }
        else {
            Write-Log "User $Username not found in Google Workspace. Email will not be sent." "WARN"
            $User | Add-Member -NotePropertyName "GWSVerified" -NotePropertyValue $false -Force
            $stats.EmailsSkipped++
        }
    }
    
    Write-Log "" "INFO"
    Write-Log "Waiting until :05 past the hour to send emails..." "INFO"
    Wait-UntilMinutePastHour -TargetMinute 5
    
    Write-Log "" "INFO"
    Write-Log "=== PHASE 2: EMAIL SENDING (at :05) ===" "INFO"
    
    foreach ($User in $verifiedUsers) {
        $Username = ($User.username).ToLower()
        
        if ($User.GWSVerified -ne $true) {
            Write-Log "Skipping email for $Username - Google Workspace not verified" "WARN"
            $stats.EmailsSkipped++
            continue
        }
        
        $password = $null
        if ($userPasswords.ContainsKey($Username)) {
            $password = $userPasswords[$Username]
        }
        elseif (-not [string]::IsNullOrWhiteSpace($User.GeneratedPassword)) {
            $password = $User.GeneratedPassword
        }
        
        if ([string]::IsNullOrWhiteSpace($password)) {
            Write-Log "ERROR: Password not found for user $Username. Cannot send email." "ERROR"
            $stats.EmailsSkipped++
            continue
        }
        
        Write-Log "Sending welcome email to $($User.personalemail) for user $Username..." "INFO"
        $emailSent = Send-WelcomeEmail -User $User -Password $password
        
        if ($emailSent) {
            $stats.EmailsSent++
            Write-Log "Email sent successfully to $($User.personalemail)" "INFO"
        }
        else {
            $stats.EmailsSkipped++
            Write-Log "Failed to send email to $($User.personalemail)" "ERROR"
        }
    }
    
    Write-Log "" "INFO"
    Write-Log "=== EMAIL SUMMARY ===" "INFO"
    Write-Log "Emails sent: $($stats.EmailsSent)" "INFO"
    Write-Log "Emails skipped: $($stats.EmailsSkipped)" "INFO"
} else {
    Write-Log "" "INFO"
    Write-Log "Google Workspace domain not configured. Skipping Google Workspace verification and email sending." "INFO"
    Write-Log "To enable email sending, configure GoogleWorkspaceDomain parameter." "INFO"
}

#endregion

# Final summary
Write-Log "" "INFO"
Write-Log "=== FINAL SUMMARY ===" "INFO"
Write-Log "Total users processed: $($stats.Total)" "INFO"
Write-Log "AD users created: $($stats.Created)" "INFO"
Write-Log "Users already existed: $($stats.AlreadyExists)" "INFO"
Write-Log "AD creation failed: $($stats.Failed)" "INFO"
Write-Log "Welcome emails sent: $($stats.EmailsSent)" "INFO"
Write-Log "Emails skipped: $($stats.EmailsSkipped)" "INFO"

Write-Log "=== Script Completed Successfully ===" "INFO"
exit 0
