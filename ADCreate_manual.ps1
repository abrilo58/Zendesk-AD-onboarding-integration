<#
.SYNOPSIS
    Creates AD users with group membership and sends HTML-formatted welcome emails from CSV
.DESCRIPTION
    Creates users from CSV, adds to groups, and sends emails with clickable links
.NOTES
    CSV must include columns: ITEquipment, RemoteAccess, and OfficeUsers
    Password column is not required - passwords are auto-generated
#>

param(
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
    [Parameter(Mandatory=$true)]
    [string]$CSVfile,
    
    [string]$LogDir = "",
    [string]$BaseDirectory = "",  # If empty, uses script location
    
    [int]$WaitMinutesBeforeEmail = 15  # Wait time before sending emails (for Google Workspace sync)
)

# Configuration
if ([string]::IsNullOrWhiteSpace($BaseDirectory)) {
    $ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptPath = $BaseDirectory
}

if ([string]::IsNullOrWhiteSpace($LogDir)) {
    $logdir = Join-Path $ScriptPath "logs"
} else {
    $logdir = $LogDir
}

# Create log directory if needed
if (-not (Test-Path $logdir)) {
    New-Item -ItemType Directory -Path $logdir -Force | Out-Null
}

$LogFile = Join-Path $logdir "scriptLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

# Function to write to log
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    Add-Content $LogFile $logMessage
    Write-Host $logMessage
}

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

Write-Log "=== Script Started ==="
Write-Log "CSV File: $CSVfile"
Write-Log "AD OU: $AD_OU"
Write-Log "AD Domain: $AD_DomainName"
Write-Log "Email Domain: $EmailDomain"

# Import CSV data
try {
    $ADUsers = Import-Csv $CSVfile
    Write-Log "Loaded $($ADUsers.Count) users from CSV"
}
catch {
    Write-Log "ERROR: Failed to import CSV file: $_"
    exit 1
}

# Main processing loop - First create all users and add to groups
foreach ($User in $ADUsers) {
    $Username = $($User.username).ToLower()
    $Password = Generate-ComplexPassword
    $Firstname = $User.firstname
    $Lastname = $User.lastname
    
    # Store password in user object for email later
    $User | Add-Member -NotePropertyName "GeneratedPassword" -NotePropertyValue $Password -Force
    
    # Get group membership flags (generic group names)
    $AddToItEquipment = $User.ITEquipment -eq "TRUE"
    $AddToRemoteAccess = $User.RemoteAccess -eq "TRUE"
    $AddToOfficeUsers = $User.OfficeUsers -eq "TRUE"
    
    Write-Log ""
    Write-Log "----- Processing $Username at $(Get-Date) -----"
    Write-Log "Group Flags - ITEquipment: $AddToItEquipment, RemoteAccess: $AddToRemoteAccess, OfficeUsers: $AddToOfficeUsers"
    
    # Check if user exists
    if (Get-ADUser -Filter {SamAccountName -eq $Username} -ErrorAction SilentlyContinue) {
        Write-Log "User $Username already exists"
        $User | Add-Member -NotePropertyName "error" -NotePropertyValue "User already exists" -Force
        continue
    }
    
    # Create new user
    try {
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
            -OtherAttributes @{'EmployeeType' = $User.employeetype; 'secondaryMail' = $User.personalemail}
        
        Write-Log "Successfully created $Username"
        
        # Add to groups - including mandatory MFA group
        try {
            Add-ADGroupMember -Identity $MfaGroupName -Members $Username -ErrorAction Stop
            Write-Log "Added to mandatory $MfaGroupName group"
        }
        catch {
            Write-Log "WARNING: Failed to add to $MfaGroupName group: $_"
        }
        
        if ($AddToItEquipment) {
            Add-ADGroupMember -Identity $ItEquipmentGroupName -Members $Username -ErrorAction SilentlyContinue
            Write-Log "Added to $ItEquipmentGroupName group"
        }
        if ($AddToRemoteAccess) {
            Add-ADGroupMember -Identity $RemoteAccessGroupName -Members $Username -ErrorAction SilentlyContinue
            Write-Log "Added to $RemoteAccessGroupName group"
        }
        if ($AddToOfficeUsers) {
            Add-ADGroupMember -Identity $OfficeLocationGroupName -Members $Username -ErrorAction SilentlyContinue
            Write-Log "Added to $OfficeLocationGroupName group"
        }
    }
    catch {
        Write-Log "ERROR creating $Username : $_"
        $User | Add-Member -NotePropertyName "error" -NotePropertyValue $_ -Force
    }
}

# Wait before sending emails (for Google Workspace sync)
if ($WaitMinutesBeforeEmail -gt 0) {
    Write-Host "Waiting $WaitMinutesBeforeEmail minutes before sending emails..." -ForegroundColor Yellow
    Write-Log "Waiting $WaitMinutesBeforeEmail minutes before sending emails..."
    Start-Sleep -Seconds ($WaitMinutesBeforeEmail * 60)
}

# Email sending loop after delay
foreach ($User in $ADUsers) {
    $Username = $($User.username).ToLower()
    $Firstname = $User.firstname
    $mail = "$Username@$EmailDomain"
    
    # Only process users that were successfully created (no error) and have personal email
    if (-not $User.error -and $User.personalemail) {
        try {
            # Check if the user account is in Google Workspace using 'gam info user'
            $GWSuitCreated = $false
            
            if (-not [string]::IsNullOrWhiteSpace($GoogleWorkspaceDomain)) {
                try {
                    $GSuitInfo = cmd /c "$GamCommand info user $mail" 2>&1 | Out-String
                    if ($GSuitInfo -match "User: $mail") {
                        if ($GSuitInfo -match '(?i)creation[\s-]time[:\s]*([0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}\.[0-9]+Z)') {
                            $creationTimeStr = $matches[1].Trim()
                            try {
                                $creationTime = [datetime]::ParseExact($creationTimeStr, "yyyy-MM-ddTHH:mm:ss.fffZ", $null)
                                $timeDifference = (Get-Date) - $creationTime
                                
                                if ($timeDifference.TotalHours -le 1) {
                                    $GWSuitCreated = $true
                                    Write-Log "Google Workspace account for $mail exists and was created in the last hour"
                                } else {
                                    Write-Log "Google Workspace account for $mail exists but was created more than 1 hour ago"
                                    continue
                                }
                            } catch {
                                Write-Log "WARNING: Could not parse creation time for ${mail}: $_"
                                continue
                            }
                        } else {
                            Write-Log "WARNING: No creationTime found for $mail"
                            continue
                        }
                    } else {
                        Write-Log "Google Workspace account for $mail not found"
                        continue
                    }
                } catch {
                    Write-Log "ERROR checking Google Workspace for ${mail}: $_"
                    continue
                }
            } else {
                # If Google Workspace domain not configured, skip check and proceed
                $GWSuitCreated = $true
                Write-Log "Google Workspace domain not configured. Skipping verification for $mail"
            }
            
            # Only proceed with email if Google Workspace check passed (or skipped)
            if ($GWSuitCreated) {
                try {
                    $securePassword = Get-Content $CredentialFile -ErrorAction Stop | ConvertTo-SecureString
                    $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SmtpUsername, $securePassword
                }
                catch {
                    Write-Log "ERROR: Failed to decrypt credential file: $_"
                    Write-Log "Solution: Run the script as the same user who created the credential file, or recreate the credential file"
                    Write-Log "To recreate: Read-Host 'Enter SMTP password' -AsSecureString | ConvertFrom-SecureString | Out-File '$CredentialFile'"
                    continue
                }
                
                $EmailTo = $User.personalemail
                $Subject = $EmailSubject
                
                # Generic HTML-formatted email body with placeholders
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
<strong>Temp Password:</strong> $($User.GeneratedPassword)</p>
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
                
                # Create and send email
                $mailMessage = New-Object System.Net.Mail.MailMessage
                $mailMessage.From = $FromEmail
                $mailMessage.To.Add($EmailTo)
                if (-not [string]::IsNullOrWhiteSpace($CcEmail)) {
                    $mailMessage.CC.Add($CcEmail)
                }
                $mailMessage.Subject = $Subject
                $mailMessage.Body = $Body
                $mailMessage.IsBodyHtml = $true
                
                $SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
                $SMTPClient.EnableSsl = $true
                $SMTPClient.Credentials = $cred
                $SMTPClient.Send($mailMessage)
                
                Write-Log "Sent HTML welcome email to $EmailTo (CC: $CcEmail) after Google Workspace verification"
            }
        }
        catch {
            Write-Log "Failed to process Google Workspace check or send email to $($User.personalemail): $_"
        }
    }
}

Write-Host "Script completed. See $LogFile for details." -ForegroundColor Cyan
Write-Log "=== Script Completed ==="
