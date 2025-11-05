# Fetch Zendesk tickets by Form field with custom fields and export to CSV
# Only fetches 'new' and 'open' tickets with one or zero comments (either public or private)
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
    
    # Email Domain (for generating email if PersonalEmail is empty)
    [Parameter(Mandatory=$true)]
    [string]$EmailDomain,  # e.g., "company.com"
    
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

Write-Log "=== Script Started ===" "INFO"
Write-Log "Script: $ScriptName" "INFO"
Write-Log "Script Path: $ScriptPath" "INFO"
Write-Log "Output CSV: $FullOutputPath" "INFO"
Write-Log "Log File: $LogFile" "INFO"
Write-Log "Fetching 'new' and 'open' tickets with form: $ZendeskFormName (one or zero comments only)..." "INFO"
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

# Function to get comment counts for a ticket
function Get-TicketCommentCounts {
    param([long]$TicketId)
    
    try {
        $commentsUrl = "https://$ZendeskUrl/api/v2/tickets/$TicketId/comments.json"
        $commentsResponse = Invoke-RestMethod -Uri $commentsUrl -Method GET -Headers $Headers -TimeoutSec 30 -ErrorAction Stop
        
        $publicCount = 0
        $privateCount = 0
        
        if ($commentsResponse.comments) {
            foreach ($comment in $commentsResponse.comments) {
                if ($comment.public -eq $true) {
                    $publicCount++
                } else {
                    $privateCount++
                }
            }
        }
        
        return @{
            Public = $publicCount
            Private = $privateCount
            Total = $publicCount + $privateCount
        }
    }
    catch {
        Write-Log "Warning: Could not fetch comments for ticket #${TicketId}: $($_.Exception.Message)" "WARN"
        return $null
    }
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

Write-Log "=== FILTERING TICKETS WITH ONE OR ZERO COMMENTS ===" "INFO"
Write-Log "Checking comment counts for $($allFilteredTickets.Count) tickets..." "INFO"

$ticketsToExport = @()
foreach ($ticket in $allFilteredTickets) {
    Write-Log "  Checking Ticket #$($ticket.id)..." "INFO"
    $commentCounts = Get-TicketCommentCounts -TicketId $ticket.id
    
    if ($null -eq $commentCounts) {
        Write-Log "    Skipped - could not fetch comments" "WARN"
        continue
    }
    
    $totalComments = $commentCounts.Total
    if ($totalComments -le 1) {
        Write-Log "    OK - $totalComments comment(s): $($commentCounts.Public) public, $($commentCounts.Private) private" "INFO"
        $ticketsToExport += $ticket
    } else {
        Write-Log "    Excluded - Has $totalComments comments: $($commentCounts.Public) public, $($commentCounts.Private) private" "INFO"
    }
}

$totalTickets = $ticketsToExport.Count

Write-Log "" "INFO"
Write-Log "Found $totalTickets tickets with one or zero comments (out of $($allFilteredTickets.Count) new/open tickets)" "INFO"
Write-Log "" "INFO"

if ($totalTickets -eq 0) {
    Write-Log "No tickets found with one or zero comments. Exiting." "WARN"
    Write-Log "=== Script Completed (No tickets to export) ===" "INFO"
    exit 0
}

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
    
    # Group memberships default to FALSE - these can be set manually in CSV or via Zendesk workflow
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
Write-Log "=== SUMMARY ===" "INFO"
Write-Log "Total tickets processed: $totalTickets ($newCount New + $openCount Open, with one or zero comments)" "INFO"
Write-Log "CSV file created: $FullOutputPath" "INFO"
Write-Log "Custom fields extracted successfully!" "INFO"

if ($csvData.Count -gt 0) {
    Write-Log "" "INFO"
    Write-Log "Sample exported data:" "INFO"
    $sampleData = $csvData | Select-Object -First 5 | Select-Object firstname, lastname, username, department, jobtitle, employeetype
    foreach ($row in $sampleData) {
        Write-Log "  $($row.firstname) $($row.lastname) ($($row.username)) - $($row.department) - $($row.jobtitle)" "INFO"
    }
}

Write-Log "=== Script Completed Successfully ===" "INFO"
exit 0
