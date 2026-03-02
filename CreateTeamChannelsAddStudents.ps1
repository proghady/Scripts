<#
.SYNOPSIS
Creates a Microsoft Teams Class Team, adds channels from interactive input, 
then adds students in bulk from a CSV.

.REQUIREMENTS
- Microsoft.Graph PowerShell module
- Permissions to create teams and add members
#>

# ---------------------------
# Helper Functions
# ---------------------------

function Read-MultiLine {
    param(
        [Parameter(Mandatory=$true)][string]$Prompt,
        [string]$StopWord = "done"
    )
    Write-Host $Prompt -ForegroundColor Cyan
    Write-Host "Type one item per line. Type '$StopWord' when finished." -ForegroundColor DarkCyan
    $items = @()
    while ($true) {
        $line = Read-Host ">"
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line.Trim().ToLower() -eq $StopWord.ToLower()) { break }
        $items += $line.Trim()
    }
    return $items
}

function Resolve-UserId {
    param([Parameter(Mandatory=$true)][string]$UserPrincipalName)
    try {
        $u = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
        return $u.Id
    } catch {
        Write-Warning "Could not resolve user: $UserPrincipalName"
        return $null
    }
}

# ---------------------------
# Connect to Microsoft Graph
# ---------------------------

$scopes = @(
    "Group.ReadWrite.All",
    "Team.Create",
    "TeamMember.ReadWrite.All",
    "Channel.Create",
    "User.Read.All"
)

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
Connect-MgGraph -Scopes $scopes | Out-Null

# ---------------------------
# Interactive Inputs
# ---------------------------

$teamName        = Read-Host "Enter Class Team Display Name"
$teamDescription = Read-Host "Enter Team Description (optional)"
$ownerUpn        = Read-Host "Enter Owner (Teacher) UPN (e.g., teacher@domain.com)"
$channels        = Read-MultiLine -Prompt "Enter channel names to create" -StopWord "done"

$csvPath = Read-Host "Enter full path to students CSV (e.g., C:\Temp\students.csv)"

if (-not (Test-Path $csvPath)) {
    throw "CSV path not found: $csvPath"
}

# ---------------------------
# Create Class Team
# ---------------------------

Write-Host "Resolving owner..." -ForegroundColor Yellow
$ownerId = Resolve-UserId -UserPrincipalName $ownerUpn
if (-not $ownerId) { throw "Owner UPN not found in tenant: $ownerUpn" }

Write-Host "Creating Class Team '$teamName'..." -ForegroundColor Yellow

# Education Class team template: "educationClass"
# (Using Graph Team template binding)
$teamBody = @{
    "template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('educationClass')"
    displayName           = $teamName
    description           = $teamDescription
    members               = @(
        @{
            "@odata.type"    = "#microsoft.graph.aadUserConversationMember"
            roles            = @("owner")
            "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$ownerId')"
        }
    )
}

#$team = New-MgTeam -BodyParameter $teamBody
#$teamId = $team.Id

# Create the team via raw request so we can read the Location header (async operation)
$uri = "https://graph.microsoft.com/v1.0/teams"

# Make sure these are strings (defensive)
$teamBody.displayName = [string]$teamBody.displayName
$teamBody.description = [string]$teamBody.description

# Convert to JSON with enough depth for members array
$jsonBody = $teamBody | ConvertTo-Json -Depth 20 -Compress

# POST and capture HTTP response headers (Location)
$response = Invoke-MgGraphRequest `
    -Uri $uri `
    -Method POST `
    -Body $jsonBody `
    -ContentType "application/json" `
    -OutputType HttpResponseMessage



# POST already done, $response is HttpResponseMessage

# Extract Location robustly (works for relative Location headers)
$location = $null

if ($response.Headers.Location) {
    $location = $response.Headers.Location.ToString()
} else {
    $vals = $null
    if ($response.Headers.TryGetValues("Location", [ref]$vals)) {
        $location = ($vals | Select-Object -First 1)
    }
}

if (-not $location) {
    throw "Create team returned no Location header (cannot track async operation)."
}

# If Location is relative, make it absolute
if ($location.StartsWith("/")) {
    $location = "https://graph.microsoft.com/v1.0$location"
}

Write-Host "Async operation URL: $location" -ForegroundColor DarkGray

# Poll operation until succeeded/failed, then read targetResourceId (TeamId)
$teamId = $null
for ($i = 1; $i -le 60; $i++) {
    $op = Invoke-MgGraphRequest -Uri $location -Method GET
    $status = $op.status

    if ($status -match 'succeeded|success') {
        $teamId = $op.targetResourceId
        break
    }
    elseif ($status -match 'failed') {
        throw "Team creation failed: $($op.error.code) - $($op.error.message)"
    }

    Start-Sleep -Seconds 5
}

if (-not $teamId) {
    throw "Timed out waiting for TeamId (targetResourceId) from async operation."
}

Write-Host "Team provisioned ✅ TeamId: $teamId" -ForegroundColor Green

#Write-Host "Team created. TeamId: $teamId" -ForegroundColor Green
#Write-Host "Waiting briefly for team provisioning..." -ForegroundColor Yellow
#Start-Sleep -Seconds 15

# ---------------------------
# Add Channels
# ---------------------------

if ($channels.Count -gt 0) {
    Write-Host "Creating channels..." -ForegroundColor Yellow
    foreach ($ch in $channels) {
        try {
            # "General" already exists by default
            if ($ch.Trim().ToLower() -eq "general") {
                Write-Host "Skipping 'General' (already exists)" -ForegroundColor DarkYellow
                continue
            }

            New-MgTeamChannel -TeamId $teamId -DisplayName $ch -MembershipType "standard" | Out-Null
            Write-Host "Created channel: $ch" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to create channel '$ch': $($_.Exception.Message)"
        }
    }
} else {
    Write-Host "No channels provided. Skipping channel creation." -ForegroundColor DarkYellow
}

# ---------------------------
# Add Students from CSV
# ---------------------------

Write-Host "Importing students from CSV..." -ForegroundColor Yellow
$students = Import-Csv $csvPath

# Accept either UserPrincipalName or Email column
$upnColumn = $null
if ($students[0].PSObject.Properties.Name -contains "UserPrincipalName") { $upnColumn = "UserPrincipalName" }
elseif ($students[0].PSObject.Properties.Name -contains "Email") { $upnColumn = "Email" }

if (-not $upnColumn) {
    throw "CSV must contain a 'UserPrincipalName' or 'Email' column."
}

Write-Host "Adding students to the team..." -ForegroundColor Yellow

foreach ($row in $students) {
    $studentUpn = ($row.$upnColumn).Trim()
    if ([string]::IsNullOrWhiteSpace($studentUpn)) { continue }

    $studentId = Resolve-UserId -UserPrincipalName $studentUpn
    if (-not $studentId) { 
        Write-Warning "Skipping (not found): $studentUpn"
        continue 
    }

    try {
        $memberBody = @{
            "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
            roles             = @() # empty roles => member
            "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$studentId')"
        }

        New-MgTeamMember -TeamId $teamId -BodyParameter $memberBody | Out-Null
        Write-Host "Added student: $studentUpn" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to add '$studentUpn': $($_.Exception.Message)"
    }
}

Write-Host "Done, Class Team created, channels added, students enrolled." -ForegroundColor Green
