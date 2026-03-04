# Create Team, Channels, and Add Students (PowerShell)

## Overview

`CreateTeamChannelsAddStudents.ps1` is a PowerShell automation script that streamlines Microsoft Teams provisioning by:

*   Creating a Microsoft Teams team
*   Creating multiple channels within the team
*   Adding students (users) to the team from a CSV file
*   Handling common input and user validation scenarios

This script is designed for **education and enterprise tenants** where bulk Team creation and enrollment is required.

***

## Use Cases

*   University / School IT onboarding
*   Course or class team provisioning
*   Bulk student enrollment into Teams
*   Reducing manual effort in Teams Admin Center

***

## Prerequisites

Before running the script, ensure the following:

### 1. PowerShell

*   PowerShell 5.1 or PowerShell 7+

### 2. Required Permissions

The account or app running the script must have permission to:

*   Create Microsoft 365 Groups / Teams
*   Add users to Teams
*   Read user objects in Microsoft Entra ID

> Typically: **Teams Administrator** or **Global Administrator**

### 3. Microsoft Graph / Teams Modules

Ensure required modules are installed and available in your environment (as used by the script).

***

## Input Files

### Students CSV File

The script reads users from a CSV file.

**Example: `students.csv`**

```csv
UserPrincipalName
student1@contoso.onmicrosoft.com
student2@contoso.onmicrosoft.com
student3@contoso.onmicrosoft.com
```

**Notes**

*   One user per line
*   Invalid or non‑existing UPNs are safely handled by the script
*   Duplicate users are ignored

***

## Script Parameters (High Level)

The script prompts for or accepts values such as:

*   **Team Display Name**
*   **Team Description**
*   **Channel Names** (one or multiple)
*   **CSV Path** containing users

> Refer to inline comments in the script for exact parameter names and usage.

***

## How It Works (High Level Flow)

1.  Creates a new Microsoft Team
2.  Waits until the Team is fully provisioned
3.  Creates standard channels inside the Team
4.  Reads users from the CSV file
5.  Adds users to the Team
6.  Skips invalid or duplicate users
7.  Outputs progress and status messages

***

## Error Handling & Validation

The script includes handling for:

*   Invalid or malformed UPNs
*   Duplicate users
*   Timing issues during Team creation
*   Missing or empty CSV files

This ensures the script can run end‑to‑end without failing on individual user errors.

***

## Example Execution

```powershell
.\CreateTeamChannelsAddStudents.ps1
```

Follow the prompts to:

*   Enter Team details
*   Provide channel names
*   Specify the CSV file path

***

## Output

*   Real‑time progress messages in the console
*   Confirmation of:
    *   Team creation
    *   Channel creation
    *   User additions

***

## Repository Structure

```text
Scripts/
│
├── CreateTeamChannelsAddStudents.ps1
├── README.md
└── students.csv (example)
```

***

## Security Notes

*   Do **not** commit real student or user data to GitHub
*   Use test accounts for validation
*   Store secrets securely if app‑based authentication is used

***


## License

This project is provided as‑is for educational and automation purposes.

***
