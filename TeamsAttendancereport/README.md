# Microsoft Teams – Meeting Attendance Report (Organizer‑Only)

This PowerShell script extracts **Microsoft Teams meeting attendance counts** for meetings **organized by specified users**, using **Microsoft Graph (application permissions)**.

The script is designed to run **without a signed‑in user** (app‑only authentication) and is suitable for **tenant‑wide reporting, automation, and compliance scenarios**.

***

## ✅ What the Script Does

*   Reads a list of users (organizers) from a CSV file
*   Retrieves **today’s meetings** from each user’s calendar
*   **Processes only meetings organized by that user** (Option B)
*   Resolves the Teams **OnlineMeetingId** using the meeting `joinWebUrl`
*   Retrieves **attendance reports** for completed meetings
*   Counts the number of attendees per meeting
*   Deduplicates results
*   Exports a clean CSV report

✅ This design avoids common Graph errors such as:  
`3003: User does not have access to lookup meeting`

***

## 🧠 Why Organizer‑Only (Option B)

Microsoft Graph enforces **meeting ownership rules**:

*   Attendance reports are reliably accessible **only for meetings where the user is the organizer (or co‑organizer)**
*   Attendee‑only access (especially via distribution lists or recurring meetings) can fail with `3003`

👉 This script intentionally processes **organizer‑owned meetings only** to ensure stability and correctness.

***

## 📁 Repository Structure

    TeamsAttendancereport/
    │
    ├── Extract Meetings attendees.ps1
    ├── ONE-Time Setup.ps1
    ├── users.csv
    └── README.md

***

## 📄 CSV Input Format

Create a CSV file containing the meeting organizers:

```csv
UserPrincipalName
admin@contoso.com
faculty1@contoso.com
faculty2@contoso.com
```

Supported column names:

*   `UserPrincipalName`
*   `UPN`

***

## 🔐 Prerequisites

### 1️⃣ Microsoft Entra ID App Registration

Create an app registration and grant **Application permissions**:

| Permission                     | Type        |
| ------------------------------ | ----------- |
| Calendars.Read                 | Application |
| OnlineMeetings.Read.All        | Application |
| OnlineMeetingArtifact.Read.All | Application |
| User.Read.All                  | Application |

✅ **Admin consent is required**

***

### 2️⃣ Teams Application Access Policy (Required)

Microsoft Teams **requires an Application Access Policy** for app‑only access to online meetings.

#### Create policy (once):

```powershell
New-CsApplicationAccessPolicy `
  -Identity "GraphOnlineMeetingsPolicy" `
  -AppIds "<APP_CLIENT_ID>" `
  -Description "Allow app-only Graph access to online meetings + attendance artifacts"
```

#### Grant globally (recommended for reporting):

```powershell
Grant-CsApplicationAccessPolicy `
  -PolicyName "GraphOnlineMeetingsPolicy" `
  -Global
```

⏳ **Important:**  
Policy propagation can take **up to 30 minutes** before Graph calls succeed.

***

## ▶️ How to Run the Script

1.  Open **PowerShell 7+** (recommended)
2.  Update the configuration section in the script:

```powershell
$TenantId     = "<TENANT_ID>"
$ClientId     = "<APP_CLIENT_ID>"
$ClientSecret = "<APP_SECRET>"
$UsersCsv     = "C:\Temp\users.csv"
$OutputFile   = "C:\Temp\Meetings_AttendeeCount_Today.csv"
```

3.  Run the script:

```powershell
.\Extract Meetings attendees.ps1
```

***

## 📤 Output Example

```csv
OrganizerUPN,MeetingSubject,MeetingStart,MeetingEnd,AttendeeCount,OnlineMeetingId
admin@contoso.com,Weekly Faculty Sync,2026-03-03T09:00,2026-03-03T10:00,27,19:meeting_...
```

Each row represents **one completed meeting session**.

***

## ⚠️ Known Limitations

*   Only **completed meetings** generate attendance reports
*   Only **organizer‑owned meetings** are processed
*   Channel meetings have limited support for attendance report retrieval
*   `attendanceReports` listing returns **up to 50 most recent reports**
*   External or cross‑tenant organizers are skipped
*   Microsoft Graph API throttling https://learn.microsoft.com/en-us/graph/throttling 

These behaviors are enforced by Microsoft Graph and are **by design**.

***

## 🛠 Troubleshooting

### ❌ `No application access policy found for this app`

✅ Ensure:

*   Policy exists
*   Correct AppId is in the policy
*   Policy is granted (`-Global` or per user)
*   Wait up to **30 minutes** after changes

***

### ❌ `3003: User does not have access to lookup meeting`

✅ Expected if:

*   User is not the organizer  
    ✅ Fixed by **Organizer‑Only (Option B)** logic

***

## 📚 References

*   Microsoft Graph – Online Meeting Attendance Reports
*   Microsoft Teams – Application Access Policy
*   Microsoft Entra ID – App‑Only Authentication

