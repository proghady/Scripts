# ==============================
# ONE-TIME SETUP (RUN AS TEAMS ADMIN)
# ==============================

Install-Module MicrosoftTeams -Force -AllowClobber
Import-Module MicrosoftTeams
Connect-MicrosoftTeams

# Your Entra App (Client) ID:
$AppId = "e1b699bb-4f06-43bc-8334-1848296c7f2f"

# Create policy
New-CsApplicationAccessPolicy `
  -Identity "GraphOnlineMeetingsPolicy" `
  -AppIds $AppId `
  -Description "Allow app-only Graph access to online meetings + attendance artifacts"

# OPTION 1 (Recommended for testing): Grant to specific users you will call in /users/{userId}/...
# Example:
Grant-CsApplicationAccessPolicy `
  -PolicyName "GraphOnlineMeetingsPolicy" `
  -Identity "admin@M365EDU063925.OnMicrosoft.com"

# OPTION 2 (Tenant-wide): Uncomment ONLY if you really want global access
# Grant-CsApplicationAccessPolicy -PolicyName "GraphOnlineMeetingsPolicy" -Global

# Verify assignment
Get-CsOnlineUser -Identity "admin@M365EDU063925.OnMicrosoft.com" |  Select-Object DisplayName, UserPrincipalName, ApplicationAccessPolicy
