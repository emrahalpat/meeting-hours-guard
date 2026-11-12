# Connects to Exchange Online and shows the settings that matter for Outlook add-ins.
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$false

Write-Host "`n== Signed in as ==" -ForegroundColor Cyan
Get-ConnectionInformation | Format-List UserPrincipalName, TenantID, Organization

Write-Host "== Add-ins allowed in this tenant? (AppsForOfficeEnabled must be True) ==" -ForegroundColor Cyan
Get-OrganizationConfig | Format-List AppsForOfficeEnabled

Write-Host "== Add-ins already deployed by an admin ==" -ForegroundColor Cyan
Get-App -OrganizationApp | Format-Table DisplayName, Enabled, ProvidedTo, DefaultStateForUser -AutoSize

Write-Host "CONNECT-DONE" -ForegroundColor Green
