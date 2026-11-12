# Checks that Meeting Hours Guard is deployed and visible to a mailbox.
param([string]$Mailbox = "emrah@alpat.com.tr")
Import-Module ExchangeOnlineManagement
if (-not (Get-ConnectionInformation)) { Connect-ExchangeOnline -ShowBanner:$false }

Write-Host "`n== Org deployment ==" -ForegroundColor Cyan
Get-App -OrganizationApp | Where-Object DisplayName -like "Meeting Hours*" |
  Format-List DisplayName, AppId, AppVersion, Enabled, ProvidedTo, UserList, DefaultStateForUser, ManifestXml

Write-Host "== Add-ins visible to $Mailbox ==" -ForegroundColor Cyan
Get-App -Mailbox $Mailbox | Where-Object DisplayName -like "Meeting Hours*" |
  Format-List DisplayName, AppId, AppVersion, Enabled, Type

Write-Host "== All add-ins visible to $Mailbox ==" -ForegroundColor Cyan
Get-App -Mailbox $Mailbox | Format-Table DisplayName, AppId, Enabled, Type -AutoSize

Disconnect-ExchangeOnline -Confirm:$false
Write-Host "VERIFY-DONE" -ForegroundColor Green
