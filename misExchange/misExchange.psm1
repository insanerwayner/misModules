# requires -Modules ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName "$($env:username)@lifepathsystems.org" -ShowProgress $True
