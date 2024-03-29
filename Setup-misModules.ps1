Write-Host "Checking for Installed Modules" -ForegroundColor Yellow
$InstalledModules = Get-Module mis* -ListAvailable
if ( $InstalledModules )
    {
    $InstalledModules | Remove-Module
    }
Write-Host "Copying Modules to Computer" -ForegroundColor Yellow
Set-Location '\\missvr2\mis\Powershell scripts\misModules'
$ModulePath = Join-Path $env:USERPROFILE '\Documents\WindowsPowershell\Modules\'
Get-Item mis* | Copy-Item -Destination $ModulePath -Recurse -Force
Set-Location $ModulePath
Write-Host "Importing Modules and Setting up" -ForegroundColor Yellow
$Modules = 'misScripting', 'misSecurity', 'misAD', 'misUserSession', 'misEncryption'
$Remove = 'misExchange'
foreach ( $Module in $Modules)
    {
    Write-Host "Setting Up $Module" -ForegroundColor Yellow
    #Import-Module $Module
    }
foreach ( $Module in $Remove )
    {
    $RemovePath = Join-Path $ModulePath $Module 
    if ( Test-Path $RemovePath )
	{
	Remove-Item $RemovePath -Recurse -Confirm:$False 
	}
    }
Write-Host "Exit now and Re-open Powershell." -ForegroundColor Yellow
