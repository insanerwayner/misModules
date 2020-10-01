if ( (Get-Host).Version.Major -ge 7 )
    {
    Write-Host "Checking for Installed Modules" -ForegroundColor Yellow
    if ( !(Get-Module ExchangeOnlineManagement) )
	{
	Write-Host "Installing ExchangeOnlineManagement Module" -ForegroundColor Yellow
	Install-Module ExchangeOnlineManagement -Confirm:$False
	}
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
    $Modules = 'misScripting', 'misSecurity', 'misAD', 'misUserSession', 'misEncryption', 'misExchange'
    foreach ( $Module in $Modules)
	{
	Write-Host "Setting Up $Module" -ForegroundColor Yellow
	#Import-Module $Module
	}
    Write-Host "Exit now and Re-open Powershell." -ForegroundColor Yellow
    }
else
    {
    Write-Host "Powershell 7 not detected. Installing."
    iex "& { $(irm https://aka.ms/install-powershell.ps1) } -UseMSI"
    Write-Host "Launch Powershell 7 as Administrator and rerun this setup file."
    }
