#requires -Modules misAD

Function Get-CurrentUser
{
<#
.Synopsis
Queries a computer to find current logged on user

.DESCRIPTION
This script uses CIMInstance to query the win32_operatingsystem class for the logged on user. This does not work for Remote Desktop Sessions.

.NOTES   
Name: Get-CurrentUser
Author: Wayne Reeves
Version: 2024.07.30

.PARAMETER ComputerName
The name of the Computer that you are qurerying. Can also be just the asset tag.

.EXAMPLE
Get-CurrentUser -ComputerName adminXXXX

Description:
Will show you the user currently logged onto adminXXXX

.EXAMPLE
Get-CurrentUser xxxx

Will show you the user currently logged into computer with asset tag xxxx

.EXAMPLE
Get-CurrentUser

Will show you the user currently logged into the local computer
#>
param(
    [Parameter(Mandatory=$True)]
    $ComputerName
    )
if ( $ComputerName -ne 'localhost' )
    {
    $ComputerName = (Select-Computer $ComputerName).name
    }
if ( Test-Connection $ComputerName -Count 1 -Quiet )
    {
    Try
        {
        $username = (Get-CIMInstance win32_computersystem -ComputerName $ComputerName -ErrorAction Stop).username.replace("CCMHMR\","") 
        $ADUser = Get-ADUser $username -Server dom01 -ErrorAction Stop
	New-DataTable -Names Name, SamAccountName, ComputerName -Data $ADUser.Name, $ADUser.SamAccountName, $ComputerName
        }
    Catch
        {
        Write-Host "No User Logged in" -ForegroundColor Yellow
        }
    }
else
    {
    Write-Host "Host Unreachable" -ForegroundColor Red
    }
}

Function Disconnect-CurrentUser
{
<#
.Synopsis
Will log off the user currently logged into the computer specified

.DESCRIPTION
Issues a logoff using the win32shutdown method from the win32_operatingsystem class. This does not work for Remote Desktop essions.

.NOTES   
Name: Disconnect-CurrentUser
Author: Wayne Reeves
Version: 2024.07.29

.PARAMETER ComputerName
The name of the Computer that you are querying. Can also be just the asset tag. 

.EXAMPLE
Disconnect-CurrentUser -ComputerName adminXXXX

Description:
Will log off the user currently logged onto adminXXXX

.EXAMPLE
Disconnect-CurrentUser xxxx

Will log off the user currently logged into computer with asset tag xxxx

.EXAMPLE
Disconnect-CurrentUser

Will log you out of the computer
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$True)]
    $ComputerName
    )
$CurrentUserData = Get-CurrentUser $ComputerName
if ( $CurrentUserData )
    {
    if (  Invoke-CimMethod -ClassName win32_operatingsystem -MethodName Win32Shutdown -Arguments @{ Flags = 4 } -ComputerName $CurrentUserData.ComputerName -ErrorAction Stop)
        {       
        Write-Host "Successfully Logged Off User: $($CurrentUserData.Name)" -ForegroundColor Green
        }
    }
}
