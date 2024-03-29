#requires -Modules misAD

Function Get-CurrentUser($ComputerName='localhost')
{
<#
.Synopsis
Queries a computer to find current logged on user

.DESCRIPTION
This script uses CIMInstance to query the win32_operatingsystem class for the logged on user

.NOTES   
Name: Get-CurrentUser
Author: Wayne Reeves
Version: 3.7.18

.PARAMETER ComputerName
The name of the Computer that you are qurerying. Can also be just the asset tag. It will default to localhost.

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
if ( $ComputerName -ne 'localhost' )
    {
    $ComputerName = (Find-ADComputer $ComputerName).name
    }
if ( Test-Connection $ComputerName -Count 1 -Quiet )
    {
    Try
        {
        $username = (Get-CIMInstance win32_computersystem -ComputerName $ComputerName -ErrorAction Stop).username.replace("CCMHMR\","") 
        Get-ADUser $username
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

Function Disconnect-CurrentUser($ComputerName='localhost')
{
<#
.Synopsis
Will log off the user currently logged into the computer specified

.DESCRIPTION
This script uses WMI to query the win32_operatingsystem class and issue a logoff

.NOTES   
Name: Disconnect-CurrentUser
Author: Wayne Reeves
Version: 3.7.18

.PARAMETER ComputerName
The name of the Computer that you are qurerying. Can also be just the asset tag. It will default to localhost.

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
if ( $ComputerName -ne 'localhost' ) 
    {
    $ComputerName = (Find-ADComputer $ComputerName).name
    }
$User = (Get-CurrentUser $ComputerName).Name
if ( $User )
    {
    if ( (gwmi win32_operatingsystem -ComputerName $ComputerName -ErrorAction Stop ).Win32Shutdown(4) )
        {       
        Write-Host "Successfully Logged Off User: $($User)" -ForegroundColor Green
        }
    }
}

function Get-UserSessions
{
<#
.Synopsis
queries a computer to check for interactive sessions

.DESCRIPTION
this script takes the output from the quser program and parses this to PowerShell objects

.NOTES   
name: Get-UserSessions
author: Jaap Brasser
version: 1.2.1
dateUpdated: 2015-09-23

.LINK
http://www.jaapbrasser.com

.PARAMETER ComputerName
the string or array of string for which a query will be executed

.EXAMPLE
get-UserSessions -ComputerName server01,server02

description:
will display the session information on server01 and server02

.EXAMPLE
'server01','server02' | Get-UserSessions

description:
will display the session information on server01 and server02
#>
param(
    [CmdletBinding()] 
    [Parameter(ValueFromPipeline=$true,
               ValueFromPipelineByPropertyName=$true)]
    [string[]]$ComputerName = 'localhost'
)
begin {
    $ErrorActionPreference = 'Stop'
}

process {
    foreach ($Computer in $ComputerName) {
        try {
            quser /server:$Computer 2>&1 | Select-Object -Skip 1 | ForEach-Object {
                $CurrentLine = $_.Trim() -Replace '\s+',' ' -Split '\s'
                $HashProps = @{
                    UserName = $CurrentLine[0]
                    ComputerName = $Computer
                }

                # If session is disconnected different fields will be selected
                if ($CurrentLine[2] -eq 'Disc') {
                        $HashProps.SessionName = $null
                        $HashProps.Id = $CurrentLine[1]
                        $HashProps.State = $CurrentLine[2]
                        $HashProps.IdleTime = $CurrentLine[3]
                        $HashProps.LogonTime = $CurrentLine[4..6] -join ' '
                        $HashProps.LogonTime = $CurrentLine[4..($CurrentLine.GetUpperBound(0))] -join ' '
                } else {
                        $HashProps.SessionName = $CurrentLine[1]
                        $HashProps.Id = $CurrentLine[2]
                        $HashProps.State = $CurrentLine[3]
                        $HashProps.IdleTime = $CurrentLine[4]
                        $HashProps.LogonTime = $CurrentLine[5..($CurrentLine.GetUpperBound(0))] -join ' '
                }

                New-Object -TypeName PSCustomObject -Property $HashProps |
                Select-Object -Property UserName,ComputerName,SessionName,Id,State,IdleTime,LogonTime,Error
            }
        } catch {
            New-Object -TypeName PSCustomObject -Property @{
                ComputerName = $Computer
                Error = $_.Exception.Message
            } | Select-Object -Property UserName,ComputerName,SessionName,Id,State,IdleTime,LogonTime,Error
        }
    }
}     
}

Function Disconnect-CitrixSession
{
<#
.Synopsis
Disconnects Active Citrix Sessions for a specified user

.DESCRIPTION
Queries all Citrix Servers to check if user has active session. Once one is located, it will Disconnect that Session. This is useful for when a user is having difficulties with reconnecting to Anasazi after a disrupted or frozen session.

.NOTES   
Name: Disconnect-CitrixSession
Author: Wayne Reeves
Version: 11.28.17

.PARAMETER Filter
This is the string filter to search for the user

.EXAMPLE
Disconnect-CitrixSession test

Description:
Will present a menu with users that match the string filter "test", then check for Citrix Sessions for that user and Disconnect them.
#>
param($filter)

$user = Select-User $filter
if ( $User )
    {
    $ctrxservers = (find-adcomputer misctrx).name
    $i = 0
    $count = $ctrxservers.count
    do 
        {
        $server = $ctrxservers[$i]
        $operation = "Checking $($server)"
        $percent = ($i/$count)*100
        Write-Progress -PercentComplete $percent -Activity 'Logging Off User' -CurrentOperation $operation
        if ( Test-Connection -ComputerName $server -Count 1 -BufferSize 1 -Quiet -ErrorAction SilentlyContinue )
            {
            $session = Get-UserSessions -ComputerName $server | Where-Object username -match $user.samaccountname
            if ( $session )
                {
                $sessionstring = $session.sessionname
                $serverstring = "/SERVER:$($server)"
                logoff $sessionstring $serverstring
                $found = $True
                }
            }
        $i++
        }
    until ( $found -or ( $i -eq $count ) )
    if ( !$found )
        {
        Write-Host "No Active Session for $($user.name)" -ForegroundColor Red
        }
    else 
        {
        Write-Host "$($user.name) Successfully Logged Off from $($server)" -ForegroundColor Yellow
        }
    }
}

