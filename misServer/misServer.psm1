# depends misAD, misScripting

Function Clear-PrintQueue
{
<#
.Synopsis
Clears print queue of a computer or server

.DESCRIPTION
Stops the print service, clears the queue, then starts the service again

.NOTES
Name: Clear-PrintQueue
Author: Wayne Reeves
Version: 9-25-18

.PARAMETER ComputerName
Name or partial name of computer.

.EXAMPLE
Clear-PrintQueue -ComputerName XXXX

Description:
Will clear the print queue of computer with Asset XXXX

.EXAMPLE
Clear-PrintQueue -ComputerName MISHERITAGE1

Description:
Will clear the print queue of MISHERITAGE1
#>
param(
    $ComputerName
    )
$ComputerName = (Select-Computer $ComputerName).Name
if ( !(Test-Connection -computername $computername -count 1 -buffersize 1 -quiet -erroraction silentlycontinue)  )
    { 
    Write-Warning "$($computername) unreachable"
    }
else
    {
    $Success = @()
    $Fail = @()
    $Printers = Get-CIMInstance win32_printer -computername $computername 
    If ( $Printers )
        {
        Foreach ( $Printer in $Printers )
            {
            if ( $Printer.ShareName -ne $null )
                {
                $PName = $Printer.ShareName
                }
            else
                {
                $Pname = $Printer.Name
                }
            Try
                {
                $Cancel = $Printer.cancelalljobs()
                $Success += $Pname
                }
            Catch
                {
                $Fail += $Pname
                }
            }
        }
    else
        {
        Write-Host "No print jobs found"
        }
    New-DataTable -Names Successful, Failed -Data $Success, $Fail
    }
}
