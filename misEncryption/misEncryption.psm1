#requires -Modules misAD, misSecurity
#Sets up PoSHKeePass if it isn't already
if ( !(Get-Module -Name PoShKeePass -ListAvailable) )
        {
        Write-Host "This module requires the 'PoShKeePass' module. Would you like to automatically set it up?"
        $answer = Read-Host '[Y,n]'
        if ( $answer -eq 'y' -or $answer -eq '' )
                {
                Install-Module PoShKeePass -ErrorAction Stop
                Import-Module PoShKeePass
                $param = @{
                        'DatabaseProfileName'='Assets';
                        'DatabasePath'='\\missvr2\mis\Computer Encryption\Database\Lifepath Systems Assets.kdbx';
                        'UseNetworkAccount'=$False
                        'UseMasterKey'=$True
                        }
                New-KeePassDatabaseConfiguration @param
                }
        else
                {
                exit
                }
        }

Function Get-Key($Asset)
        {
        <#
        .Synopsis
        This retrieves a Truecrypt or Bitlocker key from the computer specified 

        .DESCRIPTION
        This uses the PoSHKeePass Module to display the Truecrypt or Bitlocker Key stored in KeePass

        .NOTES   
        Name: Get-Key
        Author: Wayne Reeves
        Version: 11.28.17

        .PARAMETER Asset
        Either the Asset tag or the Full Computer Name

        .EXAMPLE
        Get-Key XXXX

        Description:
        Will Display the Encryption Key for asset tag XXXX

        .EXAMPLE
        Get-Key adminXXXX

        Description:
        Will Display the Encryption Key for Computer adminXXXX
        #>
        $Name = (Find-ADComputer $Asset).name
        $Password = Get-XMLPassword -Name KeePass -Type Password
        $params = @{
                'KeePassEntryGroupPath'='Lifepath Systems Assets/Network'
                'DatabaseProfileName'='Assets'
                'MasterKey'=$password
                }
        $Table = Get-KeePassEntry @params -AsPlainText -WarningAction SilentlyContinue | ? Title -eq $asset 
        $Table | Add-Member NoteProperty Name($Name)
        $Table | fl Name, Title, UserName, Password, Notes
        }

Function New-Key
        {
        <#
        .Synopsis
        Will store an encryption key for a new Asset

        .DESCRIPTION
        Will store an encryption key for a new Asset

        .NOTES   
        Name: New-Key
        Author: Wayne Reeves
        Version: 11.28.17

        .PARAMETER Asset
        Asset Tag of the Computer

        .PARAMETER Key
        The Key or Password being stored in KeePass

        .PARAMETER Notes
        Notes section in KeePass, typpicaly used to indicate whether Bitlocker or TrueCrypt

        .EXAMPLE
        New-Key -Asset XXXX -Key XXXX-XXXXX-XXXX-XXXXX -Notes "Bitlocker"

        Description:
        Will store a new Bitlocker Key for asset XXXX
        #>
        param(
        [Parameter(Mandatory=$true)]
        $Asset,
        [Parameter(Mandatory=$true)]
        $Key, 
        $Notes
        )
        $Key = ConvertTo-SecureString $Key -AsPlainText -Force
        $Password = Get-XMLPassword -Name KeePass -Type Password
        $params = @{
                'KeePassEntryGroupPath'='Lifepath Systems Assets/Network'
                'DatabaseProfileName'='Assets'
                'MasterKey'=$password
                'Title'=$Asset
                'KeePassPassword'=$Key
                'Notes'=$Notes
                }
        New-KeePassEntry @params -WarningAction SilentlyContinue
        }

Function Get-RecoveryKeyFromFile($asset)
        {
        <#
        .Synopsis
        Get Bitlocker Recovery Key from it's matching Text File stored on the Server

        .DESCRIPTION
        Get Bitlocker Recovery Key from it's matching Text File stored on the Server

        .NOTES   
        Name: Get-RecoveryKeyFromFile
        Author: Wayne Reeves
        Version: 11.28.17

        .PARAMETER Asset
        Asset Tag of the Computer

        .EXAMPLE
        Get-RecoveryKeyFromFile -Asset XXXX

        Description:
        Will retrieve Bitlocker Recovery key from text file matching the asset tag
        #>
        $file = "\\missvr2\mis\Computer Encryption\Bitlocker Recovery Keys\$($asset).txt"
        (cat $file)[12] -Replace "\s"
        }

Function Move-RecoveryKeysFromFilesToKeePass
        {
        <#
        .Synopsis
        This simply transfers any New Recovery Keys for newly encrypted machines to KeePass
        #>
        $Files = Get-ChildItem "\\missvr2\mis\Computer Encryption\BitLocker Recovery Keys\new\"
        foreach ($File in $Files)
                {
                Write-Host "Writing Recovery Key for $($Asset)"
                $Asset = $File.basename
                $Key= (Get-Content $file.FullName)[12] -Replace "\s"
                New-Key -Asset $Asset -Key $Key -Notes "Bitlocker Recovery Key"
                Write-Host "Moving Recovery Key for $($Asset)"
                Move-Item $File.fullname "\\missvr2\mis\Computer Encryption\BitLocker Recovery Keys"
                }
        }
