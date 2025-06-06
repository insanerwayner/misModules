﻿#requires -Modules ActiveDirectory, misScripting, Microsoft.Graph.Users, Microsoft.Graph.Authentication, Microsoft.Graph.Reports, Microsoft.Graph.Groups

Function Confirm-MgGraph 
    {
    <#
    .SYNOPSIS
    Ensures an active Microsoft Graph session with the required scopes.

    .DESCRIPTION
    This function checks for an active Graph session by calling Get-MgContext. If no session is found,
    or if the session is missing any required OAuth scopes, it attempts to connect (or reconnect)
    using Connect-MgGraph with the specified scopes. If the connection fails, it provides a clear error
    message indicating what is required.

    .PARAMETER RequiredScopes
    An array of required OAuth scopes for the session. Defaults to @("AuditLog.Read.All", "Directory.Read.All").
    If the current session lacks any of these, the function will attempt to request additional consent.

    .EXAMPLE
    Confirm-MgGraph
    Checks for an active Graph session with the default scopes; if none exists or if any required scopes
    are missing, it connects using the defaults.

    .EXAMPLE
    Confirm-MgGraph -RequiredScopes @("User.ReadWrite.All", "Group.ReadWrite.All")
    Checks for an active Graph session with the specified scopes and requests additional consent if needed.

    .NOTES
    Requires the Microsoft.Graph module. If connection fails, ensure that the module is installed and that
    you have consented to the required scopes.
    #>
    param(
        [string[]]$RequiredScopes = @("AuditLog.Read.All", "Directory.Read.All")
    )
    
    try {
        $ctx = Get-MgContext -ErrorAction Stop
        }
    catch 
        {
        $errorMessage = $_.Exception.Message
        Write-Error "Error retrieving Graph context: $errorMessage. Please ensure the Microsoft.Graph module is installed and that you have network connectivity."
        $ctx = $null
        }
    
    if (-not $ctx)
        {
        Write-Host "No active Graph session found. Attempting to connect with scopes: $($RequiredScopes -join ', ')"
        try 
            {
            Connect-MgGraph -Scopes $RequiredScopes -NoWelcome -ErrorAction Stop
            }
        catch 
            {
            $connectError = $_.Exception.Message
            Write-Error "Failed to connect to Microsoft Graph. Please ensure that the Microsoft.Graph module is installed and that your account has consented to the following scopes: $($RequiredScopes -join ', '). Error: $connectError"
            throw
            }
        }
    else
        {
        $missingScopes = $RequiredScopes | Where-Object { $ctx.Scopes -notcontains $_ }
        if ($missingScopes) 
            {
            Write-Host "Missing scopes: $($missingScopes -join ', '). Requesting additional consent..."
            try 
                {
                Connect-MgGraph -Scopes $RequiredScopes -NoWelcome -ErrorAction Stop
                }
            catch 
                {
                $reconnectError = $_.Exception.Message
                Write-Error "Failed to update the Microsoft Graph connection with the required scopes. Please ensure that your account has consented to: $($RequiredScopes -join ', '). Error: $reconnectError"
                throw
                }
            }
        }
    }


Function Find-ADComputer
    {
    <#
    .Synopsis
    Queries Active Director for Computers that match the Asset Tag

    .DESCRIPTION
    Queries Active Director for Computers that match the Asset Tag

    .NOTES   
    Name: Find-ADComputer
    Author: Wayne Reeves
    Version: 11.29.17

    .PARAMETER Asset
    Asset Tag of the Computer you are searching for

    .PARAMETER Server
    Name of the server you wish to query

    .EXAMPLE
    Find-ADComputer -Asset XXXX

    Description:
    Will show you computer(s) that match the Asset

    .EXAMPLE
    Find-ADcomputer -Asset XXXX -Server dom01

    Description:
    Will show you computer(s) that match the Asset, but from the Server you specify.
    #>
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [string]$Asset
    )
    DynamicParam 
        {
        $DCs = @( get-addomaincontroller -Filter { OperatingSystem -notlike "Windows Server 2003" -and OperatingSystem -notlike "Windows Server® 2008 Standard" } | Foreach-Object { $_.Name })
        New-DynamicParam -Name Server -ValidateSet $DCs
        }       

    begin 
        {
        $Server = $PsBoundParameters.Server
        }

    process
        {
        if ( $Server -ne $null )
            {
            get-adcomputer -ldapfilter "(name=*$asset*)" -Server $Server
            }
        else 
            {
            get-adcomputer -ldapfilter "(name=*$asset*)"
            }
        }
    }

    Function Get-LastBootTime($Computer=$env:COMPUTERNAME)
    {
    <#
    .Synopsis
    Queries a computer to find the last time it booted up

    .DESCRIPTION
    This script uses WMI to query the win32_operatingsystem class and return the last boot time.

    .NOTES   
    Name: Get-LastBootTime
    Author: Wayne Reeves
    Version: 11.29.17

    .PARAMETER Computer
    This can either be the Asset Tag of the Computer or the Full Computer Name

    .EXAMPLE
    Get-LastBootTime XXXX

    Description:
    Will show you the last boot time for computer with Asset Tag "XXXX"

    .EXAMPLE
    Get-LastBootTime adminXXXX

    Description:
    Will show you the last boot time for Computer Name "adminXXXX"

    .EXAMPLE
    Get-LastBootTime

    Description:
    Will show you the last boot time for your computer
    #>
    $computername = Find-ADComputer -Asset $Computer
    Get-CIMInstance win32_operatingsystem -computername $computername.name | select LastBootUpTime
    }

Function Find-ADuser
    {
    <#
    .Synopsis
    Queries Active Director for users that match a string filter

    .DESCRIPTION
    Queries Active Director for users that match a string filter

    .NOTES   
    Name: Find-ADUser
    Author: Wayne Reeves
    Version: 11.29.17

    .PARAMETER Filter
    This is the string you are searching the user with

    .PARAMETER Server
    Name of the server you wish to query

    .EXAMPLE
    Find-ADUser test

    Description:
    Will show you user(s) that match the string filter "test"

    .EXAMPLE
    Find-ADUser test -Server dom01

    Description:
    Will show you user(s) that match the string filter "test" but query from "dom01"
    #>
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [string]
        $Filter
    )
    DynamicParam 
        {
        $DCs = @( get-addomaincontroller -Filter { OperatingSystem -notlike "Windows Server 2003" -and OperatingSystem -notlike "Windows Server® 2008 Standard" } | Foreach-Object { $_.Name })
        New-DynamicParam -Name Server -ValidateSet $DCs
        }       

    begin 
        {
        $Server = $PsBoundParameters.Server
        }

    process
        {
        if ( !$Server ) { $Server = "dom01" }
        if ( $Server -ne $null )
            {
            get-aduser -ldapfilter "(|(name=*$filter*)(samaccountname=*$filter*))" -Server $Server
            }
        else 
            {
            get-aduser -ldapfilter "(|(name=*$filter*)(samaccountname=*$filter*))"
            }
        }
    }

Function Select-User
    {
    <#
    .Synopsis
    A selectable menu to select a user

    .DESCRIPTION
    This uses the output of Find-ADUser to create a menu to select a user

    .NOTES   
    Name: Select-User
    Author: Wayne Reeves
    Version: 11.29.17

    This really will only be used by other commandlettes

    .PARAMETER Filter
    This is the string filter to search for the user

    .EXAMPLE
    Select-User test

    Description:
    Will create a menu with users that match the string filter "test"
    #>
    param(
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [string]
        $Filter
        )
    $users = Find-ADuser $filter
    if ( $users.Count -ne '' )
        {
        $menu = @()
        for ($i=1;$i -le $users.count; $i++)
            {
            Write-Host "$i. $($users[$i-1].name)" -ForegroundColor Cyan
            $user = New-Object System.Object
            $user | Add-Member -MemberType NoteProperty -Name 'Index' -Value $i
            $user | Add-Member -MemberType NoteProperty -Name 'Name' -Value $($users[$i-1].name)
            $user | Add-Member -MemberType NoteProperty -Name 'SAMAccountName' -Value $($users[$i-1].samaccountname)
            $menu += $user
            }
        $selection = Read-Host 'Selection'
        try
            {
            [int]$num = $selection
            }
        catch
            {
            Write-Host "Not a Valid Entry" -ForegroundColor Red
            }
        if ( $num -lt 1 -or $num -gt $menu.Count )
            {
            Write-Host "Not a Valid Entry" -ForegroundColor Red
            }
        $selection = $menu | ? { $_.Index -eq $selection }
        Return $Selection
        }
    else
        {
        if ( $users -eq $null )
            {
            Write-Host "No Match Found" -ForegroundColor Yellow
            }
        else 
            {
            Write-Host $users.name -ForegroundColor Cyan
            $continue = Read-Host "Continue? [Y,n]"
            if ( ( $continue -eq 'y' ) -or ( $continue -eq "" ) )
                {
                Return $users
                }
            }   
        }
    }

Function Select-Computer
    {
    <#
    .Synopsis
    A selectable menu to select a computer

    .DESCRIPTION
    This uses the output of Find-ADComputer to create a menu to select a computer

    .NOTES   
    Name: Select-Computer
    Author: Wayne Reeves
    Version: 9-25-18 

    This really will only be used by other commandlettes

    .PARAMETER Filter
    This is the string filter to search for the computer

    .EXAMPLE
    Select-Computer test

    Description:
    Will create a menu with computers that match the string filter "test"
    #>
    param(
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [string]
        $Filter
        )
    $computers = Find-ADcomputer $filter
    if ( $computers.Count -ne '' )
        {
        $menu = @()
        for ($i=1;$i -le $computers.count; $i++)
            {
            Write-Host "$i. $($computers[$i-1].name)" -ForegroundColor Cyan
            $computer = New-Object System.Object
            $computer | Add-Member -MemberType NoteProperty -Name 'Index' -Value $i
            $computer | Add-Member -MemberType NoteProperty -Name 'Name' -Value $($computers[$i-1].name)
            $computer | Add-Member -MemberType NoteProperty -Name 'SAMAccountName' -Value $($computers[$i-1].samaccountname)
            $menu += $computer
            }
        $selection = Read-Host 'Selection'
        try
            {
            [int]$num = $selection
            }
        catch
            {
            Write-Host "Not a Valid Entry" -ForegroundColor Red
            }
        if ( $num -lt 1 -or $num -gt $menu.Count )
            {
            Write-Host "Not a Valid Entry" -ForegroundColor Red
            }
        $selection = $menu | ? { $_.Index -eq $selection }
        Return $Selection
        }
    else
        {
        if ( $computers -eq $null )
            {
            Write-Host "No Match Found" -ForegroundColor Yellow
            }
        else 
            {
            Write-Host $computers.name -ForegroundColor Cyan
            $continue = Read-Host "Continue? [Y,n]"
            if ( ( $continue -eq 'y' ) -or ( $continue -eq "" ) )
                {
                Return $computers
                }
            }   
        }
    }


Function Reset-Password
    {
    <#
    .Synopsis
    Will reset a users password.

    .DESCRIPTION
    Will present a menu to select the user that matches your string filter and then reset the password of that user.

    .NOTES   
    Name: Reset-Password
    Author: Wayne Reeves
    Version: 11.29.17

    .PARAMETER Filter
    This is the string filter to search for the user

    .PARAMETER Server
    Specifies the Server you would like to reset the password

    .PARAMETER DoNotChangePasswordAtLogon
    Switch to set the user to NOT change their password at logon. If this is NOT specified the user will be prompted to change their password.

    .PARAMETER GenerateRandomPassword
    Switch that will set the password to a randomly generated password and output for you to provide to user. This will also provide you with call words of the password to read to the user.

    .EXAMPLE
    Reset-Password test

    Description:
    Will create a menu for any users that matches the string filter "test" and will reset the password of the user you select.

    .EXAMPLE
    Reset-Password test -DoNotChangePasswordAtLogon

    Description:
    Will create a menu for any users that matches the string filter "test" and will reset the password of the user you select. Will NOT require the user to change password at logon.

    .EXAMPLE
    Reset-Password test -Server dom01

    Description:
    Will create a menu for any users that matches the string filter "test" and will reset the password of the user you select on the Server you specified
    
    .EXAMPLE
    Reset-Password test -GenerateRandomPassword

    Description:
    Will create a menu for any users that matches the string filter "test" and will reset the password of the user you select on the Server you specified
    #>
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [string]
        $Filter,
        [switch]
        $DoNotChangePasswordAtLogon,
        [switch]
        $GenerateRandomPassword
        )
        DynamicParam 
            {
            $DCs = @( get-addomaincontroller -Filter { OperatingSystem -notlike "Windows Server 2003" -and OperatingSystem -notlike "Windows Server® 2008 Standard" } | Foreach-Object { $_.Name })
            New-DynamicParam -Name Server -ValidateSet $DCs
            }       

        begin 
            {
            $Server = $PsBoundParameters.Server
            }
        
        process
            {
            if ( !$Server ) { $Server = "dom01" }
            $samaccountname = (Select-User $Filter).samaccountname 
            if ( $samaccountname )
                {
                If ( $GenerateRandomPassword )
                    {
                    $Password = New-RandomPassword
                    }
                Else
                    {
                    $Password = "green kitten tail"
                    }
                Set-AdAccountPassword -Identity $samaccountname -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force) -server $server
                if ( !$DoNotChangePasswordAtLogon )
                    {
                    Set-ADuser -Identity $samaccountname -ChangePasswordAtLogon $True -server $server
                    }
                Write-Host "Password set to: $($Password)" -ForegroundColor Green
                if ( $GenerateRandomPassword )
                    {
                    $CallWords = Get-CallWords -String $Password
                    Write-Host "Call Words: $($CallWords)" -ForegroundColor Yellow
                    }
                }
            }
    } 

Function Unlock-ADUser
    {
    <#
    .Synopsis
    Simplifies unlocking users

    .DESCRIPTION
    This script will give you a menu to select the user from a list generated from your string filter then it will attempt to unlock the user

    .NOTES   
    Name: Unlock-ADUser
    Author: Wayne Reeves
    Version: 2.5.18

    .PARAMETER Filter
    This is the string filter to search for the user
    
    .PARAMETER Server
    This specifies the server from which you would like to unlock the user
    
    .EXAMPLE
    Unlock-ADUser test

    Description:
    Will check each Server if User is locked out and go and Unlock them

    .EXAMPLE
    Unlock-ADUser test -Server dom01

    Description:
    Will check dom01 to see if User is locked out. If it is locked out it will unlock the user. If it is not locked it will tell you and ask if you would like to check all DCs.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,Position=0)]
        [string]$Filter
        )
    #This is to make the Parameter 'Server' have Tab Completion with valid servers that are compatible with the script
    DynamicParam 
        {
        $DCs = @( get-addomaincontroller -Filter { OperatingSystem -notlike "Windows Server 2003" -and OperatingSystem -notlike "Windows Server® 2008 Standard" } | Foreach-Object { $_.Name })
        New-DynamicParam -Name Server -ValidateSet $DCs
        }       

    begin 
        {
        $Server = $PsBoundParameters.Server
        }
    
    # Where the Script actually starts running
    process
        {
        #Function for Unlocking all Domain Controllers
        Function Unlock-DCS($User, $DCs)
            {
            Write-Host "Getting Locked DCs..." -ForegroundColor Yellow
            $count = $dcs.Count
            $i = 0
            foreach ( $DC in $DCs )
                {
                $i++
                $operation = "Checking $($dc)"
                $percent = ($i/$count)*100
                Write-Progress -PercentComplete $percent -Activity 'Unlocking User' -CurrentOperation $operation
                if ( ( get-aduser $user.samaccountname -properties lockedout -server $dc -ErrorAction SilentlyContinue ).lockedout -eq "True" )
                    {
                    $operation = "Unlocking User on $($dc)"
                    Write-Progress -PercentComplete $percent -Activity 'Unlocking User' -CurrentOperation "Unlocking from $($DC)"
                    Unlock-ADAccount $user.SAMAccountName -Server $dc
                    if ( !$locked )
                        {
                        Write-Host "Successfully Unlocked on $($DC)" -ForegroundColor Green -NoNewLine
                        }
                    else
                        {
                        Write-Host ", $($DC)" -ForegroundColor Green -NoNewLine
                        }
                    $locked += $dc
                    }
                }
            if ( $locked -eq $null ) 
                { 
                Write-Host "User Not Locked" -ForegroundColor Yellow 
                }
            }
        
        # Presents the Menu to select the User
        $user = Select-User $filter

        # WorkFlow for Deciding how to proceed with unlocking the user
        if ( $user )
            {
            if ( !$Server )
                {
                Unlock-DCS -User $user -DCs $DCs
                }
            else
                {
                if ( ( get-aduser $user.samaccountname -properties lockedout -server $Server -ErrorAction SilentlyContinue ).lockedout -eq "True" )
                    {
                    Write-Host "Unlocking" $Server.toUpper() -ForegroundColor Green
                    Unlock-ADAccount $user.SAMAccountName -Server $Server
                    }
                else
                    {
                    Write-Host "User not locked out on" $Server.toUpper() -ForegroundColor Yellow
                    $Choice = Read-Host "Check All DCs? [Y,n]"
                    if ( ( $Choice -eq 'y' ) -or ( $Choice -eq "" ) )
                        {
                        Unlock-DCs -User $user -DCs $DCs
                        }
                    }
                }
            }
        }
    }


Function Get-PasswordExpirationList
    {
    <#
    .Synopsis
    Providess a list of Users Password Expiration Days Left

    .DESCRIPTION
    Lists a countdown of users password expirations

    .NOTES   
    Name: Get-PasswordExpirationList
    Author: Wayne Reeves
    Version: 11.28.17

    .PARAMETER Office
    This parameter filters the users down to a string that matches the Office they are assigned in ActiveDirectory

    .EXAMPLE
    Get-PasswordExpirationList

    Description:
    Will show you the Password Expirations for ALL Users

    .EXAMPLE
    Get-PasswordExpirationList -Office Crisis

    Description:
    Will show you the Password Expirations for users that have an office that matches the string "Crisis"
    #>
    param($office="")
    $users = Get-ADUser -Filter * -properties office | ? { $_.office -match $office -and $_.Enabled -eq $True }
    $list = @()
    foreach ( $user in $users )
        {
        $username = $user.SamAccountName
        $searcher=New-Object DirectoryServices.DirectorySearcher
        $searcher.Filter="(&(samaccountname=$username))"
        $results=$searcher.findone()
        $lastset = [datetime]::fromfiletime($results.properties.pwdlastset[0])
        $timeleft = 90 - (( Get-Date ) - $lastset ).days
        $expires = Get-Date $lastset.adddays(90) -Format "MM/dd/yy hh:mm tt"
        $info = New-Object -TypeName PSObject
        $info | Add-Member -MemberType NoteProperty -Name Name -Value $user.Name
        $info | Add-Member -MemberType NOteProperty -Name Expires -Value $expires
        $info | Add-Member -MemberType NoteProperty -Name DaysLeft -Value $timeleft	
        $list += $info
        }
    $list | Sort-Object DaysLeft | Out-GridView
    }

Function New-LPSUser
    {
    <#
    .Synopsis
    Creates a new LifePath User

    .DESCRIPTION
    Will create new user with correct Group Memberships based off of a template user and will trigger Exchange Online to create a mailbox for the user by default, unless otherwise specified.

    .NOTES   
    Name: New-LPSUser
    Author: Wayne Reeves
    Version: 10.9.18

    .PARAMETER FirstN
    First Name of User
    
    .PARAMETER MI
    Middle Initial of User
    
    .PARAMETER LastN
    Last Name of User

    .PARAMETER Title
    User's Title
    .PARAMETER Office
    User's Office

    .PARAMETER Department
    User's Division. Can only choose between "Admin", "BH", "ECI", or "IDD"

    .PARAMETER Template
    Username of the Template User you wish to copy Group Memberships from

    .PARAMETER Office
    This parameter filters the users down to a string that matches the Office they are assigned in ActiveDirectory

    .PARAMETER HomeDirectory
    $True or $False Value to create a HomeDirectory for User (Default True)
    
    .PARAMETER Enabled
    $True or $False Value to set user to enabled state and make visible in Address Book (Default False)

    .PARAMETER LicenseGroup
    Security Group indicating what group of Azure Licenses that shoudld be applied to the user. Default is 'E3 Simple Licenses'. 
    If 'Mailbox' is set to $False, this license will be ignored.  

    .PARAMETER Mailbox
    True or False to create a mailbox for this user. Will add the user to the selected E3 License Security Group, thus prompting assigning a license which will cause Exchange Online to create a mailbox. (Default True)

    .PARAMETER EmployeeID
    The Payroll ID from HR.

    .PARAMETER Manager
    The SamAccountName of the Manager

    .EXAMPLE
    New-LPSUser -FirstN Bob -MI S -LastN Cratchet -Title Hero -Office "BH McKinney" -Department BH -Template mwarren 

    Description:
    Will create a new user and mailbox for "Bob S Cratchet." "mwarren" will be used as a template for Group Memberships. User will be disabled and hidden from Address Book.

    .EXAMPLE
    New-LPSUser -FirstN Bob -MI S -LastN Cratchet -Title Hero -Office "BH McKinney" -Department BH -Template mwarren -Mailbox $False

    Description:
    Will create a new user without a mailbox for "Bob S Cratchet." "mwarren" will be used as a template for Group Memberships. User will be disabled and hidden from Address Book.

    .EXAMPLE
    New-LPSUser -FirstN Bob -MI S -LastN Cratchet -Title Hero -Office "BH McKinney" -Department BH -Template mwarren -HomeDirectory $False -Enabled $True 

    Description:
    Will create a new user and mailbox for "Bob S Cratchet." "mwarren" will be used as a template for Group Memberships. User will be enabled and visible in Address Book.

    .EXAMPLE
    New-LPSUser -FirstN Bob -MI S -LastN Cratchet -Title Hero -Office "IDD Plano" -Department IDD -Template jbraughton

    Description:
    Will create a new user and mailbox for "Bob S Cratchet." "jbraughton" will be used as a template for Group Memberships. User will be disabled and hidden from Address Book.
    #>
    param ( 
	[Parameter(Mandatory)]
        [string]$FirstN, 
        [string]$MI,
	[Parameter(Mandatory)]
        [string]$LastN, 
	[Parameter(Mandatory)]
        [string]$Title, 
	[Parameter(Mandatory)]
        [string]$Office, 
	[Parameter(Mandatory)]
	[ValidateSet("Admin","BH","ECI","IDD")]
        [string]$Department, 
	[Parameter(Mandatory)]
        [string]$EmployeeID,
	[Parameter(Mandatory)]
        [string]$Manager,
	[Parameter(Mandatory)]
        [string]$Template, 
        [bool]$HomeDirectory=$True, 
        [bool]$Enabled=$False, 
	[bool]$Mailbox=$True,
	[string]$LicenseGroup='E3 Simple Licenses'
        )
    #User Variables
    $alias = $FirstN.toLower().substring(0,1)+$LastN.tolower().replace("-","").replace(" ","")
    $aliaswithMI = $FirstN.toLower().substring(0,1)+$MI.tolower()+$LastN.tolower().replace("-","").replace(" ","")
    $UnencryptedPassword = New-RandomPassword
    $Password = ConvertTo-SecureString $UnencryptedPassword -AsPlainText -Force
    #Write-Host "Creating New User: $FirstN $MI $LastN" -ForegroundColor White
    $HideInAddressBook=!$Enabled
    $Activity = "Creating New User: $FirstN $MI $LastN" -Replace "  ", " "
    Write-Progress -Activity $Activity -CurrentOperation $Activity 
    $UserObject = New-Object -TypeName PSObject
    Function Set-HomeDirectory($alias, $Department, $Office)
        {
        $HDActivity = "HomeDirectory"
        Write-Progress -Activity $HDActivity -CurrentOperation "Creating Share: $SharePath"
        switch ($Department) 
            {
            Admin
                {
                $FileServer = "misfs1"
                $NewFolder = Join-Path "\\$FileServer\d`$\User Shares\" $alias
                $LocalPath = "D:\User Shares\$($alias)"
                }
            BH 
                { 
                $FileServer = "misfs1"
                $NewFolder = Join-Path "\\$FileServer\d`$\User Shares\" $alias
                $LocalPath = "D:\User Shares\$($alias)"
                }
            IDD 
                {
                $FileServer = "misfs2"
                $NewFolder = Join-Path "\\$FileServer\e`$\IDD User Shares\" $alias
                $LocalPath = "E:\IDD User Shares\$($alias)"
                } 
            ECI 
                { 
                $FileServer = "misfs1"
                $NewFolder = Join-Path "\\$($FileServer)\d`$\ECI USERS\" $alias
                $LocalPath = "D:\ECI USERS\$($alias)"
                }
            Hotline { $NoHome = $True }
            }
            if ( !$NoHome )
                {
                $SharePath = "\\$($FileServer)\$($alias)"
                $UserObject | Add-Member -MemberType NoteProperty -Name HomeDirectory -Value $SharePath                
                #Write-Host "Creating Share: $SharePath" -ForegroundColor Yellow
                Write-Progress -Activity $HDActivity -CurrentOperation "Creating Share: $SharePath"
                New-Item $NewFolder -Type Directory | Out-Null
                $ScriptBlock = { param($LocalPath,$alias) Add-NTFSAccess -Path $LocalPath -Account "CCMHMR\$($alias)" -AccessRights Modify }
                Write-Progress -Activity $HDActivity -CurrentOperation "Setting File Permissions"
		sleep 5
                Invoke-Command $ScriptBlock -ArgumentList $LocalPath, $alias -ComputerName $FileServer -ErrorVariable NoPerms
		# If Add Permissions Failed, wait 5 seconds and try again 5 times.
		if ( $NoPerms )
		    {
		    $Count = 0
		    while ( $NoPerms -and $Count -lt 5 )
			{
			$NoPerms = $Null
			Write-Host "Cannot add HomeDirectory Permission for User. Trying again in 5 Seconds." -ForegroundColor Yellow
			sleep 5
			Invoke-Command $ScriptBlock -ArgumentList $LocalPath, $alias -ComputerName $FileServer -ErrorVariable NoPerms
			$Count++
			}
		    }
		if ( !$NoPerms )
		    {
		    New-SMBShare -Name $alias -Path $LocalPath -FullAccess Everyone -CimSession $FileServer	| Out-Null
		    #Write-Host "Setting N Drive to $SharePath" -ForegroundColor Yellow
		    Write-Progress -Activity $HDActivity -CurrentOperation "Setting N Drive to $SharePath"	
		    Set-AdUser -Identity $alias -HomeDirectory $SharePath -HomeDrive "N:" -Server dom01
		    Write-Progress -Activity $HDActivity -Completed
		    }
		else
		    {
		    $UserObject | Add-Member -MemberType NoteProperty -Name HomeDirectory -Value "None"
		    $UserObject | Add-Member -MemberType NoteProperty -Name Error -Value $NoPerms
		    #Write-Host "No HomeDirectory for $alias" -ForegroundColor Yellow
		    Write-Progress -Activity $HDActivity -Completed
		    }   
		}
            else
                {
                $UserObject | Add-Member -MemberType NoteProperty -Name HomeDirectory -Value "None"
                #Write-Host "No HomeDirectory for $alias" -ForegroundColor Yellow
                Write-Progress -Activity $HDActivity -Completed
                }
        }
    
    #Write-Host "Checking if username already exists" -ForegroundColor Yellow
    Write-Progress -Activity $Activity -CurrentOperation "Checking if username already exists" 
    
    If ( ( Get-ADUser -LDAPFilter "(sAMAccountName=$alias)" -Server dom01 ) -eq $null ) 
        {
        #Write-Host "Creating $alias" -ForegroundColor Yellow
        Write-Progress -Activity $Activity -CurrentOperation "Creating $alias"
        $FullN = "$FirstN $LastN"
        $principal = $alias+"@lifepathsystems.org"
        $UserObject | Add-Member -MemberType NoteProperty -Name DisplayName -Value $FullN
        $UserObject | Add-Member -MemberType NoteProperty -Name Alias -Value $alias
        }
    elseif ( ( Get-ADUser -LDAPFilter "(sAMAccountName=$aliaswithMI)" -Server dom01 ) -eq $null )
        {
        #Write-Host "$alias already exists."
        Write-Progress -Activity $Activity -CurrentOperation "$alias already exists."
        $alias = $aliaswithMI
        #Write-Host "Creating $alias" -ForegroundColor Yellow
        Write-Progress -Activity $Activity -CurrentOperation "Creating $alias"
        $FullN = "$($FirstN) $($MI). $($LastN)"
        $principal = $alias+"@lifepathsystems.org"
        $UserObject | Add-Member -MemberType NoteProperty -Name DisplayName -Value $FullN
        $UserObject | Add-Member -MemberType NoteProperty -Name Alias -Value $alias       
        }
    else 
        {
        Write-Progress -Activity $Activity -Completed
        #Write-Host "Both $alias and $aliaswithMI taken. Canceled." -ForegroundColor Red
        $Cancel = $True
        If ( $alias -eq $aliaswithMI )
            {
            $FullN = "$($FirstN) $($LastN)"
            $UserObject | Add-Member -MemberType NoteProperty -Name DisplayName -Value $FullN
            $UserObject | Add-Member -MemberType Noteproperty -Name Error -Value "User Creation Cancelled. $($alias) already exists and no MI was provided." 
            }
        else
            {
            $FullN = "$($FirstN) $($MI). $($LastN)"
            $UserObject | Add-Member -MemberType NoteProperty -Name DisplayName -Value $FullN
            $UserObject | Add-Member -MemberType NoteProperty -Name Error -Value "User Creation Cancelled. $($alias) and $($aliaswithMI) already exists."
            }
        Return $UserObject
        }
    
    If ( !$Cancel )
        {
	try 
	    {
	    New-ADUser -UserPrincipalName $principal -SamAccountName $alias -DisplayName $fulln -Name $fulln -GivenName $firstn -Surname $lastn -Title $Title -Description $Title -Department $Department -Office $Office -AccountPassword $Password -Enabled $Enabled -OtherAttributes @{'msExchHideFromAddressLists'=$HideInAddressBook; 'EmployeeID'=$EmployeeID} -Server dom01 -ErrorAction stop | Out-Null
	    Set-ADuser -Identity $alias -ChangePasswordAtLogon $True -Server dom01
        try
            {
            $ManagerName =  (Get-ADUser $Manager -ErrorAction SilentlyContinue).DisplayName
            Set-ADuser -Identity $alias -Manager $Manager -Server dom01
            }
        catch
            {
            Write-Host "Manager doesn't exist in Active Directory. Not setting the Manager property"
            $ManagerName = ""
            }
        $UserObject | Add-Member -MemberType NoteProperty -Name Manager -Value $ManagerName
	    #Write-Host "Adding Group Memberships" -ForegroundColor Yellow
	    $UserObject | Add-Member -MemberType NoteProperty -Name Template -Value $Template
	    $UserObject | Add-Member -MemberType NoteProperty -Name Password -Value $UnencryptedPassword
	    Write-Progress -Activity $Activity -CurrentOperation "Adding Group Memberships"
	    $groups = (Get-ADUser $Template -Properties memberof).memberof
	    $groups | Where-Object { $_ -notmatch $LicenseGroup} | Get-ADGroup -Server dom01 | Add-ADGroupMember -Members $alias -Server dom01
	    if ( $Mailbox )
            {
            Write-Progress -Activity $Activity -CurrentOperation "Adding Membership to $LicenseGroup"
            Get-ADGroup $LicenseGroup -Server dom01 | Add-ADGroupMember -Members $alias -Server dom01
            Write-Progress -Activity $Activity -CurrentOperation 'Setting "EmailAddress" and "mail" property in AD'
            Set-ADUser -Identity $alias -EmailAddress $principal -Add @{proxyAddresses="SMTP:$alias@lifepathsystems.org", "smtp:$alias@lifepathsystems.mail.onmicrosoft.com", "smtp:$alias@lifepathsystems.onmicrosoft.com"; mailNickName="$alias"} -Server dom01
            $UserObject | Add-Member -MemberType NoteProperty -Name Email -Value $principal
            }
            #Write-Host "Setting Logon Hours based on $($Template)" -ForegroundColor Yellow
            Write-Progress -Activity $Activity -CurrentOperation "Setting Logon Hours based on $($Template)"
            $logonHours = (Get-ADUser $Template -Properties logonHours).logonHours
            Set-ADUser $alias -Replace @{logonhours = $logonHours} -Server dom01
            Write-Progress -Activity $Activity -CurrentOperation "Setting ScriptPath based on $($Template)"
            $ScriptPath = (Get-ADUser $Template -Properties ScriptPath).ScriptPath
            Set-ADUser $alias -ScriptPath $ScriptPath -Server dom01
	    If ( $HomeDirectory )
            {
            Write-Progress -Activity $Activity -CurrentOperation "HomeDirectory"
            Set-HomeDirectory -alias $alias -Department $Department -Office $Office
            }
	    Else    
            {
            $UserObject | Add-Member -MemberType NoteProperty -Name HomeDirectory -Value "None"
            Write-Progress -Activity $Activity -Completed
            }
	    }
	catch
	    {
	    Write-Progress -Activity $Activity -CurrentOperation "Failed: $_"
	    $UserObject | Add-Member -MemberType NoteProperty -Name Error -Value "$_"
	    sleep 5
	    }
        Return $UserObject
        Write-Progress -Activity $Activity -Completed
        }
    }

Function New-LPSUsersFromCSV
    {
	<#
	.SYNOPSIS
	Creates LifePath users from a CSV file.

	.DESCRIPTION
	This function reads a CSV file containing user information and creates new users in Active Directory. It utilizes the `New-LPSUser` cmdlet for user creation and can handle the creation of users, their group memberships, and other attributes based on a specified template. 

	The function supports two output modes:
	1. Standard output, where user creation details are returned to the console.
	2. Output to a timestamped NEO file (default), where user creation details are exported to a CSV file.

	The CSV file should include columns like:
	- First Name
	- Last Name
	- Employee ID
	- Department
	- Manager
	- Template
	- Enabled (Optional)
	- Mailbox (Optional)
	- HomeDirectory (Optional)

	If the `NoOutputFile` parameter is specified, the output will not be written to a file. If not specified, the default behavior is to output to a timestamped CSV file.

	.PARAMETER FilePath
	The path to the CSV file containing user data. The CSV file should have the necessary columns (`First Name`, `Last Name`, `Employee ID`, `Department`, etc.).

	.PARAMETER OutputDirectory
	(Optional) The directory where output files will be saved. If the `NoOutputFile` parameter is not used, this defines the folder for the CSV export. If not specified, the output file will be saved in the same directory as the input CSV.

	.PARAMETER NoOutputFile
	(Optional) If specified, the function will **not** export the user creation details to a file and will output to the console instead.

	.EXAMPLE
	New-LPSUsersFromCSV -FilePath "C:\Users\admin\Desktop\New Users.csv"
	Creates users based on the data in "New Users.csv" and exports the details to a timestamped CSV file (default behavior).

	.EXAMPLE
	New-LPSUsersFromCSV -FilePath "C:\Users\admin\Desktop\New Users.csv" -NoOutputFile
	Creates users from "New Users.csv" and outputs user creation details to the console without exporting to a file.

	.EXAMPLE
	New-LPSUsersFromCSV -FilePath "C:\Users\admin\Desktop\New Users.csv" -OutputDirectory "C:\Temp" -NoOutputFile
	Creates users from "New Users.csv" and outputs user creation details to the console without exporting to a file.

	.NOTES
	Author: Wayne Reeves
	The function requires the `New-LPSUser` cmdlet to create the users and assumes the presence of the appropriate templates for group memberships.
	#>
    [cmdletBinding()]
    Param(
	[Parameter(Mandatory)]
        [System.IO.FileInfo]$FilePath,
        [System.IO.DirectoryInfo]$OutputDirectory = ( Split-Path (Resolve-Path $FilePath) -Parent ),
        [switch]$NoOutputFile
    )
    $Users = Import-CSV $FilePath
    $UserObjects = New-Object System.Collections.ArrayList
    $UserObjects | Add-Member -MemberType NoteProperty -Name DisplayName
    $UserObjects | Add-Member -MemberType NoteProperty -Name Alias
    $UserObjects | Add-Member -MemberType NoteProperty -Name Email
    $UserObjects | Add-Member -MemberType NoteProperty -Name HomeDirectory
    $UserObjects | Add-Member -MemberType NoteProperty -Name Template
    $UserObjects | Add-Member -MemberType NoteProperty -Name Password
    $UserObjects | Add-Member -MemberType NoteProperty -Name Manager
    $UserObjects | Add-Member -MemberType NoteProperty -Name Error
    foreach ( $User in $Users)
        {
        $splat = @{}
	if ( $User.Mailbox ) { $User.Mailbox = [bool]::Parse($User.Mailbox) }
    if ( $User.Enabled ) { $User.Enabled = [bool]::Parse($User.Enabled) }
    $UserParameters = (Get-Command New-LPSUser).Parameters.Keys
    $User.PSObject.Properties | 
        Where-Object { $UserParameters -Contains $_.Name } | 
            ForEach-Object { $splat[$_.Name] = $_.Value }
    $UserObject = New-LPSUser @splat
    $User | Add-Member -MemberType NoteProperty -Name Alias -Value $UserObject.Alias
    $User | Add-Member -MemberType NoteProperty -Name Email -Value $UserObject.Email
    $User | Add-Member -MemberType NoteProperty -Name HomeDirectory -Value $UserObject.HomeDirectory
    $UserObjects.Add($UserObject) | Out-Null
    $splat = $null
        }
    if ( $NoOutPutFile )
        {
        $UserObjects
        }
    else 
        {
        if ( -not [DateTime]::TryParse($OutputDirectory.name,[ref][DateTime]::MinValue) )
            {
            $date = Get-Date -Format yyyMMdd
            }
        $OutputCSV = Join-Path $OutputDirectory.FullName "CreatedUsers$($date).csv"
        if ( Test-Path -LiteralPath $OutputCSV -PathType Leaf )
            {
            Write-Warning "Appending to existing file."
            }
        $UserObjects
        $Users | Select-Object -ExcludeProperty Mailbox, Enabled | Export-Csv -Path $OutputCSV -Append -Verbose
        }
    Get-NextADSync
    }

Function Import-HRData
    {
    <#
    .Synopsis
    Imports employees' Cost Centers and Managers into Active Directory.

    .DESCRIPTION
    Imports employees' Cost Centers and Managers from a CSV file into Active Directory. The CSV file requires the following columns: "first_name", "last_name", "Employee_Code", "Department", and "Supervisor_Primary_Code". It will find a single match for the EmployeeID and update the Cost Center for that employee in Active Directory.

    .NOTES   
    Name: Import-HRData
    Author: Wayne Reeves
    Version: 2024.07.15

    .PARAMETER FilePath
    The path of the CSV File you are importing.

    .EXAMPLE
    Import-HRData -FilePath C:\temp\HRData.csv

    Description:
    In this example you are specifying a path using the FilePath Parameter

    .EXAMPLE
    Import-HRData  C:\temp\HRData.csv

    Description:
    In this example you are specifying a path without using the FilePath Parameter. It will know implicitely this is the FilePath.
    #>

    [CmdletBinding()]
    param(
	[ValidateScript(
	{
	if(-Not ($_ | Test-Path) )
	    {
	    throw "File or folder does not exist"
	    }
	if(-Not ($_ | Test-Path -PathType Leaf) )
	    {
	    throw "The Path argument must be a file. Folder paths are not allowed."
	    }
	return $true 
	})]
	[Parameter(Position=0,mandatory=$true)]
	[System.IO.FileInfo]$FilePath
        )
    $Counter = 0
    Write-Progress -Activity "Import HR Data" -CurrentOperation "Getting All Users from Active Directory" -PercentComplete 0
    $AllUsers = Get-ADUser -Filter * -Properties EmployeeID -Server Dom01
    $IDs = Import-CSV $FilePath | Select-Object @{N="Name";E={"$($_.first_name+" "+$_.last_name)"}}, @{N="EmployeeID";E={$_.Employee_Code}}, @{N="EmployeeType";E={$_.Department}}, @{N="ManagerID";E={$_.Supervisor_Primary_Code}}
    $AllIDCount = ($IDs | Measure-Object).count
    $BadMatches = @()
    Foreach ( $ID in $IDs )
        {
        ++$Counter
        $Progress = ($Counter/$AllIDCount) * 100
        Write-Progress -Activity "Import HR Data" -CurrentOperation "Importing $ID.Name, $ID.EmployeeID" -PercentComplete $Progress
        $IDMatches = $AllUsers | Where-Object EmployeeID -eq $ID.EmployeeID
        $Manager = ($AllUsers | Where-Object EmployeeID -eq $ID.ManagerID).samaccountname
        $Count = ($IDMatches | Measure-Object).count
        If ( $Count -eq 1 )
            {
            Set-ADUser $IDMatches.SAMAccountName -EmployeeID $ID.EmployeeID -Replace @{EmployeeType=$ID.EmployeeType} -Manager $Manager -Server dom01
            }
        Elseif ( $Count -gt 1 )
            {
            $ID | Add-Member -MemberType NoteProperty -Name RecommendedAction -Value "Find and eliminate duplicate EmployeeID from Active Directory"
            $BadMatches += $ID
            Write-Progress -Activity "Import HR Data" -CurrentOperation "Skipping $ID" -PercentComplete $Progress
            }
        Else
            {
            $ID | Add-Member -MemberType NoteProperty -Name RecommendedAction -Value "Add EmployeeID to appropriate match in Active Directory"
            $BadMatches += $ID
            Write-Progress -Activity "Import HR Data" -CurrentOperation "Skipping $ID" -PercentComplete $Progress
            }
            }
    If ( $BadMatches )
        {
        Write-Host "HR Data imported with the following exceptions:"
        $BadMatches | Format-Table Name, EmployeeID, RecommendedAction
        }
    Else
        {
        Write-Host "All HR Data imported successfully"
        }
    }

Function Get-NextADSync
    {
    param(
    $Server="azuresync01"
    )
    $Results = Invoke-Command -ComputerName $Server -Scriptblock { Get-ADSyncScheduler }
    $NextSyncLocalTime = (Get-Date $($Results.NextSyncCycleStartTimeInUTC)).ToLocalTime()
    Write-Host "Next AD Sync Cycle Start Time: $($NextSyncLocalTime)" -ForegroundColor Yellow
    }

Function Set-ProfilePhotos
    {
    <#
    .SYNOPSIS
    Sets profile pics of users in Entra and/or Workvivo

    .DESCRIPTION
    Uploads profile pictures of users based on files named with their EmployeeIDs to Entra and/or Workvivo

    .NOTES   
    Name: Set-ProfilePhotos
    Author: Wayne Reeves
    Version: 2024.07.15

    .PARAMETER <FolderPath>
    Path where the pictures are stored; Defaults to current folder, if not specified

    .PARAMETER <FileType>
    File extenstion type to filter for; Valid values are "All", "png", "jpg"; Defaults to "All"

    .PARAMETER <Destination>
    Where you want to upload the photos; Options are "All", "Entra", or "Workvivo"; Defaults to "All"

    #>
    param(
      $FolderPath = $(Get-Location),
      [ValidateSet("All", "png", "jpg")]
      $FileType = "All",
      [ValidateSet("All", "Entra", "Workvivo")]
      $Destination = "All"
      )

    try
        { Get-Command "curl.exe" -ErrorAction Stop | Out-Null }
    catch
        {
        Write-Error "This command requires curl to function" -ErrorAction Stop
        }
    if ( $PSVersionTable.PSEdition -eq "Desktop" )
        {
        Write-Error 'This command requires "PowerShell Core" vs "Windows PowerShell"' -ErrorAction Stop
        }

    $Bearer = Get-XMLPassword -Name "WorkvivoAPI-1000152" -Type Password -AsPlainText $True
    if ( $null -eq $Bearer )
        {
        Write-Error -ErrorAction Break -Message 'No password for Workvivo API Provided'
        }

    $EmployeeListFile = Join-Path $FolderPath "EmployeeList.csv"

    Function Get-WorkvivoUser
        {
        param(
            $userID
            )
        $headers = @{
                    "Accept" = "application/json"
                    "Workvivo-Id" = "1000152"
                    "Authorization" = "Bearer $Bearer"
                    }
        try
            {
            $response = Invoke-WebRequest -Uri "https://api.workvivo.us/v1/users/by-email/$($userID)" -Headers $headers -ErrorAction Stop
            $WorkvivoUser = (ConvertFrom-Json $response.content).data
            return $WorkvivoUser
            }
        catch
            {
            Write-Error "No workvivo user found for $UserID"
            }
        }

    Function Set-WorkvivoPhoto
        {
        param(
            $userID,
            [System.IO.FileInfo]$InFile
            )
        $WorkvivoUser = Get-WorkvivoUser -userID $userID
        $WorkvivoUser = $WorkvivoUser.external_id

        if ( $WorkvivoUser )
            {
            $uri = "https://api.workvivo.us/v1/users/by-external-id/$WorkvivoUser/profile-photo"
            $response = curl.exe -s --location --request PUT $uri `
                --header 'Workvivo-Id: 1000152' `
                --header "Authorization: Bearer $Bearer" `
                --form image=@"$($InFile.FullName)"
            $response = (ConvertFrom-Json $response).data
            if ( -not $response.avatar_url )
                {
                Write-Error "Error Writing to Workvivo: $userID"
                }
            }
        }

    Function Set-EntraPhoto
        {
        param(
            $userID,
            [System.IO.FileInfo]$InFile
            )
        if ( $($($InFile.Length)/1MB) -le 4 )
            {
            Set-MgUserPhotoContent -UserId $userId -InFile $photoPath
            }
        else
            {
            Write-Error "$($inFile.name) is too large for Entra"
            }
        }

    if ( "All", "Entra" -contains $Destination )
        {
        Write-Progress -Activity "Setting User Profile Pics" -Status "Connecting to Microsoft Graph"
        Confirm-MgGraph -RequiredScopes "User.ReadWrite.All","Group.ReadWrite.All"
        }

    # Get all user profiles
    Write-Progress -Activity "Setting User Profile Pics" -CurrentOperation "Getting all users with EmployeeIDs"
    $users = Get-ADUser -Filter * -Properties employeeid -Server dom01 | Where-Object { $null -ne $_.EmployeeID -and $_.EmployeeID.StartsWith("A") }
    $users  | Select-Object EmployeeID, GivenName, SurName | Export-CSV -Path $EmployeeListFile -UseQuotes Never
    switch ( $FileType )
        {
        "All" { [array]$FileType = ".png", ".jpg" }
        "png" { [array]$FileType = ".png" }
        "jpg" { [array]$FileType = ".jpg" }
        }
    $Files = Get-ChildItem $FolderPath -File | Where-Object { $FileType -contains $_.Extension }
    $sum = $Files.count

    for ( $i=0; $i -lt $sum; $i++ )
        {
        $CurrentFile = $Files[$i]
        $employeeID = $CurrentFile.BaseName
        $user = ($users | Where-Object { $_.EmployeeID -eq $employeeID })
        $userId = $user.UserPrincipalName
        $photoPath = $CurrentFile.FullName
        # Check if the photo has a matching employee
        if ( $userID -and $user.Enabled -eq $True )
            {
            # Update the user's profile photo
            switch ( $Destination )
                {
                "All" 	    {
                            Write-Progress -Activity "Setting User Profile Pics" -Status "Setting Entra Profile Pic for $($userID):$($employeeID)" -PercentComplete  $(($i/$sum)*100)
                            Set-EntraPhoto -UserId $userId -InFile $photoPath
                            Write-Progress -Activity "Setting User Profile Pics" -Status "Setting Workvivo Profile Pic for $($userID):$($employeeID)" -PercentComplete  $(($i/$sum)*100)
                            Set-WorkvivoPhoto -userID $userID -InFile $CurrentFile
                            }
                "Entra"     {
                            Write-Progress -Activity "Setting User Profile Pics" -Status "Setting Entra Profile Pic for $($userID):$($employeeID)" -PercentComplete  $(($i/$sum)*100)
                            Set-EntraPhoto -UserId $userID -InFile $photoPath
                            Start-Sleep 1
                            }
                "Workvivo"  {
                            Write-Progress -Activity "Setting User Profile Pics" -Status "Setting Workvivo Profile Pic for $($userID):$($employeeID)" -PercentComplete  $(($i/$sum)*100)
                            Set-WorkvivoPhoto -userID $userID -InFile $CurrentFile
                            Start-Sleep 1
                            }
                }
            }
        else
            {
            Write-Error "No match found for $employeeID"
            }
        }
    }

Function Export-EntraSigninReport
    {
    <#
    .Synopsis
    Exports Microsoft Entra for Sign-in Logs for a user account to a csv

    .DESCRIPTION
    Queries Microsoft Entra for Sign-in Logs for a user account between the dates
    specified and outputs to a csv file.

    .NOTES   
    Name: Export-EntraSigninReport
    Author: Wayne Reeves
    Version: 2025.03.14

    .PARAMETER Username
    Asset Tag of the Computer you are searching for

    .PARAMETER StartDate
    The date you want to start the query from

    .PARAMETER EndDate
    The date for the last log

    .PARAMETER FilePath
    Specify the file path and name for the report

    .EXAMPLE
    Export-EntraSigninReport -Username wreeves -StartDate "2025-03-01 07:00AM" -EndDate "4PM"

    Description:
    Will get logs from March 1st, 2025 to 4PM today. 

    .EXAMPLE
    Export-EntraSigninReport -Username wreeves -StartDate "2025-03-01 07:00AM" -EndDate "2025-03-05"

    Description:
    Will get logs from 7AM March 1st, 2025 to 0 hour of 2025-03-05. Which means that you won't get logs for the day of the 5th, but you will get until the end of the 4th. PowerShell dates without a time default to 12:00AM.   

    .EXAMPLE
    Export-EntraSigninReport -Username wreeves -StartDate "7:00AM February 9" -EndDate "5PM March 1" -FilePath C:\temp\wreeves_export.csv

    Description:
    Will get logs from February 9th at 7AM to March 1st at 5PM and output the csv file to c:\temp\wreeves_export.csv
    #>

    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        $Username,
        [parameter(Mandatory=$true)]
        [datetime]$StartDate,
        [parameter(Mandatory=$true)]
        [datetime]$EndDate,
        $FilePath = (Join-Path $pwd.path "$Username.csv")
    )

    function Convert-ToUTC
        {
        param(
        [datetime]$Date
        )
        $Date = $Date.ToUniversalTime()
        Get-Date $Date -Format yyyy-MM-ddTHH:mm:ssZ
        }

    Function Convert-ToCurrentTZ
        {
        param(
        [datetime]$Date
        )
        $NewDate = (Get-Date $Date.tostring("yyyy-MM-ddTHH:mm:ssZ") -UFormat "%F %R TZOffset:%Z").tostring().Replace("TZOffset:-05","CDT").Replace("TZOffset:-06","CST")
        return [string]$NewDate
        }

    Function Get-EntraSigninLogs
        {
        param(
        $UserPrincipalName,
        $UTCStartDate,
        $UTCEndDate
        )
        Confirm-MgGraph -RequiredScopes 'AuditLog.Read.All','Directory.Read.All'
        Get-MgAuditLogSignIn -Filter "userPrincipalName eq `'$UserPrincipalName`' and createdDateTime ge $UTCStartDate and createdDateTime le $UTCEndDate" -ErrorAction 'Stop' -ErrorVariable EntraError
        }

    try
        {
        $UserPrincipalName = (Get-ADUser $Username).userprincipalname
        }
    catch
        {
        Write-Error "$Username not found"
        throw
        }

    $CSVFile = $FilePath

    if ( ((Get-Date) - $StartDate).days -ge 31 )
        {   
        $StartDate = (Get-Date).AddDays(-30)
        Write-Host "StartDate is greater than maximum of 30 days from current date. `nSetting StartDate to $StartDate" -ForegroundColor Yellow
        }

    $TotalDays = ($EndDate - $StartDate).days + 1
    $ChunkDays = 7
    $count = 0
    $total_failure = $true
    $entralogs = @()
    while ( $StartDate -lt $EndDate -and $null -eq $abort )
        {
        $logs = $null
        $PercentComplete = ( ( $count * $ChunkDays ) / $TotalDays ) * 100
        $TempEndDate = $StartDate.adddays($ChunkDays)
        [string]$UTCStartDate = Convert-ToUTC -Date $StartDate
        if ( $TempEndDate -gt $EndDate )
        {
        [string]$UTCEndDate = Convert-ToUTC -Date $EndDate
        }
        else
        {
        [string]$UTCEndDate = Convert-ToUTC -Date $TempEndDate
        }
        Write-Progress -Activity "Sign in logs" -Status "Fetching Entra Sign-in logs from $UTCStartDate to $UTCEndDate" -PercentComplete $PercentComplete
        $tries = 1
        $success = $false
        while ( $tries -le 3 -and $success -eq $false )
            {
            try {
                $logs = Get-EntraSigninLogs -UserPrincipalName $UserPrincipalName -UTCStartDate $UTCStartDate -UTCEndDate $UTCEndDate
                $success = $true
                $total_failure = $false
                }
            catch
                {
                if ( $count -eq 0 )
                    {
                    $abort = $true 
                    }
                $tries++
                if ( $tries -le 3 ) 
                    {
                    Write-Error "Failed to retrieve logs for period $UTCSTartDate to $UTCEndDate `nWill attempt to fetch again in 3 seconds."
                    Start-Sleep -Seconds 3
                    Write-Progress -Activity "Sign in logs" -Status "Attempt $tries/3 of Fetching Entra Sign-in logs from $UTCStartDate to $UTCEndDate" -PercentComplete $PercentComplete
                    }
                else
                    {
                    Write-Error "Aborted Fetch of Entra Sign-in logs from $UTCStartDate to $UTCEndDate"
                    }
                }
            }
        Remove-Variable tries
        $entralogs += $logs
        $StartDate = $TempEndDate
        Start-Sleep -Seconds 1
        $count++
        }

    if ( $total_failure -eq $false )
        {
        Write-Progress -Activity "Sign in logs" -Status "Writing CSV"
        $entralogs | Select-Object `
            @{e={(Convert-ToCurrentTZ -Date $_.createddatetime)};label="DateTime"},
            AppDisplayName,
            clientAppUsed,
            ipAddress,
            @{label="Location";e={"$($_.location.City), $($_.location.State), $($_.Location.CountryorRegion)"}},
            @{label="DeviceName";e={$_.DeviceDetail.DisplayName}},
            @{label="OperatingSystem";e={$_.DeviceDetail.OperatingSystem}} | Sort-Object DateTime | Export-csv $CSVFile
        Write-Host "Report written to $CSVFile" -ForegroundColor Yellow
        }
    else
        {
        Write-Error "No attempts to fetch the logs were successful. Operation aborted."
        }
    }

Function Export-NEOCredentials
    {
    <#
    .SYNOPSIS
    Exports temporary credentials for new employees and resets their passwords.

    .DESCRIPTION
    The Export-NEOCredentials function reads a CSV file containing new employee information, resets their passwords to temporary values, and generates an HTML report with the new credentials. This report can be printed for distribution to new hires during their orientation.

    .PARAMETER FilePath
    The path to the CSV file containing new employee information. The CSV should have columns for DisplayName and Alias.

    .EXAMPLE
    Export-NEOCredentials -FilePath CreatedUsers.csv

    This example reads the specified CSV file, resets the passwords for the new hires listed in the file, and generates an HTML report with their temporary credentials.

    .EXAMPLE
    Export-NEOCredentials -FilePath "C:\Path\To\CreatedUsers.csv"

    This example reads the specified CSV file, resets the passwords for the new hires listed in the file, and generates an HTML report with their temporary credentials.

    .NOTES
    Author: Wayne Reeves
    #>
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullorEmpty()]
        [ValidateScript({Test-Path $_})]
        [string]
        $FilePath
        )
    
    try
        {
        $UserInfo = Import-CSV $FilePath | Select-Object DisplayName, Alias
        $html = @(
            '<style>',
            'table, tr, td {',
            '    text-align: left;',
            '}',
            '</style>',
            '<body>'
            )

        foreach ( $User in $UserInfo )
            {
            $User | Add-Member -MemberType NoteProperty -Name Password -Value (New-RandomPassword)
            Set-ADAccountPassword -Identity $User.Alias -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $User.Password -Force) -server dom01
            Set-ADUser -Identity $User.Alias -ChangePasswordAtLogon $True -Server dom01
            $html += @(
                "<h3>$($User.DisplayName)</h3>",
                "<table>",
                "<tr>",
                "<td>Username:</td><th>$($User.alias)</th>",
                "</tr>",
                "<tr>",
                "<td>Temporary Password:</td><th>$($User.Password)</th>",
                "</tr>",
                "</table>",
                "<hr>"
                )
            }

        $html += @(
            '</body>',
            '</html'
            )
        $htmlfile = New-Item Accounts.html -Force
        $html | Out-File $htmlfile
        Start-Process $htmlfile.Name -Wait | Out-Null
        Remove-Item $htmlfile -Force
        }

    catch
        {
        Write-Error "An error occurred: $_"
        }
    }

Function Set-LPSUserStatus
    {
    <#
    .SYNOPSIS
    Centralizes our leave and off-boarding/on-boarding processes in a single function.

    .DESCRIPTION
    Set an AD user's state to Terminated, FMLA, or Return in one unified workflow.

    .PARAMETER SamAccountName
    The user's sAMAccountName (or Name).
    Accepts pipeline input by value (string) or by property name (ADUser.SamAccountName).
    Alias: **Identity**.

    .PARAMETER FMLA
    Switch to place the user on FMLA hold:
    - Disables the account
    - Adds to the "FMLA Users" group

    .PARAMETER Terminated
    Switch to terminate the user:
    - Disables the account
    - Adds to the "Terminated Users" group
    - Hides from the Exchange address lists
    - Sets `AccountExpirationDate = $DateTerminated + 90 days`

    .PARAMETER Return
    Switch to restore a previously held or terminated user:
    - Enables the account
    - Clears address-list hiding
    - Removes from both "FMLA Users" and "Terminated Users" groups
    - Clears any `AccountExpirationDate`

    .PARAMETER DateTerminated
    The base date for the 90-day expiration stamp when `-Terminated` is used.
    Defaults to `(Get-Date)` if omitted.

    .PARAMETER ZoomLicense
    Defines the type of Zoom license you want to apply for a returned user.

    Valid Options:
        - Zoom Basic
        - Zoom Licensed
        - None

    .OUTPUTS
    Microsoft.PowerShell.Commands.PSCustomObject[]
    A collection of PSCustomObjects, each with these properties:
        - Name              : The display name of the user processed.
        - SamAccountName    : The username of the user operated on.
        - Enabled           : If user is enabled or not
        - Action            : One of 'FMLA', 'Terminated', or 'Returned'
        - Expiration        : Expration of account, if applied.

    .EXAMPLE
    # Put a single user on FMLA hold:
    Set-LpsUserStatus 'jdoe' -FMLA

    .EXAMPLE
    # Terminate multiple users from the pipeline:
    'alice','bob' | Set-LpsUserStatus -Terminated

    .EXAMPLE
    # Terminate a user and back-date the termination marker:
    Set-LpsUserStatus -Identity cwilson -Terminated -DateTerminated 2025-04-01

    .EXAMPLE
    # Return a user from FMLA
    Set-LpsUserStatus jdoe -Return

    Will enable their account and remove them from the "FMLA Users" group.  

    .EXAMPLE
    # Return everyone in the "FMLA Users" group to active status:
    Get-ADGroupMember "FMLA Users" | Set-LpsUserStatus -Return

    .EXAMPLE
    # A terminated staff returning to employment
    Set-LpsUserStatus chickerson -Return

    Will enable account, remove from "Terminated Users" group, add back E3 Licenses, and ask which Zoom License to apply"

    .EXAMPLE
    # A terminated staff returning to employment
    Set-LpsUserStatus chickerson -Return -ZoomLicense "Zoom Licensed"

    Will enable account, remove from "Terminated Users" group, add back E3 Licenses, and add them to the "Zoom Licensed" Security group in Entra, giving them a Zoom License
    .NOTES
    Author: Wayne Reeves
    Created: 2025-05-02
    #>
    [CmdletBinding(
        SupportsShouldProcess=$true,
        ConfirmImpact = "Medium"
        )]
    param(
    [Parameter(
        Mandatory,
        Position = 0,
        ValueFromPipeline
        )]
    [Alias('Identity')]
    [string]$sAMAccountName,
    [Parameter(Mandatory, ParameterSetName = "FMLA")]
    [switch]$FMLA,
    [Parameter(Mandatory, ParameterSetName = "Terminated")]
    [switch]$Terminated,
    [Parameter(ParameterSetName = "Terminated")]
    [DateTime]$DateTerminated =  (Get-Date -Hour 0 -Minute 0 -Second 0 -Millisecond 0),
    [Parameter(Mandatory, ParameterSetName = "Return")]
    [switch]$Return,
    [Parameter(ParameterSetName = "Return")]
    [ValidateSet("Zoom Basic", "Zoom Licensed", "None")]
    [string]$ZoomLicense
    )
    begin
        {
        if ( $PSCmdlet.ParameterSetName -ne "FMLA" )
            {
            Confirm-MgGraph -RequiredScopes Group.ReadWrite.All, User.ReadWrite.All
            }
        $UserInfoList = New-Object System.Collections.Generic.List[PSObject]
        }
    process
        {
        $Identity = Get-ADUser -Identity $sAMAccountName -Properties Created -ErrorAction SilentlyContinue -Server dom01
            $UserInfo = [PSCustomObject]@{
                Name = $Identity.Name
                SamAccountName = $Identity.SamAccountName
                Enabled = $null
                Action = $PsCmdlet.ParameterSetName
                Expiration = $null
                }
        if ( -not $Identity )
            {
            Write-Warning "Could not find user name $($Identity). Skipping."
            return
            }
        if ( $PSCmdlet.ShouldProcess($Identity.Name,"Set user as $($PSCmdlet.ParameterSetName)") )
            {
            if ( $PSCmdlet.ParameterSetName -ne "Return" )
                {
                Disable-ADAccount -Identity $Identity -Server dom01
                $UserInfo.Enabled = $False
                Add-ADGroupMember -Identity "$($PSCmdlet.ParameterSetName) Users" -Members $Identity -Server dom01
                if ( $PSCmdlet.ParameterSetName -eq "Terminated" )
                    {
                    if ( $Identity.Created -le $DateTerminated.AddDays(-30) )
                        {
                        $Expiration = $DateTerminated.AddDays(90)
                        }
                    else
                        {
                        $Expiration = Get-Date
                        }
                    $UserInfo.Expiration = $Expiration
                    Set-ADUser -Identity $Identity -add @{msExchHideFromAddressLists=$true} -Server dom01
                    Set-ADAccountExpiration -Identity $Identity -DateTime $Expiration -Server dom01
                    $E3Groups = Get-ADPrincipalGroupMembership $Identity -Server dom01 |
                        Where-Object { $_.name -like "E3*License*" -or $_.name -eq "OneDrive" }
                    $E3Groups | Remove-AdGroupMember -members $Identity.sAMAccountName -Confirm:$False -Server dom01
                    $EntraIdentity = Get-MgUser -UserId $Identity.userprincipalname
                    $ZoomGroups = Get-MgUserMemberOf -UserId $Identity.userprincipalname -all |
                        Where-Object { $_.additionalproperties['displayName'] -match "Zoom" }
                    if ( $ZoomGroups.count -gt 0 )
                        {
                        $ZoomGroups |
                            Foreach-Object {
                                Remove-MgGroupMemberbyRef -GroupID $_.Id -DirectoryObjectID $EntraIdentity.Id -Confirm:$False
                                }
                        }
                    }
                }
            else
                {
                Enable-ADAccount -Identity $Identity -Server dom01
                $UserInfo.Enabled = $True
                Set-ADUser -Identity $Identity -clear msExchHideFromAddressLists -Server dom01
                $MemberOf = Get-ADPrincipalGroupMembership $Identity -Server dom01
                if ( $MemberOf.Name -contains "Terminated Users" )
                    {
                    Write-Verbose "Adding E3 Licenses"
                    "E3 Simple Licenses", "E3 Teams License", "E3 Stream License", "OneDrive" |
                        Add-AdGroupMember -Members $Identity -Server dom01
                    if ( -not $ZoomLicense )
                        {
                        $ZoomChoiceTitle = "Zoom License Group"
                        $ZoomChoiceMessage = "Choose which Zoom license to apply to $($Identity.DisplayName)"
                        $ZoomChoiceOptions = @(
                            [System.Management.Automation.Host.ChoiceDescription]::new('Zoom &Basic')
                            [System.Management.Automation.Host.ChoiceDescription]::new('Zoom &Licensed')
                            [System.Management.Automation.Host.ChoiceDescription]::new('&None')
                            )
                        $ZoomChoiceDefault = 2
                        $ZoomChoiceSelection = $host.ui.PromptForChoice($ZoomChoiceTitle, $ZoomChoiceMessage, $ZoomChoiceOptions, $ZoomChoiceDefault)
                        $ZoomLicense = $ZoomChoiceOptions[$ZoomChoiceSelection].Label.Replace('&','')
                        }
                    if ( $ZoomLicense -match "Zoom" )
                        {
                        $ZoomGroup = Get-MgGroup -Filter "DisplayName eq `'$ZoomLicense`'"
                        $EntraIdentity = Get-MgUser -UserId $Identity.userprincipalname
                        New-MgGroupMember -GroupID $ZoomGroup.Id -DirectoryObjectID $EntraIdentity.Id -Confirm:$False
                        }
                    }
                try
                    {
                    "Terminated Users", "FMLA Users" | Remove-ADGroupMember -Members $Identity -Confirm:$False -Server dom01
                    }
                catch
                    {
                    $Error
                    }
                Clear-ADAccountExpiration -Identity $Identity -Server dom01
                }
                if ( $PSCmdlet.ParameterSetName -ne "FMLA" )
                    {
                    try
                        {
                        Remove-ADGroupMember "FMLA Users" -Members $Identity -Confirm:$False -Server dom01
                        }
                    catch
                        {
                        Write-Warning "Couldn't remove $($Identity.Displayname) from "FMLA Users" group"
                        }
                    }
                $UserInfoList.Add($UserInfo)
            }
        }
        end
            {
            $UserInfoList
            }
    }

Function Get-LPSExpiredTermination
    {
    <#
    .SYNOPSIS
        Lists terminated user accounts that have exceeded expiration date (90 days)

    .DESCRIPTION
        Will query for users that are terminated and have expired. Defaults to look for users in the "Terminated Users" group and that have exceeded 90 day expiration date.

    .PARAMETER GroupName
        Overrides the default "Terminated Users" group by specifying the name of a different security group.

    .PARAMETER ReportDate
        Overrides the comparison date from the expriation date, which defaults to today.

    .EXAMPLE
        Get-LPSExpiredTermination

        AccountExpirationDate : 1/5/2021 12:00:00 AM
        DistinguishedName     : CN=Jo Beth Collier,CN=Users,DC=ccmhmr,DC=local
        Enabled               : False
        GivenName             : Jo Beth
        Name                  : Jo Beth Collier
        ObjectClass           : user
        ObjectGUID            : 4011cc1a-6ecb-4b05-860f-7e0c64ab4cc9
        SamAccountName        : jcollier
        SID                   : S-1-5-21-848239347-1132225710-1734353810-33999
        Surname               : Collier
        UserPrincipalName     : jcollier@lifepathsystems.org

        AccountExpirationDate : 5/22/2025 10:28:01 AM
        DistinguishedName     : CN=Ikemba Dyke,CN=Users,DC=ccmhmr,DC=local
        Enabled               : False
        GivenName             : Ikemba
        Name                  : Ikemba Dyke
        ObjectClass           : user
        ObjectGUID            : a3e18ba4-24c3-4e63-ab54-ef77bf016cc0
        SamAccountName        : idyke
        SID                   : S-1-5-21-848239347-1132225710-1734353810-40526
        Surname               : Dyke
        UserPrincipalName     : idyke@lifepathsystems.org

    .EXAMPLE
        Get-LPSExpiredTermination -ReportDate 2025-04-28

        Queries for users that have expired after 2025-04-28 instead of the current date

    .EXAMPLE
        Get-LPSExpiredTermination -GroupName "Example Group"

        Queries by users in the "Example Group" Security group instead of "Terminated Users
    #>
    param (
        [string]$GroupName = "Terminated Users",
        [DateTime]$ReportDate = ( Get-Date )
    )
        $Terminated = Get-ADGroupMember "Terminated Users" -Server dom01
        $Expired = $Terminated |
            Get-ADUser -Properties AccountExpirationDate -Server dom01 |
                Where-Object { ( $_.AccountExpirationDate -lt $ReportDate ) -and
                    ( $null -ne $_.AccountExpirationDate ) }
        $Expired
    }


Function Remove-LPSUser
    {
    <#
    .SYNOPSIS
        Deletes LifePath users from Active Directory.

    .DESCRIPTION
        Deletes LifePath users from Active Directory. Outputs path to their HomeDrive so you can move/delete it yourself.

    .NOTES
        Cleaning up/moving the HomeDrive is performed manually by the IT Staff. (for now)

    .PARAMETER sAMAccountName
        Username of Active Directory user you wish to remove.

    .EXAMPLE
        Remove-LPSUser zztest

        Will delete the user account for zztest. Outputs the path to the HomeDrive so you can clean it up.

    .EXAMPLE
        Get-LPSExpiredTermination | Remove-LPSUser

        Removes all expired Terminated users (Users that are in the Terminated Users group and their account has expired).
    #>
    [cmdletBinding(
        SupportsShouldProcess = $true,
        ConfirmImpact = 'High'
        )]
    param (
        [Parameter(
            Mandatory,
            Position = 0,
            ValueFromPipeline
            )]
        [Alias('Identity')]
        [string]$sAMAccountName
        )
        begin
            {
            $UserInfoList = New-Object System.Collections.Generic.List[PSObject]
            }

        process
        {
        $Identity = Get-ADUser -Identity $sAMAccountName -ErrorAction SilentlyContinue -Server dom01 -Properties HomeDirectory
        if ( -not $Identity )
            {
            Write-Warning "Could not find user name $($Identity). Skipping."
            return
            }
        $Confirmed = $PSCmdlet.ShouldProcess($Identity.Name,"Delete Active Directory account")
        if ( $confirmed )
            {
            Write-Verbose "Removing $Identity.Name"
            Remove-ADUser $Identity
            }
        if ( $Confirmed -or $WhatIfPreference )
            {
            if ( $Identity.HomeDirectory )
                {
                $HomeDirectory = $Identity.HomeDirectory.Replace("\\","").Split("\")
                try
                    {
                    $HomeShare = Get-SMBShare -Name $HomeDirectory[1] -CimSession (New-CimSession $HomeDirectory[0]) -ErrorAction Stop
                    $HomeSharePath = $HomeShare.Path
                    }
                catch
                    {
                    $HomeSharePath = "Share not found"
                    }
                }
            $UserInfo = [PSCustomObject]@{
                Name = $Identity.Name
                SamAccountName = $Identity.SamAccountName
                HomeDirectory = $Identity.HomeDirectory
                HomeShareServer = $HomeDirectory[0]
                HomeSharePath = $HomeSharePath
                }
            $UserInfoList.Add($UserInfo)
            }
        }

        end
            {
            Write-Host "$(if ($WhatIfPreference ) { "What if: " })Removed Active Directory accounts. You man now clean/move the HomeDirectory folders."
            $UserInfoList
            }
    }

