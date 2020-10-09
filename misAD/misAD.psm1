#requires -Modules ActiveDirectory, misScripting
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
    Find-ADcomputer -Asset XXXX -Server DC01

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
    Find-ADUser test -Server DC01

    Description:
    Will show you user(s) that match the string filter "test" but query from "DC01"
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
        if ( !$Server ) { $Server = "DC01" }
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
    Will reset a users password to "mouse99!"

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
    Will create a menu for any users that matches the string filter "test" and will reset the password of the user you select

    .EXAMPLE
    Reset-Password test -DoNotChangePasswordAtLogon

    Description:
    Will create a menu for any users that matches the string filter "test" and will reset the password of the user you select. Will NOT require the user to change password at logon.

    .EXAMPLE
    Reset-Password test -Server DC01

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
            if ( !$Server ) { $Server = "DC01" }
            $samaccountname = (Select-User $Filter).samaccountname 
            if ( $samaccountname )
                {
                If ( $GenerateRandomPassword )
                    {
                    $Password = New-RandomPassword
                    }
                Else
                    {
                    $Password = "mouse99!"
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
    Unlock-ADUser test -Server DC01

    Description:
    Will check DC01 to see if User is locked out. If it is locked out it will unlock the user. If it is not locked it will tell you and ask if you would like to check all DCs.
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

    .PARAMETER SendEmail 
    True or False to send email template to yourself. ( Default True )
    
    .EXAMPLE
    New-LPSUser -FirstN Bob -MI S -LastN Cratchet -Title Hero -Office "BH McKinney" -Department BH -Template mwarren 

    Description:
    Will create a new user and mailbox for "Bob S Cratchet." "mwarren" will be used as a template for Group Memberships. User will be disabled and hidden from Address Book.
    An email template will be sent to you for sending on to staff.

    .EXAMPLE
    New-LPSUser -FirstN Bob -MI S -LastN Cratchet -Title Hero -Office "BH McKinney" -Department BH -Template mwarren -Mailbox $False

    Description:
    Will create a new user without a mailbox for "Bob S Cratchet." "mwarren" will be used as a template for Group Memberships. User will be disabled and hidden from Address Book.
    An email template will be sent to you for sending on to staff.

    .EXAMPLE
    New-LPSUser -FirstN Bob -MI S -LastN Cratchet -Title Hero -Office "BH McKinney" -Department BH -Template mwarren

    Description:
    Specify to not send an email template to you.

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
        [string]$Template, 
        [bool]$HomeDirectory=$True, 
        [bool]$Enabled=$False, 
	[bool]$Mailbox=$True,
	[string]$LicenseGroup='E3 Simple Licenses',
        [bool]$SendEmail=$True
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
		    Set-AdUser -Identity $alias -HomeDirectory $SharePath -HomeDrive "N:" -Server DC01
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
    
    Function Send-Email($DisplayName, $Alias)
        {
        $MISCreatorEmail = $ENV:username + "@lifepathsystems.org"
        $Subject = "Login information for new staff $($DisplayName)"
        $Body = @"
<p>Here is the login information for your new staff member:</p>

<p>
Computer login ID: <b>$($Alias)</b>
<br>
Computer temporary password: <b>$($UnencryptedPassword)</b>
</p>  

<p>Anasazi staff IDs:</p>
</p>
"@
        Send-MailMessage `
            -From "New LPS User <noreply@newlpsuser.org>" `
            -To $MISCreatorEmail `
            -Subject $Subject `
            -BodyAsHTML $Body `
            -SMTPServer "misexch01.ccmhmr.local"
        
        Write-Host "Sending Email to $($MISCreatorEmail)"
        }
    
    #Write-Host "Checking if username already exists" -ForegroundColor Yellow
    Write-Progress -Activity $Activity -CurrentOperation "Checking if username already exists" 
    
    If ( ( Get-ADUser -LDAPFilter "(sAMAccountName=$alias)" -Server DC01 ) -eq $null ) 
        {
        #Write-Host "Creating $alias" -ForegroundColor Yellow
        Write-Progress -Activity $Activity -CurrentOperation "Creating $alias"
        $FullN = "$FirstN $LastN"
        $principal = $alias+"@lifepathsystems.org"
        $email = $alias+"@lifepathsystems.org"
        $UserObject | Add-Member -MemberType NoteProperty -Name DisplayName -Value $FullN
        $UserObject | Add-Member -MemberType NoteProperty -Name Alias -Value $alias
        }
    elseif ( ( Get-ADUser -LDAPFilter "(sAMAccountName=$aliaswithMI)" -Server DC01 ) -eq $null )
        {
        #Write-Host "$alias already exists."
        Write-Progress -Activity $Activity -CurrentOperation "$alias already exists."
        $alias = $aliaswithMI
        #Write-Host "Creating $alias" -ForegroundColor Yellow
        Write-Progress -Activity $Activity -CurrentOperation "Creating $alias"
        $FullN = "$($FirstN) $($MI). $($LastN)"
        $principal = $alias+"@ccmhmr.local"
        $email = $alias+"@lifepathsystems.org"
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
	    New-ADUser -UserPrincipalName $principal -SamAccountName $alias -DisplayName $fulln -Name $fulln -GivenName $firstn -Surname $lastn -Title $Title -Description $Title -Department $Department -Office $Office -AccountPassword $Password -Enabled $Enabled -OtherAttribute @{'msExchHideFromAddressLists'=$HideInAddressBook} -Server DC01 -ErrorAction stop | Out-Null
	    Set-ADuser -Identity $alias -ChangePasswordAtLogon $True -Server DC01
	    #Write-Host "Adding Group Memberships" -ForegroundColor Yellow
	    $UserObject | Add-Member -MemberType NoteProperty -Name Template -Value $Template
	    $UserObject | Add-Member -MemberType NoteProperty -Name Password -Value $UnencryptedPassword
	    Write-Progress -Activity $Activity -CurrentOperation "Adding Group Memberships"
	    $groups = (Get-ADUser $Template -Properties memberof).memberof
	    $groups | Where-Object { $_.Name -ne $LicenseGroup} | Get-ADGroup -Server DC01 | Add-ADGroupMember -Members $alias -Server dc01
	    if ( $Mailbox )
		{
		Write-Progress -Activity $Activity -CurrentOperation "Adding Membership to $LicenseGroup"
		Get-ADGroup $LicenseGroup -Server DC01 | Add-ADGroupMember -Members $alias -Server DC01
		Write-Progress -Activity $Activity -CurrentOperation 'Setting "EmailAddress" and "mail" property in AD'
		Set-ADUser -Identity $alias -EmailAddress $principal -Server DC01
		}
	    #Write-Host "Setting Logon Hours based on $($Template)" -ForegroundColor Yellow
	    Write-Progress -Activity $Activity -CurrentOperation "Setting Logon Hours based on $($Template)"
	    $logonHours = (Get-ADUser $Template -Properties logonHours).logonHours
	    Set-ADUser $alias -Replace @{logonhours = $logonHours} -Server DC01
	    Write-Progress -Activity $Activity -CurrentOperation "Setting ScriptPath based on $($Template)"
	    $ScriptPath = (Get-ADUser $Template -Properties ScriptPath).ScriptPath
	    Set-ADUser $alias -ScriptPath $ScriptPath -Server DC01
	    If ( $HomeDirectory )
		{
		Write-Progress -Activity $Activity -CurrentOperation "HomeDirectory"
		Set-HomeDirectory -alias $alias -Department $Department -Office $Office
		}
	    Else    
		{
		$UserObject | Add-Member -MemberType NoteProperty -Name HomeDirectory -Value "None"
		Write-Progress -Activity $HDActivity -Completed
		}

	    If ( $SendEmail )
		{
		Send-Email -DisplayName $UserObject.DisplayName -Alias $UserObject.alias
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
    .Synopsis
    Creates LifePath Users from CSV File

    .DESCRIPTION
    Creates LifePath Users from CSV File by utilizing the New-LPSUser cmdlette

    .NOTES   
    Name: New-LPSUsersFromCSV
    Author: Wayne Reeves
    Version: 11.29.17

    .PARAMETER Path
    Either the Path to the CSV File you are pulling from or if you are in the current directory just the name of the file. The CSV requires two columns: "SAMAccountName" and "EmployeeID".

    .PARAMETER SendEmail
    True or Fasle to specifies to send email template. ( Default True )

    .EXAMPLE
    New-LPSUsersFromCSV -Path "New Users.csv"

    Description:
    In this example you are already in the current directory that the CSV File resides. It will pull in the information and create each user specified

    .EXAMPLE
    New-LPSUsersFromCSV -Path "New Users.csv" -SendEmail:$False

    Description:
    In this example you are telling the script to not send you an email template for each user.

    .EXAMPLE
    New-LPSUsersFromCSV -Path "C:\Users\jdoe\Desktop\New Users.csv"

    Description:
    In this example you are explicitely setting the Full Path to the file. It will pull in the information and create each user specified
    #>
    [cmdletBinding()]
    Param(
	[Parameter(Mandatory)]
        [string]$Path,
        [bool]$SendEmail=$True
    )
    $Users = Import-CSV $Path
    $UserObjects = New-Object System.Collections.ArrayList
    $UserObjects | Add-Member -MemberType NoteProperty -Name DisplayName
    $UserObjects | Add-Member -MemberType NoteProperty -Name Alias
    $UserObjects | Add-Member -MemberType NoteProperty -Name HomeDirectory
    $UserObjects | Add-Member -MemberType NoteProperty -Name Template
    $UserObjects | Add-Member -MemberType NoteProperty -Name Password
    $UserObjects | Add-Member -MemberType NoteProperty -Name Error
    foreach ( $User in $Users)
        {
        $splat = @{}
	if ( $User.Mailbox )
	    {
	    $User.Mailbox = [bool]::Parse($User.Mailbox)
	    }
        if ( !$SendEmail )
            {
            $User | Add-Member -Type NoteProperty -Name SendEmail -Value $False
            }
        $User.psobject.properties | ForEach-Object { $splat[$_.Name] = $_.Value }
        $UserObject = New-LPSUser @splat
        $UserObjects.Add($UserObject) | Out-Null
        $splat = $null
        }
    $UserObjects
    Get-NextADSync
    }

Function Import-AnasaziIDs
    {
    <#
    .Synopsis
    Imports employees' Anasazi IDs into Active Directory

    .DESCRIPTION
    Imports employees' Anasazi IDs from a CSV file into Active Directory. Will display the results from Active Directory after finishing.

    .NOTES   
    Name: Import-AnasaziIDs
    Author: Wayne Reeves
    Version: 11.29.17

    .PARAMETER Path
    The path of the CSV File you are importing. The Default is "\\missvr2\mis\Apps\Anasazi\Active Directory Updates\AnasaziIDs.csv"

    .EXAMPLE
    Import-AnasaziIDs

    Description:
    In this example you are simply using the default path after preloading the CSV file. No parameters are required.

    .EXAMPLE
    Import-AnasaziIDs -Path C:\temp\AnasaziIDs.csv

    Description:
    In this example you are telling the script to pull from a different path for the CSV file.
    #>

    param(
        $Path = "\\missvr2\mis\Apps\Anasazi\Active Directory Updates\AnasaziIDs.csv"
        )
    $IDs = Import-CSV $Path
    Foreach ( $ID in $IDs )
        {
        Set-ADUser $ID.SAMAccountName -EmployeeID $ID.EmployeeID -Server DC01
        }
    $List = $IDs | Foreach { Get-ADUser $_.SAMAccountName -Properties EmployeeID -Server DC01 } 
    $List | Select Name, EmployeeID
    }

Function Add-PhoneUser
    {
    <#
    .Synopsis
    Sets up an Active Directory user with settings for use with company phones.

    .DESCRIPTION
    Adds the user to the "All Hours Access" Security Group, Sets their logon hours to all hours, then enables ActiveSync on their mailbox.

    .NOTES   
    Name: Add-PhoneUser
    Author: Wayne Reeves
    Version: 10.11.19

    .PARAMETER Username
    The username of the employee you are setting up for company phone use.

    .EXAMPLE
    Add-PhoneUser -Username zztest

    Description:
    In this example you are setting up zztest for company phone use.

    .EXAMPLE
    Add-PhoneUser zztest

    Description:
    This is the same as the example above, but you are not specifying the "Username" property since it is implied.
    #>
    param(
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [string]$Username
        )
    if ( Get-ADUser $Username )
        {
        [byte[]]$LogonHours = @(255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255)
        Add-ADGroupMember -Identity "All Hours Access" -Members $Username 
        Set-ADUser $Username -Replace @{ logonhours = $LogonHours } 
        Set-CASMailbox $Username -ActiveSyncEnabled $True 
        }
    }

Function Get-NextADSync
    {
    param(
    $Server="migrate365"
    )
    $Results = Invoke-Command -ComputerName $Server -Scriptblock { Get-ADSyncScheduler }
    $NextSyncLocalTime = (Get-Date $($Results.NextSyncCycleStartTimeInUTC)).ToLocalTime()
    Write-Host "Next AD Sync Cycle Start Time: $($NextSyncLocalTime)" -ForegroundColor Yellow
    }
