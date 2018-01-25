#Exchange 2010/2013/2016/EXO mailbox delegate management module
#v1.5.1 9/7/17
#1.5.1 Added localization support for Deleted Items and Sent Items folders (automatic as needed), added permission check to owner mailbox
#1.5.0 Moved settings to config file, removed extraneous functions and need for connection to AAD, added option to disable meeting request forwarding
#1.4.7 Fixed loophole that could cause the module to use a credential object that doesn't exist
#1.4.6 Changed SID resolution for EXO to use SMTP because Translate doesn't work
#1.4.5 Fixed errors when working with EXO, added use of AAD module, remove SI/DI permission when removing delegate

#Check for EWS API
$apiPath = (($(gp -ErrorAction SilentlyContinue -Path Registry::$(dir -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services' |
			sort -Property Name -Descending | select -First 1 -ExpandProperty Name)).'Install Directory') + 'Microsoft.Exchange.WebServices.dll')
if (Test-Path -Path $apiPath)
	{
	Add-Type -Path $apiPath
	}
else
	{
	Write-Error -Message 'The Exchange Web Services Managed API is required to use this script.' -Category NotInstalled
	break
	}

#Import settings from configuration file
$sourceFileName = 'DelegateManagementSettings.xml'
$settingsFile = $PSScriptRoot + "\$sourceFileName"
if (-not(Test-Path -Path $settingsFile))
	{
    #Config file not found, so set defaults for the current session
	Write-Warning -Message "Settings file `"$sourceFileName`" cannot be found in $PSScriptRoot. Default settings will be used.  Use Get-/Set-DelegateManagementSetting to view and modify."
	$script:targetEnvironment = 'EXO'
	$script:useAutodiscover = $false
	$script:ewsUri = $null
	$script:useImpersonation = $false
	$script:applyPermToSentItems = $true
	$script:applyPermToDeletedItems = $true
	}
else
	{
	#Set session defaults from config file
    $settings = ([xml](cat -Path $settingsFile)).Settings
	$script:targetEnvironment = $settings.ConnectionMode
	$script:useAutodiscover = [System.Convert]::ToBoolean($settings.UseAutoDiscover)
	$script:ewsUri = $settings.EwsUrl
	$script:useImpersonation = [System.Convert]::ToBoolean($settings.UseImpersonation)
	$script:applyPermToSentItems = [System.Convert]::ToBoolean($settings.ApplyPermissionToSentItemsFolder)
	$script:applyPermToDeletedItems = [System.Convert]::ToBoolean($settings.ApplyPermissionToDeletedItemsFolder)
	}

#Region Helper Functions
function Test-ExchangeConnectivity
    {
    if (-not(gcm Get-Mailbox -ErrorAction SilentlyContinue))
	    {
        Write-Error "Connect to an Exchange environment before running a delegate cmdlet in the module." -Category ConnectionError
		return $false
        }
	$true
	}

function Get-EwsCredential
	{
	if (-not($EWSCredential))
		{
		$script:EWSCredential = Get-Credential -Message 'Enter the credentials to use to access Exchange Online mailboxes.'
		}
	}

function Connect-WebServices ($smtpAddress)
	{
	#The functionality of the module is not leveraging any features in a later schema,
	#so 2010 SP2 is sufficient regardless of the actual server version, whether on-prem or online
	$exchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
	$script:exchangeService = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExchangeService($exchangeVersion) 
 	if ($targetEnvironment -eq 'EXO')
        {
        Get-EwsCredential
		$exchangeService.Credentials = New-Object -TypeName Microsoft.Exchange.WebServices.Data.WebCredentials($EWSCredential)
        }
    #Use autodiscover or hard-coded URL
	if ($targetEnvironment -eq 'EXO')
        {
        $exchangeService.Url = 'https://outlook.office365.com/EWS/Exchange.asmx'
        }
    elseif ($useAutodiscover)
        {
        Write-Progress -Activity "Connecting to EWS" -Status "Performing autodiscover lookup for EWS endpoint"
        $exchangeService.AutodiscoverUrl($smtpAddress, {$true})
        Write-Progress -Activity "Connecting to EWS" -Completed -Status " "
        }
	else
		{
		$exchangeService.Url = $ewsUri
		}
	
	#Impersonate mailbox
	if ($useImpersonation)
		{
		$exchangeService.ImpersonatedUserId = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $smtpAddress)
		}
	
	New-Object -TypeName Microsoft.Exchange.WebServices.Data.Mailbox($smtpAddress)
	}
	
function Get-Delegates($EWSMailbox,$delegateToRetrieve,[switch]$includePermissions)
	{
	try
		{
		if ($delegateToRetrieve)
			{
			,$exchangeService.GetDelegates($EWSMailbox,$true,$delegateToRetrieve)
			}
		else
			{
			if ($includePermissions)
				{
				,$exchangeService.GetDelegates($EWSMailbox,$true)
				}
			else
				{
				,$exchangeService.GetDelegates($EWSMailbox,$false)
				}
			}
		}
	catch
		{
		if ($_.Exception -like '*The specified object was not found in the store.*')
			{
			Write-Error -Message 'You do not appear to have the required permission to the owner mailbox.  Please verify your permission.' -Category PermissionDenied
			}
		else
			{
			Write-Error -ErrorRecord $_
			}
		break
		}
	}
 
function Find-Mailbox ($identity)
	{
	try 
		{
		Get-Mailbox $identity -ErrorAction Stop
		}
	catch
		{
		Write-Progress -Activity 'Done' -Completed -Status " "
		Write-Error "A mailbox cannot be found that matches the input string $identity." -ErrorAction Stop -Category ObjectNotFound
		}
	}

function Get-SID($acl)
	{
	$aSID = @()
	#Use SID for on-prem accounts, SMTP for EXO
	if ($targetEnvironment -eq 'OnPremises')
		{
		$acl | % {
			$adUser = [System.Security.Principal.NTAccount]($_.User.ToString())
			$aSID += $adUser.Translate([System.Security.Principal.SecurityIdentifier]).Value
			}
		}
	else
		{
		$acl | % {
			$aSID += $_.User.ToString()
			}
		}
	$aSID
	}

function Get-FMA($identity)
	{
	Get-MailboxPermission $identity | ? {$_.IsInherited -eq $false -and $_.User -notlike 'S-1-5-21*'}
	}
	
function Get-SendAs($identity)
	{
	if ($targetEnvironment -eq 'OnPremises')
        {
        Get-ADPermission $identity | ? {$_.IsInherited -eq $false -and $_.ExtendedRights -like '*Send-As*'}
        }
    else
        {
        Get-RecipientPermission $identity | ? {$_.IsInherited -eq $false -and $_.AccessRights -like '*SendAs*'} `
            | select Identity, @{n='User';e={$_.Trustee}}, AccessControlType  
        }
	}

function Get-FolderPermission($mailbox,$folder)
	{
	Get-MailboxFolderPermission "$mailbox`:\$folder"
	}
	
function Set-FolderPermission($owner,$delegate,$folder,$role)
	{
	$folderPerm = Get-FolderPermission -mailbox $owner.Identity -folder $folder
	if ($targetEnvironment -eq 'OnPremises')
        {
        [array]$delegateFolderPerm = $folderPerm | ? {$_.User -eq $delegate.DisplayName}
        }
    else
        {
        [array]$delegateFolderPerm = $folderPerm | ? {$_.User.DisplayName -eq $delegate.DisplayName}
        }
	#Run cmdlet based on delegate presence in ACL
	if ($delegateFolderPerm.Count -eq 1)
		{
		try
			{
            Set-MailboxFolderPermission -Identity "$($owner.PrimarySMTPAddress.ToString()):\$folder" -User $delegate.PrimarySMTPAddress.ToString() -AccessRights $role -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
			}
		catch
			{
			$false
			}
		}
	else
		{
		try
			{
			Add-MailboxFolderPermission -Identity "$($owner.PrimarySMTPAddress.ToString()):\$folder" -User $delegate.PrimarySMTPAddress.ToString() -AccessRights $role -ErrorAction Stop | Out-Null
			}
		catch
			{
			$false
			}
		}
	}

function Remove-FolderPermission($owner,$delegate,$folder)
    {
	$folderPerm = Get-FolderPermission -mailbox $owner.Identity -folder $folder
	if ($targetEnvironment -eq 'OnPremises')
        {
        [array]$delegateFolderPerm = $folderPerm | ? {$_.User -eq $delegate.DisplayName}
        }
    else
        {
        [array]$delegateFolderPerm = $folderPerm | ? {$_.User.DisplayName -eq $delegate.DisplayName}
        }
	if ($delegateFolderPerm.Count -eq 1)
		{
		try
			{
			Remove-MailboxFolderPermission -Identity "$($owner.PrimarySMTPAddress.ToString()):\$folder" -User $delegate.PrimarySMTPAddress.ToString() -Confirm:$false -ErrorAction Stop | Out-Null
			}
		catch
			{
			$false
			}
		}
    }

function Convert-StringPermissionToEnum($role)
	{
	switch ($role)
		{
		'Reviewer' {[Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::Reviewer}
		'Author' {[Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::Author}
		'Editor' {[Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::Editor}
		'None' {[Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::None}
		}
	}

function Convert-StringDeliveryScopeToEnum($scope)
	{
	switch ($scope)
		{
		'DelegatesOnly' {[Microsoft.Exchange.WebServices.Data.MeetingRequestsDeliveryScope]::DelegatesOnly}
		'DelegatesAndOwner' {[Microsoft.Exchange.WebServices.Data.MeetingRequestsDeliveryScope]::DelegatesAndMe}
		'DelegatesAndInfoToOwner' {[Microsoft.Exchange.WebServices.Data.MeetingRequestsDeliveryScope]::DelegatesAndSendInformationToMe}
		'NoForward' {[Microsoft.Exchange.WebServices.Data.MeetingRequestsDeliveryScope]::NoForward}
		}
	}
	
function Convert-EnumDeliveryScopeToString($scope)
	{
	switch ($scope)
		{
		'DelegatesOnly' {'DelegatesOnly'}
		'DelegatesAndMe' {'DelegatesAndOwner'}
		'DelegatesAndSendInformationToMe' {'DelegatesAndInfoToOwner'}
		'NoForward' {'NoForward'}
		}
	}
	
function Get-FolderDisplayName($wellKnownFolderName,$emailAddress)
	{
	$folderID = New-Object -TypeName Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$wellKnownFolderName,$emailAddress)
	$folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService,$folderID)
	$folder.DisplayName
	}
#EndRegion

#Region Main Functions
function Get-DelegateManagementSetting
	{
    <#
	.Synopsis
		Get the configuration settings for the session.
	.Description
		Get the customizable settings for the delegate management module that have been set
		for the current session either from the settings file or by the module or user because
		there is no settings file.
	.Example
		Get-DelegateManagamentSetting
	.Notes
		Version: 1.0
		Date: 2/24/17
	#>
	
    $output = '' | select -Property ConnectionMode,UseAutoDiscover,EwsUrl,UseImpersonation,ApplyPermissionToSentItemsFolder,ApplyPermissionToDeletedItemsFolder
	$output.ConnectionMode = $targetEnvironment
	$output.UseAutodiscover = $useAutodiscover
	$output.EwsUrl = $ewsUri
	$output.UseImpersonation = $useImpersonation
	$output.ApplyPermissionToSentItemsFolder = $applyPermToSentItems
	$output.ApplyPermissionToDeletedItemsFolder = $applyPermToDeletedItems
	echo -InputObject $output
	}

function Set-DelegateManagementSetting
	{
	<#
	.Synopsis
		Configure the module in the current session or persistently.
	.Description
		Settings for the delegate management module are stored in
		DelegateManagementSettings.xml in the same directory as the module.  You
		can use this cmdlet to change settings that will apply to the current
		session only or also be written to the settings file.  It can also create
		a default file if one does not exist.
	.Parameter ConnectionMode
		Connect to on-premises Exchange or Exchange Online.  Valid values are EXO and OnPremises.
	.Parameter UseAutodiscover
	    Use a static EWS URL or autodiscover to determine the URL to connect to the mailbox.
		If set to False when used with Exchange on-premises, EwsUrl must be set.
	.Parameter EwsUrl
		URL to connect to Exchange Web Services.  Only used if not using autodiscover and the
		connection mode is for on-premises.
	.Parameter UseImpersonation
		Switch to connect to mailboxes using impersonation instead of full access.
	.Parameter ApplyPermissionToSentItemsFolder
		Automatically grant Author permission to the Sent Items	folder when a delegate is added.
	.Parameter ApplyPermissionToDeletedItemsFolder
		Automatically grant Author permission to the Deleted Items folder when a delegate is added.
	.Parameter Persist
		Save the specified setting(s) in the settings file.
	.Parameter CreateDefaultFile
		Create a settings file with default values.
	.Parameter Force
		When used with CreateDefaultFile, it overwrites the existing file.
	.Example
		Set-DelegateManagementSetting -ConnectionMode OnPremises -UseImpersonation $true
	.Example
		Set-DelegateManagementSetting -CreateDefaultFile
	.Notes
		Version: 1.0
		Date: 8/9/17
	#>
	[CmdletBinding(DefaultParameterSetName='save')]	
	param (
		[parameter(ParameterSetName='save')][ValidateSet('EXO','OnPremises')]$ConnectionMode,
	    [parameter(ParameterSetName='save')][bool]$UseAutodiscover,
		[parameter(ParameterSetName='save')][string]$EwsUrl,
		[parameter(ParameterSetName='save')][bool]$UseImpersonation,
		[parameter(ParameterSetName='save')][bool]$ApplyPermissionToSentItemsFolder,
		[parameter(ParameterSetName='save')][bool]$ApplyPermissionToDeletedItemsFolder,
		[parameter(ParameterSetName='save')][switch]$Persist,
		[parameter(ParameterSetName='create')][switch]$CreateDefaultFile,
		[parameter(ParameterSetName='create')][switch]$Force
		)
	if ($CreateDefaultFile)
		{
		#Check for existing file but no Force parameter
		if ((Test-Path -Path $settingsFile) -and (-not($Force)))
			{
			Write-Error "A settings file already exists.  To create a new file with the default settings, use the Force parameter." -Category InvalidArgument
			break
			}
		#Here string to save as a default settings file
		$defaultSettings = @"
<?xml version="1.0"?>
<Settings>
	<!-- Default settings.  All can be changed for the current session via Set-DelegateManagementSetting -->
	<!-- Specify EXO or OnPremises -->
	<ConnectionMode>EXO</ConnectionMode>
	<!-- Specify True or False to use autodiscover -->
	<UseAutodiscover>False</UseAutodiscover>
	<!-- If connection mode is on-premises and you don't want to use autodiscover, specify the EWS URL to use. -->
	<!-- The value is only used if the connection mode is OnPremises. -->
	<EwsUrl></EwsUrl>
	<!-- Specify True or False to use impersonation instead of full access -->
	<UseImpersonation>False</UseImpersonation>
	<!-- When adding a delegate, should Author permission be added to Sent Items and Deleted Items folders. -->
	<!-- Specific permission can also be explicitly added when adding a delegate. -->
	<ApplyPermissionToSentItemsFolder>True</ApplyPermissionToSentItemsFolder>
	<ApplyPermissionToDeletedItemsFolder>True</ApplyPermissionToDeletedItemsFolder>
</Settings>
"@
		$defaultSettings | Out-File -FilePath $settingsFile -Force
		break
		}
		
	if ($Persist)
		{
		#Check for no settings file
		if (-not(Test-Path -Path $settingsFile))
			{
			Write-Error "No settings file found in the module directory.  To create a default settings file, use the CreateDefaultFile parameter." -Category ObjectNotFound
			break
			}
		[xml]$file = cat -Path $settingsFile
		}
	#Set autodiscover choice if specified
	if ($PSBoundParameters.ContainsKey('UseAutodiscover'))
	    {
		$script:UseAutodiscover = $UseAutodiscover
		if ($Persist)
			{
			$file.Settings.UseAutodiscover = [string]$UseAutodiscover
			}
	    }
	#Set EWS URL if specified
	if ($PSBoundParameters.ContainsKey('EwsUrl'))
	    {
		$script:ewsUri = $EwsUrl
		if ($Persist)
			{
			$file.Settings.EwsUrl = $EwsUrl
			}
	    }
	#Set connection mode if specified
	if ($ConnectionMode)
	    {
	    switch ($ConnectionMode)
			{
			'OnPremises'
				{
				#Check if EWS URL is set when not using autodiscover
				if (-not($script:UseAutodiscover -and ([System.Uri]$ewsUri).AbsoluteUri))
					{
					Write-Error "An EWS URL needs to be specified when ConnectionMode is `'OnPremises`' and UseAutodiscover is `'False`'." -Category InvalidArgument
					break
					}
				$script:targetEnvironment = 'OnPremises'
				if ($Persist)
					{
					$file.Settings.ConnectionMode = 'OnPremises'
					}
				}
			'EXO'
				{
	            $script:targetEnvironment = 'EXO'
				if ($Persist)
					{
					$file.Settings.ConnectionMode = 'EXO'
					}
	            }
			}
	    }
	#Set impersonation choice if specified
	if ($PSBoundParameters.ContainsKey('UseImpersonation'))
		{
		$script:useImpersonation = $UseImpersonation
		if ($Persist)
			{
			$file.Settings.UseImpersonation = [string]$UseImpersonation
			}
		}
	#Set sent items choice if specified
	if ($PSBoundParameters.ContainsKey('ApplyPermissionToSentItemsFolder'))
		{
		$script:applyPermToSentItems = $ApplyPermissionToSentItemsFolder
		if ($Persist)
			{
			$file.Settings.ApplyPermissionToSentItemsFolder = [string]$ApplyPermissionToSentItemsFolder
			}
		}
	#Set deleted items choice if specified
	if ($PSBoundParameters.ContainsKey('ApplyPermissionToDeletedItemsFolder'))
		{
		$script:applyPermToDeletedItems = $ApplyPermissionToDeletedItemsFolder
		if ($Persist)
			{
			$file.Settings.ApplyPermissionToDeletedItemsFolder = [string]$ApplyPermissionToDeletedItemsFolder
			}
		}
	#Save settings included in command to file when specified
	if ($Persist)
		{
		$file.Save($settingsFile)
		}
	}
	
function Add-MailboxDelegate
	{
	<#
	.Synopsis
		Add a mailbox as a delegate of an owner's mailbox.
	.Description
		Add a mailbox delegate, optionally specifying permission to various folders,
		whether private items are viewable by the delegate, and if the delegate should receive
		meeting requests for the owner.
	.Parameter Owner
		Identity string of the user whose mailbox is to have the delegate.
	.Parameter Delegate
		Identity string of the user who is to be added to owner's mailbox.
	.Parameter InboxPermission
		Role to assign to the Inbox folder.  Valid roles are Reviewer, Author, and Editor.
	.Parameter CalendarPermission
		Role to assign to the Calendar folder.  Valid roles are Reviewer, Author, and Editor.
	.Parameter TasksPermission
		Role to assign to the Tasks folder.  Valid roles are Reviewer, Author, and Editor.
	.Parameter ContactsPermission
		Role to assign to the Contacts folder.  Valid roles are Reviewer, Author, and Editor.
	.Parameter SentItemsPermission
		Role to assign to the Sent Items folder.  Valid roles are Reviewer, Author, Editor, and None.
		The default permission is Author. 
	.Parameter DeletedItemsPermission
		Role to assign to the Deleted Items folder.  Valid roles are Reviewer, Author, Editor, and None.
		The default permission is Author.
	.Parameter ViewPrivateItems
		Enable the delegate to view items marked as private.
	.Parameter ReceiveMeetingRequests
		Enable the delegate to receive meeting requests for the owner.
	.Parameter MeetingRequestDeliveryScope
		Specify how meeting requests should be handled for the owner.  Valid scopes are DelegatesOnly,
		DelegatesAndOwner, DelegatesAndInfoToOwner, and NoForward.  Note that this parameter applies to all delegates.
	.Example
		Add-MailboxDelegate Username DelegateUsername -InboxPermission Editor -CalendarPermission Editor -ViewPrivateItems
	.Example
		Add-MailboxDelegate -Owner domain\username -Delegate <delegateemail> -CalendarPermission Editor -ReceiveMeetingRequests
	.Notes
		Version: 1.4
		Date: 9/7/17
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position=0,Mandatory=$true)][Alias("Manager")][string]$Owner,
		[Parameter(Position=1,Mandatory=$true)][string]$Delegate,
		[ValidateSet('Reviewer','Author','Editor')][string]$InboxPermission,
		[ValidateSet('Reviewer','Author','Editor')][string]$CalendarPermission,
		[ValidateSet('Reviewer','Author','Editor')][string]$TasksPermission,
		[ValidateSet('Reviewer','Author','Editor')][string]$ContactsPermission,
		[ValidateSet('Reviewer','Author','Editor','None')][string]$SentItemsPermission,
		[ValidateSet('Reviewer','Author','Editor','None')][string]$DeletedItemsPermission,
		[Alias("PI")][switch]$ViewPrivateItems,
		[Alias("MR")][switch]$ReceiveMeetingRequests,
		[ValidateSet('DelegatesOnly','DelegatesAndOwner','DelegatesAndInfoToOwner','NoForward')][string]$MeetingRequestDeliveryScope
		)

	#Validate Exchange cmdlets in session
    if (-not(Test-ExchangeConnectivity))
		{
		break
		}
	
	#Validate mailboxes
	Write-Progress -Activity "Adding Mailbox Delegate" -Status "Validating owner and delegate mailboxes" -PercentComplete 0
	$mbOwner = Find-Mailbox -identity $Owner
	$mbDelegate = Find-Mailbox -identity $Delegate
	
    $ownerFirstName = (Get-User -Identity $mbOwner.Identity).FirstName
	$ownerLastName = (Get-User -Identity $mbOwner.Identity).LastName
    $delegateFirstName = (Get-User -Identity $mbDelegate.Identity).FirstName
	$delegateLastName = (Get-User -Identity $mbDelegate.Identity).LastName
		
	#Get EWS mailbox reference
	Write-Progress -Activity "Adding Mailbox Delegate" -Status "Connecting to EWS" -PercentComplete 25
	$EWSMailbox = Connect-WebServices -smtpAddress $mbOwner.PrimarySMTPAddress.ToString()
	
	#Get collection of delegates, without folder permissions
	Write-Progress -Activity "Adding Mailbox Delegate" -Status "Retrieving existing delegates" -PercentComplete 50
	$currentDelegates = Get-Delegates -EWSMailbox $EWSMailbox
	
	#Check if user is already a delegate
	if ($currentDelegates.DelegateUserResponses.Count -gt 0)
		{
		foreach ($currentDelegate in $currentDelegates.DelegateUserResponses)
			{
			if ($currentDelegate.DelegateUser.UserId.PrimarySMTPAddress -eq $mbDelegate.PrimarySMTPAddress.ToString())
				{
				Write-Progress -Activity "Adding Mailbox Delegate" -Completed -Status " "
				Write-Host "$delegateFirstName $delegateLastName is already a delegate of $ownerFirstName $ownerLastName. Use Set-MailboxDelegate to update an existing delegate."
				return
				}
			}
		}

	Write-Progress -Activity "Adding Mailbox Delegate" -Status "Adding delegate" -PercentComplete 75
	#Create EWS delegate object
	$delegateUser = New-Object -TypeName Microsoft.Exchange.WebServices.Data.DelegateUser($mbDelegate.PrimarySMTPAddress.ToString())
	
	#Get localized folder name for Deleted Items and Sent Items
	$sentDisplayName = Get-FolderDisplayName -emailAddress $mbOwner.PrimarySMTPAddress.ToString() -wellKnownFolderName 'SentItems'
	$deletedDisplayName = Get-FolderDisplayName -emailAddress $mbOwner.PrimarySMTPAddress.ToString() -wellKnownFolderName 'DeletedItems'
	
	#Set private items
	if ($ViewPrivateItems)
		{
		$delegateUser.ViewPrivateItems = $true
		}
	
	#Set meeting request receipt
	if ($ReceiveMeetingRequests)
		{
		$delegateUser.ReceiveCopiesOfMeetingMessages = $true
		}
	
	#Set permissions on folders
	if ($InboxPermission)
		{
		$delegateUser.Permissions.InboxFolderPermissionLevel = Convert-StringPermissionToEnum -role $InboxPermission
		}
	if ($CalendarPermission)
		{
		$delegateUser.Permissions.CalendarFolderPermissionLevel = Convert-StringPermissionToEnum -role $CalendarPermission
		}
	if ($TasksPermission)
		{
		$delegateUser.Permissions.TasksFolderPermissionLevel = Convert-StringPermissionToEnum -role $TasksPermission
		}
	if ($ContactsPermission)
		{
		$delegateUser.Permissions.ContactsFolderPermissionLevel = Convert-StringPermissionToEnum -role $ContactsPermission
		}
	if ($SentItemsPermission)
		{
		$SIPermResponse = Set-FolderPermission -owner $mbOwner -delegate $mbDelegate -folder $sentDisplayName -role $SentItemsPermission
		}
	elseif ($applyPermToSentItems)
		{
		$SIPermResponse = Set-FolderPermission -owner $mbOwner -delegate $mbDelegate -folder $sentDisplayName -role 'Author'
		}
	if ($DeletedItemsPermission)
		{
		$DIPermResponse = Set-FolderPermission -owner $mbOwner -delegate $mbDelegate -folder $deletedDisplayName -role $DeletedItemsPermission
		}
	elseif ($applyPermToDeletedItems)
		{
		$DIPermResponse = Set-FolderPermission -owner $mbOwner -delegate $mbDelegate -folder $deletedDisplayName -role 'Author'
		}
		
	#Build delegate collection object to use in EWS method
	$delegateArray = New-Object -TypeName Microsoft.Exchange.WebServices.Data.DelegateUser[] 1 
	$delegateArray[0] = $delegateUser

	#Set new meeting request delivery scope if specified
	if ($MeetingRequestDeliveryScope)
		{
		$addResponse = $exchangeService.AddDelegates($EWSMailbox, (Convert-StringDeliveryScopeToEnum -scope $MeetingRequestDeliveryScope), $delegateArray)
		}
	else
		{
		$addResponse = $exchangeService.AddDelegates($EWSMailbox, $null, $delegateArray)
		}
		
	Write-Progress -Activity "Adding Mailbox Delegate" -Completed -Status " "
	if ($addResponse[0].Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success)
		{
		Write-Host "$delegateFirstName $delegateLastName has been added as a delegate of $ownerFirstName $ownerLastName." -ForegroundColor Green
		if ($SIPermResponse -eq $false)
			{
			Write-Host "An error occurred adding the delegate permission to the Sent Items folder." -ForegroundColor Yellow
			}
		if ($DIPermResponse -eq $false)
			{
			Write-Host "An error occurred adding the delegate permission to the Deleted Items folder." -ForegroundColor Yellow
			}
		}
	else
		{
		Write-Host "An error occurred adding $delegateFirstName $delegateLastName as a delegate of $ownerFirstName $ownerLastName." -ForegroundColor Red
		}
	}

function Set-MailboxDelegate
	{
	<#
	.Synopsis
		Update the settings for an existing delegate of an owner's mailbox.
	.Description
		Update an existing mailbox delegate, specifying any changes to folder permissions,
		whether private items are viewable by the delegate, or if the delegate should receive
		meeting requests for the owner.
	.Parameter Owner
		Identity string of the user whose mailbox has the delegate.
	.Parameter Delegate
		Identity string of the user whose delegate settings are to be updated.
	.Parameter InboxPermission
		Role to assign to the Inbox folder.  Valid roles are Reviewer, Author, Editor, and None.
	.Parameter CalendarPermission
		Role to assign to the Calendar folder.  Valid roles are Reviewer, Author, Editor, and None.
	.Parameter TasksPermission
		Role to assign to the Tasks folder.  Valid roles are Reviewer, Author, Editor, and None.
	.Parameter ContactsPermission
		Role to assign to the Contacts folder.  Valid roles are Reviewer, Author, Editor, and None.
	.Parameter SentItemsPermission
		Role to assign to the Sent Items folder.  Valid roles are Reviewer, Author, Editor, and None.
	.Parameter DeletedItemsPermission
		Role to assign to the Deleted Items folder.  Valid roles are Reviewer, Author, Editor, and None.
	.Parameter ViewPrivateItems
		Enable the delegate to view items marked as private.
	.Parameter ReceiveMeetingRequests
		Enable the delegate to receive meeting requests for the owner.
	.Parameter MeetingRequestDeliveryScope
		Specify how meeting requests should be handled for the owner.  Valid scopes are DelegatesOnly,
		DelegatesAndOwner, DelegatesAndInfoToOwner, and NoForward.  Note that this parameter applies to all delegates.
	.Example
		Set-MailboxDelegate Username DelegateUsername -InboxPermission Editor -CalendarPermission Editor -ViewPrivateItems
	.Example
		Set-MailboxDelegate -Owner domain\username -Delegate <delegateemail> -CalendarPermission Editor -ReceiveMeetingRequests
	.Notes
		Version: 1.3
		Date: 9/7/17
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position=0,Mandatory=$true)][Alias("Manager")][string]$Owner,
		[Parameter(Position=1,Mandatory=$true)][string]$Delegate,
		[ValidateSet('Reviewer','Author','Editor','None')][string]$InboxPermission,
		[ValidateSet('Reviewer','Author','Editor','None')][string]$CalendarPermission,
		[ValidateSet('Reviewer','Author','Editor','None')][string]$TasksPermission,
		[ValidateSet('Reviewer','Author','Editor','None')][string]$ContactsPermission,
		[ValidateSet('Reviewer','Author','Editor','None')][string]$SentItemsPermission,
		[ValidateSet('Reviewer','Author','Editor','None')][string]$DeletedItemsPermission,
		[Alias("PI")][switch]$ViewPrivateItems,
		[Alias("MR")][switch]$ReceiveMeetingRequests,
		[ValidateSet('DelegatesOnly','DelegatesAndOwner','DelegatesAndInfoToOwner','NoForward')][string]$MeetingRequestDeliveryScope
		)

	#Validate Exchange cmdlets in session
    if (-not(Test-ExchangeConnectivity))
		{
		break
		}
	
	#Validate mailboxes
	Write-Progress -Activity "Updating Mailbox Delegate" -Status "Validating owner and delegate mailboxes" -PercentComplete 0
	$mbOwner = Find-Mailbox -identity $Owner
	$mbDelegate = Find-Mailbox -identity $Delegate
	
    $ownerFirstName = (Get-User -Identity $mbOwner.Identity).FirstName
	$ownerLastName = (Get-User -Identity $mbOwner.Identity).LastName
    $delegateFirstName = (Get-User -Identity $mbDelegate.Identity).FirstName
	$delegateLastName = (Get-User -Identity $mbDelegate.Identity).LastName
	
	#Get EWS mailbox reference
	Write-Progress -Activity "Updating Mailbox Delegate" -Status "Connecting to EWS" -PercentComplete 25
	$EWSMailbox = Connect-WebServices -smtpAddress $mbOwner.PrimarySMTPAddress.ToString()
	
	#Get collection of delegates, with folder permissions
	Write-Progress -Activity "Updating Mailbox Delegate" -Status "Retrieving existing delegates" -PercentComplete 50
	$currentDelegates = Get-Delegates -EWSMailbox $EWSMailbox -includePermissions
	
	#Confirm user is already a delegate
	$delegateMatch = $false
	if ($currentDelegates.DelegateUserResponses.Count -eq 0)
		{
		Write-Progress -Activity "Updating Mailbox Delegate" -Completed -Status " "
		Write-Host "$ownerFirstName $ownerLastName does not have any delegates. `
				Use Add-MailboxDelegate to add a new delegate."
				return		
		}
	elseif ($currentDelegates.DelegateUserResponses.Count -gt 0)
		{
		foreach ($currentDelegate in $currentDelegates.DelegateUserResponses)
			{
			if ($currentDelegate.DelegateUser.UserId.PrimarySMTPAddress -eq $mbDelegate.PrimarySMTPAddress.ToString())
				{
				#Modify existing delegate object instead of creating new one so existing settings
				#can be preserved
				$delegateUser = $currentDelegate.DelegateUser
				$delegateMatch = $true
				}
			}
		if (-not($delegateMatch))
			{
			Write-Progress -Activity "Updating Mailbox Delegate" -Completed -Status " "
			Write-Host "$delegateFirstName $delegateLastName is not a delegate of $ownerFirstName $ownerLastName. Use Add-MailboxDelegate to add a new delegate."
			return
			}
		}
	
	Write-Progress -Activity "Updating Mailbox Delegate" -Status "Updating delegate" -PercentComplete 75
	
	#Get localized folder name for Deleted Items and Sent Items
	$sentDisplayName = Get-FolderDisplayName -emailAddress $mbOwner.PrimarySMTPAddress.ToString() -wellKnownFolderName 'SentItems'
	$deletedDisplayName = Get-FolderDisplayName -emailAddress $mbOwner.PrimarySMTPAddress.ToString() -wellKnownFolderName 'DeletedItems'	
	
	#Set private items if included
	if ($MyInvocation.BoundParameters.ContainsKey('ViewPrivateItems'))
		{
		$delegateUser.ViewPrivateItems = $ViewPrivateItems
		}
	
	#Set meeting request receipt if included
	if ($MyInvocation.BoundParameters.ContainsKey('ReceiveMeetingRequests'))
		{
		$delegateUser.ReceiveCopiesOfMeetingMessages = $ReceiveMeetingRequests
		}
	
	#Set permissions on folders
	if ($InboxPermission)
		{
		$delegateUser.Permissions.InboxFolderPermissionLevel = Convert-StringPermissionToEnum -role $InboxPermission
		}
	if ($CalendarPermission)
		{
		$delegateUser.Permissions.CalendarFolderPermissionLevel = Convert-StringPermissionToEnum -role $CalendarPermission
		}
	if ($TasksPermission)
		{
		$delegateUser.Permissions.TasksFolderPermissionLevel = Convert-StringPermissionToEnum -role $TasksPermission
		}
	if ($ContactsPermission)
		{
		$delegateUser.Permissions.ContactsFolderPermissionLevel = Convert-StringPermissionToEnum -role $ContactsPermission
		}
	if ($SentItemsPermission)
		{
		$SIPermResponse = Set-FolderPermission -owner $mbOwner -delegate $mbDelegate -folder $sentDisplayName -role $SentItemsPermission
		}
	if ($DeletedItemsPermission)
		{
		$DIPermResponse = Set-FolderPermission -owner $mbOwner -delegate $mbDelegate -folder $deletedDisplayName -role $DeletedItemsPermission
		}
		
	#Build delegate collection object to use in EWS method
	$delegateArray = New-Object -TypeName Microsoft.Exchange.WebServices.Data.DelegateUser[] 1 
	$delegateArray[0] = $delegateUser

	#Set new meeting request delivery scope if specified
	if ($MeetingRequestDeliveryScope)
		{
		$updateResponse = $exchangeService.UpdateDelegates($EWSMailbox, (Convert-StringDeliveryScopeToEnum -scope $MeetingRequestDeliveryScope), $delegateArray)
		}
	else
		{
		$updateResponse = $exchangeService.UpdateDelegates($EWSMailbox, $null, $delegateArray)
		}
		
	Write-Progress -Activity "Updating Mailbox Delegate" -Completed -Status " "
	if ($updateResponse[0].Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success)
		{
		Write-Host "$delegateFirstName $delegateLastName has been updated as a delegate of $ownerFirstName $ownerLastName." -ForegroundColor Green
		if ($SIPermResponse -eq $false)
			{
			Write-Host "An error occurred adding the delegate permission to the Sent Items folder." -ForegroundColor Yellow
			}
		if ($DIPermResponse -eq $false)
			{
			Write-Host "An error occurred adding the delegate permission to the Deleted Items folder." -ForegroundColor Yellow
			}
		}
	else
		{
		Write-Host "An error occurred updating $delegateFirstName $delegateLastName as a delegate of $ownerFirstName $ownerLastName." -ForegroundColor Red
		}
	}
	
function Remove-MailboxDelegate
	{
	<#
	.Synopsis
		Remove a delegate from an owner's mailbox.
	.Description
		Remove a supplied mailbox delegate from a supplied owner's mailbox.
	.Parameter Owner
		Identity string of the user whose mailbox has the delegate.
	.Parameter Delegate
		Identity string of the user who is to be removed from the owner's mailbox.
	.Example
		Remove-MailboxDelegate user@domain.com delegate@domain.com
	.Example
		Remove-MailboxDelegate -Owner domain\username -Delegate <delegatealias>
	.Notes
		Version: 1.3
		Date: 9/7/17
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position=0,Mandatory=$true)][Alias("Manager")][string]$Owner,
		[Parameter(Position=1,Mandatory=$true)][string]$Delegate
		)

	#Validate Exchange cmdlets in session and also AAD module exists if EXO
    if (-not(Test-ExchangeConnectivity))
		{
		break
		}
	
	#Validate mailboxes
	Write-Progress -Activity "Removing Mailbox Delegate" -Status "Validating owner and delegate mailboxes" -PercentComplete 0
	$mbOwner = Find-Mailbox -identity $Owner
	$mbDelegate = Find-Mailbox -identity $Delegate

    $ownerFirstName = (Get-User -Identity $mbOwner.Identity).FirstName
	$ownerLastName = (Get-User -Identity $mbOwner.Identity).LastName
    $delegateFirstName = (Get-User -Identity $mbDelegate.Identity).FirstName
	$delegateLastName = (Get-User -Identity $mbDelegate.Identity).LastName
	
	#Get EWS mailbox reference
	Write-Progress -Activity "Removing Mailbox Delegate" -Status "Connecting to EWS" -PercentComplete 25
	$EWSMailbox = Connect-WebServices -smtpAddress $mbOwner.PrimarySMTPAddress.ToString()
	
	#Get collection of delegates, without folder permissions
	Write-Progress -Activity "Removing Mailbox Delegate" -Status "Retrieving delegates" -PercentComplete 50
	$currentDelegates = Get-Delegates -EWSMailbox $EWSMailbox
	if ($currentDelegates.DelegateUserResponses.Count -eq 0)
		{
		Write-Progress -Activity "Removing Mailbox Delegate" -Completed -Status " "
		Write-Host $ownerFirstName $ownerLastName "does not have any delegates."
		}
	else
		{
		$delegateToRemove = @()
		$delegateMatch = $false
		Write-Progress -Activity "Removing Mailbox Delegate" -Status "Removing delegate" -PercentComplete 75
		foreach ($currentDelegate in $currentDelegates.DelegateUserResponses)
			{
			if ($currentDelegate.DelegateUser.UserId.PrimarySMTPAddress -eq $mbDelegate.PrimarySMTPAddress.ToString())
				{
				#Add userid object to collection of delegates to remove
				$delegateToRemove += $currentDelegate.DelegateUser.UserId
				$delegateMatch = $true
				}
			}	
		if (-not($delegateMatch))
			{
			Write-Progress -Activity "Removing Mailbox Delegate" -Completed -Status " "
			Write-Host $mbDelegate.PrimarySMTPAddress "is not a delegate of" $mbOwner.PrimarySMTPAddress "."
			}
		else
			{
			#Remove delegate from owner's mailbox
			$removeResponse = $exchangeService.RemoveDelegates($EWSMailbox, $delegateToRemove)
			Write-Progress -Activity "Removing Mailbox Delegate" -Completed -Status " "
			if ($removeResponse[0].Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success)
				{
				Write-Host "$delegateFirstName $delegateLastName has been removed as a delegate of $ownerFirstName $ownerLastName." -ForegroundColor Green
				}
			else
				{
				Write-Host "An error occurred removing $delegateFirstName $delegateLastName as a delegate of $ownerFirstName $ownerLastName." -ForegroundColor Red
				}
			#Get localized folder name for Deleted Items and Sent Items
			$sentDisplayName = Get-FolderDisplayName -emailAddress $mbOwner.PrimarySMTPAddress.ToString() -wellKnownFolderName 'SentItems'
			$deletedDisplayName = Get-FolderDisplayName -emailAddress $mbOwner.PrimarySMTPAddress.ToString() -wellKnownFolderName 'DeletedItems'            
			#Remove SI and DI folder permission
			$SIPermResponse = Remove-FolderPermission -owner $mbOwner -delegate $mbDelegate -folder $sentDisplayName
            $DIPermResponse = Remove-FolderPermission -owner $mbOwner -delegate $mbDelegate -folder $deletedDisplayName
		    if ($SIPermResponse -eq $false)
			    {
			    Write-Host "An error occurred removing the delegate permission to the Sent Items folder." -ForegroundColor Yellow
			    }
		    if ($DIPermResponse -eq $false)
			    {
			    Write-Host "An error occurred removing the delegate permission to the Deleted Items folder." -ForegroundColor Yellow
			    }
			}
		}
	}

function Get-MailboxDelegate
	{
	<#
	.Synopsis
		Display a mailbox's delegates and permissions.
	.Description
		Retrieve the list of delegates for a mailbox and display the mailbox permission,
		folder permissions, meeting invite settings, and (optionally) whether the
		delegate has Send As permission.
	.Parameter Identity
		Identity string of the user whose mailbox has the delegates.  Owner and Manager are valid
		aliases for this parameter.
	.Parameter Delegate
		Identity string of the delegate you want to retrieve.  If omitted, all delegates
		are retrieved.
	.Parameter IncludeSendAs
		Switch to indicate that you want Send As permission to be included.
	.Example
		Get-MailboxDelegate user@domain.com -includesendas
	.Example
		Get-MailboxDelegate domain\username
	.Notes
		Version: 1.9
		Date: 9/7/17
	#>
	param (
		[Parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
		[ValidatePattern('(?# Wildcards not allowed)^[^\*]+$')][Alias("Owner")][Alias("Manager")][string]$Identity,
		[Parameter(Position=1)][string]$Delegate,
		[Alias("SA")][switch]$IncludeSendAs #Perform Send As lookup (takes longer)	
		)
	
	process
		{
		#Validate Exchange cmdlets in session and also AAD module exists if EXO
	    if (-not(Test-ExchangeConnectivity))
			{
			break
			}

        #Validate owner mailbox
		Write-Progress -Activity "Getting Mailbox Delegate" -Status "Validating owner mailbox" -PercentComplete 0
		$mbOwner = Find-Mailbox -identity $Identity
		
		#Validate delegate mailbox
		if ($Delegate)
			{
			Write-Progress -Activity "Getting Mailbox Delegate" -Status "Validating delegate mailbox" -PercentComplete 5
			$mbDelegate = Find-Mailbox -identity $Delegate
			$delegateFirstName = (Get-User -Identity $mbDelegate.Identity).FirstName
			$delegateLastName = (Get-User -Identity $mbDelegate.Identity).LastName
            }

		$ownerFirstName = (Get-User -Identity $mbOwner.Identity).FirstName
		$ownerLastName = (Get-User -Identity $mbOwner.Identity).LastName

		#Get EWS mailbox reference
		Write-Progress -Activity "Getting Mailbox Delegate" -Status "Connecting to EWS" -PercentComplete 10
		$EWSMailbox = Connect-WebServices -smtpAddress $mbOwner.PrimarySMTPAddress.ToString()

		#Get list of delegates and permissions from EWS
		Write-Progress -Activity "Getting Mailbox Delegate" -Status "Retrieving delegates" -PercentComplete 20
		#Get collection of delegates, with folder permissions
		if ($mbDelegate)
			{
			#Retrieve only the specified delegate
			$delegateUser = New-Object -TypeName Microsoft.Exchange.WebServices.Data.UserId($mbDelegate.PrimarySMTPAddress.ToString())
			$delegateArray = New-Object -TypeName Microsoft.Exchange.WebServices.Data.UserId[] 1 
			$delegateArray[0] = $delegateUser
			$currentDelegates = Get-Delegates -EWSMailbox $EWSMailbox -delegateToRetrieve $delegateArray -includePermissions
			}
		else
			{
			$currentDelegates = Get-Delegates -EWSMailbox $EWSMailbox -includePermissions
			}

		#Get list of users with full mailbox access
		Write-Progress -Activity "Getting Mailbox Delegate" -Status "Retrieving Full Access list" -PercentComplete 40
		$fullMailboxAccess = Get-FMA -identity $mbOwner.Identity
		$fmaSID = Get-SID -acl $fullMailboxAccess #Convert username to SID
		
		#Get list of users with Send As permission from AD
		if ($IncludeSendAs)
			{
			Write-Progress -Activity "Getting Mailbox Delegate" -Status "Retrieving Send As List" -CurrentOperation "(This part takes the longest.)" -PercentComplete 70
			$sendAs = Get-SendAs -identity $mbOwner.Identity
			$saSID = Get-SID -acl $sendAs #Convert username to SID
			}
		
		#Get permissions for additional folders
		Write-Progress -Activity "Getting Mailbox Delegate" -Status "Retrieving additional folder permissions" -PercentComplete 90
		#Get localized folder name for Deleted Items and Sent Items
		$sentDisplayName = Get-FolderDisplayName -emailAddress $mbOwner.PrimarySMTPAddress.ToString() -wellKnownFolderName 'SentItems'
		$deletedDisplayName = Get-FolderDisplayName -emailAddress $mbOwner.PrimarySMTPAddress.ToString() -wellKnownFolderName 'DeletedItems'
	
		$deletedItemsPerm = Get-FolderPermission -mailbox $mbOwner.Identity -folder $deletedDisplayName
		$sentItemsPerm = Get-FolderPermission -mailbox $mbOwner.Identity -folder $sentDisplayName
		Write-Progress -Activity "Getting Mailbox Delegate" -Completed -Status " "

		#Build output
		$outputArr = New-Object -TypeName System.Collections.ArrayList
		#Loop through list of delegates
		if ($currentDelegates.DelegateUserResponses.Count -gt 0)
			{
			if ($mbDelegate -and $currentDelegates.DelegateUserResponses[0].ErrorCode -eq 'ErrorNotDelegate')
				{
				Write-Host "$delegateFirstName $delegateLastName is not a delegate of $ownerFirstName $ownerLastName."
				}
			else
				{
				$currentDelegates.DelegateUserResponses | % {
					#Create custom object with property names
					$outputItem = "" | select -Property Owner,Delegate,MeetingHandling,FullAccess,SendAs,Calendar,Inbox,Contacts,Tasks,DeletedItems,SentItems,ReceiveMeetings,ViewPrivate,Error,ErrorNote
					$outputItem.Owner = $mbOwner.DisplayName
					$outputItem.MeetingHandling = Convert-EnumDeliveryScopeToString -scope $currentDelegates.MeetingRequestsDeliveryScope
					if ($_.ErrorMessage -eq 'The delegate does not map to a user in the Active Directory.')
						{#Delegate account deleted in AD but still listed in list
						$outputItem.Error = "Orphan"
						$outputItem.ErrorNote = "Check NON_IPM_SUBTREE\Freebusy Data\LocalFreebusy.eml property 0x684A101E to determine orphan entry."
						}
					elseif ($_.ErrorMessage -eq 'Delegate is not configured properly.')
						{
						$outputItem.Error = "Misconfigured"
						$outputItem.ErrorNote = "Missing from Freebusy Data folder or publicDelegates attribute."
						}
					elseif ($_.Result -eq 'Error')
						{
						$outputItem.Error = "UnknownError"
						$outputItem.ErrorNote = $_.ErrorMessage
						}
					else
						{
						$delegateDisplayName = $_.delegateuser.userid.displayname
						$outputItem.Delegate = $delegateDisplayName
						#Use SID comparison for on-prem accounts, SMTP for EXO
						if ($targetEnvironment -eq 'OnPremises')
							{
							$delegateSid = $_.delegateuser.UserId.SID
							}
						else
							{
							$delegateSid = $_.delegateuser.UserId.PrimarySmtpAddress
							}
						if ($fmaSID -match $delegateSid)
							{
							$outputItem.FullAccess = $true
							}
						else
							{
							$outputItem.FullAccess = $false
							}
						if ($includeSendAs)
							{
							if ($saSID -match $delegateSid)
								{
								$outputItem.SendAs = $true
								}
							else
								{
								$outputItem.SendAs = $false
								}
							}
						$outputItem.Calendar = $_.DelegateUser.Permissions.CalendarFolderPermissionLevel.ToString()
						$outputItem.Inbox = $_.DelegateUser.Permissions.InboxFolderPermissionLevel.ToString()
						$outputItem.Contacts = $_.DelegateUser.Permissions.ContactsFolderPermissionLevel.ToString()
						$outputItem.Tasks = $_.DelegateUser.Permissions.TasksFolderPermissionLevel.ToString()
						
						#Construct Deleted Items permission output
						if ($targetEnvironment -eq 'OnPremises')
                            {
                            [array]$delegateDIPerm = $deletedItemsPerm | ? {$_.User -eq $delegateDisplayName}
                            }
                        else
                            {
                            [array]$delegateDIPerm = $deletedItemsPerm | ? {$_.User.DisplayName -eq $delegateDisplayName}
                            }
						if ($delegateDIPerm.Count -eq 1)
							{
							$delegateDIPermValue = $delegateDIPerm[0].AccessRights[0].ToString()
							}
						else
							{
							$delegateDIPermValue = 'None'
							}
						$outputItem.DeletedItems = $delegateDIPermValue
						
						#Construct Sent Items permission output
						if ($targetEnvironment -eq 'OnPremises')
                            {
                            [array]$delegateSIPerm = $sentItemsPerm | ? {$_.User -eq $delegateDisplayName}
                            }
                        else
                            {
                            [array]$delegateSIPerm = $sentItemsPerm | ? {$_.User.DisplayName -eq $delegateDisplayName}
                            }
						if ($delegateSIPerm.Count -eq 1)
							{
							$delegateSIPermValue = $delegateSIPerm[0].AccessRights[0].ToString()
							}
						else
							{
							$delegateSIPermValue = 'None'
							}
						$outputItem.SentItems = $delegateSIPermValue
						
						$outputItem.ReceiveMeetings = $_.delegateuser.receivecopiesofmeetingmessages
						$outputItem.ViewPrivate = $_.delegateuser.viewprivateitems
						
						$outputArr.Add($outputItem) | Out-Null
						}
					}
				}
			}
		else
			{
			Write-Host "$ownerFirstName $ownerLastName has no delegates."
			}
		$outputArr
		}
	}
#EndRegion

sal amd Add-MailboxDelegate
sal smd Set-MailboxDelegate
sal rmd Remove-MailboxDelegate
sal gmd Get-MailboxDelegate
Export-ModuleMember -Function *-DelegateManagementSetting, *-MailboxDelegate
Export-ModuleMember -Alias *
