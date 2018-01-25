if ( $ExchangeSession -eq $null -and ( test-connection misexch01 -count 1 -ErrorAction SilentlyContinue ) )
	{
	$cred = Get-XMLPassword -Name Exchange -Type Credential 
	$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://misexch01/PowerShell/ -Authentication Kerberos -Credential $Cred
	Import-PSSession $ExchangeSession -DisableNameChecking | Out-Null
	}
