#Checks for Local Security Repository ( Where the Passwords are Stored )
$ModulePath = "$env:USERPROFILE\Documents\WindowsPowershell\Modules\misSecurity\"
$SecurityPath =  Get-Content ( Join-Path $ModulePath "securitypath.txt" )
if ( $SecurityPath -eq 'default' )
        {
        $Security = ( Join-Path $ModulePath "Security" )
        }
if ( !(Test-Path $Security ) )
        {
        Write-Host "Security Path Doesn't Exist. Create?" -ForegroundColor White
        $answer = Read-Host "[Y,n]"
        if ( $answer -eq 'y' -or $answer -eq '' )
                {
                $answer = $null
                Write-Host "Use Default path? [ $Security ]"
                $answer = Read-Host '[Y,n]'
                if ( $answer -eq 'y' -or $answer -eq '' )
                        {
                        New-Item -ItemType Directory -Path $Security
                        }
                else
                        {
                        $Security = Read-Host "Folder Path:"
                        New-Item -ItemType Directory -Path $Security
                        $Security | Out-File $SecurityPath
                        }            
                }
        else
                {
                Write-Host "Module will not be loaded." -ForegroundColor Yellow
                }
        }

Function ConvertFrom-SecureStringToPlaintext($SecureString)
    {
    <#
    .Synopsis
    Converts a SecureString to PlainText

    .NOTES   
    Name: ConvertFrom-SecureStringToPlaintext
    Author: Wayne Reeves
    Version: 11.28.17
    #>
    [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString))
    }

Function New-XMLPassword
{
<#
.Synopsis
This stores a Password or Credential in the "Security Repository." This will mostly be used by other Commandlettes but can be used to store your own Passwords or secure information. 

.DESCRIPTION
This will mostly be used by other Commandlettes but can be used to store your own Passwords or secure information. The XML files generated are only accessible by YOUR User on THIS Machine

.NOTES   
Name: New-XMLPassword
Author: Wayne Reeves
Version: 11.28.17

.PARAMETER Type
Type of credential being stored. Is it a Windows Credential or a just a Password?
[Credental, Password]

.EXAMPLE
New-XMLPassword -Type Password -Name Example

Description:
Will ask you the password to store and store the XML File with the name specified

.EXAMPLE
New-XMLPassword -Type Credential -Name Example

Description:
Will ask you for Windows Credential to store and store the XML File with the name specified
#>
Param(
[Parameter(Mandatory=$True)]
[ValidateSet('Credential', 'Password')]
[string]$Type,
[Parameter(Mandatory=$True)]
[String]$Name,
[String]$Path = ( Join-Path "$Security"  ( $name + ".xml" ) )
)
switch ($type) {
    Credential { $password = Get-Credential }
    Password { $password = Read-Host "Password" -AsSecureString }
}
try 
    {
    $Password | Export-Clixml $Path 
    }
catch 
    {
    Write-Host $error -ForegroundColor Red
    }
finally
    {
    Write-Host "Successfully created XML Password File:" $Path -ForegroundColor Yellow
    }

}

Function Get-XMLPassword
    {
    <#
    .Synopsis
    Retrieves Password or Credential already stored

    .DESCRIPTION
    Retrieves Password or Credential already stored. If it doesn't exist it will ask if you would like to create it.

    .NOTES   
    Name: Get-XMLPassword
    Author: Wayne Reeves
    Version: 11.28.17

    .PARAMETER Type
    Type of credential being stored. Is it a Windows Credential or a just a Password? This is only required if the Password is not already stored. Typically only used by other commandlettes.
    [Credental, Password]

    .EXAMPLE
    Get-XMLPassword -Name Example

    Description:
    Will retrieve the stored password or Windows Credential from the Local Security Repository. If it doesn't exist, will ask you the Type and the Password or Windows Credential and store it for you.

    .EXAMPLE
    Get-XMLPassword -Type Password -Name Example

    Description:
    Will retrieve the stored password from the Local Security Repository. If it doesn't exist, will ask for the password and store it for you.

    .EXAMPLE
    Get-XMLPassword -Type Credential -Name Example

    Description:
    Will retrieve the stored Credential from the Local Security Repository. If it doesn't exist, will ask for the Windows Credential and store it for you.
    #>
    
    param (
        [parameter(Mandatory=$True)]
        [String]$Name,
        [String]$Type,
        [bool]$AsPlainText = $false
    )
    $Path = "$($Security)/$($Name).xml"
    If ( Test-Path $Path )
        {
        if ( $AsPlainText )
            {
            Return ( ConvertFrom-SecureStringToPlaintext ( Import-Clixml -Path $Path ) )
            }
        else 
            {
            Return ( Import-Clixml -Path $Path )
            }
        
        }
    else
        {
        Write-Host "No Password File for $($Name.ToUpper()) found. Would you like to create it?" -ForegroundColor Yellow   
        $Answer = Read-Host '[Y,n]'
        if ( ( $answer -eq 'y' ) -or ( $answer -eq '' ) )
            {
            if ( !$Type )
                {
                Write-Host "Type not specified." -ForegroundColor Yellow
                Write-Host "1. Credential" -ForegroundColor Cyan
                Write-HOst "2. Password" -ForegroundColor Cyan
                $Choice = Read-Host "Selection"
                switch ( $Choice ) 
                    {
                    1 { $Type = "Credential" }
                    2 { $Type = "Password" }
                    }
                }
            New-XMLPassword -Name $Name -Type $Type
            Get-XMLPassword -Name $Name -Type $Type
            }
        }
    }

Function New-RandomPassword
    {
    <#
    .Synopsis
    Generates a Random String to Be Used as a Password

    .DESCRIPTION
    A script to generate a random string of characters to be used for a password.

    .PARAMETER CallWords
    Switch to display call words to read each character to user.

    .EXAMPLE
    New-RandomPassword

    Description:
    Generates a randomy password string.

    .EXAMPLE
    New-RandomPassword -CallWords

    Description:
    Generates a random password string and also displays call words for reading the characters to the user.

    .NOTES
    Name: New-RandomPassword
    Author: Wayne Reeves
    Version: 4.3.19
    #>
    param(
        [switch]
        $CallWords
        )
    Function Get-RandomCharacters($length, $characters)
        { 
        $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length } 
        $private:ofs="" 
        return [String]$characters[$random]
        }

    Function Scramble-String([string]$inputString)
        {
        $characterArray = $inputString.ToCharArray()   
        $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
        $outputString = -join $scrambledStringArray
        return $outputString 
        }

    # Get Random Strings
    $password = Get-RandomCharacters -length 5 -characters 'abcdefghikmnprstuvwxyz'
    $password += Get-RandomCharacters -length 1 -characters 'ABCDEFGHKLMNPRSTUVWXYZ'
    $password += Get-RandomCharacters -length 1 -characters '2345678'
    $password += Get-RandomCharacters -length 1 -characters '!@#$%&+?'
    $password = Scramble-String -inputstring $password
    If ( $CallWords )
        {
        Write-Host $Password
        Write-Host "$(Get-CallWords $password)"
        }
    Else
        {
        return $password
        }
    }

Function Get-Callwords
    {
    <#
    .Synopsis
    Prints out call words for each character of a given string

    .PARAMETER String
    The string you provide to be spelled out with call words.

    .NOTES
    Name: Get-CallWords
    Author: Wayne Reeves
    Version: 4.23.19
    #>
    param(
        [string]
        $String
        )
    $CallWords = [char[]]$String | foreach {
        switch -CaseSensitive ( $_ )
            {
            "a" { "alpha" }
            "b" { "beta" }
            "c" { "charlie" }
            "d" { "delta" }
            "e" { "echo" }
            "f" { "foxtrot" }
            "g" { "golf" }
            "h" { "hotel" }
            "i" { "india" }
            "j" { "juliett" }
            "k" { "kilo" }
            "l" { "lima" }
            "m" { "mike" }
            "n" { "november" }
            "o" { "oscar" }
            "p" { "papa" }
            "q" { "quebec" }
            "r" { "romeo" }
            "s" { "sierra" }
            "t" { "tango" }
            "u" { "uniform" }
            "v" { "victor" }
            "w" { "whiskey" }
            "x" { "x-ray" }
            "y" { "yankee" }
            "z" { "zulu" }     
            "A" { "ALPHA" }
            "B" { "BETA" }
            "C" { "CHARLIE" }
            "D" { "DELTA" }
            "E" { "ECHO" }
            "F" { "FOXTROT" }
            "G" { "GOLF" }
            "H" { "HOTEL" }
            "I" { "INDIA" }
            "J" { "JULIETT" }
            "K" { "KILO" }
            "L" { "LIMA" }
            "M" { "MIKE" }
            "N" { "NOVEMBER" }
            "O" { "OSCAR" }
            "P" { "PAPA" }
            "Q" { "QUEBEC" }
            "R" { "ROMEO" }
            "S" { "SIERRA" }
            "T" { "TANGO" }
            "U" { "UNIFORM" }
            "V" { "VICTOR" }
            "W" { "WHISKEY" }
            "X" { "X-RAY" }
            "Y" { "YANKEE" }
            "Z" { "ZULU" }    
            "0" { "zero" }
            "1" { "one" }
            "2" { "two" }
            "3" { "three" }
            "4" { "four" }
            "5" { "five" }
            "6" { "six" }
            "7" { "seven" }
            "8" { "eight" }
            "9" { "nine" }
            "!" { "exclamation" }
            "@" { "at-sign" }
            '#' { "number-sign" }
            "$" { "dollar-sign" }
            "%" { "percent" }
            "&" { "ampersand" }
            "+" { "plus-sign" }
            "?" { "question-mark" }
            }
        }
    return $CallWords
    }
