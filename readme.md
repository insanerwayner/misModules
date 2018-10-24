# misModules

These PowerShell modules were created to help MIS with repetitive tasks that we do everyday. Each module contains a group of Powershell Scripts that are used for common tasks.

## Installation

**Note: Requires Windows 10**

Run the **Setup-misModules.ps1** Script with **Elevated Administrator Privileges**. This will copy the modules to your computer then walk you through storing your _encrypted_ passwords for the scripts.  


**Note:**

When entering credentials for exchange don't forget the **domain\username**

### Setup misExchange to Load on PowerShell Startup

In order to get your **Remote PowerShell Exchange Session** activated you have to run the command  
`Import-Module misExchange -DisableNameChecking`.  

You can add this to your PowerShell Profile (after you have run **Setup-misModules.ps1**) by:

- Open your Powershell Profile: `notepad $profile`
- Add the line: `Import-Module misExchange -DisableNameChecking`
- Save the File
- Close and Reopen PowerShell

## Commands

To get a list of commands for each module you can use:
`Get-Command -Module <ModuleName>`  

**Example:**

`Get-Command -Module misAD`

Returns:  

```
CommandType     Name                                               Version    Source
-----------     ----                                               -------    ------
Function        Find-ADComputer                                    0.0        misad
Function        Find-ADuser                                        0.0        misad
Function        Get-LastBootTime                                   0.0        misad
Function        Get-PasswordExpirationList                         0.0        misad
Function        Reset-Password                                     0.0        misad
Function        Select-User                                        0.0        misad
Function        Unlock-ADUser                                      0.0        misad
```

## Help

### Get-Help

To get help with a command you can use the `Get-Help` command.

**Example:**

`Get-Help Unlock-ADUser`

Returns:

```
NAME
    Unlock-ADUser

SYNOPSIS
    Simplifies unlocking users


SYNTAX
    Unlock-ADUser [-Filter] <String> [<CommonParameters>]


DESCRIPTION
    This script will give you a menu to select the user from a list generated from your string filter then it will attempt to unlock the user


RELATED LINKS

REMARKS
    To see the examples, type: "get-help Unlock-ADUser -examples".
    For more information, type: "get-help Unlock-ADUser -detailed".
    For technical information, type: "get-help Unlock-ADUser -full".
```

### Getting Examples

You can also get a list of examples for the command by using `Get-Help <CommandName> -Examples`

**Example:**

`Get-Help Unlock-ADUser -Examples`

Returns:

```
NAME
    Unlock-ADUser

SYNOPSIS
    Simplifies unlocking users


    -------------------------- EXAMPLE 1 --------------------------

    PS C:\>Unlock-ADUser test

    Description:
    Will check each Server if User is locked out and go and Unlock them




    -------------------------- EXAMPLE 2 --------------------------

    PS C:\>Unlock-ADUser test -Server DC01

    Description:
    Will check DC01 to see if User is locked out. If it is locked out it will unlock the user. 
    If it is not locked it will tell you and ask if you would like to check all DCs.
```

### Detailed Help

You can also use `Get-Help <CommandName> -Detailed` to get the main help with the examples all at once.


## Modules

### misAD

Has some extended commands to manage Active Directory Users and Computers

### misUserSession

Manages Users' Windows and Citrix Sessions

### misEncryption

Used to retrieve and add computer encryption passwords quickly.

### misExchange

Connects a **Remote PowerShell Session** to Exchange so you can run Exchange commands from your computer.

### misSecurity

Used to store and retrieve passwords for other commands.

### misServer

Used to manage Windows Servers' various services easily.

### misScripting

Cmdlets to make some advanced scripting easier. For use in some of the other Modules.
