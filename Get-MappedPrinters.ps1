<#
.SYNOPSIS
Checks for printers current user is mapped to on the network (must be logged on).

.PARAMETER ComputerName
Specifies the computer user is on.

.INPUTS
None. You cannot pipe objects to Get-MappedPrinters.ps1.

.OUTPUTS
System.Object

.EXAMPLE
.\Get-MappedPrinters -ComputerName TESTPCNAME

.NOTES
Author: Edward Champa
Date Modified: 24 August 2020
#>

[Cmdletbinding()]
Param(

    [Parameter(Mandatory)]
    [string]
    $ComputerName
)

$ID = Get-WmiObject -Class win32_computersystem -ComputerName $ComputerName |
Select-Object -ExpandProperty Username |
ForEach-Object { ([System.Security.Principal.NTAccount]$_).Translate([System.Security.Principal.SecurityIdentifier]).Value }

$Path = "Registry::\HKEY_USERS\$ID\Printers\Connections\"

$Result = Invoke-Command -Computername $ComputerName -ScriptBlock {
    param($Path)
    
    Get-Childitem $Path | Select PSChildName
} -ArgumentList $Path | Select -Property *

Write-Host "`nPrint Server/Printer:"

$Result = $Result.PSChildName -Split(",")

$Result.Trim()
