Function Get-MappedPrinters {

<#
.SYNOPSIS
Checks for printers current user is mapped to on the network.

.DESCRIPTION
Checks for printers current user is mapped to on the network. User must be logged on.

.PARAMETER ComputerName
Mandatory. Specifies the computer user is logged on to.

.EXAMPLE
.\Get-MappedPrinters -ComputerName TESTPCNAME

.NOTES
Author: Edward Champa
Date Modified: 25 August 2020
#>

    [Cmdletbinding()]
    
    Param(
    
        [Parameter(Mandatory)]
        [string]
        $ComputerName
    )
    
    Begin {
    }

    Process {

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
    }

    End {
    }
}
