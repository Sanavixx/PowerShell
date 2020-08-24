<#
.SYNOPSIS
Queries Registry for specified software and logs results.

.PARAMETER PCListPath
Specifies the path to list of PCs.

.PARAMETER LogFolder
Specifies the path to folder to store logs.

.PARAMETER Software
Specifies the software to search for. Use DisplayName from Registry.

.PARAMETER Version
Specifies the software version. Use DisplayVersion from Registry.

.INPUTS
None. You cannot pipe objects to Software-Checker.ps1.

.OUTPUTS
Csv logs based on results.

.EXAMPLE
.\Software-Query -PCListPath C:\temp\PCs.txt -LogFolder C:\temp\Log -Software Mozilla -Version 79.0

.NOTES
Author: Edward Champa
Date Modified: 23 August 2020
#>


[CmdletBinding()]
param (

    [Parameter(Mandatory)]
    [string]
    $PCListPath,

    [Parameter(Mandatory)]
    [string]
    $LogFolder,

    [Parameter(Mandatory)]
    [string]
    $Software,

    [Parameter(Mandatory)]
    [string]
    $Version
)

$List = Get-Content -Path $PCListPath

#Clears Log Folder
If (Test-Path $LogFolder) {
    Remove-Item $LogFolder -Recurse -Force
}

foreach ($Computer in $List) {

    #Tests Connection To Computer       
    if (!($Computer -like "*.*.*.*")) {

        if (Test-Connection $Computer -quiet -count 1) {

            Write-Host "Connection to $Computer is successful. Checking software..." -ForegroundColor Green

            #Searches Registry For Software/Version
            $Result = Invoke-Command -ComputerName $Computer -ScriptBlock {

                param($Computer,$Software,$Version)

                if(Get-ChildItem -Path HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall,
                    HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall |
                    Get-ItemProperty | Where-Object {$_.DisplayVersion -like "$Version" -and $_.DisplayName -like "$Software*"}) {

                    Write-Host "    $Computer has $Software. Adding to Logs..." -ForegroundColor Green

                    $Return = "Yes"
                }

                else {

                    Write-Host "    $Computer does not have $Software. Adding to Logs..." -ForegroundColor Red

                    $Return = "No"
                }

                $Return
            } -ArgumentList ($Computer,$Software,$Version)
        }

        else {

            #No Connection To Computer
            Write-Host "Could not connect to $Computer. Adding to Logs..." -ForegroundColor Red

            if (!(Test-Path $LogFolder)){New-Item -ItemType Directory -Force -Path $LogFolder}

            New-Object -TypeName PSCustomObject -Property @{
            ComputerName = $Computer
            } | Export-Csv -NoTypeInformation -Path "$LogFolder\No Connection PCs.csv" -Append -Force

            Sleep 2
        }
    }

    #Creates Logs Based On Results
    if ($Result -contains "Yes") {

        if (!(Test-Path $LogFolder)){New-Item -ItemType Directory -Force -Path $LogFolder}

        New-Object -TypeName PSCustomObject -Property @{
        ComputerName = $Computer
        } | Export-Csv -NoTypeInformation -Path "$LogFolder\PCs With $Software.csv" -Append -Force

        Sleep 2
    }

    else {

        if (!(Test-Path $LogFolder)){New-Item -ItemType Directory -Force -Path $LogFolder}

        New-Object -TypeName PSCustomObject -Property @{
        ComputerName = $Computer
        } | Export-Csv -NoTypeInformation -Path "$LogFolder\PCs Without $Software.csv" -Append -Force

        Sleep 2
    }
}
