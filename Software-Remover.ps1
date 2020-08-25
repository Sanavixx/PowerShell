Function Software-Remover {

<#
.SYNOPSIS
Queries Registry for specified software, then uninstalls if found.

.DESCRIPTION
Queries Registry for specified software, then uninstalls if found. Also logs results.

.PARAMETER PCListPath
Mandatory. Specifies the path to list of PCs.

.PARAMETER LogFolder
Mandatory. Specifies the path to folder to store logs.

.PARAMETER Software
Mandatory. Specifies the software to search for. Use DisplayName from Registry.

.PARAMETER Version
Mandatory. Specifies the software version. Use DisplayVersion from Registry.

.EXAMPLE
.\Software-Remover -PCListPath C:\temp\PCs.txt -LogFolder C:\temp\Log -Software Mozilla -Version 79.0

.NOTES
Author: Edward Champa
Date Modified: 25 August 2020
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
    

    Begin {

        $List = Get-Content -Path $PCListPath
        
        #Clears Log Folder
        if (Test-Path $LogFolder) {
        
            Remove-Item $LogFolder -Recurse -Force
        }
    }

    Process {
    
        foreach ($Computer in $List) {
        
            #Tests Connection To Computer       
            if (!($Computer -like "*.*.*.*")) {
        
                if (Test-Connection $Computer -quiet -count 1) {
        
                    Write-Host "Connection to $Computer is successful. Checking software..." -ForegroundColor Green
        
                    $Result = Invoke-Command -ComputerName $Computer -ScriptBlock {
                        param ($Computer,$Software,$Version)
        
                        if (Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall,
                            HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall |
                            Get-ItemProperty |
                            Where-Object {$_.DisplayVersion -like "$Version" -and $_.DisplayName -like "$Software*"} |
                            Select-Object UninstallString) {
        
                            $Strings = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall,
                                HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall |
                                Get-ItemProperty |
                                Where-Object {$_.DisplayVersion -like "$Version" -and $_.DisplayName -like "$Software*"} |
                                Select-Object UninstallString
        
                            $Strings = $Strings.UninstallString
        
                            $Strings = $Strings.Trim()
        
                            foreach ($Item in $Strings) {
                            
                                C:\Windows\System32\cmd.exe /C $Item /s
                                Write-Host "        Uninstalling String $Item" -ForegroundColor Green
        
                                Sleep 2
                            }
        
                            $Return = "Yes"
                        }
        
                        else {
        
                            Write-Host "$Computer does not have $Software $SoftwareVersion." -ForegroundColor Red
        
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
    }

    End {
    }
}
