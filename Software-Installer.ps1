<#
.SYNOPSIS
Generates dynamic batch file for software installation based on file type, or in the case of .Msu unpacks into .Cab.
Copies software (and batch file if required) to remote computer for installation through WMI.

.PARAMETER PCListPath
Specifies the path to list of PCs.

.PARAMETER SoftwareFolder
Specifies the path to software folder.

.PARAMETER LogFolder
Specifies the path to folder to store logs.

.INPUTS
None. You cannot pipe objects to Software-Installer.ps1.

.OUTPUTS
Csv logs based on results.

.EXAMPLE
.\Software-Installer -PCListPath C:\temp\PCs.txt -SoftwareFolder C:\temp\Software\FireFox -LogFolder C:\temp\Log

.EXAMPLE
.\Software-Installer -PCListPath C:\temp\PCs.txt -SoftwareFolder C:\temp\Updates\KB4566516 -LogFolder C:\temp\Log

.NOTES
Author: Edward Champa
Date Modified: 24 August 2020
#>


[CmdletBinding()]
param (

    [Parameter(Mandatory)]
    [string]
    $PCListPath,

    [Parameter(Mandatory)]
    [string]
    $SoftwareFolder,

    [Parameter(Mandatory)]
    [string]
    $LogFolder
)

$List = Get-Content -Path $PCListPath
$Batfile = "$SoftwareFolder\Install.bat"
$BatList = Get-ChildItem $SoftwareFolder -Recurse

#Clears Log Folder
if (Test-Path $LogFolder) {

    Remove-Item $LogFolder -Recurse -Force
}

#Clears Bat File
if (Test-Path $Batfile) {
    
    Remove-Item $Batfile  -Force
    $Batfile = New-Item -Path $SoftwareFolder -Name "Install.bat" -ItemType "file" -Value ""
}

foreach ($BatItem in $BatList) {

    if ($BatItem | where {$_.extension -eq ".exe"}) {

       Add-Content $BatFile "Start """" ""C:\temp\Software\$BatItem"" /s" # .exe are volatile, might have to play with the arguments based on the software
       $Return = "1"
    }

    elseif ($BatItem | where {$_.extension -eq ".msu"}) {

       C:\Windows\System32\cmd.exe /C expand "$Source\$BatItem" "$Source" -F:*

       $Return = "2"
    }

    elseif ($BatItem | where {$_.extension -eq ".msp"}) {

       Add-Content $BatFile "msiexec /p ""C:\temp\Software\$BatItem"" REINSTALLMODE=""ecmus"" REINSTALL=""ALL"" /quiet /norestart"
       $Return = "1"
    }

    elseif ($BatItem | where {$_.extension -eq ".msi"}) {

       Add-Content $BatFile "msiexec.exe /i ""C:\temp\Software\$BatItem"" /qn"
       $Return = "1"
    }
}

foreach ($Computer in $List) {

    $Source = $SoftwareFolder
    $Destination = "\\$Computer\c$\temp\Software"

    #Tests Connection To Computer       
    if (!($Computer -like "*.*.*.*")) {

        if (Test-Connection $Computer -quiet -count 1) {

            Write-Host "Connection to $Computer is successful. Installing software..." -ForegroundColor Green

            #Clears Software Folder On Computer
            if (Test-Path $Destination) {

                Remove-Item $Destination -Recurse -Force
            }

            #Disables SmartScreen
            Invoke-Command -ComputerName $Computer -ScriptBlock {

                Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" -Name "SmartScreenEnabled" -Type String -Value "Off"
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppHost" -Name "EnableWebContentEvaluation" -Type DWORD -Value 0
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System\" -Name "EnableSmartScreen" -Type DWORD -Value 0
            }

            if ($Return -contains "1") {

                Copy-Item $Source $Destination -Recurse -Force
                Copy-Item $BatFile $Destination -Recurse -Force

                Sleep 5

                $InstallString = "C:\temp\Software\Install.bat"
                Write-Host $InstallString
                ([WMICLASS]"\\$Computer\ROOT\CIMV2:Win32_Process").Create($InstallString)
            }

            else {

                Copy-Item $Source $Destination -Recurse -Force

                Sleep 5

                $Cab = Get-ChildItem -Path $Source | Where-Object {$_.extension -eq ".cab" -and $_.Name -like "Windows*"}

                $InstallString = "C:\temp\Software\$Cab"
                ([WMICLASS]"\\$Computer\ROOT\CIMV2:Win32_Process").Create("DISM.exe /Online /Add-Package /PackagePath:$InstallString /Quiet")
            }

            if (!(Test-Path $LogFolder)){New-Item -ItemType Directory -Force -Path $LogFolder}

            New-Object -TypeName PSCustomObject -Property @{
            ComputerName = $Computer
            } | Export-Csv -NoTypeInformation -Path "$LogFolder\Successful Install PCs.csv" -Append -Force

            Sleep 5

            #Enables SmartScreen
            Invoke-Command -ComputerName $Computer -ScriptBlock {

                Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" -Name "SmartScreenEnabled" -Type String -Value "On"
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppHost" -Name "EnableWebContentEvaluation" -Type DWORD -Value 1
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System\" -Name "EnableSmartScreen" -Type DWORD -Value 1
            }
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
}
