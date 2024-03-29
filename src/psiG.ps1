# psiG Script
# Version 1.1
$scriptVersion = "v1.1"
# Computer Information Retrieval Script :>
# https://github.com/devjvx/psiG

############################
#                 _ ______ #
#     ____  _____(_) ____/ #
#    / __ \/ ___/ / / __   #
#   / /_/ (__  ) / /_/ /   #
#  / .___/____/_/\____/    #
# /_/                      #
############################

#########################################################################

Clear-Host

# Set Console Window's Background and Foreground Color
function Set-ConsoleColor ($bc, $fc) {
    $Host.UI.RawUI.BackgroundColor = $bc
    $Host.UI.RawUI.ForegroundColor = $fc
    Clear-Host
}

$ScriptDateTime = Get-Date -Format "yyyy/MM/dd HH:mm"
$ScriptDateTimeNumbers = Get-Date -Format "yyyy-MM-dd_HH-mm"

Write-Output "`n"
Write-Output " [psiG] Running psiG Script $scriptVersion"

Start-Sleep 3

Set-ConsoleColor 'black' 'white'

Write-Output "      ___           ___                       ___     "
Write-Output "     /\  \         /\  \          ___        /\  \    "
Write-Output "    /::\  \       /::\  \        /\  \      /::\  \   "
Write-Output "   /:/\:\  \     /:/\ \  \       \:\  \    /:/\:\  \  "
Write-Output "  /::\~\:\  \   _\:\~\ \  \      /::\__\  /:/  \:\  \ "
Write-Output " /:/\:\ \:\__\ /\ \:\ \ \__\  __/:/\/__/ /:/__/_\:\__\"
Write-Output " \/__\:\/:/  / \:\ \:\ \/__/ /\/:/  /    \:\  /\ \/__/"
Write-Output "      \::/  /   \:\ \:\__\   \::/__/      \:\ \:\__\  "
Write-Output "       \/__/     \:\/:/  /    \:\__\       \:\/:/  /  "
Write-Output "                  \::/  /      \/__/        \::/  /   "
Write-Output "                   \/__/                     \/__/    "

Start-Sleep 2

Write-Output "`n"
Write-Output " Hello there."

Start-Sleep 2

# Write-Output "`n"
Write-Output " I'm a script that collects information about your computer that's summarized (and complete)!"

Start-Sleep 2

Write-Output "`n"
Write-Output " This script is brought to you by Janzent Vapor :>"

Start-Sleep 3

Write-Output "`n"
Write-Output " First off, what's your name? "
$UserName = Read-Host " Full Name"

Start-Sleep 3

Write-Output "`n"
Write-Output " Great! Nice to meet you, $UserName."

Start-Sleep 3

Write-Output "`n"
Write-Output " [psiG] I'll be running the script for you now."

Start-Sleep 5

#########################################################################

Clear-Host

Write-Output " =========================="
Write-Output "                  _ ______ "
Write-Output "      ____  _____(_) ____/ "
Write-Output "     / __ \/ ___/ / / __   "
Write-Output "    / /_/ (__  ) / /_/ /   "
Write-Output "   / .___/____/_/\____/    "
Write-Output "  /_/                       $scriptVersion"
Write-Output " =========================="

Start-Sleep 5

# Computer Name Retrieval
$ComputerName = Get-WMIObject Win32_ComputerSystem| Select-Object -ExpandProperty Name
Write-Output "`n"
Write-Output " [psiG] Computer Name: $ComputerName"

# Operating System Retrieval
Write-Output "`n"
Write-Output " [psiG] Retrieving Operating System Data..."
$ComputerOS = (Get-WmiObject Win32_OperatingSystem).Version
$ComputerOSBuild = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -Property Build*,OSType,ServicePack*

Write-Output " [psiG] Formatting OS Data..."
switch -Wildcard ($ComputerOS) {
    "6.1.7600" {$summaryOS = "Windows 7"; break}
    "6.1.7601" {$summaryOS = "Windows 7 SP1"; break}
    "6.2.9200" {$summaryOS = "Windows 8"; break}
    "6.3.9600" {$summaryOS = "Windows 8.1"; break}
    "10.0.*" {$summaryOS = "Windows 10"; break}
    default {$summaryOS = "Unknown Operating System"; break}
}
#OS is the final output
Write-Output " [psiG] Done retrieving Operating System information."

Start-Sleep 2

# Motherboard Retrieval
Write-Output "`n"
Write-Output " [psiG] Retrieving Motherboard Data..."
$ComputerMobo = Get-WmiObject -Class Win32_BaseBoard | 
        Format-List Manufacturer, Model, Product, SerialNumber, Version
$ComputerMoboManufacturer = Get-WmiObject -Class Win32_BaseBoard | Select-Object -ExpandProperty Manufacturer
$ComputerMoboProduct = Get-WmiObject -Class Win32_BaseBoard | Select-Object -ExpandProperty Product
Write-Output " [psiG] Done retrieving motherboard information."

Start-Sleep 2

# BIOS Retrieval
Write-Output "`n"
Write-Output " [psiG] Retrieving BIOS Data..."
$ComputerBIOS = Get-CimInstance -ClassName Win32_BIOS |
        Format-List SMBIOSBIOSVersion, Manufacturer, SerialNumber, Version
$ComputerBIOSManufacturer = Get-WmiObject -Class Win32_BIOS | Select-Object -ExpandProperty Manufacturer
$ComputerBIOSVersion = Get-WmiObject -Class Win32_BIOS | Select-Object -ExpandProperty Version
Write-Output " [psiG] Done retrieving BIOS info."

Start-Sleep 2

# Processor Retrieval       
Write-Output "`n"
Write-Output " [psiG] Retrieving Processor Data..." 
$ComputerCPU = Get-WmiObject win32_processor | Format-List
$ComputerCPUName = Get-WmiObject win32_processor | Select-Object -ExpandProperty Name
Write-Output " [psiG] Done retrieving processor information."

Start-Sleep 2
        
# RAM Retrieval
Write-Output "`n"
Write-Output " [psiG] Retrieving Memory Data..."
$ComputerRAM = Get-WmiObject Win32_PhysicalMemory |
        Select-Object DeviceLocator,Manufacturer,PartNumber,@{n="Capacity";e={[math]::Round($_.Capacity/1GB,2)}},Speed | Format-List
$ComputerRAMCapacity = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum /1gb

# RAM Capacity Retrieval
# Write-Output " [psiG] Retrieving Memory Capacity Data..."
# $ComputerRam_Total = Get-WmiObject Win32_PhysicalMemoryArray |
#         Select-Object MemoryDevices,MaxCapacity | Format-List
Write-Output " [psiG] Done getting the memory info."

Start-Sleep 2

# Storage Retrieval
Write-Output "`n"
Write-Output " [psiG] Retrieving Storage Disks Data..."
$ComputerDisks = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3" |
        Select-Object DeviceID,VolumeName,@{n="Size";e={[math]::Round($_.Size/1GB,2)}},@{n="FreeSpace";e={[math]::Round($_.FreeSpace/1GB,2)}} | Format-Table -AutoSize
# Storage not retrieved for summary.
Write-Output " [psiG] Done retrieving the storage data."


Start-Sleep 2

# Graphics Retrieval
Write-Output "`n"
Write-Output " [psiG] Retrieving Graphics Data..."
try {
    Start-Sleep 2
    Write-Output " [psiG] Attempting to retrieve Graphics Data."
    $ComputerGPU = Get-WmiObject win32_VideoController | 
        Select-Object SystemName, Name, CurrentHorizontalResolution, CurrentVerticalResolution, MaxRefreshRate, MinRefreshRate | 
        Format-List

    $ComputerGPURaw = Get-WmiObject win32_VideoController | Select-Object -ExpandProperty Name

    # Split the output into lines
    $lines = $ComputerGPURaw -split "`r?`n"

    # Remove newlines from each line and join with commas
    $ComputerGPUList = $lines -join ", "

    Write-Output " [psiG] Success! Graphics data retrieval successful."
} catch {
    Write-Output " [psiG] Error. Failed to retrieve Graphics Data."
}

Write-Output "`n"
Write-Output " [psiG] Consolidating all data gathered..."

Start-Sleep 5

$fileExportChoice = 0

while($fileExportChoice -eq 0) {

    Clear-Host

    Write-Output " =========================="
    Write-Output "                  _ ______ "
    Write-Output "      ____  _____(_) ____/ "
    Write-Output "     / __ \/ ___/ / / __   "
    Write-Output "    / /_/ (__  ) / /_/ /   "
    Write-Output "   / .___/____/_/\____/    "
    Write-Output "  /_/                       $scriptVersion"
    Write-Output " =========================="

    Start-Sleep 2

    Write-Output "`n"
    Write-Output " Quick question, in what format would you like this to be?"
    Write-Output " Option 1: Text File (Detailed; Regular Text)"
    Write-Output " Option 2: CSV File (Summarized; Excel/Spreadsheet Data)"
    Write-Output " Option 3: Export both"
    Write-Output "`n"

    $fileExportType = Read-Host " Enter your choice (1-3)"

    switch($fileExportType) {
        "1" {
            $fileExportChoice = 1

            try {
                "[psiG Script]`nComputer Information Retrievel Script $scriptVersion (https://github.com/devjvx)" | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt
                "`nComputer is used by: $UserName" | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "`nScript ran at Date/Time: $ScriptDateTime `n" | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "======================== OPERATING SYSTEM ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "`nOperating System: $summaryOS" | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                $ComputerOSBuild | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "======================== MOTHERBOARD ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                $ComputerMobo | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "======================== BIOS ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                $ComputerBIOS | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "======================== CPU ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                $ComputerCPU  | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "======================== MEMORY ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                $ComputerRAM  | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                # $ComputerRAM_Total  | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "======================== STORAGE ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                $ComputerDisks  | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "======================== GRAPHICS ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                $ComputerGPU  | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "-------------------------------- END OF FILE --------------------------------" | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
            
                Write-Output " [psiG] Successfully exported detailed information! See the location of the script."
            } catch {
                Write-Output " [psiG] Error. Unable to export consolidated data. Check file permissions."
            }
        }
        "2" {
            $fileExportChoice = 2

            try {
                $csvData = [PSCustomObject]@{
                    "Computer Name" = $ComputerName
                    "Operating System" = $summaryOS
                    "CPU" = $ComputerCPUName
                    "Motherboard Manufacturer" = $ComputerMoboManufacturer
                    "Motherboard Product" = $ComputerMoboProduct
                    "BIOS Manufacturer" = $ComputerBIOSManufacturer
                    "BIOS Version" = $ComputerBIOSVersion
                    "Memory Capacity" = $ComputerRAMCapacity
                    "GPU" = $ComputerGPUList
                }
                
                $csvData | Export-Csv -Path psiG-Output-CSV-$Username-$ScriptDateTimeNumbers.csv -NoTypeInformation

                Write-Output " [psiG] Exported summarized CSV file! See the location of the script."
            } catch {
                Write-Output " [psiG] Error. Unable to export consolidated data. Check file permissions."
            }
        }
        "3" {
            $fileExportChoice = 3

            try {
                "[psiG Script]`nComputer Information Retrievel Script $scriptVersion (https://github.com/devjvx)" | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt
                "`nComputer is used by: $UserName" | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "`nScript ran at Date/Time: $ScriptDateTime `n" | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "======================== OPERATING SYSTEM ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "`nOperating System: $summaryOS" | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                $ComputerOSBuild | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "======================== MOTHERBOARD ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                $ComputerMobo | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "======================== BIOS ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                $ComputerBIOS | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "======================== CPU ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                $ComputerCPU  | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "======================== MEMORY ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                $ComputerRAM  | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                # $ComputerRAM_Total  | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "======================== STORAGE ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                $ComputerDisks  | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "======================== GRAPHICS ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                $ComputerGPU  | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
                "-------------------------------- END OF FILE --------------------------------" | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
            
                $csvData = [PSCustomObject]@{
                    "Computer Name" = $ComputerName
                    "Operating System" = $summaryOS
                    "CPU" = $ComputerCPUName
                    "Motherboard Manufacturer" = $ComputerMoboManufacturer
                    "Motherboard Product" = $ComputerMoboProduct
                    "BIOS Manufacturer" = $ComputerBIOSManufacturer
                    "BIOS Version" = $ComputerBIOSVersion
                    "Memory Capacity" = $ComputerRAMCapacity
                    "GPU" = $ComputerGPUList
                }

                $csvData | Export-Csv -Path psiG-Output-CSV-$Username-$ScriptDateTimeNumbers.csv -NoTypeInformation

                Write-Output " [psiG] Exported both detailed text file and summarized CSV file! See the location of the script."
            } catch {
                Write-Output " [psiG] Error. Unable to export consolidated data. Check file permissions."
            }
        }
        default {
            
            Write-Output "`n"
            Write-Output " Hmmmm. You seem to have selected an invalid choice. Please choose one :D"
            Write-Output " Hold on, refreshing the choices..."

            Start-Sleep 5
        }
    }
}

Start-Sleep 4

Clear-Host

Write-Output "`n"
Write-Output " [psiG] All done!"
Write-Output " [psiG] Please send the output text file(s) to the person requesting it."

Write-Output "`n"

Write-Output " psiG Script, out. Peace! (You can exit me now.)"
Write-Output "`n"

Start-Sleep 30