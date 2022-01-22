# psiG Script
# Version 1.0
# Brought to you by Janzent Vapor :>>
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
Write-Output " [psiG] Running psiG Script v1.0"

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
Write-Output " Hello there!"

Start-Sleep 2

Write-Output "`n"
Write-Output " This script is brought to you by Janzent Vapor :>>"

Start-Sleep 3

Write-Output "`n"
Write-Output " First off, what's your name? "
$UserName = Read-Host " Full Name"

Start-Sleep 3

Write-Output "`n"
Write-Output " Great! Nice to meet you, $UserName."
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
Write-Output "  /_/                        by Janzent V. :>>"
Write-Output " =========================="
Write-Output " v1.0 "

Start-Sleep 5

# Operating System Retrieval
Write-Output "`n"
Write-Output " [psiG] Retrieving Operating System Data..."
$ComputerOS = (Get-WmiObject Win32_OperatingSystem).Version
$ComputerOSBuild = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -Property Build*,OSType,ServicePack*

Write-Output " [psiG] Formatting OS Data..."
switch -Wildcard ($ComputerOS) {
    "6.1.7600" {$OS = "Windows 7"; break}
    "6.1.7601" {$OS = "Windows 7 SP1"; break}
    "6.2.9200" {$OS = "Windows 8"; break}
    "6.3.9600" {$OS = "Windows 8.1"; break}
    "10.0.*" {$OS = "Windows 10"; break}
    default {$OS = "Unknown Operating System"; break}
}
Write-Output " [psiG] Done!"

Start-Sleep 2

# Motherboard Retrieval
Write-Output "`n"
Write-Output " [psiG] Retrieving Motherboard Data..."
$ComputerMobo = Get-WmiObject -Class Win32_BaseBoard | 
        Format-Table Manufacturer, Model, Product, SerialNumber, Version

$ComputerBIOS = Get-CimInstance -ClassName Win32_BIOS
Write-Output " [psiG] Done!"

Start-Sleep 2

# Processor Retrieval       
Write-Output "`n"
Write-Output " [psiG] Retrieving Processor Data..." 
$ComputerCPU = Get-WmiObject win32_processor |
        Select-Object DeviceID,Name | Format-Table -AutoSize
Write-Output " [psiG] Done!"

Start-Sleep 2
        
# RAM Retrieval
Write-Output "`n"
Write-Output " [psiG] Retrieving Memory Data..."
$ComputerRAM = Get-WmiObject Win32_PhysicalMemory |
        Select-Object DeviceLocator,Manufacturer,PartNumber,@{n="Capacity";e={[math]::Round($_.Capacity/1GB,2)}},Speed | Format-Table -AutoSize

# RAM Capacity Retrieval
Write-Output " [psiG] Retrieving Memory Capacity Data..."
$ComputerRam_Total = Get-WmiObject Win32_PhysicalMemoryArray |
        Select-Object MemoryDevices,MaxCapacity | Format-Table -AutoSize
Write-Output " [psiG] Done!"

Start-Sleep 2

# Storage Retrieval
Write-Output "`n"
Write-Output " [psiG] Retrieving Storage Disks Data..."
$ComputerDisks = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3" |
        Select-Object DeviceID,VolumeName,@{n="Size";e={[math]::Round($_.Size/1GB,2)}},@{n="FreeSpace";e={[math]::Round($_.FreeSpace/1GB,2)}} | Format-Table -AutoSize
Write-Output " [psiG] Done!"

Start-Sleep 2
# Graphics Retrieval
Write-Output "`n"
Write-Output " [psiG] Retrieving Graphics Data..."
try {
    Start-Sleep 2
    Write-Output " [psiG] Attempting to retrieve Graphics Data."
    $ComputerGPU = Get-WmiObject win32_VideoController
    Write-Output " [psiG] Success! Graphics Data retrieval successful."
    Write-Output " [psiG] Done!"
} catch {
    Write-Output " [psiG] Error. Failed to retrieve Graphics Data."
}

Start-Sleep 5

Write-Output "`n"
Write-Output " [psiG] Consolidating all data gathered..."
try {
    "[psiG Script]`nBrought to you by Janzent Vapor :>> (https://github.com/devjvx)" | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt
    "`nComputer is used by: $UserName" | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    "`nScript Date/Time: $ScriptDateTime" | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    "======================== OPERATING SYSTEM ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    "Operating System: $OS" | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    $ComputerOSBuild  | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    "======================== MOTHERBOARD ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    $ComputerMobo | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    $ComputerBIOS | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    "======================== CPU ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    $ComputerCPU  | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    "======================== MEMORY ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    $ComputerRAM_Total  | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    $ComputerRAM  | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    "======================== STORAGE ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    $ComputerDisks  | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    "======================== GRAPHICS ======================== " | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    $ComputerGPU  | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append
    "-------------------------------- END OF FILE --------------------------------" | Out-File -FilePath psiG-Output-$Username-$ScriptDateTimeNumbers.txt -Append

    Write-Output " [psiG] Done!"
} catch {
    Write-Output " [psiG] Error. Unable to export consolidated data. Check permissions."
}

Start-Sleep 4

Write-Output "`n"
Write-Output " [psiG] All done!"
Write-Output " [psiG] Please send the output text file to the person requesting it."

Write-Output "`n"

Write-Output " psiG Script, out. Peace! (You can exit me now.)"
Write-Output "`n"

Start-Sleep 60