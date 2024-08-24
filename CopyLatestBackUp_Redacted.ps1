##Import Active Directory Module
Import-Module -Name ActiveDirectory

##Uncomment if transcript is needed. (Useful for troubleshooting if run from task manager or otherwise without visible shell interface.)
#Start-Transcript -path C:\Users\your_username\Desktop\CustomTools\CopyLatestBackUp_TRANSCRIPT.txt

##Set $WhatIf variable in next line to $true to generate a list of the most recent backup files without actually performing copy operations.
$WhatIf = $true

##Target machine and share location for file copy operations is set here. (i.e. The machine the target drive is plugged into, and its corresponding share name.)
$targetMachine = "SET MACHINE NAME HERE..."
$targetShare = "F"

##Target drive letter. (Need this for checking free space. Do NOT include ':' with drive letter. [e.g. 'F'])
$targetDriveLetter = "F"

##Initialize lists
$targetFiles = New-Object System.Collections.Generic.List[System.Object]
$nullResults = New-Object System.Collections.Generic.List[System.Object]
$errors = New-Object System.Collections.Generic.List[System.Object]

##Initialize $lastBackupCopied variable
$lastBackupCopied = "NONE"

##Use the line below instead of the Get-ADObject cmdlet [2-4 lines below this comment] to reference a list of computers instead of querying AD.
#$computers = Get-Content 'C:\Users\your_username\computers.txt'
Get-ADObject -LDAPFilter "(objectClass=computer)" |
    Where-Object { $_.Name -notlike "redacted_machine_name*" -and $_.Name -notlike "redacted_machine_name3*" -and $_.Name -notlike "redacted_machine_name2*" } |
    Select-Object Name | Sort-Object Name | Set-Variable -Name computers

##Counter for outputting number of machines processed to console
$numberProcessed = 0

##Generate list of files to copy, and [if $WhatIf is $false], copy files to target destination.
Write-Host "`r`n"
ForEach ($computer in $computers) {
   If ($computer.Name) { $computerName = $computer.Name } Else { $computerName = $computer }
   Write-Host "$numberProcessed Machines Processed. $($computers.Count - $numberProcessed) Machines Remaining. Working on $computerName..."
   ##Reset $latestFullBackup_NAS03 from previous iteration
   $latestFullBackup_NAS03 = $null
   Try {
       ##Search redacted_machine_name401 for this node's latest backup
       $latestFullBackup_NAS01 = Get-ChildItem -Path "\\redacted_machine_name401\Images\*" -Recurse -ErrorAction Stop -Filter "*.tib*" |
                                 Where-Object { $_.Name -like "$computerName*.tib*" -and $_.FullName -notlike "*hold*" -and $_.FullName -notmatch "[a-zA-Z_0-9]*D\d\.tib*" -and $_.Length -gt 102400 }
                                #| Sort-Object LastWriteTime -Descending | Select-Object -First 1
       
       ##Preferentially select EBR501 directory if backups exist therein, and filter down to most recent. (NAS01)
       $latestFullBackup_NAS01_EBR501 = $latestFullBackup_NAS01 | Where-Object { $_.FullName -like "*EBR501*" }
       If ($latestFullBackup_NAS01_EBR501) { $latestFullBackup_NAS01 = $latestFullBackup_NAS01_EBR501 | Sort-Object LastWriteTime -Descending | Select-Object -First 1 }
       Else { $latestFullBackup_NAS01 = $latestFullBackup_NAS01 | Sort-Object LastWriteTime -Descending | Select-Object -First 1 }
       
       ##If not found on redacted_machine_name401, try redacted_machine_name403 - Can uncomment this conditional if we are certain backups don't exist in both locations.
       #If (!$latestFullBackup_NAS01) {
           $latestFullBackup_NAS03 = Get-ChildItem -Path "\\redacted_machine_name403\Images\*" -Recurse -ErrorAction Stop -Filter "*.tib*" |
           Where-Object { $_.Name -like "$computerName*.tib*" -and $_.FullName -notlike "*hold*" -and $_.FullName -notmatch "[a-zA-Z_0-9]*D\d\.tib*" -and $_.Length -gt 102400 }
          #| Sort-Object LastWriteTime -Descending | Select-Object -First 1

           ##Preferentially select EBR501 directory if backups exist therein, and filter down to most recent. (NAS03)
           $latestFullBackup_NAS03_EBR501 = $latestFullBackup_NAS03 | Where-Object { $_.FullName -like "*EBR501*" }
           If ($latestFullBackup_NAS03_EBR501) { $latestFullBackup_NAS03 = $latestFullBackup_NAS03_EBR501 | Sort-Object LastWriteTime -Descending | Select-Object -First 1 }
           Else { $latestFullBackup_NAS03 = $latestFullBackup_NAS03 | Sort-Object LastWriteTime -Descending | Select-Object -First 1 }
       #}
       ##Depending on where backups are found (NAS01 vs. NAS03), assign $latestFullBackup the newest backup available between the two locations.
       switch ($latestFullBackup_NAS01) {
           { !$_ -and !$latestFullBackup_NAS03 } { $latestFullBackup = $null ; Break }
           { $_ -and $latestFullBackup_NAS03 } { If ($_.LatestWriteTime -gt $latestFullBackup_NAS03.LastWriteTime) { $latestFullBackup = $_ } Else { $latestFullBackup = $latestFullBackup_NAS03 } ; Break }
           { $_ -and !$latestFullBackup_NAS03 } { $latestFullBackup = $_ ; Break }
           { !$_ -and $latestFullBackup_NAS03 } { $latestFullBackup = $latestFullBackup_NAS03 ; Break }
       }
       ##If it is found, add to $targetFiles list.
       If ($latestFullBackup) { $targetFiles.Add(@{'Hostname' = $computerName ;
                                'Latest_Backup' = $latestFullBackup.FullName ;
                                'LastWriteTime' = $latestFullBackup.LastWriteTime ;
                                'Size_KB' = $latestFullBackup.Length / 1024 ;
                                'Copy_Result' = $(If (!$WhatIf) { 'SUCCESS' } Else { 'N/A' })})
       }
       ##Copy latest backup to target destination if $WhatIf is $false and a backup was identified.
       If (!$WhatIf -and $latestFullBackup) {
           ##Check to make sure enough free space exists. If not, prompt user to enter new drive.
           $targetDriveFreeSpace = (Get-WmiObject Win32_LogicalDisk -ComputerName $targetMachine -Filter "DeviceID='$($targetDriveLetter):'").FreeSpace
           While ($targetDriveFreeSpace -lt $latestFullBackup.Length) {
               Write-Host "`nInsufficient free space on target drive. Replace drive and enter drive letter below to resume. (Do NOT include ':' character. [e.g. 'F']"
               Write-Host "(Can leave blank if drive letter unchanged.)"
               Write-Host "`nLast Backup Copied: $lastBackupCopied"
               $newDriveLetter = Read-Host "`nDrive Letter"
               If ($newDriveLetter) { $targetDriveLetter = $newDriveLetter }
               $targetDriveFreeSpace = (Get-WmiObject Win32_LogicalDisk -ComputerName $targetMachine -Filter "DeviceID='$($targetDriveLetter):'").FreeSpace
           }
           Try {
               Copy-Item $latestFullBackup.FullName -Destination "\\$targetMachine\$targetShare\$computerName"
               $lastBackupCopied = $computerName
           }
           Catch {
               $errors.Add( @{ 'Hostname' = $computerName ; 'Exception' = $_.Exception.Message } )
               $targetFiles[$targetFiles.Count - 1].Copy_Result = "FAILED (See Error Log)"
           }
       }
       Elseif (!$latestFullBackup) { $nullResults.Add("$($computer): Null Result. No Backups Found.") }
   }
   Catch { $errors.Add( @{ 'Hostname' = $computerName ; 'Exception' = $_.Exception.Message } ) }

   $numberProcessed++
}

##Generate list of latest backups, indicate copy result status if applicable (i.e. if $WhatIf is $false).
$targetFiles | Select-Object @{n='Hostname' ; e={$_.Hostname}},
                             @{n='Latest Backup' ; e={$_.Latest_Backup}},
                             @{n='LastWriteTime' ; e={$_.LastWriteTime}},
                             @{n='Size (KB)' ; e={$_.Size_KB}},
                             @{n='Copy Result' ; e={$_.Copy_Result}} |
               Export-CSV C:\Users\your_username\Desktop\CustomTools\CopyLatestBackUp.csv -NoTypeInformation
$nullResults | Add-Content C:\Users\your_username\Desktop\CustomTools\CopyLatestBackUp.csv

##Generate error log
$errors | Select-Object @{ n = 'Unavailable Hosts' ; e = {$_.Hostname}},
                        @{ n = 'Exceptions Generated' ; e = {$_.Exception}} |
          Out-File $PSScriptRoot\CopyLatestBackUpERRORS.csv

##Uncomment if transcript is needed. (Useful for troubleshooting if run from task manager or otherwise without visible shell interface.)
#Stop-Transcript
