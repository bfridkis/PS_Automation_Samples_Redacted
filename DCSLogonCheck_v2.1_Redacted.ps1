#==========================================================================
#
# PowerScript Source File
#
# NAME: DCSLogonCheck_v2.1.ps1
# Version: 2.1
#
# COMMENT:
#
# This file will run QWINSTA against a list of given machines via "Computers.txt"
# located in the same directory as the script file and return the list of logged
# on users to each machine.  It will identify users logged on locally, remotely,
# or abandoned/disconnected sessions.  The script will replace any user names found
# with a string of further description or full name.
#
# Additional Files:
#
# computers.txt [REQUIRED]
# This is a comma delimited file, first param is the name of the Computer, second
# parameter is a description of the location to be added to the output file.
#
# users.txt [OPTIONAL]
# This is a semicolon delimited file, 2 parameters, 1st parameter is the userID,
# 2nd parameter is the replacement text for the output file (full name with ID)
#
# Version 2 Revision Notes : 2/10/2022
# -Added error message if a node is unreachable.
# -Added html output.
#
# Version 2.1 Revision Notes : 8/29/2023
# -Made bug fixes to account for login counts exceeding 3 per machine, as ConvertTo-HTML module seems to have an issue outputting tabular results in excess of 4 columns. (Need to rebuild the remaining columns when applicable manually.)
#
#==========================================================================

#Initialize error list
$errors = New-Object System.Collections.Generic.List[System.Object]
$errorMessage = "ERROR"

#Initialize sessionList and sessionListObject variables
$sessionList = $null
$sessionListObject = New-Object System.Collections.Generic.List[System.Object]
$sessionListObjectTemp = $null

# Get today's date for the report
$today = Get-Date
 
#Report Title
$title = "DCS Logon Report - " + $today
$title2 = "`r`n---------------------------------------------`r`n"

# Create a fresh variable to collect the results.
$sessionList = $title + $title2

# Read list of computers to query from text file
$computers = Get-Content 'computers.txt'
If (Test-Path .\users.txt -PathType Leaf) { $ADUsers = Get-Content 'users.txt' }

# Loop through the list to query each computer for login sessions
ForEach ($computernamea in $computers) {

    # When running interactively, uncomment the Write-Host line below to show which server is being queried
    # Write-Host "Querying $computernamea"
    $computernamea,$computernameb = $computernamea.Split(",")

    # Run the qwinsta.exe and parse the output
    $queryResults = qwinsta /server:$computernamea
    $qwinstaError = !$?
    $queryResults = $queryResults | foreach { (($_.trim() -replace "\s+",","))} | ConvertFrom-Csv
    
    #Determine if a "person" is logged onto the machine. Used to omit listing machines with no "person" sessions in text output. (Comment out the following 2 lines and delete the '-and $hasPersonSession' from the If statement below the following two lines if all machine names should be output to the text file regardless of presence of a "person" session.)
    $hasPersonSession = $false
    ForEach ($queryResult in $queryResults) { If ($queryResult.STATE -ne $null -or ($queryResult.ID -eq "Disc" -and $queryResult.SESSIONNAME -ne "services")) { $hasPersonSession = $true ; break } }
          
    #If results are registered, add machine name and query info to the session output. (Else add machine to error list and output error message to session output.)
    If ($queryResults -and $hasPersonSession) {
    #If ($queryResults) {
        if ($computernameb -eq $null) {$sessionList += "`r`n" + "*$computernamea" + "`r`n"}
        Else {$sessionList += "`r`n" + "*$computernamea - " + $computernameb + "`r`n"}

        # Pull the session information from each instance. If no reportable sessions, log NO SESSIONS. (To omit NO SESSIONS results from output, comment out line 95.)
        $noReportableSessions = $true
        ForEach ($queryResult in $queryResults) {
            $RDPUser = $queryResult.USERNAME
            $sessionType = $queryResult.SESSIONNAME
            $State = $queryResult.STATE
            $ID = $queryResult.ID
            #$sessionList += "Debug - " + " - " + $RDPUser + " - " +  $sessionType + " - " + $State + " - " + $ID + "`r`n"

            # We only want to display where a "person" is logged in. Otherwise unused sessions show up as USERNAME as a number
            If ($State -eq $NULL -and $ID -eq "Disc" -and $sessionType -notmatch "services") {$sessionList += "User logged on: " + $sessionType + " - Disconnected" + "`r`n" ; $noReportableSessions = $false}
            ElseIf ($State -eq "Active" -and $sessionType -match "rdp") {$sessionList += "User logged on: " + $RDPUser + " - Active (RDP)" + "`r`n" ; $noReportableSessions = $false}
            ElseIf ($State -eq "Active" -and $sessionType -match "console") {$sessionList += "User logged on: " + $RDPUser + " - Active (Local)" + "`r`n" ; $noReportableSessions = $false}
        }
        #If ($noReportableSessions) {$sessionList += "&--NO SESSIONS--" + "`r`n"}
    }
    Elseif ($qwinstaError) {
        $errors.Add($computernamea)
        if ($computernameb -eq $null) {$sessionList += "`r`n" + "*$computernamea" + "`r`n" + "&$errorMessage" + "`r`n"}
        Else {$sessionList += "`r`n" + "*$computernamea  - " + "$computernameb" + "`r`n" + "&$errorMessage" + "`r`n"}
    }
}

# Output results to file
If (Test-Path .\users.txt -PathType Leaf) {
    ForEach ($ADUser in $ADUsers){
        $ADuserID,$ADUserName = $ADUser.Split(";")
        $sessionList = $sessionList -Replace($ADuserID,$ADUserName)
    }
}

#Create SessionList Object for HTML output, then output HTML file results
$sessionListObjectTemp = (($sessionList.Substring($title.Length+$title2.Length+2)).Replace('User logged on: ', '&')) | ConvertFrom-String -Delimiter "$([regex]::Escape("*"))"
$sessionListObjectTemp | Get-Member -MemberType NoteProperty | ForEach-Object {
    If ($sessionListObjectTemp.($_.name) -ne "") {
        $sessionListObject.Add(($sessionListObjectTemp.($_.name) | ConvertFrom-String -Delimiter "&" -PropertyNames Hostname, S1, S2, S3, S4, S5, S6, S7, S8, S9, S10, S11, S12, S13, S14, S15, S16, S17, S18, S19, S20, S21, S22, S23, S24, S25, S26, S27, S28, S29, S30, S31, S32, S33, S34, S35, S36, S37, S38, S39, S40, S41, S42, S43, S44, S45, S46, S47, S48, S49, S50))
    }
}
$sessionListObject | Sort-Object Hostname | ConvertTo-Html -Head $title -CssUri "CurrentDCSLogons.css" | Out-File -FilePath .\CurrentDCSLogonsTEMP.htm

#Delete all newlines and empty cells generated by ConvertTo-HTML (then delete the TEMP file). (Replace method used below does not work well with newline characters.)
$tempResults = Select-String -Pattern "." -Path .\CurrentDCSLogonsTEMP.htm
[string]::join("",($tempResults.line.split("`r`n"))) -replace "<td></td>", ""  | Set-Content -path .\CurrentDCSLogons.htm
Remove-Item .\CurrentDCSLogonsTEMP.htm

#Due to incomplete HTML file generated by ConvertTo-HTML command (which is only an issue when outputting a table, not a list using the -As List flag), add remaining data as needed.
$maxSessionCountUpTo3 = 1
$headerCount = 2
ForEach ($_sessionList in $sessionListObject) {
    If ($_sessionList) {
        $thisSessionCount = ($_sessionList | Get-Member -MemberType NoteProperty).Count - 1
        If ($thisSessionCount -lt 4 -and $thisSessionCount -gt $maxSessionCountUpTo3) { $maxSessionCountUpTo3 = $thisSessionCount }
        If ($thisSessionCount + 1 -gt $headerCount) { $headerCount = $thisSessionCount + 1 }
    }
}
##Add additional headers
##ConvertTo-HTML seems to only generate a maximum of 4 table columns, so need to build replacement strings based on how many columns it generates...
If ($maxSessionCountUpTo3 -eq 1) { $headerStrToReplace = "<th>S1</th>" }
ElseIf ($maxSessionCountUpTo3 -eq 2) { $headerStrToReplace = "<th>S1</th><th>S2</th>" }
Else { $headerStrToReplace = "<th>S1</th><th>S2</th><th>S3</th>" }
$headers = "<th>S1</th>"
For ($i=2;$i -lt $headerCount;$i++) { $headers += "<th>S$i</th>" }
(Get-Content .\CurrentDCSLogons.htm).replace($headerStrToReplace, $headers) | Set-Content .\CurrentDCSLogons.htm

##Add additional session results
ForEach ($_sessionList in $sessionListObject){
    If ($_sessionList) {
        $results = "<td>$($_sessionList.Hostname)</td><td>$($_sessionList.S1)</td>"
        $_sessionList | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -ne "Hostname" -and $_.Name -ne "S1" } | ForEach-Object { $results += "<td>$($_.Definition.split("=")[1])</td>" }
        #Replace method does not work well on `r`n charcters. Strip all newlines from strings used for updating.
        $results = [string]::join("",($results.split("`r`n")))
        $sessionCount = ($_sessionList | Get-Member -MemberType NoteProperty | Measure-Object).Count - 1
        If ($maxSessionCountUpTo3 -eq 1) { $tempStr = [string]::join("",("<td>$($_sessionList.Hostname)</td><td>$($_sessionList.S1)</td>").split("`r`n")) }
        ElseIf ($maxSessionCountUpTo3 -eq 2) { $tempStr = [string]::join("",("<td>$($_sessionList.Hostname)</td><td>$($_sessionList.S1)</td><td>$($_sessionList.S2)</td>").split("`r`n")) }
        Else { $tempStr = [string]::join("",("<td>$($_sessionList.Hostname)</td><td>$($_sessionList.S1)</td><td>$($_sessionList.S2)</td><td>$($_sessionList.S3)</td>").split("`r`n")) }
        (Get-Content .\CurrentDCSLogons.htm).replace($tempStr, $results) | Set-Content .\CurrentDCSLogons.htm
    }
}

#Output Text file results (First remove & delimiter used for creating object needed for HTML output.)
$sessionList = $sessionList.Replace('&', '')
#Comment out the next line if no .txt file is needed.
$sessionList | Out-File -FilePath .\CurrentDCSLogons.txt -Encoding UNICODE

#Uncomment the following if a separate error log is needed. Saves to working directory.
If ($errors.Count -gt 0) {
    New-Item -Path .\CurrentDCSLogons_ErrorLog.txt -ItemType "File" -Value "Errors querying the following nodes:`r`n`r`n" -Force | Out-Null
    $errors | Add-Content .\CurrentDCSLogons_ErrorLog.txt
    Add-Content .\CurrentDCSLogons_ErrorLog.txt -Value "`r`n$errorMessage"
}
