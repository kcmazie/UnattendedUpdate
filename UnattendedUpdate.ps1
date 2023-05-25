param(
    [Switch]$Console = $False,        #--[ Set to true to enable local console result display. Defaults to false ]--
    [Switch]$Debug = $False,          #--[ Set to true to only send results to debug email address. Default to false ]--
    [Switch]$Manual = $False,         #--[ Use to run update off-schedule ]--
    [Switch]$Status = $False,         #--[ If set to true checks for results after the reboot and emails, then goes idle. ]--
    [Switch]$Deploy = $False,         #--[ If set to true will copy this script to the other members of the peer group ]--
    [Switch]$UpTimeCheck = $False,    #--[ If set to true will send an email alert if no restart occurs in preset # of days ]--
    [Switch]$NoRestart = $False       #--[ Stops restart from occurring. Restart may still occur if update determines it's needed. ]--
    )
<#======================================================================================
          File Name : UnattendedUpdate.ps1
    Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com)
                    :
        Description : Will scan the Windows Update site and install all missing updates.
                    :
          Operation : Requires PowerShell v5. Requires NuGet and PSWindowsUpdate modules. Will auto-install them if needed.
                    : Reboots system following patching to assure new updates are applied. Creates a LOCAL scheduled task
                    : automatically on first run with script version in name. Will delete and recreate the task if
                    : a script version change is detected. The patch schedule is set to randomize the run
                    : within a 90 minute window. Update routine is governed by the week days noted in the config file.
                    : The scheduled task has two triggers, run time, and on restart. An HTML report is sent on every restart.
                    : Antivirus is disabled during the update run. Kaspersky is configure so change that as needed.
                    : Requires a config file in XML format located in the same folder as the script. See example at bottom.
                    :
          Arguments : Normal operation is with no command line options.
                    : -Console $true (Will enable local console output)
                    : -Debug $true (Not used)
                    : -Manual $true (forces a manual run bypassing the schedule)
                    : -Status $true (forces a status email to be sent)
                    : -Deploy $true (forces the script and config file to be copied to the identical
                    : location on the other peer servers listed in the config file)
                    : -UpTimeCheck $true (If set to true will send an email alert if no restart occurs in preset # of days)
                    : -NoRestart $true (Stops restart from occurring. Restart may still occur if update determines it's needed.)
                    :
           Warnings : Uses local SYSTEM user context for tasks. Install LOCALLY, not remotely.
                    : Adjust the task schedule(s) to conform to your maintenance window.
                    :
              Legal : Public Domain. Modify and redistribute freely. No rights reserved.
                    : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF
                    : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.
                    :
            Credits : Code snippets and/or ideas came from many sources including but
                    : not limited to the following:
                    : https://www.powershellgallery.com/packages/PSWindowsUpdate/1.5.2.2
                    :
     Last Update by : Kenneth C. Mazie (kcmjr AT kcmjr.com)
                    :
    Version History : v1.00 - 12-29-16 - Original
     Change History : v2.00 - 01-13-17 - Added forced status option for Sundays.
                    : v2.10 - 01-17-17 - Added regkey delete.
                    : v2.20 - 04-06-17 - Fixed row data to start clean at each loop
                    : v2.30 - 09-28-17 - Turned off extra email after patching.
                    : Moved config file out. Added script replication option.
                    : Added run day from config file option.
                    : v2.40 - 10-19-17 - Added randomizer for reboot. Added check for no data.
                    : v2.50 - 12-01-17 - Fixed runday detection
                    : v2.60 - 02-27-18 - Eliminated second schedule to send status at restart due to bugs in PS
                    : task schedule commandlets not detecting all schedules.
                    : v2.70 - 02-28-18 - Added reboot detection for automatic status report.
                    : v2.80 - 03-01-18 - Added automatic scheduling adjustmant. Added restart bypass.
                    : v2.90 - 03-05-18 - Fixed registry key removal error.
                    : v3.00 - 10-12-18 - Changed AV disable to default to false to stop false emails. Added option
                    : to select wsus (in-house) or windows update (internet) as update source
                    :
                    #>
     $Script:ScriptVer = "3.00"
                    <#
=======================================================================================#>
<#PSScriptInfo
.VERSION 3.00
.AUTHOR Kenneth C. Mazie (kcmjr AT kcmjr.com)
.DESCRIPTION
Automatically applies current patches to a single Windows system, then reboots. Emails a status report upon restart. Should be run from a scheduled task. Can deploy itself to "peer" systems.
#>
#Requires -version 5.0
clear-host

if ($Console){$Script:Console = $true}
if ($Debug){$Script:Debug = $true}
if ($Manual){$Script:Manual = $true}
if ($Status){$Script:Status = $true}
if ($Deploy){
    $Script:Deploy = $true
    $Script:Console = $true
}
if ($UpTimeCheck){$Script:UpTimeCheck = $true}
if ($NoRestart){$Script:NoRestart = $true}

$ErrorActionPreference = "SilentlyContinue"
$Now = Get-Date -Format "MM-dd-yyyy_HHmm"
$Script:ThisComputer = $Env:Computername
$Script:MessageBody = @() 
$Today = (get-date).dayofweek
$Script:Attach = $false
$Script:ResultLog = "$PSScriptRoot\Results-$Now.csv"

#--[ This next line is used to stop the local antivirus client. Edit the function below to support your Av client ]--
#--[ If the AV changes or is not used leaving this enabled will cause an extra junk email to go out ]--
$Script:KillKaspersky = $false            
#---------------------------------------------------------------------------------------------------------------

$Script:ScriptName = $MyInvocation.MyCommand.Name 
$Script:ScriptFullPath = $PSScriptRoot+"\"+$MyInvocation.MyCommand.Name 
$ConfigFile = $Script:ScriptFullPath.Split(".")[0]+".xml"
$Script:UserContext = [Security.Principal.WindowsIdentity]::GetCurrent()

#==[ Functions ]================================================================

Function LoadConfig {
#--[ Read and load configuration file ]-------------------------------------
    if (!(Test-Path $ConfigFile)){                       #--[ Error out if configuration file doesn't exist ]--
        $Script:HTMLData = "MISSING CONFIG FILE. Script aborted."
        if ($Script:Log){Add-content -Path "$PSScriptRoot\debug.txt" -Value "MISSING CONFIG FILE. Script aborted."}
        write-host "CONFIGURATION FILE $ConfigFile NOT FOUND - EXITING" -ForegroundColor Red 
        break;break;break
    }else{
        [xml]$Script:Configuration = Get-Content $ConfigFile  #--[ Read & Load XML ]--
        $Script:PeerList = $Script:Configuration.Settings.General.PeerList
        $Script:RunDays = $Script:Configuration.Settings.General.RunDays
        $Script:RunTime = $Script:Configuration.Settings.General.RunTime
        $Script:UpdateSource = $Script:Configuration.Settings.General.UpdateSource
        $Script:DebugEmail = $Script:Configuration.Settings.Email.Debug 
        $Script:eMailRecipient = $Script:Configuration.Settings.Email.To
        $Script:eMailFrom = $ThisComputer+'_'+$Script:Configuration.Settings.Email.From    
        $Script:eMailHTML = $Script:Configuration.Settings.Email.HTML
        $Script:eMailSubject = $ThisComputer+' '+($Script:Configuration.Settings.Email.Subject)
        $Script:SmtpServer = $Script:Configuration.Settings.Email.SmtpServer
        $Script:UserName = $Script:Configuration.Settings.Credentials.Username
        $Script:EncryptedPW = $Script:Configuration.Settings.Credentials.Password
        $Script:Base64String = $Script:Configuration.Settings.Credentials.Key
        $Script:ReportName = $ThisComputer+' '+($Script:Configuration.Settings.General.ReportName)
        $Script:UpTimeDays = $Script:Configuration.Settings.General.UpTimeDays
    }
    $ByteArray = [System.Convert]::FromBase64String($Script:Base64String);
    $Script:Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Script:UserName, ($Script:EncryptedPW | ConvertTo-SecureString -Key $ByteArray)
    $Script:Password = $Script:Credential.GetNetworkCredential().Password
}

function SendEmail {
    $SMTP = new-object System.Net.Mail.SmtpClient($Script:SmtpServer)
    $Email = New-Object System.Net.Mail.MailMessage
    $Email.Body = $Script:MessageBody
    $Email.IsBodyHtml = $Script:eMailHTML
    #If ($Script:Debug){
        $Email.To.Add($Script:DebugEmail)
    #}Else{
    # $Email.To.Add($Script:eMailRecipient)
    #}
    if ($Script:Attach){
        $Attachment = New-Object System.Net.Mail.Attachment -ArgumentList $Script:ResultLog, 'Application/Octet'
        $Email.Attachments.Add($Attachment)
    }    
    $Email.From = $Script:eMailFrom 
    $Email.Subject = $Script:eMailSubject
    $SMTP.Send($Email)
    $Email.Dispose()
    $SMTP.Dispose()
}

function GetResults {
    $Script:ResultOut = ""
    $Script:ResultLine = ""
    $RowFlag = $false
    #--[ Add header to html output file ]--------------------
    $Script:MessageBody += '<tr><th>KB Number</th><th>Status</th><th>Results</th><th>Date / Time</th><th>Messages</th></tr>'
    #--[ HTML row Settings ]----------------------------------------------------
    $BGColor = "#dfdfdf"                                                    #--[ Grey default cell background ]--
    $BGColorRed = "#ff0000"                                                 #--[ Red background for alerts ]--
    $BGColorOra = "#ff9900"                                                 #--[ Orange background for alerts ]--
    $BGColorYel = "#ffd900"                                                 #--[ Yellow background for alerts ]--
    $FGColor = "#000000"                                                    #--[ Black default cell foreground ]--
    
    #--[ Only keep 10 of the last runtime logs ]------------------------------------
    Get-ChildItem -Path $PSScriptRoot | Where-Object {(-not $_.PsIsContainer) -and ($_.Name -like "*results*.csv")} | Sort-Object -Descending -Property LastTimeWrite | Select-Object -Skip 10 | Remove-Item 
    
    #--[ Scan Eventlogs for Event 19 & 20 ]--
    $Script:LogDump = Get-WinEvent -FilterHashtable @{LogName = "System";ID=19,20} 
    $Script:KBTracker = @()
    foreach ($Script:LogItem in $Script:LogDump ){
        if ($Script:LogItem.ProviderName -eq 'Microsoft-Windows-WindowsUpdateClient'){
            $Script:LogItemStat = $Script:LogItem.message.split(":")[0]
            if ($Script:LogItem.Message -like "*(KB*"){
                $Script:LogItemKB = ($Script:LogItem.message.split("(")[1]).split(")")[0]
                if (!($Script:LogItemKB -like "KB*")){
                    $Script:LogItemKB = ($Script:LogItem.message.split("(")[2]).split(")")[0]
                }
            }else{
                $Script:LogItemKB = $Script:LogItem.message.split(":")[2] #"---------"
            }  

            If ($Script:KBTracker -notcontains ($Script:LogItemKB+" "+$Script:LogItemStat)){            #--[ Have not seen this KB yet ]--
                $RowFlag = $true
                $RowData = '<tr>'                                                                                                           #--[ Start table row ]--

                if ($Script:Console){write-host $Script:LogItemKB" " -ForegroundColor Red -NoNewline }
                $RowData += '<td bgcolor=' + $BGColor + '><font color=' + $FGColor + '>' + $Script:LogItemKB + '</td>'                      #--[ KB number to html table ]--
           
                if ($Script:Console){write-host $Script:LogItem.LevelDisplayName" " -ForegroundColor yellow -NoNewline }
                if ($Script:Console){write-host $Script:LogItemStat" " -ForegroundColor Cyan -NoNewline }
                if ($Script:LogItem.LevelDisplayName -like "Error"){ 
                    $RowData += '<td bgcolor=' + $BGColor + '><font color=#800000>' + $Script:LogItem.LevelDisplayName + '</td>'            #--[ Status (if error) to html table ]--
                    $RowData += '<td bgcolor=' + $BGColor + '><font color=#800000>' + $Script:LogItemStat + '</td>'                         #--[ Results to html table ]--
                }else{    
                    $RowData += '<td bgcolor=' + $BGColor + '><font color=' + $FGColor + '>' + $Script:LogItem.LevelDisplayName + '</td>'   #--[ Status to html table ]--
                    $RowData += '<td bgcolor=' + $BGColor + '><font color=' + $FGColor + '>' + $Script:LogItemStat + '</td>'                #--[ Results to html table ]--
                }

                if ($Script:Console){write-host $Script:LogItem.TimeCreated" " -ForegroundColor yellow -NoNewline }                     #--[ Date/time to html table ]--
                $RowData += '<td bgcolor=' + $BGColor + '><font color=' + $FGColor + '>' + $Script:LogItem.TimeCreated + '</td>'
                      
                if ($Script:Console){write-host $Script:LogItem.Message" " -ForegroundColor green}
                $RowData += '<td bgcolor=' + $BGColor + '><font color=' + $FGColor + '>' + $Script:LogItem.message + '</td>'                #--[ Result message to html table ]--
            
                $Script:ResultLine = $Script:LogItemKB+","+$Script:LogItem.LevelDisplayName+","+$Script:LogItem.TimeCreated+","+$Script:LogItemStat+","+$Script:LogItem.Message
                $Script:ResultOut = $Script:ResultOut+$Script:ResultLine+"`n"  
                $RowData += '</tr>'

                $Script:KBTracker += ($Script:LogItemKB+" "+$Script:LogItemStat)
                $Script:MessageBody += $RowData
                $Script:KBTracker
            }
        }
    }

    If ($RowFlag){
        $Script:Attach = $true
    }Else{
        $Script:MessageBody += '<td colspan=5 bgcolor=' + $BGColor + '><font color=' + $FGColor + '><center>No New Data to Report </center></td>'  
    }
    
    Add-Content -value $Script:ResultOut -path $Script:ResultLog
    $Script:MessageBody += '</table><br>'
    $Script:MessageBody += "<font size=3 face='times new roman'>- Done. Emailing results...<br>"
    if ($Script:Console){Write-Host "`n- Done. Emailing results...`n" -ForegroundColor yellow }
    SendEmail
}

Function ServiceMgr ($Svc, $SvcStatus) {
    $Script:MessageBody += "- Processing: $Svc<br>"
    if ($Script:Console){Write-Host "`n- Processing :$Svc" -ForegroundColor cyan }
    $Count = 0
    #--[ State prior to start/stop process ]--
    $Script:SvcState = (Get-Service -Name $Svc).Status
    $Script:MessageBody += "-- $Svc Initial Status: $Script:SvcState<br>"
    if ($Script:Console){Write-Host "-- $Svc Initial Status: $Script:SvcState" -ForegroundColor yellow }
    $Script:MessageBody += "--- Pausing while attemtping to set service state to: $SvcStatus...<br>" 
    if ($Script:Console){Write-Host "--- Pausing while attemtping to set service state to: $SvcStatus..." -ForegroundColor yellow }
    while ((Get-Service -Name $Svc).Status -ne $SvcStatus){
        Get-Service -Name $Svc | Set-Service -Status $SvcStatus        
        sleep -Milliseconds 500
        $Count ++
        if ($Count -ge 20){
            if ($Script:Console){Write-Host "-- There was an error setting the "$Svc" service to the "$SvcStatus" state..." -ForegroundColor red }
            $Script:MessageBody += '-- There was an error setting the "$Svc" service to the "$SvcStatus" state...<br>'
            break
        }
    }
    #--[ State after start/stop process ]--
    $Script:SvcState = (Get-Service -Name $Svc).Status
    $Script:MessageBody += "-- $Svc Final Status: $Script:SvcState<br>"
    if ($Script:Console){Write-Host "-- $Svc Final Status: $Script:SvcState" -ForegroundColor yellow }
    Sleep -Seconds 1
    if ($Script:SvcState -eq "stopped"){
        RegKill
    }
}

function RegKill {
    if ($Script:Console){Write-Host "`n- Removing Windows Update registry key..." -ForegroundColor cyan }
    try{
        Clear-ItemProperty -Name 'WUServer' -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -Force 
        Clear-ItemProperty -Name 'WUStatusServer' -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -Force 
        Remove-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -Force -Recurse 
    }catch{
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        $Script:MessageBody += "- Failed to remove Windows Update Key(s). $ErrorMessage<br>"
        if ($Script:Console){Write-Host "- Failed to remove Windows Update Key(s). $ErrorMessage" -ForegroundColor Red }
    }
    if (Test-Path -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'){
        if ($Script:Console){Write-Host "-- FAILED to remove Windows Update registry key..." -ForegroundColor red }
        $Script:MessageBody += "- Failed to remove Windows Update Key(s).<br>"
    }else{
        if ($Script:Console){Write-Host "-- Verified removal of Windows Update registry key..." -ForegroundColor green }
        $Script:MessageBody += "- Verified removal of Windows Update registry key(s).<br>"
    }
}

function PatchIt {
    if ($Script:KillKaspersky){
        ServiceMgr "klnagent" "stopped"  #--[ Stops Kaspersky agent prior to running update. Comment out above if not applicable. ]--
    }else{
        $SvcState = "stopped"
    }
    if ($SvcState -eq "stopped"){
        #--[ NuGet is required to pull the module from the MS repository ]--
        if (!(Get-PackageProvider NuGet)){
            if (!(Get-ChildItem -Path "C:\Program Files\PackageManagement\ProviderAssemblies\nuget" -Filter "Microsoft.PackageManagement.NuGetProvider.dll" -Recurse)){
                try{
                    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -ErrorAction SilentlyContinue -Confirm:$false -Force:$true
                    $Script:MessageBody += "- NuGet provider is being installed.<br>"
                    if ($Script:Console){Write-Host "- NuGet provider is being installed." -ForegroundColor yellow }
                }catch{
                    $ErrorMessage = $_.Exception.Message
                    $FailedItem = $_.Exception.ItemName
                    $Script:MessageBody += "- NuGet module install FAILED. $ErrorMessage<br>"
                    if ($Script:Console){Write-Host "- NuGet module install FAILED. $ErrorMessage" -ForegroundColor Red }
                }    
            }
        }

        #--[ Install the update module if it's not already loaded ]--
        if (!(Get-Module PSWindowsUpdate)){
            try{
                $Script:MessageBody += '- "PSWindowsUpdate" module is being installed.<br>'
                if ($Script:Console){Write-Host '- "PSWindowsUpdate" module is being installed...' -ForegroundColor yellow }
                Install-Module PSWindowsUpdate -ErrorAction Stop -Confirm:$false -Force:$true
            }catch{    
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                $Script:MessageBody += '- "PSWindowsUpdate" module install FAILED. $ErrorMessage<br>'
                if ($Script:Console){Write-Host '- "PSWindowsUpdate" module install FAILED. $ErrorMessage' -ForegroundColor Red }
            }
        }

        #--[ Register to use the Microsoft Update Service, as opposed to just the default Windows Update Service. ]--
        if (!((Get-WUServiceManager).ServiceID -contains "7971f918-a847-4430-9279-4a52d1efe18d")){
            try{
                $Script:MessageBody += "- Service Manager ID is being registered.<br>"
                if ($Script:Console){Write-Host "- Service Manager ID is being registered." -ForegroundColor yellow }
                Add-WUServiceManager -ServiceID '7971f918-a847-4430-9279-4a52d1efe18d' -Confirm:$false -ErrorAction Stop
            }catch{
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                $Script:MessageBody += "- Service Manager ID registration FAILED. $ErrorMessage<br>"
                if ($Script:Console){Write-Host "- Service Manager ID registration FAILED. $ErrorMessage" -ForegroundColor Red }
            }    
        }

        #--[ Run the update in unattended mode, check for all updates on the MS update site, accept all EULAs, reboot if needed ]--
        $Script:MessageBody += "- Checking for and installing new updates.<br>"
        $Script:MessageBody += "-- A reboot is required to register new patches, as such the system will be rebooted shortly.<br>"
        $Script:MessageBody += "-- A summary report will be dispatched shortly after the system comes back online.<br>"
        if ($Script:Console){
            Write-Host "`n- Checking for and installing new updates." -ForegroundColor cyan
            Write-Host "-- Note that this process produces no console output." -ForegroundColor yellow 
            Write-Host "-- A reboot is required to register new patches, as such the system will be rebooted shortly." -ForegroundColor yellow 
            Write-Host "-- A summary report will be dispatched shortly after the system comes back online." -ForegroundColor yellow         
        }
        $Script:Attach = $false 
        #SendEmail #--[ Disabled so that the emails only go out are after reboot or on svc error ]--

        #--[ Select update source, either WSUS or MS Update web site. Selected from config file ]--
        If ($Script:UpdateSource -eq "wsus"){
            Get-WUInstall -WindowsUpdate -AcceptAll -Confirm:$false -AutoReboot:$true
        }Else{    
            Get-WUInstall -MicrosoftUpdate -AcceptAll -Confirm:$false -AutoReboot:$true
        }
    }else{
        $Script:MessageBody += "-- AntiVirus service failed to stop -- ABORTING --"
        if ($Script:Console){Write-Host "`n-- $Svc Failed to stop -- ABORTING --`n" -ForegroundColor Red }
        SendEmail
    }
    
    #--[ Things to do if no reboot aurtomatically occurs after running update. Comment items out if not applicable. ]--
    Sleep -Seconds 60  #--[ Wait to see if no reboot has occurred ]--
    if ($Script:KillKaspersky){
        ServiceMgr "klnagent" "running"    #--[ Restart it if we don't reboot ]--
    }else{
        $SvcState = "running"
    }
    
    #--[ Force a reboot with random delay if none occurred. ]--
    $RndArray = @(300,600,900)   #--[ 300 seconds = 5 minutes]--
    $Rnd = (new-object System.Random)
    $RndDelay = $Array[ $Rnd.Next( $Array.Count ) ]
    $RndDelay = 5
    if ($Script:Console){Write-Host `n'--- REBOOTING COMPUTER --- ('$RndDelay' second delay)...' -ForegroundColor red }
    Sleep -Seconds $RndDelay
    If ($NoRestart){
        if ($Script:Console){Write-Host `n'--- REBOOT BYPASS ENABLED --- ' -ForegroundColor yellow }    
    }Else{
        Try{
            #Restart-Computer -Credential $Script:Credential -Confirm $false -force #-WhatIf #--[ Not working for the local PC ]--
            Invoke-Command -ComputerName "localhost" -Credential $Script:Credential -ScriptBlock {shutdown -r -t 3}
        }Catch{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            if ($Script:Console){Write-Host '- System Restart has failed to execute... $ErrorMessage' -ForegroundColor Red }
        }
    }    
}
    
Function ScheduledTask ($ActiveTask){
    #$ActiveTask = "UnattendedUpdate-2.9" #--------------------------- for testing ---------------------------------
    $CreateTask = $True
    $Script:TaskMessage = ""
    $ExistingTasks = (Get-ScheduledTask | Where-Object {$_.TaskName -like ('*'+$ActiveTask.Split("-")[0]+'*')})

    ForEach ($FoundTask in $ExistingTasks){
        If (($FoundTask.taskname.Split("-")[1]) -match '\d'){
            If ($FoundTask.TaskName -eq $ActiveTask ){
                if ($Script:Console){Write-Host '- Task "'$FoundTask.TaskName'" already exists... IGNORING' -ForegroundColor Green }
                $Script:TaskMessage += '- Task "'+$FoundTask.TaskName+'" already exists...<br>'
                $CreateTask = $False
            }Else{
                if ($Script:Console){Write-Host '- Task "'$FoundTask.TaskName'" is an incorrect version... REMOVING' -ForegroundColor red }
                $Script:TaskMessage += '- Task "'+$FoundTask.TaskName+'" is an incorrect version... REMOVING<br>'
                Try{
                    #$Command = Get-ScheduledTask | Where-Object {$_.TaskName -eq $FoundTask.TaskName}
                    #Unregister-ScheduledTask -taskname $Command -taskpath "\" #-Confirm:$false #--[ not working [--
                    $Result = Invoke-Expression ("schtasks.exe /delete /s "+$Env:ComputerName+" /tn "+$FoundTask.TaskName+" /F")
                    if ($Script:Console){Write-Host "-- "$Result -ForegroundColor white }
                    $Script:TaskMessage += '-- "'+$Result+'<br>'
                    $CreateTask = $true 
                }Catch{
                    $_.Exception.Message
                    $_.Exception.ItemName
                }
            }
        }Else{
            if ($Script:Console){Write-Host '- Task "'$FoundTask.TaskName'" is unknown... IGNORING' -ForegroundColor yellow }
        }            
    }
 
    If ($CreateTask){
        If ($Script:Console){Write-Host '- Creating new scheduled task "'$ActiveTask' "...' -ForegroundColor Cyan }
        $Script:TaskMessage += '- Creating new scheduled task "'+$ActiveTask+'"...<br>'
        #--[ Task Parameters ]--------------------
        $Principal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType S4U -RunLevel Highest 
        $PSArgument = '-WindowStyle Hidden -Noninteractive -noprofile -nologo -executionpolicy Bypass -Command "&{'+$Script:ScriptFullPath+'}"'
        $Action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $PSArgument 
        $Trigger = @()    #--[ Allows creation of multiple triggers ]--
        $Trigger += New-ScheduledTaskTrigger -Daily -RandomDelay (New-TimeSpan -Minutes 90) -At $Script:RunTime   #--[ Creates patch task with 90 minute random delay ]--
        $Trigger += New-ScheduledTaskTrigger -AtStartup 
        $Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Principal $Principal -Settings (New-ScheduledTaskSettingsSet) 
        #--[ Task Parameters ]--------------------
        try{
            $Result = ($Task | Register-ScheduledTask -TaskName $ActiveTask -Force -ErrorAction Stop )  
            if ($Script:Console){Write-Host "-- Created"$Result -ForegroundColor green }
            $Script:TaskMessage += '-- Created"'+$Result+'<br>'
        }catch{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            $Script:TaskMessage += '- Scheduled Task "'+$ActiveTask+'" failed to be created... $ErrorMessage<br>'
            if ($Script:Console){Write-Host '- Scheduled Task "'$ActiveTask'" failed to be created... $ErrorMessage' -ForegroundColor Red }
        } 
    }Else{
        $Script:TaskMessage += '- Scheduled Task "'+$ActiveTask+'" detected... No Action... <br>'
        if ($Script:Console){Write-Host '- Scheduled Task "'$ActiveTask'" detected... No Action Required... ' -ForegroundColor Green }    
    }    
}    

Function DetectRestart {
    $OS = Get-WmiObject win32_operatingsystem -ComputerName $Env:ComputerName -ErrorAction SilentlyContinue
    $LastBoot = [DateTime]$OS.ConvertToDateTime($OS.LastBootUpTime)
    $TimeNow = (Get-Date)
    $UpTime = New-TimeSpan -End $TimeNow -Start $LastBoot
    If (($UpTime.Days -lt 1) -and ($UpTime.Hours -lt 1) -and ($UpTime.Minutes -le 30)){     #--[ If last restart was less than 30 minutes ago start a status run ]--
        $Script:Status = $true             
        If ($Script:Console){
            Write-Host "- Detected a recent restart..." -ForegroundColor cyan
            Write-Host "-- Last boot : "$LastBoot -foregroundcolor Yellow
            Write-Host "-- Uptime : "$UpTime.Days" Days "$UpTime.Hours" Hours "$UpTime.Minutes" Minutes" -ForegroundColor Yellow
            Write-Host '-- Setting "status" mode...' -ForegroundColor Yellow
        }    
    }
   
    If ($UpTimeCheck -and ($UpTime.Days -ge $Script:UpTimeDays)){        #--[ A secondary check. If system restart has exceeded 5 days this optional email may be sent to warn admins ]--
        If ($Script:Console){Write-Host `n"--- No reboot has occurred in over "$UpTime.Days" ---..."`n -ForegroundColor red}
        $Script:MessageBody += "<br><br>WARNING: Computer $Script:ThisComputer has exceeded the reboot window of $Script:UpTimeDays days.<br>Please investigate or disable this feature of the AUTOUPDATE script.<br><br>"
    }
}

Function DeployScript{    #--[ Copies this script to all other systems noted in the config file. ]--
    if ($Script:Console){Write-Host '--[ Deploying Updated Script to Peer Hosts ]------------' -ForegroundColor Yellow}
    foreach ($Target in $Script:PeerList.Split(",")){
    if ($Script:Console){Write-Host `n'--[ Deploying to'($Target.ToUpper())' ]-------------------------' -ForegroundColor Cyan} 
        if ($Target -ne $ThisComputer){
            if (Test-Path "\\$Target\c$\scripts\unattendedupdate.ps1"){
                if ($Script:Console){Write-Host " -- Existing files found..." -ForegroundColor green 
                    try{
                        Get-ChildItem -Path "\\$Target\c$\scripts\" | where{$_.Name -match "unattendedupdate.*"} | Remove-Item -Force:$true -Confirm:$false
                    }catch{
                        if ($Script:Console){Write-Host " -- File delete on $Target FAILED..." -ForegroundColor Red}
                        if ($Script:Console){Write-Host " -- Error Message = "$_.Exception.Message}
                        if ($Script:Console){Write-Host " -- Error Item = "$_.Exception.ItemName}
                        break 
                    }    
                }
        
                if (!(Test-Path "\\$Target\c$\scripts\unattendedupdate.ps1")){
                    if ($Script:Console){Write-Host " -- Deletion validated. Files no longer detected..." -ForegroundColor green}
                }else{
                    if ($Script:Console){Write-Host " -- Deletion FAILED..." -ForegroundColor red}
                }
    
                try{
                    Copy-Item -Path $Script:ScriptFullPath -Destination "\\$Target\c$\scripts\" -Force -Confirm:$false
                    Copy-Item -Path ($PSScriptRoot+'\'+$Script:ScriptName.split('.')[0]+'.xml') -Destination "\\$Target\c$\scripts\" -Force -Confirm:$false
                }catch{
                    if ($Script:Console){Write-Host " -- File copy to $Target FAILED..." -ForegroundColor Red}
                    if ($Script:Console){Write-Host " -- Error Message = "$_.Exception.Message}
                    if ($Script:Console){Write-Host " -- Error Item = "$_.Exception.ItemName}
                    break 
                }

                if (Test-Path "\\$Target\c$\scripts\unattendedupdate.ps1"){
                    if ($Script:Console){Write-Host " -- Verified PS1 copy to $Target..." -ForegroundColor Green}
                }
                if (Test-Path "\\$Target\c$\scripts\unattendedupdate.xml"){
                    if ($Script:Console){Write-Host " -- Verified XML copy to $Target..." -ForegroundColor Green}
                }
            }
        }else{
            if ($Script:Console){Write-Host ' -- Bypassing local system'($Target.ToUpper()) -ForegroundColor yellow}        
        }    
    }
    if ($Script:Console){Write-Host `n"--- COMPLETED ---" -ForegroundColor Red }
    break
}

#==[ End of Functions / Start of Main Process ]===============================================
if ($Script:Console){Write-Host `n"--[ Beginning Run ]-----------------------------------`n" -ForegroundColor cyan }
LoadConfig                                               #--[ Load the external config file ]--
if ($Script:Deploy){DeployScript}                        #--[ Check for a script deployment command, then exit ]--
ScheduledTask "UnattendedUpdate-$Script:ScriptVer"       #--[ Check for existance of scheduled task named for current version, create if missing ]--
DetectRestart                                            #--[ Check for last restart to determin RUN or STATUS mode ]--

#--[ Create header for html output file ]--
$Script:MessageBody += '
<style Type="text/css">
    table.myTable { border:5px solid black;border-collapse:collapse; }
    table.myTable td { border:2px solid black;padding:5px}
    table.myTable th { border:2px solid black;padding:5px;background: #949494 }
    table.bottomBorder { border-collapse:collapse; }
    table.bottomBorder td, table.bottomBorder th { border-bottom:1px dotted black;padding:5px; }
    tr.noBorder td {border: 0; }
</style>'

$Script:MessageBody += 
'<table class="myTable">
    <tr class="noBorder"><td colspan=5><center><h1>- ' + $Script:eMailSubject + ' -</h1></td></tr>
    <tr class="noBorder"><td colspan=5><center>The following report displays all recently installed patches on the system.</center></td></tr>
    <tr class="noBorder"><td colspan=5></tr>
    <tr class="noBorder"><td colspan=5>- Script Executed by: ' + $Script:UserContext.Name + '</tr>
    <tr class="noBorder"><td colspan=5>- Script Version : ' + $Script:ScriptVer + '</tr><br>
'

$Script:MessageBody += $Script:TaskMessage      

if ($Script:Status){ 
    if ($Script:Console){Write-Host "- Collecting results..." -ForegroundColor Cyan }
    $Script:MessageBody += '<tr class="noBorder"><td colspan=5>- Collecting results....</tr><br>'
    GetResults
}elseif (($Script:Manual) -or ($Script:RunDays -Match $Today)){   #--[ Update routine is governed by the week days noted in the config file ]--
    if(Test-Path $PSScriptRoot\Results.csv){Remove-Item -path $PSScriptRoot\Results.csv -confirm:$false }
    if ($Script:Console){Write-Host "- Running update routine..." -ForegroundColor Cyan }
    $Script:MessageBody += "- Running update routine....<br>"
    PatchIt
}else{
    if ($Script:Console){Write-Host "`n-- Nothing scheduled for today --`n" -ForegroundColor Cyan }
}

if ($Script:Console){Write-Host "`n--- COMPLETED ---" -ForegroundColor Red }


<#==================================================================================================
#--[ Sample XML config file. Should use the same name as this script and be in the same folder. ]--
 
<!-- Settings & configuration file -->
<Settings>
    <General>
        <ReportName>Patch Processing</ReportName>
        <PeerList>server1,server2,server3,server4</PeerList>
        <RunDays>Monday,Wednesday,Friday</RunDays>
        <RunTime>1am</RunTime>
        <UpTimeDays>5</UpTimeDays>
        <UpdateSource></UpdateSource> <!-- Set to wsus for internal. anything else for internet -->
    </General>
    <Email>
        <from>UnattendedUpdate@domain.com</from>
        <To>you@domain.com</To>
        <Debug>you@domain.com</Debug>
        <Subject>Automated Patch Processing</Subject>
        <HTML>$true</HTML>
        <SmtpServer>10.10.10.1</SmtpServer>
    </Email>
    <Credentials>
        <UserName>domain\serviceaccount</UserName>
        <Password>76492d1ws656ertg116743a534AGIATbuTeJ7I7MAAwADhH087hnA25MgB8AHIAegB2AHUTGYT6ghjYAZQAxAGQAZANgBiADA8gBaAHwAYwAzADQANgA0AGEAANwBkADEANAA4AGQAZgA3ADIAYQAwA8gBaAHwAYwAzADQANgA0AGEADYAZAA3AGUAZgBkAGYAZAA=</Password>
        <Key>kdhCO+HCvL87nsdXN0E6/AWnHhQAZgB7812qh7IObie8mE=</Key>
    </Credentials>
</Settings>