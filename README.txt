# UnattendedUpdate
Automatically applies current patches to a single Windows system, then reboots.


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
