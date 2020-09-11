# PRTG Autoscheduler

The PRTG Auto Scheduler is a utility program for the monitoring tool PRTG. It automatically updates the set schedule according to predefined templates.

  - Holiday settings are applied automatically
  - Custom days are applied automatically
  - Standard template definable for during the week and for the weekend

#### Program Components

  - PRTG-autoscheduler-runner -> Powershell script that runs daily
  - PRTG-autoscheduler-configurator -> GUI tool to configure the Runner
  - PRTG-autoscheduler-config.xlsx -> Excel where the Settings are stored, can be updated manually


#### Important Information

The formatting in Excel must not be adjusted, otherwise the Runner can no longer read the values.

- The Runner must run on a server that has access to PRTG.
- The automated task should run at 00:01AM.
- The [ImportExcel](https://github.com/dfinke/ImportExcel) module is required.
- The [PRTGAPI](https://github.com/lordmilko/PrtgAPI) module is required.


# Instructions and Setup

On the server where the Runner should run, the module ImportExcel and the module PrtgAPI must be installed via Powershell. Use the following commands in Powershell (as admin)

    Install-Package PrtgAPI
    Install-Package ImportExcel

On the server where the Runner should run, create a folder PRTG-Tool on C:\ and copy the Runner prtg-autoscheduler-01.ps1 into this folder.

It is best to do this on the same server where a Runner is already running. Copy the configuration Excel into a subfolder CONFIG and rename it logically like the Runner: prtg-schedulename-config.xlsx
![Layout](https://github.com/mrpowershell/prtg-autoscheduler/raw/master/Images/Layout2.png)
