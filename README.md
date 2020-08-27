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
