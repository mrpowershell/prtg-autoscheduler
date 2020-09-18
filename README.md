
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
![layout](https://github.com/mrpowershell/prtg-autoscheduler/raw/master/Images/Layout2.png)

Next you need to open the Runner ps1 file and edit the Settings there. i would recommend to create a new user for the PRTG Autoscheduler. The user needs admin privileges to edit the schedule(s). Note: For the PRTGSERVER do not enter the IP, enter the FQDN like https://monitoring.company.com If your PRTG Server has no valid SSL Certificate, check the next step.

![runner](https://github.com/mrpowershell/prtg-autoscheduler/raw/master/Images/runner_configuration.png)

**If you have no valid SSL in PRTG:**

Search for the following line in the Runner Script:

    Connect-PrtgServer $prtgserver (New-Credential $prtguser $prtguserhash) -PassHash -Force

Edit it to the following:

    Connect-PrtgServer $prtgserver (New-Credential $prtguser $prtguserhash) -PassHash -Force -IgnoreSSL

To get the ScheduleID for the Runner, open PRTG and go to the Schedule which you want to assign to the Runner. In the URL is the RunnerID

![url](https://github.com/mrpowershell/prtg-autoscheduler/raw/master/Images/url_example.png)

Next create and schedule in Windows which runs at 00:01 every day

![autoschedule](https://github.com/mrpowershell/prtg-autoscheduler/raw/master/Images/autoschedule.png)


**Tip**: You should change the schedule name in PRTG to show that this schedule is managed by the tool and manual entries are overwritten by the Tool.

**PRTG Autoscheduler Templates **

The PRTG Auto Scheduler Runner uses predefined templates for weekdays, holidays, special days, etc. These templates can be customized with the configurator. **NOTE**: The template is used for each day where it has been defined.

To edit the default template, which is **ALWAYS** set if the day is not further defined as a holiday, custom etc., open the Configurator, import the configuration Excel for the Runner/Schedule to be edited, and select the default template:

![GUI1](https://github.com/mrpowershell/prtg-autoscheduler/raw/master/Images/GUI_1.png)

The current setting is now loaded from Excel. All blue times represent the hours that are activated. The template can now be adjusted as desired. To save the setting, click on Save Template. You can see if it was loaded correctly by clicking Load again or by opening the Excel manually.

#### Edit PRTG Autoscheduler Holiday and Custom Template(s)

To change the default holiday or custom 1 and custom 2 settings, follow the same procedure as for the standard template. Here, however, no distinction is made between MO-FR, SA-SO, only the hours. The reason for this is that these templates only refer to one day.

![GUI2](https://github.com/mrpowershell/prtg-autoscheduler/raw/master/Images/GUI_2.png)

#### PRTG Autoscheduler: Define Holidays and Special Days

**Variant 1: With the Configurator**

To set holidays or special days for the autoscheduler, open the Configurator and open the desired Excel by editing the holidays.

Then select the day you want to define in the calendar and choose the template that should be applied to the day. The entry must be saved with Apply Template. In Current configured Dates the day will now appear that was defined:

![GUI3](https://github.com/mrpowershell/prtg-autoscheduler/raw/master/Images/GUI_3.png)


**Delete entry**
 
To delete an entry or set it to default, simply select the tag and apply the default template. The tag will then disappear from the list.

**Overwrite entry** 

See Delete entry, Simply select the tag again and choose another template. 

Variant 2: Directly in Excel



