<#
The Programm requires the following Modules:

Install-Package PrtgAPI
Install-Package ImportExcel

This programm is supposed to run daily at 00:01AM to update the schedule in PRTG
#>


#### GLOBAL VARS ####

#Enter the Path of the configuration excel
$excelpath = "c:\temp\scheduler_config.xlsx"

#define the ScheduleID which should be updated
$prtg_scheduleid = "2226"

#Define the Servername, User and the passhash which the programm uses to connect to PRTG
$prtgserver = "127.0.0.1"
$prtguser="prtgadmin"
$prtguserhash = "3477709542"


################  MAIN CODE, DO NOT EDIT THE CODE  ################

$prtg_reference = Import-Excel -WorksheetName REFERENCE -path $excelpath
$prtg_custom = Import-Excel -WorksheetName CUSTOM -path $excelpath
$prtg_specialdaysetting = Import-Excel -WorksheetName PRTGSPECIALDAY_SETTING -path $excelpath
$prtg_default = Import-Excel -WorksheetName PRTGDEFAULT_SETTING -path $excelpath
$specialdays = Import-Excel -WorksheetName SPECIALDAYS -path $excelpath


function getprtgvars($day,$daysetting)
{
    if ($daysetting -eq 0)
    {
        foreach ($entry in $prtg_default)
        {

            if ($day -ne "Saturday" -or $day -ne "Sunday" -and $entry.MOFR -eq "1")
            {          
                if ($runner -eq $null) {[string]$runner = $prtg_reference | where Hour -eq $entry.Hour | Select-Object -Property $day -ExpandProperty $day}
                else {$runner = "$runner" + "," +($prtg_reference | where Hour -eq $entry.Hour | Select-Object -Property $day -ExpandProperty $day)}  
            }
            if ($day -eq "Saturday" -or $day -eq "Sunday" -and $entry.SASO -eq "1")
            {
                if ($runner -eq $null) {[string]$runner = $prtg_reference | where Hour -eq $entry.Hour | Select-Object -Property $day -ExpandProperty $day}
                else {$runner = "$runner" + "," +($prtg_reference | where Hour -eq $entry.Hour | Select-Object -Property $day -ExpandProperty $day)}              

            }
        }

    }
    if ($daysetting -eq 2)
    {
        foreach ($entry in $prtg_custom)
        {
            if ($entry.Enabled -eq "1")
            {
                if ($runner -eq $null) {[string]$runner = $prtg_reference | where Hour -eq $entry.Hour | Select-Object -Property $day -ExpandProperty $day}
                else {$runner = "$runner" + "," +($prtg_reference | where Hour -eq $entry.Hour | Select-Object -Property $day -ExpandProperty $day)}            

            }

        }
    }
    if ($daysetting -eq 1)
    {
        foreach ($entry in $prtg_specialdaysetting)
        {

            if ($day -ne "Saturday" -or $day -ne "Sunday" -and $entry.Enabled -eq "1")
            {
                if ($runner -eq $null) {[string]$runner = $prtg_reference | where Hour -eq $entry.Hour | Select-Object -Property $day -ExpandProperty $day}
                else {$runner = "$runner" + "," +($prtg_reference | where Hour -eq $entry.Hour | Select-Object -Property $day -ExpandProperty $day)}            

            }
            if ($day -eq "Saturday" -or $day -eq "Sunday" -and $entry.Enabled -eq "1")
            {
                if ($runner -eq $null) {[string]$runner = $prtg_reference | where Hour -eq $entry.Hour | Select-Object -Property $day -ExpandProperty $day}
                else {$runner = "$runner" + "," +($prtg_reference | where Hour -eq $entry.Hour | Select-Object -Property $day -ExpandProperty $day)}             

            }
        }
    }
    
    return $runner

}

<#
    Connect to PRTG Server with Passhash Methode, 
    -Ignoressl is for testing, without it the connection will fail if
    the trust fails
#>

if(!(Get-PrtgClient))
{
    Connect-PrtgServer $prtgserver (New-Credential $prtguser $prtguserhash) -PassHash -IgnoreSSL -Force
}

#Saving the Scheduler Ids

$schedulers = Get-PrtgSchedule

#Getting Date Configuration, the hard way :)

$currentdayname = get-date -UFormat %A
$currentdate = get-date -Uformat %d.%m.%Y

#No lets get the dates of the other days

if ($currentdayname -eq "Montag")
{
    $MO = $currentdate
    $DI = (get-date).AddDays(1).ToString("dd.MM.yyy")
    $MI = (get-date).AddDays(2).ToString("dd.MM.yyy")
    $DO = (get-date).AddDays(3).ToString("dd.MM.yyy")
    $FR = (get-date).AddDays(4).ToString("dd.MM.yyy")
    $SA = (get-date).AddDays(5).ToString("dd.MM.yyy")
    $SO = (get-date).AddDays(6).ToString("dd.MM.yyy")
}
if ($currentdayname -eq "Dienstag")
{
    $MO = (get-date).AddDays(-1).ToString("dd.MM.yyy")
    $DI = $currentdate
    $MI = (get-date).AddDays(1).ToString("dd.MM.yyy")
    $DO = (get-date).AddDays(2).ToString("dd.MM.yyy")
    $FR = (get-date).AddDays(3).ToString("dd.MM.yyy")
    $SA = (get-date).AddDays(4).ToString("dd.MM.yyy")
    $SO = (get-date).AddDays(5).ToString("dd.MM.yyy")
}
if ($currentdayname -eq "Mittwoch")
{
    $MO = (get-date).AddDays(-2).ToString("dd.MM.yyy")
    $DI = (get-date).AddDays(-1).ToString("dd.MM.yyy")
    $MI = $currentdate
    $DO = (get-date).AddDays(1).ToString("dd.MM.yyy")
    $FR = (get-date).AddDays(2).ToString("dd.MM.yyy")
    $SA = (get-date).AddDays(3).ToString("dd.MM.yyy")
    $SO = (get-date).AddDays(4).ToString("dd.MM.yyy")
}
if ($currentdayname -eq "Donnerstag")
{
    $MO = (get-date).AddDays(-3).ToString("dd.MM.yyy")
    $DI = (get-date).AddDays(-2).ToString("dd.MM.yyy")
    $MI = (get-date).AddDays(-1).ToString("dd.MM.yyy")
    $DO = $currentdate
    $FR = (get-date).AddDays(1).ToString("dd.MM.yyy")
    $SA = (get-date).AddDays(2).ToString("dd.MM.yyy")
    $SO = (get-date).AddDays(3).ToString("dd.MM.yyy")
}
if ($currentdayname -eq "Freitag")
{
    $MO = (get-date).AddDays(-4).ToString("dd.MM.yyy")
    $DI = (get-date).AddDays(-3).ToString("dd.MM.yyy")
    $MI = (get-date).AddDays(-2).ToString("dd.MM.yyy")
    $DO = (get-date).AddDays(-1).ToString("dd.MM.yyy")
    $FR = $currentdate
    $SA = (get-date).AddDays(1).ToString("dd.MM.yyy")
    $SO = (get-date).AddDays(2).ToString("dd.MM.yyy")
}
if ($currentdayname -eq "Samstag")
{
    $MO = (get-date).AddDays(-5).ToString("dd.MM.yyy")
    $DI = (get-date).AddDays(-4).ToString("dd.MM.yyy")
    $MI = (get-date).AddDays(-3).ToString("dd.MM.yyy")
    $DO = (get-date).AddDays(-2).ToString("dd.MM.yyy")
    $FR = (get-date).AddDays(-1).ToString("dd.MM.yyy")
    $SA = $currentdate
    $SO = (get-date).AddDays(1).ToString("dd.MM.yyy")
}
if ($currentdayname -eq "Sonntag")
{
    $MO = (get-date).AddDays(-6).ToString("dd.MM.yyy")
    $DI = (get-date).AddDays(-5).ToString("dd.MM.yyy")
    $MI = (get-date).AddDays(-4).ToString("dd.MM.yyy")
    $DO = (get-date).AddDays(-3).ToString("dd.MM.yyy")
    $FR = (get-date).AddDays(-2).ToString("dd.MM.yyy")
    $SA = (get-date).AddDays(-1).ToString("dd.MM.yyy")
    $SO = $currentdate
}

#Define which Template needs to be assigned

if ($specialdays -match $mo)
{
    $temp = $specialdays | where DATE -eq $mo | Select-Object -Property CUSTOMSETTINGS -ExpandProperty CUSTOMSETTINGS
    $schedulevalues = (getprtgvars "Monday" "$temp")
}else {$schedulevalues = getprtgvars "Monday" "0"}

if ($specialdays -match $di)
{
    $temp = $specialdays | where DATE -eq $di | Select-Object -Property CUSTOMSETTINGS -ExpandProperty CUSTOMSETTINGS
    $schedulevalues = "$schedulevalues" + "," + (getprtgvars "Tuesday" "$temp")

}else {$schedulevalues = "$schedulevalues" + "," + (getprtgvars "Tuesday" "0")}

if ($specialdays -match $mi)
{
    $temp = $specialdays | where DATE -eq $mi | Select-Object -Property CUSTOMSETTINGS -ExpandProperty CUSTOMSETTINGS
    $schedulevalues = "$schedulevalues" + "," + (getprtgvars "Wednesday" $temp) 
}else {$schedulevalues = "$schedulevalues" + "," + (getprtgvars "Wednesday" "0")}

if ($specialdays -match $do)
{
    $temp = $specialdays | where DATE -eq $do | Select-Object -Property CUSTOMSETTINGS -ExpandProperty CUSTOMSETTINGS
    $schedulevalues = "$schedulevalues" + "," + (getprtgvars "Thursday" $temp) 
}else {$schedulevalues = "$schedulevalues" + "," + (getprtgvars "Thursday" "0")}

if ($specialdays -match $fr)
{
    $temp = $specialdays | where DATE -eq $fr | Select-Object -Property CUSTOMSETTINGS -ExpandProperty CUSTOMSETTINGS
    $schedulevalues = "$schedulevalues" + "," + (getprtgvars "Friday" $temp) 
}else {$schedulevalues = "$schedulevalues" + "," + (getprtgvars "Friday" "0")}

if ($specialdays -match $sa)
{
    $temp = $specialdays | where DATE -eq $sa | Select-Object -Property CUSTOMSETTINGS -ExpandProperty CUSTOMSETTINGS
    $schedulevalues = "$schedulevalues" + "," + (getprtgvars "Saturday" $temp) 
}else {$schedulevalues = "$schedulevalues" + "," + (getprtgvars "Saturday" "0")}

if ($specialdays -match $so)
{
    $temp = $specialdays | where DATE -eq $so | Select-Object -Property CUSTOMSETTINGS -ExpandProperty CUSTOMSETTINGS
    $schedulevalues = "$schedulevalues" + "," + (getprtgvars "Sunday" $temp) 
}else {$schedulevalues = "$schedulevalues" + "," + (getprtgvars "Sunday" "0")}


#Converting String to Array and to INT which is needed to set the PRTG Schedule
[array]$schedulevalues = "$schedulevalues".split(",")
[array]$schedulevaluesINT = foreach($number in $schedulevalues) {([int]::parse($number))}


#Updating the PRTG Schedule
Get-PrtgSchedule -Id $prtg_scheduleid | Set-ObjectProperty -Force -RawParameters @{
    timetable = $schedulevaluesINT
    timetable_ = ""
}
