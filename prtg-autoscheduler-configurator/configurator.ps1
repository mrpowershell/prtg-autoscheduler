
<#
The Programm requires the following Modules:

Install-Package PrtgAPI
Install-Package ImportExcel

#>


#### GLOBAL VARS ####


$excelpath = "c:\temp\scheduler_config.xlsx"

<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    PRTG Scheduler GUI V2
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$form_PRTGSCHEDULER              = New-Object system.Windows.Forms.Form
$form_PRTGSCHEDULER.ClientSize   = New-Object System.Drawing.Point(788,637)
$form_PRTGSCHEDULER.text         = "PRTG Scheduler GUI"
$form_PRTGSCHEDULER.TopMost      = $false

$l_selectscheduler               = New-Object system.Windows.Forms.Label
$l_selectscheduler.text          = "Scheduler auswählen"
$l_selectscheduler.AutoSize      = $true
$l_selectscheduler.width         = 25
$l_selectscheduler.height        = 10
$l_selectscheduler.location      = New-Object System.Drawing.Point(15,26)
$l_selectscheduler.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lb_mo_to_fr                     = New-Object system.Windows.Forms.ListBox
$lb_mo_to_fr.text                = "listBox"
$lb_mo_to_fr.width               = 80
$lb_mo_to_fr.height              = 337
@('00:00','01:00','02:00','03:00','04:00','05:00','06:00','07:00','08:00','09:00','10:00','11:00','12:00','13:00','14:00','15:00','16:00','17:00','18:00','19:00','20:00','21:00','22:00','23:00') | ForEach-Object {[void] $lb_mo_to_fr.Items.Add($_)}
$lb_mo_to_fr.location            = New-Object System.Drawing.Point(113,125)
$lb_mo_to_fr.SelectionMode       = 'MultiExtended'

$l_optionchoose                  = New-Object system.Windows.Forms.Label
$l_optionchoose.text             = "Option wählen"
$l_optionchoose.AutoSize         = $true
$l_optionchoose.width            = 25
$l_optionchoose.height           = 10
$l_optionchoose.location         = New-Object System.Drawing.Point(418,26)
$l_optionchoose.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lb_sa_to_su                     = New-Object system.Windows.Forms.ListBox
$lb_sa_to_su.text                = "listBox"
$lb_sa_to_su.width               = 80
$lb_sa_to_su.height              = 337
@('00:00','01:00','02:00','03:00','04:00','05:00','06:00','07:00','08:00','09:00','10:00','11:00','12:00','13:00','14:00','15:00','16:00','17:00','18:00','19:00','20:00','21:00','22:00','23:00') | ForEach-Object {[void] $lb_sa_to_su.Items.Add($_)}
$lb_sa_to_su.location            = New-Object System.Drawing.Point(206,125)
$lb_sa_to_su.SelectionMode       = 'MultiExtended'

$cb_action                       = New-Object system.Windows.Forms.ComboBox
$cb_action.text                  = "Aktion wählen"
$cb_action.width                 = 155
$cb_action.height                = 20
@('Standard editieren','Feiertag editieren','Custom Eintrag editieren') | ForEach-Object {[void] $cb_action.Items.Add($_)}
$cb_action.location              = New-Object System.Drawing.Point(514,24)
$cb_action.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$cb_scheduler                    = New-Object system.Windows.Forms.ComboBox
$cb_scheduler.text               = "Scheduler auswählen"
$cb_scheduler.width              = 241
$cb_scheduler.height             = 20
$cb_scheduler.location           = New-Object System.Drawing.Point(158,24)
$cb_scheduler.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$b_save_settings                 = New-Object system.Windows.Forms.Button
$b_save_settings.text            = "Einstellung übernehmen"
$b_save_settings.width           = 169
$b_save_settings.height          = 30
$b_save_settings.location        = New-Object System.Drawing.Point(114,488)
$b_save_settings.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$b_load_option                   = New-Object system.Windows.Forms.Button
$b_load_option.text              = "Option laden"
$b_load_option.width             = 154
$b_load_option.height            = 30
$b_load_option.location          = New-Object System.Drawing.Point(515,51)
$b_load_option.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$l_mo_to_fr_or_custom            = New-Object system.Windows.Forms.Label
$l_mo_to_fr_or_custom.text       = "MO-FR"
$l_mo_to_fr_or_custom.AutoSize   = $true
$l_mo_to_fr_or_custom.width      = 25
$l_mo_to_fr_or_custom.height     = 10
$l_mo_to_fr_or_custom.location   = New-Object System.Drawing.Point(130,96)
$l_mo_to_fr_or_custom.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


$l_sa_to_su                      = New-Object system.Windows.Forms.Label
$l_sa_to_su.text                 = "SA-SO"
$l_sa_to_su.AutoSize             = $true
$l_sa_to_su.width                = 25
$l_sa_to_su.height               = 10
$l_sa_to_su.location             = New-Object System.Drawing.Point(221,96)
$l_sa_to_su.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lb_specialday                   = New-Object system.Windows.Forms.ListBox
$lb_specialday.text              = "listBox"
$lb_specialday.width             = 80
$lb_specialday.height            = 337
@('00:00','01:00','02:00','03:00','04:00','05:00','06:00','07:00','08:00','09:00','10:00','11:00','12:00','13:00','14:00','15:00','16:00','17:00','18:00','19:00','20:00','21:00','22:00','23:00') | ForEach-Object {[void] $lb_specialday.Items.Add($_)}
$lb_specialday.location          = New-Object System.Drawing.Point(21,125)
$lb_specialday.Visible           = $false
$lb_specialday.SelectionMode     = 'MultiExtended'

$l_specialday                    = New-Object system.Windows.Forms.Label
$l_specialday.text               = "Feiertag"
$l_specialday.AutoSize           = $true
$l_specialday.width              = 25
$l_specialday.height             = 10
$l_specialday.location           = New-Object System.Drawing.Point(35,96)
$l_specialday.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$l_specialday.Visible            = $false

$cb_template                     = New-Object system.Windows.Forms.ComboBox
$cb_template.text                = "Template wählen"
$cb_template.width               = 155
$cb_template.height              = 20
@('Feiertag','Custom') | ForEach-Object {[void] $cb_template.Items.Add($_)}
$cb_template.location            = New-Object System.Drawing.Point(514,125)
$cb_template.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$b_save_date                     = New-Object system.Windows.Forms.Button
$b_save_date.text                = "Einstellung speichern"
$b_save_date.width               = 154
$b_save_date.height              = 30
$b_save_date.location            = New-Object System.Drawing.Point(515,153)
$b_save_date.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$calendar = New-Object Windows.Forms.MonthCalendar -Property @{
    ShowTodayCircle   = $true
    MaxSelectionCount = 1
    location = New-Object System.Drawing.Point(310,124)
}

$form_PRTGSCHEDULER.controls.AddRange(@($b_save_date,$cb_template,$l_specialday,$lb_specialday,$calendar,$l_selectscheduler,$lb_mo_to_fr,$l_optionchoose,$lb_sa_to_su,$cb_action,$cb_scheduler,$b_save_settings,$b_load_option,$l_mo_to_fr_or_custom,$l_sa_to_su))

#END OF GUICODE

################  Steps before the GUI loads  ################

<#
    Connect to PRTG Server with Passhash Methode, 
    -Ignoressl is for testing, without it the connection will fail if
    the trust fails
#>

if(!(Get-PrtgClient))
{
    Connect-PrtgServer 127.0.0.1 (New-Credential prtgadmin 3477709542) -PassHash -IgnoreSSL -Force
}

#Filling Scheduler Selection in GUI Form

foreach ($schedule in (Get-PrtgSchedule))
{
    $cb_scheduler.Items.Add($schedule)
   
}






################  MAIN FUNCTIONS  ################

function clearselections()
{
    #clearing Selections
    $lb_mo_to_fr.ClearSelected()   #Creates an error but works, ignore atm
    $lb_SA_to_SU.ClearSelected()   #Creates an error but works, ignore atm
    $lb_specialday.ClearSelected() #Creates an error but works, ignore atm
}

function updateview($setting) #Updates the GUI depending on the current option
{
    if ($setting -eq "0")
    {
        $l_specialday.text = "Feiertag"
        $l_specialday.Visible = $true
        $lb_specialday.Visible = $true
        $lb_sa_to_su.Visible = $false
        $lb_MO_to_FR.Visible = $false
        $l_mo_to_fr_or_custom.Visible = $false
        $l_sa_to_su.Visible = $false
    }
    if ($setting -eq "1")
    {
        $l_specialday.Visible = $false
        $lb_specialday.Visible = $false
        $lb_sa_to_su.Visible = $true
        $lb_MO_to_FR.Visible =$true
        $l_mo_to_fr_or_custom.Visible = $true
        $l_sa_to_su.Visible = $true
    }
    if ($setting -eq "2")
    {
        $l_specialday.Text = "Custom Eintrag"
        $l_specialday.Visible = $true
        $lb_specialday.Visible = $true
        $lb_sa_to_su.Visible = $false
        $lb_MO_to_FR.Visible = $false
        $l_mo_to_fr_or_custom.Visible = $false
        $l_sa_to_su.Visible = $false
    }
}


function setmarks()
{
    #Load the PRTG Defaults
    if ($cb_action.SelectedItem -eq "Standard editieren")
    {
        updateview(1)
        $exceldata = Import-Excel -WorksheetName PRTGDEFAULT_SETTING -path c:\temp\scheduler_config.xlsx
        clearselections

        #Select the Hours according to the Values in the Excel
        foreach ($entry in $exceldata)
        {
            if ($entry.MOFR -eq "1")
            {
                $lb_mo_to_fr.SelectedItem = $entry.Hour

            }
            if ($entry.SASO -eq "1")
            {
                $lb_sa_to_su.SelectedItem = $entry.Hour

            }
        }

    }

    #Load Selections for the Special Days
    if ($cb_action.SelectedItem -eq "Feiertag editieren")
    {
        clearselections
        updateview(0)

        $exceldata = Import-Excel -WorksheetName PRTGSPECIALDAY_SETTING -path c:\temp\scheduler_config.xlsx
        
        #Select the Hours according to the Values in the Excel
        foreach ($entry in $exceldata)
        {
            if ($entry.ENABLED -eq "1")
            {
                $lb_specialday.SelectedItem = $entry.HOUR
            }
        }
    }

    if ($cb_action.SelectedItem -eq "Custom Eintrag editieren")
    {
        clearselections
        updateview(2)

        $exceldata = Import-Excel -WorksheetName CUSTOM -path c:\temp\scheduler_config.xlsx
        
        #Select the Hours according to the Values in the Excel
        foreach ($entry in $exceldata)
        {
            if ($entry.ENABLED -eq "1")
            {
                $lb_specialday.SelectedItem = $entry.HOUR
            }
        }
    }
     
    
    
}

function writebacktoexcel() #Writes back the Settings to the Excel according to the Userinput
{
    if ($cb_action.SelectedItem -eq "Standard editieren")
    {
        $exceldata = Import-Excel -WorksheetName PRTGDEFAULT_SETTING -path c:\temp\scheduler_config.xlsx

        #Get Userselection from GUI
        $selectionsMOFR = $lb_mo_to_fr.SelectedItems
        $selectionsSASO = $lb_sa_to_su.SelectedItems

        foreach ($entry in $exceldata)
        {
            $searchresultMOFR = $lb_mo_to_fr.SelectedItems | Select-String -Pattern $entry.Hour #Check if selection matches
            $searchresultSASU = $lb_sa_to_su.SelectedItems | Select-String -Pattern $entry.Hour 

            #Writeback data to Excelobject
            if ($searchresultSASU -ne $null)
            {
                $entry.SASO = "1"
            }
            else
            {
                $entry.SASO = "0"
            }
            if ($searchresultMOFR -ne $null)
            {
                $entry.MOFR = "1"
            }
            else
            {
                $entry.MOFR = "0"
            }
            
        }
        Export-Excel -inputobject $exceldata -WorksheetName PRTGDEFAULT_SETTING -path c:\temp\scheduler_config.xlsx   
    }

    if ($cb_action.SelectedItem -eq "Feiertag editieren" -or $cb_action.SelectedItem -eq "Custom Eintrag editieren") #Same routine for both as they are similiar
    {
        if ($cb_action.SelectedItem -eq "Feiertag editieren" ){ $exceldata = Import-Excel -WorksheetName PRTGSPECIALDAY_SETTING -path c:\temp\scheduler_config.xlsx }
        if ($cb_action.SelectedItem -eq "Custom Eintrag editieren" ){ $exceldata = Import-Excel -WorksheetName CUSTOM -path c:\temp\scheduler_config.xlsx }

        #Get Userselection from GUI
        $selections_sd = $lb_specialday.SelectedItems


        foreach ($entry in $exceldata)
        {
            $searchresultsd = $lb_specialday.SelectedItems | Select-String -Pattern $entry.Hour #Check if selection matches

            #Writeback data to Excelobject
            if ($searchresultsd  -ne $null)
            {
                $entry.ENABLED = "1"
            }
            else
            {
                $entry.ENABLED = "0"
            }
            
        }

        if ($cb_action.SelectedItem -eq "Feiertag editieren" ){ Export-Excel -inputobject $exceldata -WorksheetName PRTGSPECIALDAY_SETTING -path c:\temp\scheduler_config.xlsx }
        if ($cb_action.SelectedItem -eq "Custom Eintrag editieren" ){ Export-Excel -inputobject $exceldata -WorksheetName CUSTOM -path c:\temp\scheduler_config.xlsx } 
    }    
}

#Function to Save custom Dates and the Settings to Excel
function savedate($date)
{
    $exceldata = Import-Excel -WorksheetName SPECIALDAYS -path c:\temp\scheduler_config.xlsx
    $search = $exceldata | Where-Object {$_.DATE -eq $date}

    if (!$search)
    {
        $exceldata = $exceldata + @{DATE="$date";CUSTOMSETTINGS="0"}
        $search = $exceldata | Where-Object {$_.DATE -eq $date}
        if ($cb_template.SelectedItem -eq "Feiertag") {$search.CUSTOMSETTINGS = 1}
        if ($cb_template.SelectedItem -eq "Custom") {$search.CUSTOMSETTINGS = 2}
    }
    if ($cb_template.SelectedItem -eq "Feiertag") {$search.CUSTOMSETTINGS = 1}
    if ($cb_template.SelectedItem -eq "Custom") {$search.CUSTOMSETTINGS = 2}

    Export-Excel -inputobject $exceldata -WorksheetName SPECIALDAYS -path c:\temp\scheduler_config.xlsx

}

################  BUTTON FUNCTIONS  ################$

$b_load_option.Add_Click(
{
    if ($cb_action.SelectedItem -eq "Standard editieren")
    {
        setmarks
    }
    if ($cb_action.SelectedItem -eq "Feiertag editieren")
    {
        setmarks
    }
    if ($cb_action.SelectedItem -eq "Custom Eintrag editieren")
    {
        setmarks
    }    

    
})

$b_save_settings.Add_Click(
{
    writebacktoexcel
})

$b_save_date.Add_Click(
{
    $dateinput = $calendar.SelectionStart
    $dateshort = $dateinput.ToShortDateString()
    savedate($dateshort)
    
})






#Show the GUI after the inital code
[void]$form_PRTGSCHEDULER.ShowDialog()