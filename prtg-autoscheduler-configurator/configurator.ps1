
<#
The Programm requires the following Modules:

Install-Package PrtgAPI
Install-Package ImportExcel

#>


#### GLOBAL VARS ####

#Define the Servername, User and the passhash which the programm uses to connect to PRTG

$prtgserver = "127.0.0.1"
$prtguser="prtgadmin"
$prtguserhash = "3477709542"


################  MAIN CODE, DO NOT EDIT THE CODE BELOW THIS LINE   ################

# This form was created using POSHGUI.com  a free online gui designer for PowerShell


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$form_PRTGSCHEDULER              = New-Object system.Windows.Forms.Form
$form_PRTGSCHEDULER.ClientSize   = New-Object System.Drawing.Point(788,637)
$form_PRTGSCHEDULER.text         = "PRTG Scheduler GUI"
$form_PRTGSCHEDULER.TopMost      = $false

$l_selectscheduler               = New-Object system.Windows.Forms.Label
$l_selectscheduler.text          = "Select Schedule:"
$l_selectscheduler.AutoSize      = $true
$l_selectscheduler.width         = 25
$l_selectscheduler.height        = 10
$l_selectscheduler.location      = New-Object System.Drawing.Point(41,26)
$l_selectscheduler.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lb_mo_to_fr                     = New-Object system.Windows.Forms.ListBox
$lb_mo_to_fr.width               = 80
$lb_mo_to_fr.height              = 337
@('00:00','01:00','02:00','03:00','04:00','05:00','06:00','07:00','08:00','09:00','10:00','11:00','12:00','13:00','14:00','15:00','16:00','17:00','18:00','19:00','20:00','21:00','22:00','23:00') | ForEach-Object {[void] $lb_mo_to_fr.Items.Add($_)}
$lb_mo_to_fr.location            = New-Object System.Drawing.Point(138,97)
$lb_mo_to_fr.SelectionMode     = 'MultiExtended'

$lb_sa_to_su                     = New-Object system.Windows.Forms.ListBox
$lb_sa_to_su.width               = 80
$lb_sa_to_su.height              = 337
@('00:00','01:00','02:00','03:00','04:00','05:00','06:00','07:00','08:00','09:00','10:00','11:00','12:00','13:00','14:00','15:00','16:00','17:00','18:00','19:00','20:00','21:00','22:00','23:00') | ForEach-Object {[void] $lb_sa_to_su.Items.Add($_)}
$lb_sa_to_su.location            = New-Object System.Drawing.Point(231,97)
$lb_sa_to_su.SelectionMode     = 'MultiExtended'

$cb_select_template              = New-Object system.Windows.Forms.ComboBox
$cb_select_template.text         = "Template"
$cb_select_template.width        = 155
$cb_select_template.height       = 20
@('Default Template','Public Holiday Template','Custom Template 1') | ForEach-Object {[void] $cb_select_template.Items.Add($_)}
$cb_select_template.location     = New-Object System.Drawing.Point(52,31)
$cb_select_template.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$cb_scheduler                    = New-Object system.Windows.Forms.ComboBox
$cb_scheduler.width              = 241
$cb_scheduler.height             = 20
$cb_scheduler.location           = New-Object System.Drawing.Point(158,24)
$cb_scheduler.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$b_save_settings                 = New-Object system.Windows.Forms.Button
$b_save_settings.text            = "Save Template"
$b_save_settings.width           = 169
$b_save_settings.height          = 30
$b_save_settings.location        = New-Object System.Drawing.Point(99,462)
$b_save_settings.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$b_load_template                 = New-Object system.Windows.Forms.Button
$b_load_template.text            = "Load"
$b_load_template.width           = 75
$b_load_template.height          = 30
$b_load_template.location        = New-Object System.Drawing.Point(234,26)
$b_load_template.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$l_mo_to_fr_or_custom            = New-Object system.Windows.Forms.Label
$l_mo_to_fr_or_custom.text       = "MO-FR"
$l_mo_to_fr_or_custom.AutoSize   = $true
$l_mo_to_fr_or_custom.width      = 25
$l_mo_to_fr_or_custom.height     = 10
$l_mo_to_fr_or_custom.location   = New-Object System.Drawing.Point(157,68)
$l_mo_to_fr_or_custom.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$l_sa_to_su                      = New-Object system.Windows.Forms.Label
$l_sa_to_su.text                 = "SA-SU"
$l_sa_to_su.AutoSize             = $true
$l_sa_to_su.width                = 25
$l_sa_to_su.height               = 10
$l_sa_to_su.location             = New-Object System.Drawing.Point(248,68)
$l_sa_to_su.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lb_specialday                   = New-Object system.Windows.Forms.ListBox
$lb_specialday.width             = 80
$lb_specialday.height            = 337
@('00:00','01:00','02:00','03:00','04:00','05:00','06:00','07:00','08:00','09:00','10:00','11:00','12:00','13:00','14:00','15:00','16:00','17:00','18:00','19:00','20:00','21:00','22:00','23:00') | ForEach-Object {[void] $lb_specialday.Items.Add($_)}
$lb_specialday.location          = New-Object System.Drawing.Point(46,97)
$lb_specialday.SelectionMode     = 'MultiExtended'

$l_specialday                    = New-Object system.Windows.Forms.Label
$l_specialday.text               = "Holiday"
$l_specialday.AutoSize           = $true
$l_specialday.width              = 25
$l_specialday.height             = 10
$l_specialday.location           = New-Object System.Drawing.Point(62,68)
$l_specialday.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$l_notification                  = New-Object system.Windows.Forms.Label
$l_notification.AutoSize         = $false
$l_notification.visible          = $false
$l_notification.width            = 25
$l_notification.height           = 20
$l_notification.Anchor           = 'left'
$l_notification.Textalign        = 'MiddleCenter'
$l_notification.Dock             = 'Bottom'
$l_notification.location         = New-Object System.Drawing.Point(363,602)
$l_notification.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$l_notification.ForeColor        = [System.Drawing.ColorTranslator]::FromHtml("#ff0000")

$gb_savetemplate                 = New-Object system.Windows.Forms.Groupbox
$gb_savetemplate.height          = 529
$gb_savetemplate.width           = 361
$gb_savetemplate.text            = "Edit Template"
$gb_savetemplate.location        = New-Object System.Drawing.Point(14,70)

$tb_excelpath                        = New-Object system.Windows.Forms.TextBox
$tb_excelpath.multiline              = $false
$tb_excelpath.width                  = 201
$tb_excelpath.height                 = 20
$tb_excelpath.location               = New-Object System.Drawing.Point(559,24)
$tb_excelpath.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$b_select_excel                  = New-Object system.Windows.Forms.Button
$b_select_excel.text             = "Select Excel"
$b_select_excel.width            = 112
$b_select_excel.height           = 30
$b_select_excel.location         = New-Object System.Drawing.Point(426,20)
$b_select_excel.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$gb_update_date                  = New-Object system.Windows.Forms.Groupbox
$gb_update_date.height           = 307
$gb_update_date.width            = 376
$gb_update_date.text             = "Apply Template to Date"
$gb_update_date.location         = New-Object System.Drawing.Point(390,70)

$cb_select_template_day          = New-Object system.Windows.Forms.ComboBox
$cb_select_template_day.text     = "Template"
$cb_select_template_day.width    = 155
$cb_select_template_day.height   = 20
@('Default Template','Public Holiday Template','Custom Template 1') | ForEach-Object {[void] $cb_select_template_day.Items.Add($_)}
$cb_select_template_day.location  = New-Object System.Drawing.Point(209,31)
$cb_select_template_day.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$b_save_date                     = New-Object system.Windows.Forms.Button
$b_save_date.text                = "Apply Template"
$b_save_date.width               = 155
$b_save_date.height              = 30
$b_save_date.location            = New-Object System.Drawing.Point(209,63)
$b_save_date.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$calendar = New-Object Windows.Forms.MonthCalendar -Property @{
    ShowTodayCircle   = $true
    MaxSelectionCount = 1
    location = New-Object System.Drawing.Point(8,30)
}

$form_PRTGSCHEDULER.controls.AddRange(@($l_selectscheduler,$cb_scheduler,$l_notification,$gb_savetemplate,$tb_excelpath,$b_select_excel,$gb_update_date))
$gb_savetemplate.controls.AddRange(@($lb_mo_to_fr,$lb_sa_to_su,$cb_select_template,$b_save_settings,$b_load_template,$l_mo_to_fr_or_custom,$l_sa_to_su,$lb_specialday,$l_specialday))
$gb_update_date.controls.AddRange(@($calendar,$cb_select_template_day,$b_save_date))


#END OF GUICODE

################  MAIN FUNCTIONS  ################

#Function to get Filepath of Excels
Function Get-FileName($initialDirectory)
{   
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
    Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.filter = "xlsx (*.xlsx)| *.xlsx"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

#Function to get and sed Excelpath from GUI
function Get-Excelpath
{
        $tb_excelpath.Text = Get-FileName
}

#Function to write an Error or Notification to the GUI
function writenotification($text)

{
    $l_notification.Visible = $true
    $l_notification.Text = $text
}

#Function to clear the Selections in the GUI
function clearselections()
{
    #clearing Selections
    $lb_mo_to_fr.ClearSelected()   #Creates an error but works, ignore atm
    $lb_SA_to_SU.ClearSelected()   #Creates an error but works, ignore atm
    $lb_specialday.ClearSelected() #Creates an error but works, ignore atm
}

#Updates the GUI depending on the current option
function updateview($setting) 
{
    if ($setting -eq "0")
    {
        $l_specialday.Visible = $false
        $lb_specialday.Visible = $false
        $lb_sa_to_su.Visible = $true
        $lb_MO_to_FR.Visible = $true
        $l_mo_to_fr_or_custom.Visible = $true
        $l_sa_to_su.Visible = $true
    }
    if ($setting -eq "1")
    {
        $l_specialday.text = "Setting per Day"
        $l_specialday.Visible = $true
        $lb_specialday.Visible = $true
        $lb_sa_to_su.Visible = $false
        $lb_MO_to_FR.Visible =$false
        $l_mo_to_fr_or_custom.Visible = $false
        $l_sa_to_su.Visible = $false
    }
}

#Function to Set the Marks according from the Options in the Excel File
function setmarks()
{
    #Load the PRTG Defaults
    if ($cb_select_template.SelectedItem -eq "Default Template")
    {
        updateview(0)
        $exceldata = Import-Excel -WorksheetName PRTGDEFAULT_SETTING -path $tb_excelpath.Text
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

    #Load Selections for the Public Holidays
    if ($cb_select_template.SelectedItem -eq "Public Holiday Template")
    {
        clearselections
        updateview(1)

        $exceldata = Import-Excel -WorksheetName PRTGSPECIALDAY_SETTING -path $tb_excelpath.Text
        
        #Select the Hours according to the Values in the Excel
        foreach ($entry in $exceldata)
        {
            if ($entry.ENABLED -eq "1")
            {
                $lb_specialday.SelectedItem = $entry.HOUR
            }
        }
    }

    if ($cb_select_template.SelectedItem -eq "Custom Template 1")
    {
        clearselections
        updateview(2)

        $exceldata = Import-Excel -WorksheetName CUSTOM -path $tb_excelpath.Text
        
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

#Writes back the Settings to the Excel according to the User selection
function writebacktoexcel() 
{
    if ($cb_select_template.SelectedItem -eq "Default Template")
    {
        $exceldata = Import-Excel -WorksheetName PRTGDEFAULT_SETTING -path $tb_excelpath.Text

        foreach ($entry in $exceldata)
        {
 
            #Searching for Selection
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
        Export-Excel -inputobject $exceldata -WorksheetName PRTGDEFAULT_SETTING -path $tb_excelpath.Text   
    }

    if ($cb_select_template.SelectedItem -eq "Public Holiday Template" -or $cb_select_template.SelectedItem -eq "Custom Template 1") #Same routine for both as they are similiar
    {
        if ($cb_select_template.SelectedItem -eq "Public Holiday Template" ){ $exceldata = Import-Excel -WorksheetName PRTGSPECIALDAY_SETTING -path $tb_excelpath.Text }
        if ($cb_select_template.SelectedItem -eq "Custom Template 1" ){ $exceldata = Import-Excel -WorksheetName CUSTOM -path $tb_excelpath.Text }

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

        if ($cb_select_template.SelectedItem -eq "Public Holiday Template" ){ Export-Excel -inputobject $exceldata -WorksheetName PRTGSPECIALDAY_SETTING -path $tb_excelpath.Text }
        if ($cb_select_template.SelectedItem -eq "Custom Template 1" ){ Export-Excel -inputobject $exceldata -WorksheetName CUSTOM -path $tb_excelpath.Text } 
    }    
}

#Function to Save custom Dates and the Settings to Excel
function savedate($date)
{
    $exceldata = Import-Excel -WorksheetName SPECIALDAYS -path $tb_excelpath.Text
    $search = $exceldata | Where-Object {$_.DATE -eq $date}

    if (!$search)
    {
        $exceldata = $exceldata + @{DATE="$date";CUSTOMSETTINGS="0"}
        $search = $exceldata | Where-Object {$_.DATE -eq $date}
        if ($cb_select_template_day.SelectedItem -eq "Public Holiday Template") {$search.CUSTOMSETTINGS = 1}
        if ($cb_select_template_day.SelectedItem -eq "Custom Template 1") {$search.CUSTOMSETTINGS = 2}
    }
    if ($cb_select_template_day.SelectedItem -eq "Public Holiday Template") {$search.CUSTOMSETTINGS = 1}
    if ($cb_select_template_day.SelectedItem -eq "Custom Template 1") {$search.CUSTOMSETTINGS = 2}
    if ($cb_select_template_day.SelectedItem -eq "Default Template") {$search.CUSTOMSETTINGS = 0}

    Export-Excel -inputobject $exceldata -WorksheetName SPECIALDAYS -path $tb_excelpath.Text

}

##### PRIMARY CODE #####

<#
    Connect to PRTG Server with Passhash Methode, 
    -Ignoressl is for testing, without it the connection will fail if
    the trust fails
#>

<# DISABELD FOR THE MOMENT, Function not yet implemented
if(!(Get-PrtgClient))
{
    Connect-PrtgServer $prtgserver (New-Credential $prtguser $prtguserhash) -PassHash -IgnoreSSL -Force
}

#Filling Scheduler Selection in GUI Form
try
{
    foreach ($schedule in (Get-PrtgSchedule))
    {
        $cb_scheduler.Items.Add($schedule)
   
    }
}
catch
{
    writenotification("Unable to fetch Schedules, Please check the PRTG Server settings")
}#>




################  BUTTON FUNCTIONS  ################

$b_load_template.Add_Click(
{
    if ($tb_excelpath.Text -ne "") {setmarks}
    else { writenotification("Please select Excel first.") }
})

$b_save_settings.Add_Click(
{
    if ($tb_excelpath.Text -ne "") {writebacktoexcel}
    else { writenotification("Please select Excel first.") }
})

$b_save_date.Add_Click(
{
    if ($tb_excelpath.Text -ne "")
    {
        $dateinput = $calendar.SelectionStart
        $dateshort = $dateinput.ToShortDateString()
        savedate($dateshort)
    }
    else { writenotification("Please select Excel first.") }
    
})

$b_select_excel.Add_Click(
{
    Get-Excelpath    
})


#Show the GUI after the inital code
[void]$form_PRTGSCHEDULER.ShowDialog()