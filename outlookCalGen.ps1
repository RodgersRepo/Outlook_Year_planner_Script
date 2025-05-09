<#
.SYNOPSIS
  Name: outlookCalGen.ps1
  Create an HTML year planner style table.
  Displays appointments on the year planner
  with start times
 
.DESCRIPTION
  Takes the year and a named calender from Outlook
  (Classic NOT OFFICE 365!!), generates a basic year planner
  from appointments in this calender.
  Based on the excelent VBScript provided by
  http://niveauverleih.blogspot.com/. I have used this
  for years but thought it was time to port to PowerShell.
  This script does not do the flash stuff the origanal
  VBScipt does, just appointments and start times no colours.
 
.NOTES
Copyright (C) 2025  RITT
 
 
     This program is free software: you can redistribute it and/or modify
     it under the terms of the GNU General Public License as published by
     the Free Software Foundation, version 3.
 
     This program is distributed in the hope that it will be useful,
     but WITHOUT ANY WARRANTY; without even the implied warranty of
     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
     GNU General Public License for more details.
 
     To view the GNU General Public License, see <http://www.gnu.org/licenses/>.

    Release Date: 08/05/2025       
   
    Change comments:
    Initial realease - RITT
   
   
  Author: RITT
       
.EXAMPLE
  Run .\outlookCalGen.ps1 <no arguments needed>
  Or create shortcut to:
  "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\<PathToYourScripts>\outlookCalGen.ps1"
  Then just double click the shortcut like you are starting an application.

#>

#----------------[ Declarations ]-----------------------------------------------------#

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"                                     # What to do if an unrecovable error occures
$global:scriptPath = Split-Path -Path $MyInvocation.MyCommand.Path  # This scripts path
$global:scriptName = $MyInvocation.MyCommand.Name                   # This scripts name
$wpf = @{ }                                                         # A hash table to store node names from the XAML below
# hash tables are key/value stored arrays, each
# value in the array has a key

Add-Type -AssemblyName presentationframework, presentationcore      # Add these assemblys,
# components of the Windows Presentation Foundation (WPF)

######################################################################################
#       Here-String with the eXAppMarkupLang (XAML) needed to display the GUI        #
######################################################################################

# A here-string of type xml
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        ResizeMode="CanResizeWithGrip" Name="outlookCalGenGui"
        Title="Outlook Year Planner Creator V1" Height="800" Width="700"
        Background="#FFFFFFFF" FontSize="15" FontFamily="Segoe UI">

    <Window.Resources> <!--Match name with the root element in this case Window-->
        <!--Setting default styling for all buttons-->  
        <Style TargetType="Button">
         <Setter Property="Width" Value="143" />
         <Setter Property="Height" Value="32" />
         <Setter Property="Margin" Value="10" />
         <Setter Property="FontSize" Value="18" />
         <Setter Property="Background" Value="#FFB8B8B8" />
        </Style>
        <Style TargetType="TextBox">
         <Setter Property="Background" Value="#FFB8B8B8" />
         <Setter Property="Height" Value="32" />
        </Style>
        <Style TargetType="ComboBox">
         <Setter Property="Background" Value="#FFB8B8B8" />
         <Setter Property="Height" Value="32" />
        </Style>
     </Window.Resources>

    <Grid Name="MainGrid">
     
      <Grid.RowDefinitions>
        <RowDefinition Name ="Row0" Height="33*"/><!--Row 0 Row Heights as percentage of entire window-->
        <RowDefinition Name="Row1" Height="27*"/> <!--Row 1-->
        <RowDefinition Name="Row2" Height="30*"/> <!--Row 2-->
        <RowDefinition Name="Row3" Height="10*"/> <!--Row 3-->
      </Grid.RowDefinitions>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>           <!--Column 0-->
      </Grid.ColumnDefinitions>

       <DockPanel>
        <Menu DockPanel.Dock="Top" Background="#FFFFFFFF">
            <MenuItem Header="_File">
                <MenuItem Header="_About" Name="menuItemAbout"/>
                <Separator />
                <MenuItem Header="_Exit" Name="menuItemExit"/>
            </MenuItem>
        </Menu>
       </DockPanel>

        <GroupBox Name="InstructionsGrpBox" Grid.Row="0" Grid.Column="0" Header="Instructions" BorderThickness="0.5" HorizontalAlignment="Stretch" Margin="15,20,15,0" VerticalAlignment="Stretch" >
           <StackPanel>
            <TextBlock Name="instructionsTxtBlk" TextWrapping="Wrap" VerticalAlignment="Top"
                Text = "&#x0a;
                This script will only execute if you have Outlook Classic&#x0a;
                installed on this computer. An error appears in the Console&#x0a;
                Messages if Outlook is not found.&#x0a;&#x0a;
                Pick a Calender from the drop down list and supply&#x0a;
                the year ( for example 2025 ), the script will generate&#x0a;
                a year planner web page with appointments from your &#x0a;
                selected calender.&#x0a;
                Press OK to continue. Press Cancel to quit.                
                "/>
           </StackPanel>
        </GroupBox>

        <GroupBox Name="UserInputGrpBox" Grid.Row="1" Grid.Column="0" Header="Please complete the following" BorderThickness="0.5" HorizontalAlignment="Stretch" Margin="15,20,15,0"  VerticalAlignment="Stretch">
            <Grid Name="UserInputGrid">
             <StackPanel>
              <TextBlock Padding="10">Pick a calender from your Outlook classic app:</TextBlock>
              <ComboBox Name="OutlookCalsComboBox"  HorizontalAlignment="Stretch">
                <!--Combo box populated with calenders from outlook-->
              </ComboBox>
              <TextBlock Padding="10">Type the year in the text box below:</TextBlock>
              <TextBox Name="yearTxtBox" />
             </StackPanel>
            </Grid>
        </GroupBox>

        <GroupBox Name="resultsGrpBox" Grid.Row="2" Grid.Column="0" Header="Console Messages" BorderThickness="0.5" HorizontalAlignment="Stretch" Margin="15,20,15,0"  VerticalAlignment="Stretch" >
            <Grid Name="resultsGrid">
             <ScrollViewer>
                <TextBlock Name="Output_TxtBlk" TextWrapping="Wrap" TextAlignment="Left" VerticalAlignment="Stretch" />
             </ScrollViewer>
            </Grid>
            <!---->
        </GroupBox>

        <StackPanel Name="OkCancelStackPanel" Grid.Row="3" Grid.Column="0" HorizontalAlignment="Right" Orientation="Horizontal" >
            <Button Name="oKButton1" Content="OK"  />
            <Button Name="myCancelButton" Content="Cancel" />
        </StackPanel>

    </Grid>
</Window>
"@

#######################################################################################
#        HTML that will be used to display the calender                               #
#######################################################################################

# Store the HTML as a here string within a script
# block. That way can expand the script block later when the $year
# variable is set. At this point in the script $year is not
# yet known so an error would occur

$html = {
    @"
<!DOCTYPE html>
<html>
<head>
    <title>Outlook Calendar - $year</title>
    <style>
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid black; padding: 5px; text-align: center; }
        th { background-color: #f2f2f2; }
        .month-header { font-weight: bold; background-color: #d9d9d9; width: 120px; }
        .weekend-column { background-color: #d9d9d9; }
    </style>
</head>
<body>
    <h2>Outlook Calendar - $year</h2>
    <table>
        <tr><th>Month</th>
"@
}

#------------------[ Functions ]------------------------------------------------------#
#######################################################################################
#        Function to exit the script, called by the cancel buttons                    #
#######################################################################################

function Close-AndExit {
    $wpf.outlookCalGenGui.Close()    
}

#######################################################################################
#      Function to check that the user has either entered some text in the txt        #
#                 boxes or seleced from the drop down lists                           #
#######################################################################################

function Test-UserInput {
    $resultBool = $true

    # If outlook not installed give the user a warning
    if (! $outLookInstalled) {
        $wpf.Output_TxtBlk.Foreground = "Red"
        $wpf.Output_TxtBlk.Text = "Unable to locate Outlook Classic`nOn this computer`nPlease check it is installed, then re-execute this script`n"
        exit # Exit script but dont close the GUI, if you do they cant read the error!!
    }

    # Checking for valid user input
    if ( ([string]::IsNullOrEmpty($wpf.yearTxtBox.Text)) -or
         ([string]::IsNullOrEmpty($wpf.OutlookCalsComboBox.Text))) {
        $resultBool = $false
        $wpf.Output_TxtBlk.Foreground = "Red"
        $wpf.Output_TxtBlk.Text = "One of the following is missing`nA Calender from the drop down list`nA Year typed into the text box`n"
    }

    if ($wpf.yearTxtBox.Text) {
        # Check if the year is a valid 4 digit number
        # The regex will match a string that contains exactly 4 digits
        if ( ! ($wpf.yearTxtBox.Text -match "^\d{4}$") ) {
            $resultBool = $false
            $wpf.Output_TxtBlk.Foreground = "Red"
            $wpf.Output_TxtBlk.Text = "You have entered an invalid year. Entry must be numeric i.e. 2025!`n"
        }
    }

    $resultBool
}

#######################################################################################
#      Function to collect all the appointments from the users selected calender      #
#                   sort the result                                                   #
#######################################################################################

function Get-AllAppointments ( $selectedCalendar ) {
    # Get the year from the GUI
    $year = $wpf.yearTxtBox.Text

    # Retrieve appointments including recurring appointments
    $Appointments = $selectedCalendar.Items
    $Appointments.IncludeRecurrences = $true
    $Appointments.Sort("[Start]")

    # Define the start and end date for the filter (one year)
    $startDate = (Get-Date -Year $year -Month 1 -Day 1).ToString("g")
    $endDate = (Get-Date -Year $year -Month 12 -Day 31).ToString("g")

    # Filter the calendar items by the date range
    $dateFilter = "[Start] >= '$startDate' AND [End] <= '$endDate'"
    $Appointments = $Appointments.Restrict($dateFilter)
    $Appointments.Sort("[Start]")

    # PowerShell unrolls arrays/objects, use a unitary comma
    # to return $Appointments intact
    return , $Appointments
}

#######################################################################################
#             Function that generates then saves the HTML calender                    #
#                                                                                     #
#######################################################################################

function New-CalHtml ( $Appointments ) {
    # Get the year of interest from the GUI form
    $year = $wpf.yearTxtBox.Text

    # Expand the HTML here string defined at the top
    # of this script. Use the call operator & to
    # execute the script block containing this here string
    $html = & $html
    
    # Add days of the week headers for each row of days in the HTML table, but stop at 31 columns
    for ($i = 1; $i -le 5; $i++) {
        $html += "<th>Monday</th><th>Tuesday</th><th>Wednesday</th><th>Thursday</th><th>Friday</th><th>Saturday</th><th>Sunday</th>"
        # Compensate for some months having less than 31 days
        if ($i -eq 5) {
            $html += "<th>Monday</th><th>Tuesday</th>"
        }
    }

    $html += "</tr>"

    # Generate rows for each month
    for ($month = 1; $month -le 12; $month++) {
        $monthName = (Get-Culture).DateTimeFormat.GetMonthName($month)
        $daysInMonth = (Get-Date -Year $year -Month $month -Day 1).AddMonths(1).AddDays(-1).Day

        # Start row with month name
        $html += "<tr><td class='month-header'>$monthName</td>"

        # Initialize empty slots for first week alignment
        $firstDayOfMonth = (Get-Date -Year $year -Month $month -Day 1).DayOfWeek
    
        # If first day of month is Sunday, PS will return 0, needs to be 7 for the maths to work
        if ([int]$firstDayOfMonth -eq 0) {
            [int]$firstDayOfMonth = 7
        }
        $emptySlots = ([int]$firstDayOfMonth - 1) % 7

        for ($i = 0; $i -lt $emptySlots; $i++) {
            $html += "<td></td>"
        }

        # Loop through each day in the month
        for ($day = 1; $day -le $daysInMonth; $day++) {
               
            $currentDate = Get-Date -Year $year -Month $month -Day $day
            $nextDate = $currentDate.AddDays(1)
            $strRestriction = "[Start] >= '$($currentDate.ToString("dd/MM/yyyy"))' AND [End] < '$($nextDate.ToString("dd/MM/yyyy"))'"
            $dayAppointments = $Appointments.Restrict($strRestriction)
            $dayAppointments.Sort("[Start]")

            # Check if text is present in $dayAppoitments
            if ($dayAppointments | ForEach-Object { $null -ne $_.Subject -and $_.Subject -ne "" }) {
                $appointmentText = $dayAppointments | ForEach-Object { "$($_.Start.ToShortTimeString()) - $($_.Subject)" }
                $appointmentText = $appointmentText -join "<br>"
            }
            else {
                $appointmentText = "<br>"
            }    

            # Determine if the column is a weekend column
            if ($emptySlots -ge 0) {
                $emptySlotsCompensated = $emptySlots - 1
            }
            else {
                $emptySlotsCompensated = $emptySlots
            }
            
            $columnClass = if ((($day + $emptySlotsCompensated) % 7) -eq 5 -or (($day + $emptySlotsCompensated) % 7) -eq 6) { "weekend-column" } else { "" }
            $html += "<td class='$columnClass'>$day<br>$appointmentText</td>"
        }
    
        # Fill remaining empty spaces if the month ends midweek, but only if it's not the last day of the month
        if ($day - 1 -ne $daysInMonth) {
            $remainingSlots = (7 - ((($daysInMonth + $emptySlots) % 7)) % 7)
            for ($i = 0; $i -lt $remainingSlots; $i++) {
                $html += "<td></td>"
            }
        }
    
        $html += "</tr>"
    
    }

    $html += "</table></body></html>"

    # Save HTML file
    $html | Out-File -FilePath "$env:TEMP\OutlookCalendar.html"
    $wpf.Output_TxtBlk.Text += "HTML file created: $env:TEMP\OutlookCalendar.html`nOpen using a browser"
}
#----------------[ Main Execution ]---------------------------------------------------#

#######################################################################################
#               Read the XAML needed for the GUI. Populate drop down                  #
#######################################################################################

$reader = New-Object System.Xml.XmlNodeReader $xaml
$myGuiForm = [Windows.Markup.XamlReader]::Load($reader)

# Collect the Node names of buttons, txt boxes etc.

$namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
$namedNodes | ForEach-Object { $wpf.Add($_.Name, $myGuiForm.FindName($_.Name)) }

# Populate the calender drop down list
# if outlook classic not installed set 
# $outLookInstalled to $false

try {
    # Create Outlook COM object
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")

    # Get all calendars
    $CalendarFolder = $Namespace.Folders | Where-Object { $_.Name -eq $Namespace.CurrentUser.Name }
    $Calendars = $CalendarFolder.Folders | Where-Object { $_.DefaultItemType -eq 1 }

    # Populate the calender drop down
    foreach ($cal in $Calendars) {
        $wpf.OutlookCalsComboBox.Items.Add($cal.Name)
    }
    $outLookInstalled = $true
}
catch {
    $outLookInstalled = $false
}

#######################################################################################
#               This code runs when the Menu Item about button is clicked             #
#######################################################################################

$wpf.menuItemAbout.Add_Click({
        #Show the help synopsis in a GUI
        Get-Help "$global:scriptPath\$global:scriptName" -ShowWindow
    })

#######################################################################################
#               This code runs when the Menu Item exit button is clicked              #
#######################################################################################

$wpf.menuItemExit.Add_Click({
        #Call the close and exit function
        Close-AndExit
    })

#######################################################################################
#               This code runs when the Cancel buttons are clicked                    #
#######################################################################################

$wpf.myCancelButton.Add_Click({
        #Call the close and exit function
        Close-AndExit
    })

#######################################################################################
#               This code runs when the OK 1 button is clicked                        #
#######################################################################################

$wpf.oKButton1.Add_Click({
    
        $wpf.Output_TxtBlk.Foreground = "Black"
        $wpf.Output_TxtBlk.Text = ""
        if (Test-UserInput) {
            # User has inputed valid info and outlook confirmed as installed
            # on this computer
            $selectedCalendar = $Calendars | Where-Object { $_.Name -eq $wpf.OutlookCalsComboBox.SelectedItem }
            
            # Update console text then refresh the dispatcher
            # This is needed to update the GUI from a non-GUI thread
            $wpf.Output_TxtBlk.Text = "Fetching calender entries for $($wpf.OutlookCalsComboBox.SelectedItem)`n"
            $wpf.Output_TxtBlk.Text += "Please wait...`n"
            $wpf.outlookCalGenGui.Dispatcher.Invoke([action]{},"Render")

            # Call function that retrieves all appointments from the selected calender
            $allMyAppointments = Get-AllAppointments $selectedCalendar

            # Call the function that generates the HTML calender
            New-CalHtml $allMyAppointments

        } 
    })
 
#######################################################################################
#               Show the GUI window by name                                           #
#######################################################################################

$wpf.outlookCalGenGui.ShowDialog() | out-null # null dosn't show false on exit
