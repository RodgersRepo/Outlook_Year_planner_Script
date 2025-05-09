# Outlook Year Planner Powershell Script

A Powershell script with a GUI for generating an HTML year planner from a selected calender in you Outlook. Only works for Outlook classic not Office 365!
Based on the excelent VBScript published by NiveauVerleih way back in 2008 (see credits below). Have used that script for years but felt it needed porting tto Powershell. This
script does not have the bells and whistles that NiveauVerleih script had, no colours etc. just start time and appointment name.

## Installation
Click on the `outlookCalGen.ps1` link for the script above. When the PowerShell code page appears click the **Download Raw file** button top right.
Once downloaded to your computer have a read of the script in your prefered editor. All the information for executing the script will be in the script synopsis.

## Usage
From a Powershell prompt:
`Run .\outlookCalGen.ps1 <no arguments needed>`
  Or create shortcut to:
  `"powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\<PathToYourScripts>\outlookCalGen.ps1"`
  Then just double click the shortcut like you are starting an application.8316b)

## Caveats
This script has only been tested on Microsoft Windows 11 Pro 10.0.26100 N/A Build 26100. The powershell version tested was 5.1.26100.3912.

## Credits and references
[NiveauVerleih origanal VBScript that this is based on](http://niveauverleih.blogspot.com/)

----


