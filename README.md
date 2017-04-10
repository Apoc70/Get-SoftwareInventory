# Get-SoftwareInventory.ps1
Script to generate an html reports of installed software, installed updates and installed components on a remote computer

## Description
This script utilizes the slightly modified function Get-RemoteProgram by Jaap Brasser to fetch information about installed software.

## Parameters
### RemoteComputer
Name of the computer to fetch data from (default: local computer)  

### SendMail
Switch to send an Html report

### MailFrom
Email address of report sender

### MailTo
Email address of report recipient

### MailServer
SMTP Server for email report

## Examples
```
.\Get-SoftwareInventory.ps1 -SendMail -MailFrom postmaster@mcsmemail.de -MailTo it-support@mcsmemail.de -MailServer mymailserver.mcsmemail.de -RemoteComputer SERVER01,SERVER02
```
Get software information about SERVER01, SERVER02 and send email report

```
.\Get-SoftwareInventory.ps1 -RemoteComputer SERVER01,SERVER02
```
Get software information about SERVER01, SERVER02 and write report to disk only

## Note
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## TechNet Gallery
Find the script at TechNet Gallery
* not yet published to TechNet Gallery

## Credits
Written by: Thomas Stensitzki

## Social 

* My Blog: http://justcantgetenough.granikos.eu
* Twitter: https://twitter.com/stensitzki
* LinkedIn:	http://de.linkedin.com/in/thomasstensitzki
* Github: https://github.com/Apoc70

For more Office 365, Cloud Security and Exchange Server stuff checkout services provided by Granikos

* Blog: http://blog.granikos.eu/
* Website: https://www.granikos.eu/en/
* Twitter: https://twitter.com/granikos_de


Additional Credits:
* Jaap Brasser
* Georges Zwingelstein, https://gallery.technet.microsoft.com/scriptcenter/site/profile?userName=Georges%20Zwingelstein
