<#
    Predefined script to fetch software inventory and send email report

    Not a real script, but Copyright (c) 2016 Thomas Stensitzki
#>

.\Get-SoftwareInventory.ps1 -SendMail -MailFrom postmaster@mcsmemail.de -MailTo it@mcsmemail.de -MailServer myserver@mcsmemail.de -RemoteComputer SERVER01,SERVER02