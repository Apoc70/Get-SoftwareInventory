<#
    Script to fetch software inventory across all Exchange servers and some additional servers
    Must be executed in Exchange Management Shell (EMS)

    Not fancy script, but Copyright (c) 2016 Thomas Stensitzki
#>

$server = @('SERVER01','server02')

Get-ExchangeServer | Select-Object Name | %{$server += $_.Name} | Sort-Object

.\Get-SoftwareInventory.ps1 -SendMail -MailFrom postmaster@mcsmail.de -MailTo it@mcsmail.de -MailServer mobile.mcsmail.de -RemoteComputer $server