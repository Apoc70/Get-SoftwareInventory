<# 
    .SYNOPSIS 
    Script to generate an html reports of installed software, installed updates and installed components on a remote computer

    Thomas Stensitzki 

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

    Version 1.0, 2016-06-03

    Please send ideas, comments and suggestions to support@granikos.eu 

    .LINK 
    More information can be found at http://www.granikos.eu/en/scripts

    .DESCRIPTION 
    This script utilizes the slightly modified function Get-RemoteProgram by Jaap Brasser to fetch information about installed software.

    The script generates outputs for 
    * Installed programs
    * Installed updates
    * Installed components
    
    These three outputs are composed to a single html emails. All three outputs are 
    exported to CSV files which are added to the generated email as attachments. The
    html report is added as attachment as well.    
     
    .NOTES 
    Requirements 
    - Windows Server 2012 R2  
    - Remote Registry
    
    Revision History 
    -------------------------------------------------------------------------------- 
    1.0      Initial community release 

    .CREDITS
    Jaap Brasser
    Georges Zwingelstein, https://gallery.technet.microsoft.com/scriptcenter/site/profile?userName=Georges%20Zwingelstein

    .PARAMETER RemoteComputer
    Name of the computer to fetch data from (default: local computer)  

    .PARAMETER SendMail
    Switch to send an Html report

    .PARAMETER MailFrom
    Email address of report sender

    .PARAMETER MailTo
    Email address of report recipient

    .PARAMETER MailServer
    SMTP Server for email report

    .EXAMPLE 
    Get software information about SERVER01, SERVER02 and send email report
    
    .\Get-SoftwareInventory.ps1 -SendMail -MailFrom postmaster@mcsmemail.de -MailTo it-support@mcsmemail.de -MailServer mymailserver.mcsmemail.de -RemoteComputer SERVER01,SERVER02

    .EXAMPLE 
    Get software information about SERVER01, SERVER02 and write report to disk only
    
    .\Get-SoftwareInventory.ps1 -RemoteComputer SERVER01,SERVER02
#>
param(
    [Parameter(ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Position=0)]
        [string[]]
            $RemoteComputer = $env:COMPUTERNAME,
    [parameter(Mandatory=$false, HelpMessage='Send report as Html email')]
        [switch] $SendMail,
    [parameter(Mandatory=$false, HelpMessage='Sender address for result summary')]
        [string]$MailFrom = "",
    [parameter(Mandatory=$false, HelpMessage='Recipient address for result summary')]
        [string]$MailTo = "",
    [parameter(Mandatory=$false, HelpMessage='SMTP Server address for sending result summary')]
        [string]$MailServer = ""
)

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$now = Get-Date -Format F
$reportTitle = "Software Inventory Report"
$count = 1

<#
    .NOTES   
    Name: Get-RemoteProgram
    Author: Jaap Brasser
    Version: 1.2.1
    DateCreated: 2013-08-23
    DateUpdated: 2015-02-28
    Blog: http://www.jaapbrasser.com

    .LINK
    http://www.jaapbrasser.com

    .LINK
    https://gallery.technet.microsoft.com/scriptcenter/Get-RemoteProgram-Get-list-de9fd2b4 
#>
Function Get-RemoteProgram {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Position=0)]
        [string[]]
            $ComputerName = $env:COMPUTERNAME,
        [Parameter(Position=0)]
        [string[]]$Property 
    )

    begin {
        $RegistryLocation = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\',
                            'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\'
        $HashProperty = @{}
        $SelectProperty = @('ProgramName','ComputerName')
        if ($Property) {
            $SelectProperty += $Property
        }
    }

    process {
        foreach ($Computer in $ComputerName) {
            try
            {
                $RegBase = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$Computer)
                foreach ($CurrentReg in $RegistryLocation) {
                if ($RegBase) {
                    $CurrentRegKey = $RegBase.OpenSubKey($CurrentReg)
                    if ($CurrentRegKey) {
                        $CurrentRegKey.GetSubKeyNames() | ForEach-Object {
                            if ($Property) {
                                foreach ($CurrentProperty in $Property) {
                                    $HashProperty.$CurrentProperty = "$(($RegBase.OpenSubKey("$CurrentReg$_")).GetValue($CurrentProperty))"
                                }
                            }
                            $HashProperty.ComputerName = $Computer
                            $HashProperty.ProgramName = ($DisplayName = ($RegBase.OpenSubKey("$CurrentReg$_")).GetValue('DisplayName'))
                            if ($DisplayName) {
                                New-Object -TypeName PSCustomObject -Property $HashProperty |
                                Select-Object -Property $SelectProperty
                            }
                        }
                    }
                }
            }
            }
            catch {}
        }
    }
}

### MAIN ######################################################
# Query strig courtesy of Georges Zwingelstein https://gallery.technet.microsoft.com/scriptcenter/site/profile?userName=Georges%20Zwingelstein 

# Some CSS to get a pretty report
$head = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$($reportTitle)</title>
<style type=”text/css”>
<!–
body {
    font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
}
h1{ clear: both; 
    font: 1.2em sans-serif;
    color:#000000; 
}
h2{ 
    clear: both; 
    font: 1.1em sans-serif;
    color:#354B5E; 
}
h3{
    clear: both;
    font-size: 75%;
    margin-left: 20px;
    margin-top: 30px;
    color:#475F77;
}
table{
    border-collapse: collapse;
    border: solid;
    font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
    color: black;
    margin-bottom: 10px;
}

table tr {
    border: solid;
}
 
table td{
    font-size: 12px;
    padding-left: 5px;
    padding-right: 5px;
    text-align: left;
}
 
table th {
    font-size: 12px;
    font-weight: bold;
    padding-left: 5px;
    padding-right: 5px;
    text-align: left;

}
->
</style>
"@


foreach($computer in $RemoteComputer) {
    $computer = $computer.ToUpper()

    Write-Progress -Activity 'Fetching software inventory' -Status "Working on server $($computer) [$($count)/$($RemoteComputer.Count)]" -PercentComplete (($count/$RemoteComputer.Count)*100)

    $htmlreport = @()
    $attachments = @()
    $reportTitle = "Software Inventory Report [$($computer)] - $($now)"
    $csvFile = 

    # Fetch installed programs
    $InstalledPrograms = Get-RemoteProgram -ComputerName $computer -Property Publisher,InstallDate,DisplayVersion,IsMinorUpgrade,ReleaseType,ParentDisplayName,SystemComponent | Where-Object {[string]$_.SystemComponent -ne 1 -and ![string]$_.IsMinorUpgrade -and ![string]$_.ReleaseType -and ![string]$_.ParentDisplayName} | Sort-Object ProgramName 
    if( $InstalledPrograms -ne $null) {
        $csvInstalledPrograms = "$myDir\InstalledPrograms-$($computer).csv"
        $InstalledPrograms | Export-Csv -Path $csvInstalledPrograms -Force -Encoding UTF8 -NoTypeInformation -Delimiter ';'
        $htmlReport += $InstalledPrograms | ConvertTo-Html -Fragment -PreContent "<h1>$($computer)</h1><h2>Installed Programs</h2>"
        $attachments += $csvInstalledPrograms
    }
    else {
        $htmlReport += "<h1>$($computer)</h1><h2>Installed Programs</h2><p>Error fetching installed programs</p>"
    }

    # Fetch installed updates
    $InstalledUpdates = Get-RemoteProgram -ComputerName $computer -Property Publisher,InstallDate,DisplayVersion,IsMinorUpgrade,ReleaseType,ParentDisplayName,SystemComponent | Where-Object {[string]$_.SystemComponent -ne 1 -and ([string]$_.IsMinorUpgrade -or [string]$_.ReleaseType -or [string]$_.ParentDisplayName)} | Sort-Object ParentDisplayName,ProgramName
    if($InstalledUpdates -ne $null) {
        $csvInstalledUpdates = "$myDir\InstalledUpdates-$($computer).csv"
        $InstalledUpdates | Export-Csv -Path $csvInstalledUpdates -Force -Encoding UTF8 -NoTypeInformation -Delimiter ';'
        $htmlReport += $InstalledUpdates | ConvertTo-Html -Fragment -PreContent '<h2>Installed Updates</h2>'
        $attachments += $csvInstalledUpdates
    }
    else {
        $htmlReport += "<h1>$($computer)</h1><h2>Installed Updates</h2><p>Error fetching installed updates</p>"
    }

    # Fetch installed components
    $InstalledComponents = Get-RemoteProgram -ComputerName $computer -Property Publisher,InstallDate,DisplayVersion,IsMinorUpgrade,ReleaseType,ParentDisplayName,SystemComponent | Where-Object {[string]$_.SystemComponent -eq 1} | Sort-Object ProgramName
    if($InstalledComponents -ne $null) {
        $csvInstalledComponents = "$myDir\InstalledComponents-$($computer).csv"
        $InstalledComponents | Export-Csv -Path $csvInstalledComponents -Force -Encoding UTF8 -NoTypeInformation -Delimiter ';'
        $htmlReport += $InstalledComponents | ConvertTo-Html -Fragment -PreContent '<h2>Installed Components</h2>'
        $attachments += $csvInstalledComponents
    }
    else {
        $htmlReport += "<h1>$($computer)</h1><h2>Installed Components</h2><p>Error fetching installed components</p>"
    }

    # Generate full html report
    [string]$htmlreport = ConvertTo-Html -Body $htmlreport -Head $head -Title $reportTitle 
    $reportfile = "$myDir\SoftwareInventory-$($computer).html"
    $htmlreport | Out-File -Encoding utf8 -FilePath $reportfile -Force
    $attachments += $reportfile

    if($SendMail) {
        Send-MailMessage -From $MailFrom -To $MailTo -SmtpServer $MailServer -Body $htmlreport -Subject $reportTitle -Attachments $attachments -BodyAsHtml
    }
    $count++
}

Return 0