<#
  Set parameters
#>
$mailparams = @{
  "mailserver" = "servername.contoso.com";
  "recipientaddress" = "emailrecipient@contoso.com";
  "senderaddress" = "ExchangeHealth@contoso.com";
  "port" = 25;
}

<#
  Load the exchange cmdlets.
#>
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
<#
  Load functions and static code.
#>
$css = 'h1 {
          margin-left: auto;
          margin-right: auto;
          text-transform: uppercase;
          text-align: center;
          font-size: 13pt;
          font-weight: bold;
          }

          h2 {
              margin-left: auto;
              margin-right: auto;
              text-transform: capitalize;
              text-align: center;
              font-family: "Segoe UI";
              font-size: 14pt;
              font-weight: bold;
          }

          body {
              margin-left: auto;
              margin-right: auto;
              text-align: center;
              font-family: "Segoe UI";
              font-weight: lighter;
              font-size: 9pt;
              color:#2f2f2f;
              background-color: white;
          }

          table {
              margin-left: auto;
              margin-right: auto;
              border-width: 1px;
              border-style: solid;
              border-color: #2f2f2f;
              border-collapse: collapse;
          }

          th {
              font-family: "Segoe UI";
              font-weight: lighter;
              color: white;
              text-transform: capitalize;
              margin-left: auto;
              margin-right: auto;
              border-width: 1px;
              border-style: solid;
              border-color: #2f2f2f;
              background-color: #d32f2f;
          }

          td {
              margin-left: auto;
              margin-right: auto;
              border-width: 1px;
              border-style: solid;
              border-color: #2f2f2f;
              background-color: white;
          }

'
function Send-Mail {
    [CmdletBinding()]
    param(
    [Parameter(ValueFromPipeline=$True)]$body,
    [Parameter(ValueFromPipeline=$True)]$bodyashtml,
    $mailserver,
    $recipientaddress,
    $senderaddress,
    $subject,
    $mailcredential,
    $port
    )
    if(!($mailcredential)){
            $pass = ConvertTo-SecureString "whatever" -asplaintext -force
            $creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "NT AUTHORITY\ANONYMOUS LOGON", $pass
    }
    ELSE {$creds = $mailcredential}
    if ($bodyashtml){
            send-mailmessage -From $senderaddress -To $recipientaddress -SmtpServer $MailServer -Subject $subject -bodyashtml $bodyashtml -port $port -Credential $creds
    }
    else {
            send-mailmessage -From $senderaddress -To $recipientaddress -SmtpServer $MailServer -Subject $subject -body $body -port $port -Credential $creds
    }
}
<#
  Start Processing
#>
$servers = (Get-ExchangeServer).Name
$Output = (convertto-html -head "<style>$css</style>" -body "<h2>Exchange Health Information</h2>")
$Output += "<center>"
#list servers
$Output += Get-ExchangeServer | select Name,ServerRole,AdminDisplayVersion | convertto-html -fragment
$Output += "</br>"
# Check database copy status.
$Output += (Get-ExchangeServer).Name |
  % { (Get-MailboxDatabaseCopyStatus -server $_ |
      select Name,Status,CopyQueueLength,ContentIndexState |
      convertto-html -fragment) + "</br>"
    }

#

# Get Disk info for each server
$output +=  $servers | % {
  Invoke-Command -computername $_ -command  {
      (gwmi win32_logicaldisk | ? {$_.Drivetype -eq 3} |
        select @{"Name"="Server";Expression = {$env:computername}},DeviceID,VolumeName,@{"Name" = "Size";Expression = {'{0:N2}' -f ($_.Size /1GB)}},@{"Name"="FreeSpace";Expression = {'{0:N2}' -f ($_.Freespace /1GB)}} | convertto-html -fragment) + "</br>"
  }

}

$output += (Get-MailboxDatabase -status | select Name,AvailableNewMailboxSpace,LastFullBackup,LastIncrementalBackup,BackupInProgress | convertto-html -fragment) + "</br>"

#Get average mailbox size.
$databases = (Get-MailboxDatabase).Name
$totalsizes = $databases | % {
  (((Get-MailboxStatistics -Database $_).TotalItemSize).Value).tobytes()
} | % {$total += [int64]$_ /1GB}
#Get Deleted Items Size in GB
$databases = (Get-MailboxDatabase).Name
$databases | % {(((Get-MailboxStatistics -Database $_).TotalDeletedItemSize).value).tobytes()} | % {$deltotal += [int64]$_ /1GB}
$delitems = '{0:N2}' -f $deltotal

#Get total Mailboxes
$mailboxes = (Get-CASMailbox * -ResultSize unlimited).count

# Get mailbox information
$Object = @()
$object += New-object psobject -property @{
  "Total Mailboxes" = $mailboxes;
  "Average Mailbox Size(GB)" = ('{0:N2}' -f ($total / $mailboxes));
  "Total Mailbox Size(GB)" = ('{0:N2}' -f $total);
  "Total Deleted Items Size(GB)" = $delitems;

}
$output += $object | convertto-html -fragment

$output += "</center>"
Send-Mail @mailparams -subject "ExchangeHealth Report" -bodyashtml $([string]$output)
