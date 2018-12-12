$Status=(Get-Content "C:\Program Files (x86)\Redemption\Kontakte\SyncLog.txt" | select-object -last 1)
$Heute=(get-date -Format d)

if($Status -like "*$Heute*Done with Sync Run*") 
{Send-MailMessage -to "serveradmin@arges.de" -from "administrator@hermes.arges.local" -subject "ContactSync OK" -body $Status -SmtpServer localhost} 
else
{Send-MailMessage -to "serveradmin@arges.de" -from "administrator@hermes.arges.local" -subject "ContactSync Problem - please check on HERMES" -body $Status -SmtpServer localhost}

