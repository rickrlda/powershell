

$Mailboxes = Import-Csv -Path C:\TEMP\EmailAddress.csv
$grupoccr = "@grupoccr.intranet"
$smtp = "smtp:$($Mailboxes.Alias)$grupoccr"


foreach($Object in $Mailboxes){

   Set-Mailbox -Identity $Mailboxes.Alias -EmailAddressPolicyEnabled $false | Out-Null
   Set-Mailbox -Identity $Mailboxes.Alias -EmailAddresses @{remove=$smtp} | Out-Null
   Write-Host "Desablitado EmailAddressPolicyEnabled do usuário $($Object.Alias) e removido o smtp $($Object.Alias)$grupoccr"
}