#<#
#.SYNOPSIS
#    Verifica se o usuário possui permissão FullAcces & SendAs, caso negativo adiciona.
#
#
#    .DESCRIPTION
#    Este script tem como finalidade verificar se o usuário tem permissão na caixa de correio compartilhada, 
#    caso negativo ele adiciona a permissão de FullAccess e SendAs 
#    para os usuários que utilizam a caixa compartilhada.
#    Script desenvolvido para uso exclusivo do Tribunal de Justiça de SP.
#    
#       Obs: Deve-se ser criado um arquivo com a extensão .csv no seguinte caminho do Windows C:\Temp\O365
#       a primeira linha deve conter um nome chamado Alias
#    
#    Version 1.0, 18 de Março, 2019
#    Autor: Ricardo Azevedo
#   $ScriptFromGithHub = Invoke-WebRequest https://raw.githubusercontent.com/rickrlda/powershell/O365/ValidatePermission.ps1
#   Invoke-Expression $($ScriptFromGithHub.Content
#>

$Identity = Read-Host "Digite o alias, email ou DisplayName da Caixa Compartilhada"
$Users = Import-Csv -Path C:\TEMP\O365\Permission.csv

foreach ($Object in $Users){
if ((Get-MailboxPermission -Identity $identity -User $Object.Alias) -eq $null -and (Get-RecipientPermission -Identity $identity -Trustee $Object.Alias) -eq $null){
    
   Add-MailboxPermission -Identity $identity -User $Object.Alias -AccessRights FullAccess -AutoMapping $false | Out-Null
   Add-RecipientPermission -Identity $identity -Trustee $Object.Alias -AccessRights SendAs -Confirm:$false | Out-Null
   Write-Host "Permissão de FullAccess e SendAs foi adicionado para o usuário $($Object.Alias) na Caixa Compartilhada: $identity" -ForegroundColor green
}
elseif ((Get-MailboxPermission -Identity $identity -User $Object.Alias) -eq $null -and (Get-RecipientPermission -Identity $identity -Trustee $Object.Alias) -ne $null){

    Add-MailboxPermission -Identity $identity -User $Object.Alias -AccessRights FullAccess -AutoMapping $false | Out-Null
    Write-Host "O usuário $($Object.Alias) já possui permissão de SendAs porém foi adicionado a permissão de FullAccess na Caixa Compartilhada: $identity" -ForegroundColor green
}
elseif ((Get-MailboxPermission -Identity $identity -User $Object.Alias) -ne $null -and (Get-RecipientPermission -Identity $identity -Trustee $Object.Alias) -eq $null){

    Add-RecipientPermission -Identity $identity -Trustee $Object.Alias -AccessRights SendAs -Confirm:$false | Out-Null
    Write-Host "O usuário $($Object.Alias) já possui permissão de FullAccess porém foi adicionado a permissão de SendAs na Caixa Compartilhada: $identity" -ForegroundColor green
}
else {
    Write-Host "O usuário $($Object.Alias) já possui permissão tanto de FullAccess quanto de SendAs na Caixa Compartilhada: $identity" -ForegroundColor yellow
}
}
