<#
.SYNOPSIS
    Verifica se o usuário possui permissão FullAcces & SendAs, caso negativo adiciona.


    .DESCRIPTION
    Este script tem como finalidade verificar se o usuário tem permissão na caixa de correio compartilhada, 
    caso negativo ele adiciona a permissão de FullAccess e SendAs 
    para os usuários que utilizam a caixa compartilhada.
    Script desenvolvido para uso exclusivo do Tribunal de Justiça de SP.
    
    
    Version 1.0, 18 de Março, 2019
    Autor: Ricardo Azevedo

#>

$Identity = Read-Host "Digite o alias, email ou Displayname da Caixa Compartilhada"
$User = Read-Host "Digite o alias, email ou DisplayName do Usuário"

if ((Get-MailboxPermission -Identity $identity -User $User) -eq $null){
    
   Add-MailboxPermission -Identity $identity -User $User -AccessRights FullAccess -AutoMapping $false | Out-Null
   Add-RecipientPermission -Identity $identity -Trustee $User -AccessRights SendAs -Confirm:$false | Out-Null
   Write-Host "Permissão de FullAccess e SendAs adicionado para o usuário $User na Caixa Compartilhada: $identity" -ForegroundColor green
}
else {
    Write-Host "O usuário $user já possui permissão na Caixa Compartilhada $identity" -ForegroundColor yellow
}