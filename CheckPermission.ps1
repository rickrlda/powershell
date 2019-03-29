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

if ((Get-MailboxPermission -Identity $identity -User $User) -eq $null -and (Get-RecipientPermission -Identity $identity -Trustee $User) -eq $null){
    
    Write-Host "O usuário $User não possui nenhum tipo de permissão na caixa compartilhada: $identity" -ForegroundColor red -BackgroundColor black
}
elseif ((Get-MailboxPermission -Identity $identity -User $User) -eq $null -and (Get-RecipientPermission -Identity $identity -Trustee $User) -ne $null){

    Write-Host "O usuário $user já possui permissão de SendAs, porém não tem permissão de FullAccess na caixa compartilhada: $identity" -ForegroundColor Yellow -BackgroundColor black
}
elseif ((Get-MailboxPermission -Identity $identity -User $User) -ne $null -and (Get-RecipientPermission -Identity $identity -Trustee $User) -eq $null){

    Write-Host "O usuário $user já possui permissão de FullAccess, porém não possui permissão de SendAs na caixa compartilhada: $identity" -ForegroundColor Yellow -BackgroundColor black
}
else {
    Write-Host "O usuário $user já possui permissão tanto de FullAccess quanto de SendAs na caixa compartilhada: $identity" -ForegroundColor green
}