#<
#.SYNOPSIS
#    Migra caixa de e-mails to/from Exchange Online.
#
#
#    .DESCRIPTION
#    Este script tem como finalidade migrar uma caixa de emai que não foi migrada, trazer uma caixa que foi migrada para o Office365
#    e finaizar a migração de uma caixa que encontra-se totalmente sincronizada com Status de Synced.
#    Script desenvolvido para uso exclusivo da empresa AUTOBAN - CCR.
#    
#    .VERSION 1.0, 26 de Junho, 2019
#    Autor: Ricardo Azevedo
#
#    .GihtHubUrl
#   $ScriptFromGithHub = Invoke-WebRequest https://raw.githubusercontent.com/rickrlda/powershell/O365/MoveRequest.ps1
#   Invoke-Expression $($ScriptFromGithHub.Content
#>



$X = Read-Host "Esse script executa as seguintes ações. Migração On-Boarding (move a caixa para o O365), migração Off-Boarding (traz a caixa do O365) e finaliza um usuário já sincronizado.

Selecione agora o tipo de ação desejada!

1 = On-Boarding
2 = Off-Boarding
3 = Finalizar Synced

"


if($x -eq 1){
    
    $Identity = Read-Host "Digite o endereço email do usuário"
    $AdminCred = Get-Credential -Message "Digite os dados da conta do usuário com permissão de Exchange Organization Management" 
    Write-host "Movendo a caixa do usuário $Identity para o Office 365"
    New-MoveRequest -Remote -Identity $Identity -RemoteHostName 'webmail.grupoccr.com.br' -RemoteCredential $AdminCred -TargetDeliveryDomain 'grupoccr.mail.onmicrosoft.com' -BadItemLimit 100 -LargeItemLimit 100
    Write-host "Para acompanhar o andamento da migração, utilize o seguinte cmdlet: Get-MoveRequestStatistics -Identity $Identity"
}
elseif($x -eq 2){
    
    $Identity = Read-Host "Digite o endereço email do usuário"
    $AdminCred = Get-Credential -Message "Digite os dados da conta do usuário com permissão de Exchange Organization Management"
    $DB = Read-Host "Por favor digite o nome do banco de dados do Exchange você deseja deixar a caixa do usuário"
    Write-host "Movendo a caixa do usuário $Identity para o Exchange On-Premises"
    New-MoveRequest -OutBound -Identity $Identity -RemoteTargetDatabase $DB -RemoteHostName 'webmail.grupoccr.com.br' -RemoteCredential $AdminCred -TargetDeliveryDomain 'grupoccr.com.br'
    Write-host "Para acompanhar o andamento da migração, utilize o seguinte cmdlet: Get-MoveRequestStatistics -Identity $Identity"
}
elseif($x -eq 3){
    
    $Identity = Read-Host "Digite o endereço email do usuário"
    Write-host "Finalizando a migração do usuário $Identity"
    Set-MoveRequest -Identity $Identity -CompleteAfter (Get-Date)
    Resume-MoveRequest -Identity $Identity
    Write-host "Para acompanhar o andamento da migração, utilize o seguinte cmdlet: Get-MoveRequestStatistics -Identity $Identity"
}
else{
    
    Write-host "
Selecione uma das opções acima!"
}
