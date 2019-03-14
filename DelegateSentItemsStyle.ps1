#<#
#.SYNOPSIS
#    Adiciona chave de registro Outlook
#
#
#    .DESCRIPTION
#    Este script tem com finalidade adicionar a chave de regsitro DelegateSentItemsStyle para usuários que tem permissão de SnedAs e 
#    utilizam caixa compartilhada e enviam mensagem atraves dela.
#    Script desenvolvido para uso exclusivo do Tribunal de Justiça de SP.
#    
#    
#    Version Created: 1.0, 11 de Março, 2019
#    Version Updated (IF and ELSEIF): 1.1, 14 de Março, 2019
#    Autor: Ricardo Azevedo
#    KB Microsoft: https://support.microsoft.com/en-us/help/2843677/messages-sent-from-a-shared-mailbox-aren-t-saved-to-the-sent-items-fol
#
#>

[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$OfficeVersion = [Microsoft.VisualBasic.Interaction]::InputBox("Digite a versão do Microsoft Office instalada nesse equipamento (16.0 = Office 2016, 15.0 = Office 2013, 14.0 = Office 2010)", "Versão do MS Office")
$registryPath = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Outlook\Preferences"
$Name = "DelegateSentItemsStyle"
$value = "1"

IF ((Test-Path $registryPath) -eq $true)
    {
    New-ItemProperty -Path $registryPath -Name $name -Value $value `
    -PropertyType DWORD -Force | Out-Null
    Write-Host "Chave de registro criada com sucesso!" -ForegroundColor green
}
ELSE {
    Write-Host "P.S.:" -ForegroundColor yellow -BackgroundColor black -NoNewline
    Write-Host " Não foi possível localizar o caminho $registryPath porque ele não existe, por favor verifique se a versão que foi informada do Microsoft Office corresponde com a versão instalada neste computador." -ForegroundColor red -BackgroundColor black
}
