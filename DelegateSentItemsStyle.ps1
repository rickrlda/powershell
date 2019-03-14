﻿<#
.SYNOPSIS
    Adiciona chave de registro Outlook


    .DESCRIPTION
    Este script tem com finalidade adicionar a chave de regsitro DelegateSentItemsStyle para usuários que tem permissão de SnedAs e 
    utilizam caixa compartilhada e enviam mensagem atraves dela.
    Script desenvolvido para uso exclusivo do Tribunal de Justiça de SP.
    
    
    Version 1.0, 11 de Março, 2019
    Autor: Ricardo Azevedo

#>

$OfficeVersion = Read-Host ("Digite a versão do Microsoft Office instalada nesse equipamento. 
Para a versão 2016 = 16.0
Para a versão 2013 = 15.0
Para a versão 2010 = 14.0
Digite aqui a versão")
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
    Write-Host (" Não foi possível localizar o caminho $registryPath porque ele não existe," +
        " por favor verifique se a versão que foi informada do Microsoft Office corresponde com a versão instalada neste computador.") -ForegroundColor red -BackgroundColor black
}
