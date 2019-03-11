#﻿<#
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
#    Version 1.0, 11 de Março, 2019
#    Autor: Ricardo Azevedo
#
#>

[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$OfficeVersion = [Microsoft.VisualBasic.Interaction]::InputBox("Type the version of Office installed (16.0 = Office 2016, 15.0 = Office 2013, 14.0 = Office 2010)", "Office Version")
$registryPath = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Outlook\Preferences"
$Name = "DelegateSentItemsStyle"
$value = "1"
New-ItemProperty -Path $registryPath -Name $name -Value $value `
    -PropertyType DWORD -Force | Out-Null
