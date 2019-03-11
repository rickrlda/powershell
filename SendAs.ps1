#<#
#.SYNOPSIS
#    Adiciona Permissão de SendAs
#
#
#    .DESCRIPTION
#    Este script tem com finalidade adicionar permissão de SendAs na caixa de correio compartilhada.
#    Script desenvolvido para uso exclusivo do Tribunal de Justiça de SP.
#    
#    
#    Version 1.0, 25 de Fevereiro, 2019
#    Autor: Ricardo Azevedo
#
#>

[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$Alias = [Microsoft.VisualBasic.Interaction]::InputBox("Entre com o alias ou email da Caixa Compartilhada", "SharedMailbox")
$Users = Import-Csv -Path C:\TEMP\O365\Permission.csv
foreach($Object in $Users){
    Add-RecipientPermission -Identity $Alias `
        -Trustee $Object.Alias `
        -AccessRights SendAs `
        -Confirm:$false -Verbose
    
}
