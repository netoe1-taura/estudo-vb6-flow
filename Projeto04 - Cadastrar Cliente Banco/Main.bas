Attribute VB_Name = "Main"
'------------------------------------Main.bas---------------------------------
' Nome: Ely Torres Neto
' Objetivo: Centralizar todas as funções principais do programa
' Data: 29/09/2025


'---------------------------------------Função: void FecharForms()----------------------------------------------------'
' Essa função é capaz de fechar todos os forms do programa. Geralmente,  o uso dele está relacionado a problemas de erro.
Public Sub FecharForms()

    Dim f As Form
    ' Percorre todos os formulários abertos
    For Each f In Forms
        Unload f       ' Fecha o formulário
        Set f = Nothing ' Limpa da memória
    Next f
    ' Encerra o programa
    End
    
End Sub

Public Sub Main()
    MenuPrincipal.Show()
End Sub

