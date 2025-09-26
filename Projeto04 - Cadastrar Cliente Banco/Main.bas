Attribute VB_Name = "Main"
'------------------------------------Main.bas---------------------------------
' Nome: Ely Torres Neto
' Objetivo: Centralizar todas as fun��es principais do programa
' Data: 29/09/2025


'---------------------------------------Fun��o: void FecharForms()----------------------------------------------------'
' Essa fun��o � capaz de fechar todos os forms do programa. Geralmente,  o uso dele est� relacionado a problemas de erro.
Public Sub FecharForms()

    Dim f As Form
    ' Percorre todos os formul�rios abertos
    For Each f In Forms
        Unload f       ' Fecha o formul�rio
        Set f = Nothing ' Limpa da mem�ria
    Next f
    ' Encerra o programa
    End
    
End Sub

Public Sub Main()
    MenuPrincipal.Show()
End Sub

