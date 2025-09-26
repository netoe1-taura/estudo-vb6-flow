Attribute VB_Name = "mdlMain"
Option Explicit
'------------------------------------mdlMain.bas---------------------------------
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

Sub Main()

    SetDebugFlag (True) ' Ativa o módulo de debug!
    Call DebugPrint("mdlMain.bas", "Main", "Iniciando Programa!")
    frmMenuPrincipal.Show ' Abre o módulo principal do programa.
End Sub

