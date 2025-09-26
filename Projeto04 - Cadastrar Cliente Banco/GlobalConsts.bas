Attribute VB_Name = "GlobalConsts"
'------------------------------------GlobalConsts.bas---------------------------------
' Nome: Ely Torres Neto
' Objetivo: Centralizar todas as defini��es de constantes e fun��es globais em um arquivo
' Data: 29/09/2025

Public ENABLE_DEBUG_MSG As Boolean
' Defini��o de uma vari�vel global para configurarmos o enable.
' Quando falso, n�o d� display de mensagens no programa, � algo que precisamos definir na Sub Main().
' Voc� usar� a fun��o do m�dulo de Debug.bas, SetDebugFlag(), para acessar essa fun��o


'---------------------------------------Fun��o: void FecharForms()----------------------------------------------------'
' Essa fun��o � capaz de fechar todos os forms do programa e encerra o programa de forma for�ada. Geralmente, o uso dele est� relacionado a problemas de erro.
Public Sub FecharForms()
    On Error GoTo Err
    Dim f As Form
    ' Percorre todos os formul�rios abertos
    For Each f In Forms
        Unload f       ' Fecha o formul�rio
        Set f = Nothing ' Limpa da mem�ria
    Next f
    ' Encerra o programa
    End
Err:
 MsgBox "Ocorreu um erro: " & Err.Number & " - " & Err.Description
End Sub






