Attribute VB_Name = "GlobalConsts"
'------------------------------------GlobalConsts.bas---------------------------------
' Nome: Ely Torres Neto
' Objetivo: Centralizar todas as definições de constantes e funções globais em um arquivo
' Data: 29/09/2025

Public ENABLE_DEBUG_MSG As Boolean
' Definição de uma variável global para configurarmos o enable.
' Quando falso, não dá display de mensagens no programa, é algo que precisamos definir na Sub Main().
' Você usará a função do módulo de Debug.bas, SetDebugFlag(), para acessar essa função


'---------------------------------------Função: void FecharForms()----------------------------------------------------'
' Essa função é capaz de fechar todos os forms do programa e encerra o programa de forma forçada. Geralmente, o uso dele está relacionado a problemas de erro.
Public Sub FecharForms()
    On Error GoTo Err
    Dim f As Form
    ' Percorre todos os formulários abertos
    For Each f In Forms
        Unload f       ' Fecha o formulário
        Set f = Nothing ' Limpa da memória
    Next f
    ' Encerra o programa
    End
Err:
 MsgBox "Ocorreu um erro: " & Err.Number & " - " & Err.Description
End Sub






