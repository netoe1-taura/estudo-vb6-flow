Attribute VB_Name = "GlobalConsts"
' Nome: Ely Torres Neto
' Arquivo: GlobalConsts.bas
' Objetivo: Centralizar todas as defini��es de constantes em um arquivo;
' Data: 29/09/2025

' Defini��o de uma vari�vel global para configurarmos o enable.
' Quando falso, n�o d� display de mensagens no programa, � algo que precisamos definir na Sub Main()
Global ENABLE_DEBUG_MSG As Boolean

'----------------------------------------Fun��o: void EnableDebugMsg() -------------------------------------------------'
' Fun��o que controla o enable, serve para acessarmos a vari�vel global
' Par�metros: enable -> bool
' Se True, ir� modificar a vari�vel,por consequ�ncia, ativar as mensagens de Debug do sistema!
' Se false, realizar� o contr�rio, bloquear� as mensagens.
Public Sub EnableDebugMsg(enable As Boolean)
    ENABLE_DEBUG_MSG = enable
End Sub

'---------------------------------------Fun��o: void DebugPrint()----------------------------------------------------'
' O debug print ir� gerenciar todas essas defini��es, desde o Enable at� a mensagem
' Iremos adotar o padr�o de erro em uma string: ARQUIVO: Fun��o: Erro
' Ex:
' bash@foo:~ Pessoa.cls: mostrarPessoa(): A pessoa n�o existe!
'
Public Sub DebugPrint(file As String, fn As String, msg As String)
    If ENABLE_DEBUG_MSG = True Then
        if(file = )
        Debug.Print file & ":" & fn & "():" & msg
    End If
End Sub



