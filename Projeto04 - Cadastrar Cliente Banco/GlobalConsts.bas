Attribute VB_Name = "GlobalConsts"
' Nome: Ely Torres Neto
' Arquivo: GlobalConsts.bas
' Objetivo: Centralizar todas as definições de constantes em um arquivo;
' Data: 29/09/2025

' Definição de uma variável global para configurarmos o enable.
' Quando falso, não dá display de mensagens no programa, é algo que precisamos definir na Sub Main()
Global ENABLE_DEBUG_MSG As Boolean

'----------------------------------------Função: void EnableDebugMsg() -------------------------------------------------'
' Função que controla o enable, serve para acessarmos a variável global
' Parâmetros: enable -> bool
' Se True, irá modificar a variável,por consequência, ativar as mensagens de Debug do sistema!
' Se false, realizará o contrário, bloqueará as mensagens.
Public Sub EnableDebugMsg(enable As Boolean)
    ENABLE_DEBUG_MSG = enable
End Sub

'---------------------------------------Função: void DebugPrint()----------------------------------------------------'
' O debug print irá gerenciar todas essas definições, desde o Enable até a mensagem
' Iremos adotar o padrão de erro em uma string: ARQUIVO: Função: Erro
' Ex:
' bash@foo:~ Pessoa.cls: mostrarPessoa(): A pessoa não existe!
'
Public Sub DebugPrint(file As String, fn As String, msg As String)
    If ENABLE_DEBUG_MSG = True Then
        if(file = )
        Debug.Print file & ":" & fn & "():" & msg
    End If
End Sub



