Attribute VB_Name = "Debug"
'------------------------------------Debug.bas---------------------------------
' Nome: Ely Torres Neto
' Objetivo: Centralizar todas as defini��es e fun��es pertencentes aos modulos Debug.
' Data: 29/09/2025


'----------------------------------------Fun��o: void SetDebugFlag() -------------------------------------------------'
' Fun��o que controla o enable, serve para acessarmos a vari�vel global
' Par�metros: enable -> bool
' Se True, ir� modificar a vari�vel,por consequ�ncia, ativar as mensagens de Debug do sistema!
' Se false, realizar� o contr�rio, bloquear� as mensagens.
' Input: SetDebugFlag()
' Output: None

Public Sub SetDebugFlag(enable As Boolean)
    On Error GoTo Err
    ENABLE_DEBUG_MSG = enable
Err:
 MsgBox "Ocorreu um erro: " & Err.Number & " - " & Err.Description
End Sub

'---------------------------------------Fun��o: void DebugPrint()----------------------------------------------------'
' O debug print ir� gerenciar todas essas defini��es, desde o Enable at� a mensagem
' Iremos adotar o padr�o de erro em uma string: ARQUIVO: Fun��o: Erro
' Ex uso:
' Input: DebugPrint("Pessoa.cls","mostrarPessoa","A pessoa n�o existe!")
' Output: bash@foo:~ Pessoa.cls: mostrarPessoa(): A pessoa n�o existe!

Public Sub DebugPrint(file As String, fn As String, msg As String)
    
    On Error GoTo Err
    
    ' Verifica se a flag foi definida com sucesso globalmente.
    ' Caso n�o seja, o programa ir� fechar totalmente. � uma flag necess�ria para ser acessada.

    If ENABLE_DEBUG_MSG = False Then
    
        Debug.Print "GlobalConsts.bas:DebugPrint(): Voc� n�o definiu a vari�vel global de forma correta."
        FecharForms ' Fun��o declarada e documentada em Main.bas
        
    End If

    ' Caso a flag seja verdadeira
    
    If ENABLE_DEBUG_MSG = True Then
    
        Debug.Print file & ":" & fn & "():" & msg
        
    End If
    
Err:
 MsgBox "Ocorreu um erro: " & Err.Number & " - " & Err.Description
End Sub
