Attribute VB_Name = "Debug"
'------------------------------------Debug.bas---------------------------------
' Nome: Ely Torres Neto
' Objetivo: Centralizar todas as definições e funções pertencentes aos modulos Debug.
' Data: 29/09/2025


'----------------------------------------Função: void SetDebugFlag() -------------------------------------------------'
' Função que controla o enable, serve para acessarmos a variável global
' Parâmetros: enable -> bool
' Se True, irá modificar a variável,por consequência, ativar as mensagens de Debug do sistema!
' Se false, realizará o contrário, bloqueará as mensagens.
' Input: SetDebugFlag()
' Output: None

Public Sub SetDebugFlag(enable As Boolean)
    On Error GoTo Err
    ENABLE_DEBUG_MSG = enable
Err:
 MsgBox "Ocorreu um erro: " & Err.Number & " - " & Err.Description
End Sub

'---------------------------------------Função: void DebugPrint()----------------------------------------------------'
' O debug print irá gerenciar todas essas definições, desde o Enable até a mensagem
' Iremos adotar o padrão de erro em uma string: ARQUIVO: Função: Erro
' Ex uso:
' Input: DebugPrint("Pessoa.cls","mostrarPessoa","A pessoa não existe!")
' Output: bash@foo:~ Pessoa.cls: mostrarPessoa(): A pessoa não existe!

Public Sub DebugPrint(file As String, fn As String, msg As String)
    
    On Error GoTo Err
    
    ' Verifica se a flag foi definida com sucesso globalmente.
    ' Caso não seja, o programa irá fechar totalmente. É uma flag necessária para ser acessada.

    If ENABLE_DEBUG_MSG = False Then
    
        Debug.Print "GlobalConsts.bas:DebugPrint(): Você não definiu a variável global de forma correta."
        FecharForms ' Função declarada e documentada em Main.bas
        
    End If

    ' Caso a flag seja verdadeira
    
    If ENABLE_DEBUG_MSG = True Then
    
        Debug.Print file & ":" & fn & "():" & msg
        
    End If
    
Err:
 MsgBox "Ocorreu um erro: " & Err.Number & " - " & Err.Description
End Sub
