Attribute VB_Name = "ControleDB"
Global cn As ADODB.Connection
Global rs As ADODB.Recordset
Global SQL_QUERY As String
Global STATUS_CONN As Boolean
Global ENABLE_DEBUG_MSG As Boolean
' Nome: Ely Torres Neto
' Data: '26/06/2025'
' Sub: AbrirConex�o
' Objetivo: Criar uma fun��o capaz de cuidar da conex�o com o banco de dados, de forma aut�noma,
'           tratando poss�veis erros.
Public Sub AbrirConexao()
    'O If for�a o programa a apenas abrir uma vez.
    If STATUS_CONN = True Then
        
        Debug.Print ("ControleDB:AbrirConexao(): A conex�o j� foi iniciada, abortando...");
        
    End If
End Sub


