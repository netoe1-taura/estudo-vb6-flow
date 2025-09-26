Attribute VB_Name = "ControleDB"
Global cn As ADODB.Connection
Global rs As ADODB.Recordset
Global SQL_QUERY As String
Global STATUS_CONN As Boolean
Global ENABLE_DEBUG_MSG As Boolean
' Nome: Ely Torres Neto
' Data: '26/06/2025'
' Sub: AbrirConexão
' Objetivo: Criar uma função capaz de cuidar da conexão com o banco de dados, de forma autônoma,
'           tratando possíveis erros.
Public Sub AbrirConexao()
    'O If força o programa a apenas abrir uma vez.
    If STATUS_CONN = True Then
        
        Debug.Print ("ControleDB:AbrirConexao(): A conexão já foi iniciada, abortando...");
        
    End If
End Sub


