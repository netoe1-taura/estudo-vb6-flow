Attribute VB_Name = "ControleDB"
'------------------------------------ControleDB.bas---------------------------------
' Nome: Ely Torres Neto
' Objetivo: Centralizar todas as funções de controle de banco de dados, de forma simples
' Data: 29/09/2025

' Definindo objetos globais para iniciar o projeto
Global connectionDb_gl As ADODB.Connection
Global recordSet_gl As ADODB.Recordset
Global CREDENTIALS_CONNDB As String

'----------------------------------------Função: void DefinirCredenciais() -------------------------------------------------'
' Função que encapsula o CREADENTIALS_CONN, definindo as credenciais para SQL SERVER.
' Parâmetros: buffer :string, que seria o buffer para conexão.
' Ex Uso:
' strconn = "Provider=SQLOLEDB;" & _
'            "Data Source=MEU_SERVIDOR\MINHA_INSTANCIA;" & _
'            "Initial Catalog=MEU_BANCO;" & _
'            "User ID=meuUsuario;" & _
'            "Password=minhaSenha;"
' DefinirCredenciais(strconn)


Public Sub DefinirCredenciais(buffer As String)

   If CREDENTIALS_CONNDB Is Nothing Then
   
        CREEDENTIALS_CONNDB = buffer
        
    End If
    
End Sub

Public Sub abrirConexao()
    
End Sub

Public Sub FecharConexao()

End Sub







