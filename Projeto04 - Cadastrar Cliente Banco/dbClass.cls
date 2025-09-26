VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dbClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------dbClass.cls---------------------------------
' Nome: Ely Torres Neto
' Objetivo: Centralizar todas as funções de controle de banco de dados, em formato de classe
' Data: 29/09/2025


'----------------------------------------Função: void DefinirCredenciaisRam() -------------------------------------------------'
' Função que encapsula o CREADENTIALS_CONN, definindo as credenciais para SQL SERVER.
' Parâmetros: buffer :string, que seria o buffer para conexão.
' Ex Uso:
' strconn = "Provider=SQLOLEDB;" & _
'            "Data Source=MEU_SERVIDOR\MINHA_INSTANCIA;" & _
'            "Initial Catalog=MEU_BANCO;" & _
'            "User ID=meuUsuario;" & _
'            "Password=minhaSenha;"
' DefinirCredenciais(strconn)

Private OPEN_CNN As Boolean
Public Sub DefinirCredenciaisRaw(buffer As String)

   If CREDENTIALS_CONNDB Is Nothing Then
   
        CREDENTIALS_CONNDB = buffer
        
    End If
    
End Sub

'----------------------------------------Função: void DefinirCredenciais() -------------------------------------------------'
' Função mais formalizada para definir as credenciais do programa, com validações, evitando erros.
' Parâmetros:   provider        :   string,
'               dataSource      :   string,
'               provider        :   string,
'               dataSource      :   string,
'               initialCatalog  :   string,
'               UserId          :   string,
'               password        :   string)
' Ex Uso:
' DefinirCredenciais("SQLOLEDB","MEU_BANCO","meuUsuario","minhaSenha")


Public Sub DefinirCredenciais(provider As String, _
                            dataSource As String, _
                            initialCatalog As String, _
                            UserId As String, _
                            password As String)
                            
    

End Sub

Public Sub abrirConexao()

    If OPEN_CNN = False Then
    
        
    
    End If
    
End Sub

Public Sub FecharConexao()
    If OPEN_CNN = True Then
    
        CREDENTIALS_CONNDB = ""
    
    End If
    
    
    
End Sub









