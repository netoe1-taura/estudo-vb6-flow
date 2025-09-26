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


Private OPEN_CNN As Boolean ' Variável para controle de banco de dados aberto'.
Public connstr As String    ' Variável para realizar a query'

'----------------------------------------Função: void DefinirCredenciaisRam() -------------------------------------------------'
' Função que encapsula o CREADENTIALS_CONN, definindo as credenciais para SQL SERVER.
' Ela aceita os parâmetros de forma direta, por isso têm o formato RAW.
' Parâmetros: buffer :string, que seria o buffer para conexão.
' Ex Uso:
' strconn = "Provider=SQLOLEDB;" & _
'            "Data Source=MEU_SERVIDOR\MINHA_INSTANCIA;" & _
'            "Initial Catalog=MEU_BANCO;" & _
'            "User ID=meuUsuario;" & _
'            "Password=minhaSenha;"
' DefinirCredenciaisRaw(strconn)


Public Sub DefinirCredenciaisRaw(buffer As String)
    On Error GoTo Err
    buffer = Trim(buffer) ' Realiza um trim das credenciais.
    
    If CREDENTIALS_CONNDB Is Nothing Then ' Verifica se a string de conexão foi iniciada.
    
         CREDENTIALS_CONNDB = buffer
         
     End If
Err:
 MsgBox "Ocorreu um erro: " & Err.Number & " - " & Err.Description
End Sub

'----------------------------------------Função: void DefinirCredenciais() -------------------------------------------------'
' Função mais formalizada para definir as credenciais do programa, em uma estrutura formal, para evitar erros.
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
                            userId As String, _
                            password As String)
                            
    On Error GoTo Err
    ' Realizando os Trim's nas strings!
    
    provider = Trim(provider)
    dataSource = Trim(dataSource)
    initialCatalog = Trim(initialCatalog)
    userId = Trim(userId)
    password = Trim(password)
Err:
 MsgBox "Ocorreu um erro: " & Err.Number & " - " & Err.Description
End Sub

'----------------------------------------Função: abrirConexao() -------------------------------------------------'
' Função para abrir conexão com o banco de dados. Deve ser usada após definir as credenciais!
' Ela irá salvar a conexão internamente na classe.
' Parâmetros: -
' Ex Uso:
' DefinirCredenciais("SQLOLEDB","MEU_BANCO","meuUsuario","minhaSenha")
' abrirConexao()

Public Sub abrirConexao()
    On Error GoTo Err
    
    If OPEN_CNN = True Then
    
        Call DebugPrint("dbClass.cls", "abrirConexao", "Você não pode abrir uma conexão já iniciada na mesma classe!")
        FecharForms
        Exit Sub
        
    End If
    
    If connstr Is Nothing Then ' Garante que a conexão nunca foi aberta!
        
            Call DebugPrint("dbClass.cls", "abrirConexao", "A string para as credenciais não foi completada.") ' Envia o Debug.Print Imediato!
            FecharForms ' Fecha todos os forms forçado!
            Exit Sub
    End If

Err:
 MsgBox "Ocorreu um erro: " & Err.Number & " - " & Err.Description
End Sub

'----------------------------------------Função: fecharConexao() -------------------------------------------------'
' Função para fechar conexão com o banco de dados. Deve ser usada após abrir conexão
' Ex Uso:
' DefinirCredenciais("SQLOLEDB","MEU_BANCO","meuUsuario","minhaSenha")
' abrirConexao()
' ...codigo...
' fecharConexao()

Public Sub FecharConexao()
    On Error GoTo Err
    If OPEN_CNN = False Then ' Verifica se a conexão já foi fechada
    
        Call DebugPrint("dbClass.cls", "fecharConexao", "Você não pode fechar uma conexão que já foi fechada ou nunca iniciada!") ' Mensagem para Immediate.
        FecharForms ' Fecha o programa.
        Exit Sub
        
    End If
        
    If connectionDb_gl.State = adStateOpen Then connectionDb_gl.Close ' Verifica se o utilitário do banco de dados fechou a conexão
    
        CREDENTIALS_CONNDB = ""
        Set cn = Nothing ' libera memória
        
    End If
    
Err:
 MsgBox "Ocorreu um erro: " & Err.Number & " - " & Err.Description
End Sub









