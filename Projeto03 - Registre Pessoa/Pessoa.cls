VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pessoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Nome As String      ' Define o nome do cliente em tipo String.
Public Sobrenome As String ' Define o sobrenome do cliente em tipo String.
    
    
'Esse método realizar o print dos dados em uma msgbox.
Public Sub ExibirDados()
    Debug.Print "ExibirDados(): Imprimindo dados de um cliente: " & vbClFr & "Nome: " & Nome & vbClFr & "Sobrenome: " & Sobrenome
    MsgBox "Nome: " & Nome & vbClFr & "Sobrenome: " & Sobrenome
End Sub


