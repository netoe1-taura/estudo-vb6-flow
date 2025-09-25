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
' Declaração obrigatória das variáveis usadas na classe.
Option Explicit

' Atributos da classe
Private pNome As String
Private pIdade As Integer

' Set
Public Property Let nome(nome As String)
    pNome = nome
End Property
Public Property Let idade(idade As Integer)
    pIdade = idade
End Property
' Get
Public Property Get idade() As Integer
    idade = pIdade
End Property

Public Property Get nome() As String
    nome = pNome
End Property

Public Sub MostrarDados()
    MsgBox "Nome: " & pNome & vbCrLf & "Idade: " & pIdade
End Sub




