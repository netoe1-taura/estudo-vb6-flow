Attribute VB_Name = "Module1"
Option Explicit

Public pessoaObj As pessoa
Public inst As Boolean

Public Sub Init()
    ' s� cria uma vez
    If pessoaObj Is Nothing Then
        Set pessoaObj = New pessoa
    End If
End Sub
