Attribute VB_Name = "mdlMain"
Option Explicit
'------------------------------------mdlMain.bas---------------------------------
' Nome: Ely Torres Neto
' Objetivo: Centralizar todas as fun��es principais do programa
' Data: 29/09/2025


Sub Main()

    SetDebugFlag (True) ' Ativa o m�dulo de debug!
    Call DebugPrint("mdlMain.bas", "Main", "Iniciando Programa!")
    frmMenuPrincipal.Show ' Abre o m�dulo principal do programa.
End Sub

