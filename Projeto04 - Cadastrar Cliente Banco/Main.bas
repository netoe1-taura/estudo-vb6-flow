Attribute VB_Name = "mdlMain"
Option Explicit
'------------------------------------mdlMain.bas---------------------------------
' Nome: Ely Torres Neto
' Objetivo: Centralizar todas as funções principais do programa
' Data: 29/09/2025


Sub Main()

    SetDebugFlag (True) ' Ativa o módulo de debug!
    Call DebugPrint("mdlMain.bas", "Main", "Iniciando Programa!")
    frmMenuPrincipal.Show ' Abre o módulo principal do programa.
End Sub

