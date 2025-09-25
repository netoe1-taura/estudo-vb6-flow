VERSION 5.00
Begin VB.Form FormVerPessoa 
   Caption         =   "VerPessoa"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox inputNome 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton cancelarBtn 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton confirmarBtn 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Digite o nome da pessoa:"
      Height          =   555
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "FormVerPessoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelarBtn_Click()
    Unload Me
End Sub

