VERSION 5.00
Begin VB.Form FormAddPessoa 
   Caption         =   "Adicionar nova pessoa"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cancelarBtn 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton confirmarBtn 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox inputSobrenome 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox inputNome 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "SOBRENOME"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label labelNome 
      Caption         =   "NOME:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "FormAddPessoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cancelarBtn_Click()
    Unload Me
End Sub

Private Sub confirmarBtn_Click()
    
End Sub
