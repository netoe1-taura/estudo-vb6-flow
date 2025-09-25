VERSION 5.00
Begin VB.Form SetNome 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnVerNome 
      Caption         =   "Ver Nome"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton btnSubmit 
      Caption         =   "Submeter"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox inputNome 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Text            =   "Digite o seu nome:"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblEntradaNome 
      Alignment       =   2  'Center
      Caption         =   "Digite o seu nome:"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "SetNome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnVerNome_Click()
    If inputNome.Text = "" Then
        Exit Sub
    End If
    SetNome.Show
End Sub

Private Sub Command1_Click()
    SetNome.Show
End Sub

Private Sub Form_Load()
    inputNome.Text = ""
End Sub

Private Sub btnSubmit_Click()
    
    If inputNome.Text = "" Then
        MsgBox "Você precisa digitar algum nome!"
        Exit Sub
    End If
    
    MsgBox "O seu nome é " & inputNome.Text
    
    nome = inputNome.Text
End Sub


