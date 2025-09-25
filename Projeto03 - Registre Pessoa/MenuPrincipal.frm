VERSION 5.00
Begin VB.Form FormPrincipal 
   Caption         =   "FormPrincipal"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton sairBtn 
      Caption         =   "Sair"
      Height          =   435
      Left            =   1680
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton verPessoaBtn 
      Caption         =   "Ver Pessoa"
      Height          =   855
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cadastrarPessoaBtn 
      Caption         =   "Cadastrar pessoa"
      Height          =   855
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label labelMenuPrincipal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Menu Principal"
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   1065
   End
End
Attribute VB_Name = "FormPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cadastrarPessoaBtn_Click(Index As Integer)
    FormAddPessoa.Show
End Sub

Private Sub sairBtn_Click()

End Sub

Private Sub verPessoaBtn_Click(Index As Integer)
    FormVerPessoa.Show
End Sub
