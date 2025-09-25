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
' Seta o valor do Form_Load para zero, resetando
Private Sub Form_Load()
    FormVerPessoa.inputNome = ""
End Sub
' Fecha a janela do form
Private Sub cancelarBtn_Click()
    Unload Me
End Sub
' Essa fun��o � sobre o bot�o de confirmar.
' Ap�s receber um nome no inputNome, ela procura por uma pessoa com aquele nome.
' Se n�o cadastrada, d� um MsgBox. Se cadastrada, o pr�prio m�todo Pessoa.ExibeDados � executado internamente do m�todo searchPessoa

Private Sub confirmarBtn_Click()
    ' Verifica se o input est� vazio.
    
    If FormVerPessoa.inputNome = "" Then
        MsgBox "O campo de busca de nome, n�o pode ser inv�lido!"
        Exit Sub
    End If
    
    ' Verifica se a pessoa foi encontrada.
    ' Se n�o, imprime uma mensagem de pessoa n�o encontrada.
    If searchPessoa(FormVerPessoa.inputNome) = False Then
        MsgBox "Nenhuma pessoa foi encontrada!"
    End If
    
End Sub
