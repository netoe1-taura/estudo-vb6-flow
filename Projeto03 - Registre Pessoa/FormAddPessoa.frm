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

' Nesse form load, temos que garantir que os dados sejam resetados.
' Ele pega os campos de form e zera o valores deles, forçadamente, para evitar problemas.

Private Sub Form_Load()
 FormAddPessoa.inputNome = ""
 FormAddPessoa.inputSobrenome = ""
End Sub


Private Sub cancelarBtn_Click()
    Unload Me
End Sub

' Confirmar_Btn é uma função que irá pegar os dados digitados nos inputNome e inputSobrenome,
' salvar na classe global; ele realiza as validações principais.

Private Sub confirmarBtn_Click()
    Debug.Print "confirmarBtn_Click(): Entrou na função"
    ' Validando para ver se o formulário está vazio '
    If FormAddPessoa.inputNome = "" Or FormAddPessoa.inputSobrenome = "" Then
        MsgBox "Você não pode iniciar um input vazio!"
        Exit Sub
    End If
    
    
    ' Confirmação de Índice de pessoas no sistema:
    'Debug.Print "-----confirmarBtn_Click-----"
    'Debug.Print "Valor de ix global:" & ix
    
    Debug.Print "confirmarBtn_Click(): Verificando se o registro está dentro do range do vetor"
    ' Verificando se o registro está dentro do range do vetor
    If ix >= 10 Then
        MsgBox "Você não pode registrar mais pessoas!"
        Unload Me
    Exit Sub
    End If
    
    ' Verifica se já existe uma pessoa no mesmo nome.
    ' Esse sistema apenas suporta nomes diferentes.
    'Verificando se o registro está dentro do range do vetor
    
    Debug.Print "confirmarBtn_Click(): Verifica se já existe uma pessoa no mesmo nome."
    If searchPessoa(FormAddPessoa.inputNome) = True Then
        MsgBox "Já existe uma pessoa com esse nome!"
        Exit Sub
    End If
    
    ' Incrementando índice global, verificando se é possível realizar o incremento.
    ' Na ideia, ele sempre acesso o espaço
    
    Debug.Print "confirmarBtn_Click(): Incrementando índice global, verificando se é possível realizar o incremento."
    If ix + 1 < 10 Then
        ix = ix + 1
    End If
    
    ' Adicionando pessoa:
    Debug.Print "confirmarBtn_Click(): Criando pessoa."
    pessoas(ix).Nome = FormAddPessoa.inputNome
    pessoas(ix).Sobrenome = FormAddPessoa.inputSobrenome
    pessoas(ix).ExibirDados
    Unload Me
End Sub

