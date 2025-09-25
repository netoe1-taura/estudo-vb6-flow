VERSION 5.00
Begin VB.Form FormGremio 
   Caption         =   "Form2"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnVoltar 
      Caption         =   "Voltar"
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label label_serieB 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "VOCÊ FOI REBAIXADO PARA A SÉRIE B!"
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   1200
      Width           =   3315
   End
End
Attribute VB_Name = "FormGremio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnVoltar_Click()
    Unload Me
End Sub

