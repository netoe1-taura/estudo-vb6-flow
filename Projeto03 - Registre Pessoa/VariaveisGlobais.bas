Attribute VB_Name = "VariaveisGlobais"
Public pessoas(1 To 10) As Pessoa
Dim inited As Boolean

Public Sub iniciarPessoas()
    If inited = False Then
        inited = True
        Dim i As Integer
        i = 1
        Do While i < 10
            Set pessoas(i) = New Pessoa
        Loop
    End If
    
End Sub
