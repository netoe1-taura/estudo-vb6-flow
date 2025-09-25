Attribute VB_Name = "VariaveisGlobais"
Public pessoas(1 To 10) As Pessoa   ' Nosso Banco de dados improvisado: Um vetor de objetos, utilizado para salvar as pessoas.
Public ix As Integer                ' O index em que se encontra esse vetor de objetos
Private inited As Boolean           ' Flag para controlar as inst�ncia de objeto

' Search pessoa � uma fun��o que procura uma pessoa dentro do array global.
' Ela recebe um Nome para procurar, sendo uma String; retorna falso se achar a pessoa, dando display em um MsgBox.
' Se verdadeiro, ela mostra os dados da pessoa.

Public Function searchPessoa(Nome As String) As Boolean
    Debug.Print "searchPessoa():  Verifica se o parametro de busca iniciado na fun��o � v�lido"
    ' Verifica se o parametro de busca iniciado na fun��o � v�lido,
    ' e sai da fun��o, retornando "false" .
    
    If Nome = "" Then
        MsgBox "O nome para procurar � inv�lido!"
        searchPessoa = False
        Exit Function
    End If
    
    ' Inicia o contador para realizar uma busca sequencial.
    Debug.Print "searchPessoa(): Inicia o contador para realizar uma busca sequencial."
    Dim i As Integer
    i = 1
    
    ' Realiza uma busca sequencial no vetor de objetos global. Se acha, mostra o nome,
    ' atrav�s do m�todo, Exibir dados; depois, sai da fun��o, retornando true.
      Debug.Print "searchPessoa(): Realiza uma busca sequencial no vetor de objetos global."
    Do While i < 10
        If Nome = pessoas(i).Nome Then
            searchPessoa = True
            pessoas(i).ExibirDados
         
            Exit Function
        End If
        i = i + 1
    Loop
    
    
End Function

' Iniciar pessoa � um m�todo respos�vel por criar os par�metros de controles normais da classe.
' Ele vai iniciar o vetor de objetos pessoas, declarado acima; tamb�m ir� controlar se o nosso cursor de objeto estar� nos limites permitidos.

Public Sub iniciarPessoas()
    
    ' Inicia a pessoa, caso a sua flag de in�cio seja falsa. Quando Instanciada, ela � True e n�o instanciada.
    If inited = False Then
        ' Identifica que o objeto foi criado, e que n�o deve ser inicializado novamente.
        inited = True
        ' Seta o nosso cursor para 0.
        ix = 0
        
        ' Realiza o Print dos valores iniciais, apenas para debug.
        'Debug.Print "---INICIAR_PESSOAS---"
        'Debug.Print "Inited Value: " & inited
        'Debug.Print "Index Array:" & ix
        
        ' Para facilitar a instancia��o dos objetos, vamos utilizar um While.
        ' Criamos esse contador simples, come�a em 1 e termina em 0.
        
        Dim i As Integer
        i = 1
        
        ' Instancia a classe, agora virando um objeto.
        Do While i < 10
            Set pessoas(i) = New Pessoa
            i = i + 1
        Loop
    End If
    
End Sub
