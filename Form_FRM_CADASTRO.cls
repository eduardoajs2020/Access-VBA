VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FRM_CADASTRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public NumCod As Integer 'onde será gravado o código da sub rotina GerarCodigo

Private Sub cmdAlterar_Click()
'Trabalhando os eventos do botão alterar com as instruções SQL UPDATE

comando = "update TabCadGames set Titulo ='" & txtTitulo & "', Desenvolvedora='" & txtDesenvolvedora & "', Descricao='" & txtDescricao & "', Data = " & txtData & " where Codigo = " & txtCodigo  'sequencia para atualizar o código informado na consulta

banco.Execute (comando)

MsgBox ("Alteração realizada com Sucesso"), vbInfomation + vbOKOnly, "Confirmação de Alteração"

'neste ponto, vamos limpar as caixa de texto, depois vamos habilitar os botoes cadastro e consulta e desbilitar os botões alterar e excluir

zerarCx 'Metodo para limpar formulários

    txtCodigo.Enabled = True

    cmdAlterar.Enabled = False

    cmdExcluir.Enabled = False

    cmdCadastrar.Enabled = True

    cmdConsultar.Enabled = True


End Sub

Private Sub cmdCadastrar_Click()

    GerarCodigo 'estamos chamando o metodo que gera o codigo de cadastro
    comando = " Insert into TabCadGames (Codigo, Titulo, Desenvolvedora, Descricao, Data)values(" & NumCod & " , ' " & txtTitulo & " ' , ' " & txtDesenvolvedora & " ' , ' " & txtDescricao & " ' , " & txtData & " ) "
    banco.Execute (comando) 'executa as ações SQL da variavel comando, inserindo os dados da tabela virtual no banco de dados local
    MsgBox ("Game cadastrado com sucesso!"), vbInformation + vbOKOnly, "Cadastro"
    zerarCx 'metodo para limpar os formulários
End Sub


Private Sub cmdConsultar_Click()
    If txtCodigo <> "" Then

    comando = "Select * From TabCadGames where Codigo = " & txtCodigo 'seleciona todos os dados contidos no cadastro solicitado

    valida_selecao 'pega as informações solicitadas acima e preenche o dataset

'abaixo vamos verificar e validar o dataset

If dataset.RecordCount <> 0 Then

            
            txtTitulo = dataset("Titulo")
            txtDesenvolvedora = dataset("Desenvolvedora")
            txtData = dataset("Data")
            txtDescricao = dataset("Descricao")
            
            txtCodigo.Enabled = False
            
            'vamos habilitar e desabilitar os botões de comando conforme o resultado da pesquisa


            cmdAlterar.Enabled = True
            cmdExcluir.Enabled = True

            cmdCadastrar.Enabled = False
            cmdConsultar.Enabled = False
            
            Else 'caso nenhum registro seja encontrado
            
            MsgBox ("Não foram encontrados registros, por favor tente novamente."), vbInformation + vbOKOnly, "Nenhum Registro"
            
            End If
            
            Else
            MsgBox ("É necessário inserir um Código para realizar a pesquisa."), vbInformation + vbOKOnly, "Código Necessário"

    End If
    
    
    
End Sub

Private Sub cmdExcluir_Click()
'método excluir, vamos limpar os dados da tabela quando solicitado pelo usuário

'vamos criar uma variável que recebe a resposta do usuário
Dim resposta

    resposta = MsgBox("deseja realmente deletar o registro", vbQuestion + vbYesNo, "Exclusão de Registro")

If resposta = vbYes Then
    
    comando = "delete * from TabCadGames where Codigo = " & txtCodigo
    
    banco.Execute (comando)
    
    MsgBox ("Exclusão realizada com Sucesso"), vbInfomation + vbOKOnly, "Confirmação de Exclusão"
    
        zerarCx
    
    txtCodigo.Enabled = True
            
        cmdAlterar.Enabled = False
        
        cmdExcluir.Enabled = False
        
        cmdCadastrar.Enabled = True
        
        cmdConsultar.Enabled = True


Else
    
        zerarCx
        
    txtCodigo.Enabled = True
            
        cmdAlterar.Enabled = False
        
        cmdExcluir.Enabled = False
        
        cmdCadastrar.Enabled = True
        
        cmdConsultar.Enabled = True


End If

End Sub


Private Sub cmdLimpar_Click()

zerarCx

txtCodigo.Enabled = True
            
        cmdAlterar.Enabled = True
        
        cmdExcluir.Enabled = True
        
        cmdCadastrar.Enabled = True
        
        cmdConsultar.Enabled = True

End Sub

Private Sub cmdRelatorio_Click()
' Vamos editar as opções de relatório


Dim relatorio

relatorio = MsgBox("Clique SIM para imprimir ou NÃO para visualizar", vbQuestion + vbYesNo, "Opções de Relatório")

    If resposta = vbYes Then

        DoCmd.OpenReport "GamesCadastrados", acViewPreview 'para impressão

    Else
        
        DoCmd.OpenReport "GamesCadastrados", acViewReport 'apenas visualização

    
    End If
End Sub

Private Sub Form_Load()

    conecta 'Ao carregar, ele fará a conexão com o banco de dados local e os comandos que inserimos
    
End Sub

Private Sub GerarCodigo()

    comando = "select * from TabCadGames order by Codigo Desc" 'coloca os registros em ordem decrescente
    valida_selecao 'vamos utilizar para acessar a tabela virtual de acordo com a tabela do banco local
    
    If dataset.BOF = True Then 'se não houver registro o NumCod será igual a 1, caso exista, será o valor +1

        NumCod = 1

Else
        NumCod = dataset("Codigo") + 1

End If

End Sub

Sub zerarCx()

txtCodigo = Empty

txtTitulo = Empty

txtDesenvolvedora = Empty

txtDescricao = Empty

txtData = Empty


End Sub
