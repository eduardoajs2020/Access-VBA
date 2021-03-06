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
Public NumCod As Integer 'onde ser� gravado o c�digo da sub rotina GerarCodigo

Private Sub cmdAlterar_Click()
'Trabalhando os eventos do bot�o alterar com as instru��es SQL UPDATE

comando = "update TabCadGames set Titulo ='" & txtTitulo & "', Desenvolvedora='" & txtDesenvolvedora & "', Descricao='" & txtDescricao & "', Data = " & txtData & " where Codigo = " & txtCodigo  'sequencia para atualizar o c�digo informado na consulta

banco.Execute (comando)

MsgBox ("Altera��o realizada com Sucesso"), vbInfomation + vbOKOnly, "Confirma��o de Altera��o"

'neste ponto, vamos limpar as caixa de texto, depois vamos habilitar os botoes cadastro e consulta e desbilitar os bot�es alterar e excluir

zerarCx 'Metodo para limpar formul�rios

    txtCodigo.Enabled = True

    cmdAlterar.Enabled = False

    cmdExcluir.Enabled = False

    cmdCadastrar.Enabled = True

    cmdConsultar.Enabled = True


End Sub

Private Sub cmdCadastrar_Click()

    GerarCodigo 'estamos chamando o metodo que gera o codigo de cadastro
    comando = " Insert into TabCadGames (Codigo, Titulo, Desenvolvedora, Descricao, Data)values(" & NumCod & " , ' " & txtTitulo & " ' , ' " & txtDesenvolvedora & " ' , ' " & txtDescricao & " ' , " & txtData & " ) "
    banco.Execute (comando) 'executa as a��es SQL da variavel comando, inserindo os dados da tabela virtual no banco de dados local
    MsgBox ("Game cadastrado com sucesso!"), vbInformation + vbOKOnly, "Cadastro"
    zerarCx 'metodo para limpar os formul�rios
End Sub


Private Sub cmdConsultar_Click()
    If txtCodigo <> "" Then

    comando = "Select * From TabCadGames where Codigo = " & txtCodigo 'seleciona todos os dados contidos no cadastro solicitado

    valida_selecao 'pega as informa��es solicitadas acima e preenche o dataset

'abaixo vamos verificar e validar o dataset

If dataset.RecordCount <> 0 Then

            
            txtTitulo = dataset("Titulo")
            txtDesenvolvedora = dataset("Desenvolvedora")
            txtData = dataset("Data")
            txtDescricao = dataset("Descricao")
            
            txtCodigo.Enabled = False
            
            'vamos habilitar e desabilitar os bot�es de comando conforme o resultado da pesquisa


            cmdAlterar.Enabled = True
            cmdExcluir.Enabled = True

            cmdCadastrar.Enabled = False
            cmdConsultar.Enabled = False
            
            Else 'caso nenhum registro seja encontrado
            
            MsgBox ("N�o foram encontrados registros, por favor tente novamente."), vbInformation + vbOKOnly, "Nenhum Registro"
            
            End If
            
            Else
            MsgBox ("� necess�rio inserir um C�digo para realizar a pesquisa."), vbInformation + vbOKOnly, "C�digo Necess�rio"

    End If
    
    
    
End Sub

Private Sub cmdExcluir_Click()
'm�todo excluir, vamos limpar os dados da tabela quando solicitado pelo usu�rio

'vamos criar uma vari�vel que recebe a resposta do usu�rio
Dim resposta

    resposta = MsgBox("deseja realmente deletar o registro", vbQuestion + vbYesNo, "Exclus�o de Registro")

If resposta = vbYes Then
    
    comando = "delete * from TabCadGames where Codigo = " & txtCodigo
    
    banco.Execute (comando)
    
    MsgBox ("Exclus�o realizada com Sucesso"), vbInfomation + vbOKOnly, "Confirma��o de Exclus�o"
    
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
' Vamos editar as op��es de relat�rio


Dim relatorio

relatorio = MsgBox("Clique SIM para imprimir ou N�O para visualizar", vbQuestion + vbYesNo, "Op��es de Relat�rio")

    If resposta = vbYes Then

        DoCmd.OpenReport "GamesCadastrados", acViewPreview 'para impress�o

    Else
        
        DoCmd.OpenReport "GamesCadastrados", acViewReport 'apenas visualiza��o

    
    End If
End Sub

Private Sub Form_Load()

    conecta 'Ao carregar, ele far� a conex�o com o banco de dados local e os comandos que inserimos
    
End Sub

Private Sub GerarCodigo()

    comando = "select * from TabCadGames order by Codigo Desc" 'coloca os registros em ordem decrescente
    valida_selecao 'vamos utilizar para acessar a tabela virtual de acordo com a tabela do banco local
    
    If dataset.BOF = True Then 'se n�o houver registro o NumCod ser� igual a 1, caso exista, ser� o valor +1

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
