Attribute VB_Name = "ModuloConexao"
Option Compare Database

Public comando As String 'vari�vel onde s�o colocados os comandos SQL como consultar, inserir, atualizar e etc.

Public banco As Database 'variavel que ir� fazer a conex�o da aplica��o com o banco de dados permitindo inserir, deletar e alterar (insert, delete e update respectivamente)

Public dataset As Recordset 'variavel que permite acesso a tabela do banco de dados pela mem�ria ram, sempre que for necess�rio manipular tabelas(Gera um rascunho virtual do BD)


Function conecta()
    
    Set banco = CurrentDb 'inicializa a vari�vel banco, para a conex�o com o banco local

End Function

Function valida_selecao()
    
    'Set dataset = CurrentDb.OpenRecordset(comando, dbOpenDynaset)
    Set dataset = banco.OpenRecordset(comando, dbOpenDynaset) 'inicializa o dataset, possibilitando executar comandos SQL por interm�dio da vari�vel Comando, ou seja, d� acesso para altera��es gerais
    
End Function




