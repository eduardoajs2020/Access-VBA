Attribute VB_Name = "ModuloConexao"
Option Compare Database

Public comando As String 'variável onde são colocados os comandos SQL como consultar, inserir, atualizar e etc.

Public banco As Database 'variavel que irá fazer a conexão da aplicação com o banco de dados permitindo inserir, deletar e alterar (insert, delete e update respectivamente)

Public dataset As Recordset 'variavel que permite acesso a tabela do banco de dados pela memória ram, sempre que for necessário manipular tabelas(Gera um rascunho virtual do BD)


Function conecta()
    
    Set banco = CurrentDb 'inicializa a variável banco, para a conexão com o banco local

End Function

Function valida_selecao()
    
    'Set dataset = CurrentDb.OpenRecordset(comando, dbOpenDynaset)
    Set dataset = banco.OpenRecordset(comando, dbOpenDynaset) 'inicializa o dataset, possibilitando executar comandos SQL por intermédio da variável Comando, ou seja, dá acesso para alterações gerais
    
End Function




