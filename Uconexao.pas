unit Uconexao;

//  rel_1,  Administração -> Amaro
//  rel_2,  Administração -> NF Difal
//  rel_3,  Veículos -> Faturamento
//  rel_4,  Veículos -> Vendas Perdidas
//  rel_5,  Veículos -> Leads
//  rel_6,  Veículos -> Fluxo de Loja(Atendimentos)
//  rel_7,  Veículos -> Estatísticas
//  rel_8,  Veículos -> Ações CRM
//  rel_9,  Administração -> ICMS ST
//  rel_10, Teste
//  rel_11; Veículos -> Dashboard

interface

//Aqui iremos instanciar a biblioteca ADODB.

uses ADODB;

//Aqui iremos declarar nossa procedure de conexão como publica para as outras forms.

procedure conectar();

var

//Agora vamos declarar duas variáveis, uma para conexão (TADOConnection) e outra para consultas (TADODataSet);

oConexao:TADOConnection;

oConsulta:TADODataSet;

implementation

//Criaremos uma procedure de conexão

procedure conectar();

Begin

//Instancia a variável de conexão

oConexao:=TADOConnection.create(nil);

//Declara onde está localizado o BD (SQL Server Express)

oConexao.ConnectionString:='Provider=MSDAORA.1;User ID=cnp;Data Source=bravos';

//Declara onde está localizado o BD (Microsoft Access)

//oConexao.ConnectionString:=’Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:bancosBDTeste.mdb;Persist Security Info=False’;

//Abre o banco de dados com o método Open() da variável de conexão

oConexao.Open();

//Este comando só é executado quando utilizar BD SQL Server Express, para que o a identificação das aspas duplas seja //desabilitada.

oConexao.Execute('Set Quoted_Identifier Off');

//Instancia a variável de conexão

oConsulta:= TADODataSet.Create(nil);

//Declara a localização do BD utilizando o caminho da variável de conexão

oConsulta.Connection := oConexao;

end;

end.
 