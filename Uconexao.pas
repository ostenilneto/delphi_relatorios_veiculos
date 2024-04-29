unit Uconexao;

//  rel_1,  Administra��o -> Amaro
//  rel_2,  Administra��o -> NF Difal
//  rel_3,  Ve�culos -> Faturamento
//  rel_4,  Ve�culos -> Vendas Perdidas
//  rel_5,  Ve�culos -> Leads
//  rel_6,  Ve�culos -> Fluxo de Loja(Atendimentos)
//  rel_7,  Ve�culos -> Estat�sticas
//  rel_8,  Ve�culos -> A��es CRM
//  rel_9,  Administra��o -> ICMS ST
//  rel_10, Teste
//  rel_11; Ve�culos -> Dashboard

interface

//Aqui iremos instanciar a biblioteca ADODB.

uses ADODB;

//Aqui iremos declarar nossa procedure de conex�o como publica para as outras forms.

procedure conectar();

var

//Agora vamos declarar duas vari�veis, uma para conex�o (TADOConnection) e outra para consultas (TADODataSet);

oConexao:TADOConnection;

oConsulta:TADODataSet;

implementation

//Criaremos uma procedure de conex�o

procedure conectar();

Begin

//Instancia a vari�vel de conex�o

oConexao:=TADOConnection.create(nil);

//Declara onde est� localizado o BD (SQL Server Express)

oConexao.ConnectionString:='Provider=MSDAORA.1;User ID=cnp;Data Source=bravos';

//Declara onde est� localizado o BD (Microsoft Access)

//oConexao.ConnectionString:=�Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:bancosBDTeste.mdb;Persist Security Info=False�;

//Abre o banco de dados com o m�todo Open() da vari�vel de conex�o

oConexao.Open();

//Este comando s� � executado quando utilizar BD SQL Server Express, para que o a identifica��o das aspas duplas seja //desabilitada.

oConexao.Execute('Set Quoted_Identifier Off');

//Instancia a vari�vel de conex�o

oConsulta:= TADODataSet.Create(nil);

//Declara a localiza��o do BD utilizando o caminho da vari�vel de conex�o

oConsulta.Connection := oConexao;

end;

end.
 