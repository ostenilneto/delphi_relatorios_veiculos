unit rel_7;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, Mask, Grids, ComObj, DBGrids, ComCtrls,
  Vcl.ExtCtrls, DateUtils, Vcl.Buttons, Vcl.DBCtrls, Datasnap.DBClient, SimpleDS,
  Data.SqlExpr, Data.FMTBcd, Datasnap.Provider;

type
  TForm8 = class(TForm)
    Panel1: TPanel;
    Label3: TLabel;
    Label4: TLabel;
    Label1: TLabel;
    DateTimePicker3: TDateTimePicker;
    DateTimePicker4: TDateTimePicker;
    DBLookupComboBox1: TDBLookupComboBox;
    BitBtn1: TBitBtn;
    BitBtn4: TBitBtn;
    BitBtn2: TBitBtn;
    DataSource3: TDataSource;
    ADOQuery3: TADOQuery;
    ADOQuery3DEPARTAMENTO: TIntegerField;
    ADOQuery3NOME: TStringField;
    BitBtn3: TBitBtn;
    DataSource2: TDataSource;
    ADOQuery2: TADOQuery;
    ADOConnection1: TADOConnection;
    DataSource5: TDataSource;
    ADOQuery5: TADOQuery;
    DataSource6: TDataSource;
    ADOQuery6: TADOQuery;
    DataSource7: TDataSource;
    ADOQuery7: TADOQuery;
    ADOQuery8: TADOQuery;
    ADOQuery9: TADOQuery;
    ADOQuery10: TADOQuery;
    DataSource8: TDataSource;
    DataSource9: TDataSource;
    DataSource10: TDataSource;
    Dados: TPageControl;
    Atendimentos: TTabSheet;
    Label7: TLabel;
    DBGrid4: TDBGrid;
    TabSheet2: TTabSheet;
    Label5: TLabel;
    DBGrid6: TDBGrid;
    DBGrid2: TDBGrid;
    Faturamento: TTabSheet;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label2: TLabel;
    Label6: TLabel;
    DBEdit3: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit1: TDBEdit;
    DBEdit5: TDBEdit;
    DBEdit6: TDBEdit;
    DBGrid5: TDBGrid;
    DBEdit4: TDBEdit;
    DBEdit7: TDBEdit;
    ADOQuery1: TADOQuery;
    DataSource1: TDataSource;
    DataSource4: TDataSource;
    ADOQuery4: TADOQuery;
    ADOQuery11: TADOQuery;
    ADOQuery12: TADOQuery;
    ADOQuery13: TADOQuery;
    ADOQuery14: TADOQuery;
    Label11: TLabel;
    procedure FormShow(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure DimensionarGrid(dbg: TDBGrid);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure DBGrid4DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid6DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid2DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);


    

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form8: TForm8;

implementation
uses Princ;
{$R *.dfm}

{          **  Tabela  **
Query1 - Estat�ticas Total Ativos
Query2 | Query13 - Faturamento
Query3 - Filtro departamento
Query4 - Estat�ticas Total Passivos
Query5 | Query11 - Atendimentos
Query6 | Query14 - Estat�sticas Loja
Query7 | Query12 - Venda Perdida
Query8 - Estat�ticas Total Fluxo Total(Atendimentos)
Query9 - Estat�ticas Total Faturamento
Query10 - Estat�ticas Total Venda Perdida
                                     }


procedure TForm8.BitBtn1Click(Sender: TObject);
begin
  DBGrid2.DataSource.DataSet.Close;
  DBGrid4.DataSource.DataSet.Close;
  DBGrid5.DataSource.DataSet.Close;
  DBGrid6.DataSource.DataSet.Close;
end;

procedure TForm8.BitBtn2Click(Sender: TObject);
begin
  Form2.Panel12.Width := 150;
  Form2.Caption := 'PROGRAMA DE APOIO';
  Form8.Close;
  DBGrid2.DataSource.DataSet.Close;
  DBGrid4.DataSource.DataSet.Close;
  DBGrid5.DataSource.DataSet.Close;
  DBGrid6.DataSource.DataSet.Close;
end;

procedure TForm8.BitBtn3Click(Sender: TObject);
Var
linha, coluna: integer;
planilha: variant;
valorCampo: string;
begin
   planilha:= CreateOleObject('Excel.Application');
   planilha.Workbooks.add;
   planilha.WorkSheets[1].DisplayPageBreaks:=False;
   planilha.Cells.Select;
   planilha.Selection.NumberFormat := '@';
   planilha.caption:= 'Exporta��o de dados para o excel';
   planilha.visible:= true;
   AdoQuery5.First;
   for linha:= 0 to AdoQuery6.RecordCount-1 do
   begin
     for coluna:= 1 to AdoQuery6.FieldCount do
      begin
         valorCampo:= AdoQuery6.Fields[coluna-1].AsString;
         planilha.cells[linha+2,coluna]:= valorCampo;
      end;

     AdoQuery6.Next;

   end;
   for coluna:=1 to AdoQuery6.FieldCount do
   begin
      valorCampo:= AdoQuery6.Fields[coluna-1].DisplayLabel;
      planilha.cells[1,coluna]:= valorCampo;
   end;
   planilha.columns.AutoFit;

   planilha.Sheets.Add;
   planilha.Cells.Select;
   planilha.Selection.NumberFormat := '@';
   planilha.WorkSheets[2].DisplayPageBreaks:=False;
   AdoQuery2.First;
   for linha:= 0 to AdoQuery2.RecordCount-1 do
   begin
     for coluna:= 1 to AdoQuery2.FieldCount do
      begin
         valorCampo:= AdoQuery2.Fields[coluna-1].AsString;
         planilha.cells[linha+2,coluna]:= valorCampo;
      end;

     AdoQuery2.Next;

   end;
   for coluna:=1 to AdoQuery2.FieldCount do
   begin
      valorCampo:= AdoQuery2.Fields[coluna-1].DisplayLabel;
      planilha.cells[1,coluna]:= valorCampo;
   end;
   planilha.columns.AutoFit;

   planilha.Sheets.Add;
   planilha.Cells.Select;
   planilha.Selection.NumberFormat := '@';
   planilha.WorkSheets[3].DisplayPageBreaks:=False;
   AdoQuery7.First;
   for linha:= 0 to AdoQuery7.RecordCount-1 do
   begin
     for coluna:= 1 to AdoQuery7.FieldCount do
      begin
         valorCampo:= AdoQuery7.Fields[coluna-1].AsString;
         planilha.cells[linha+2,coluna]:= valorCampo;
      end;

     AdoQuery7.Next;

   end;
   for coluna:=1 to AdoQuery7.FieldCount do
   begin
      valorCampo:= AdoQuery7.Fields[coluna-1].DisplayLabel;
      planilha.cells[1,coluna]:= valorCampo;
   end;
   planilha.columns.AutoFit;

   planilha.Sheets.Add;
   planilha.Cells.Select;
   planilha.Selection.NumberFormat := '@';
   planilha.WorkSheets[4].DisplayPageBreaks:=False;
   AdoQuery5.First;
   for linha:= 0 to AdoQuery5.RecordCount-1 do
   begin
     for coluna:= 1 to AdoQuery5.FieldCount do
      begin
         valorCampo:= AdoQuery5.Fields[coluna-1].AsString;
         planilha.cells[linha+2,coluna]:= valorCampo;
      end;

     AdoQuery5.Next;

   end;
   for coluna:=1 to AdoQuery5.FieldCount do
   begin
      valorCampo:= AdoQuery5.Fields[coluna-1].DisplayLabel;
      planilha.cells[1,coluna]:= valorCampo;
   end;
   planilha.columns.AutoFit;

   planilha.WorkSheets[1].Name:='Atendimentos';
   planilha.WorkSheets[2].Name:='Vendas Perdidas';
   planilha.WorkSheets[3].Name:='Faturamento';
   planilha.WorkSheets[4].Name:='Estat�sticas';

end;

procedure TForm8.BitBtn4Click(Sender: TObject);
var
a, b, c, f, g:integer;
d, e: Double;

begin
  Screen.Cursor := crHourglass;
  Form8.ADOQuery1.Close;
  Form8.ADOQuery1.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
  Form8.ADOQuery1.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.Date);
  if DBLookupCombobox1.KeyValue = 100 then
  begin
      Form8.ADOQuery1.Parameters.ParamByName('departamento').Value := (100);
      Form8.ADOQuery1.Parameters.ParamByName('departamento2').Value := (110);
      Form8.ADOQuery1.Open;
      DimensionarGrid( DBGrid2 );
  end
  else
  begin
      Form8.ADOQuery1.Parameters.ParamByName('departamento').Value := (200);
      Form8.ADOQuery1.Parameters.ParamByName('departamento2').Value := (210);
      Form8.ADOQuery1.Open;
      DimensionarGrid( DBGrid2 );
  end;

    Form8.ADOQuery13.Close;
  Form8.ADOQuery2.Close;
  if DBLookupCombobox1.KeyValue = 100 then
  begin
      Form8.DataSource2.DataSet := ADOQuery2;
      Form8.ADOQuery2.Parameters.ParamByName('inicio').Value := DatetoStr			(DateTimePicker3.DateTime);
      Form8.ADOQuery2.Parameters.ParamByName('fim').Value := DatetoStr			(DateTimePicker4.DateTime);
      Form8.ADOQuery2.Open;
      DimensionarGrid( DBGrid2 );
  end
  else
  begin
      Form8.DataSource2.DataSet := ADOQuery13;
      Form8.ADOQuery13.Parameters.ParamByName('inicio').Value := DatetoStr			(DateTimePicker3.DateTime);
      Form8.ADOQuery13.Parameters.ParamByName('fim').Value := DatetoStr			(DateTimePicker4.DateTime);
      Form8.ADOQuery13.Open;
      DimensionarGrid( DBGrid2 );
  end;

    Form8.ADOQuery4.Close;
  Form8.ADOQuery4.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
  Form8.ADOQuery4.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.Date);
  if DBLookupCombobox1.KeyValue = 100 then
  begin
    Form8.ADOQuery4.Parameters.ParamByName('departamento').Value := (100);
    Form8.ADOQuery4.Parameters.ParamByName('departamento2').Value := (110);
      Form8.ADOQuery4.Open;
      DimensionarGrid( DBGrid2 );
  end
  else
  begin
      Form8.ADOQuery4.Parameters.ParamByName('departamento').Value := (200);
      Form8.ADOQuery4.Parameters.ParamByName('departamento2').Value := (210);
      Form8.ADOQuery4.Open;
      DimensionarGrid( DBGrid2 );
  end;

  Form8.ADOQuery11.Close;
  Form8.ADOQuery5.Close;
  if DBLookupCombobox1.KeyValue = 100 then
  begin
      Form8.DataSource5.DataSet := ADOQuery5;
      Form8.ADOQuery5.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.DateTime);
      Form8.ADOQuery5.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.DateTime);
      Form8.ADOQuery5.Open;
      DimensionarGrid( DBGrid4 );
  end
  else
  begin
      Form8.DataSource5.DataSet := ADOQuery11;
      Form8.ADOQuery11.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.DateTime);
      Form8.ADOQuery11.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.DateTime);
      Form8.ADOQuery11.Open;
      DimensionarGrid( DBGrid4 );
  end;

    Form8.ADOQuery12.Close;
  Form8.ADOQuery7.Close;
  if DBLookupCombobox1.KeyValue = 100 then
  begin
      Form8.DataSource7.DataSet := ADOQuery7;
      Form8.ADOQuery7.Parameters.ParamByName('inicio').Value := DatetoStr			(DateTimePicker3.DateTime);
      Form8.ADOQuery7.Parameters.ParamByName('fim').Value := DatetoStr			(DateTimePicker4.DateTime);
      Form8.ADOQuery7.Open;
      DimensionarGrid( DBGrid6 );
  end
  else
  begin
      Form8.DataSource7.DataSet := ADOQuery12;
      Form8.ADOQuery12.Parameters.ParamByName('inicio').Value := DatetoStr			(DateTimePicker3.DateTime);
      Form8.ADOQuery12.Parameters.ParamByName('fim').Value := DatetoStr			(DateTimePicker4.DateTime);
      Form8.ADOQuery12.Open;
      DimensionarGrid( DBGrid6 );
  end;

    Form8.ADOQuery8.Close;
  Form8.ADOQuery8.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.DateTime);
  Form8.ADOQuery8.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.DateTime);
  if DBLookupCombobox1.KeyValue = 100 then
  begin
    Form8.ADOQuery8.Parameters.ParamByName('departamento').Value := (100);
    Form8.ADOQuery8.Parameters.ParamByName('departamento2').Value := (110);
    Form8.ADOQuery8.Open;
  end
  else
  begin
      Form8.ADOQuery8.Parameters.ParamByName('departamento').Value := (200);
      Form8.ADOQuery8.Parameters.ParamByName('departamento2').Value := (210);
      Form8.ADOQuery8.Open;
  end;

  Form8.ADOQuery9.Close;
  Form8.ADOQuery9.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.DateTime);
  Form8.ADOQuery9.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.DateTime);
  if DBLookupCombobox1.KeyValue = 100 then
  begin
    Form8.ADOQuery9.Parameters.ParamByName('departamento').Value := (100);
    Form8.ADOQuery9.Parameters.ParamByName('departamento2').Value := (110);
    Form8.ADOQuery9.Open;
  end
  else
  begin
      Form8.ADOQuery9.Parameters.ParamByName('departamento').Value := (200);
      Form8.ADOQuery9.Parameters.ParamByName('departamento2').Value := (210);
      Form8.ADOQuery9.Open;
  end;

  Form8.ADOQuery10.Close;
  Form8.ADOQuery10.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.DateTime);
  Form8.ADOQuery10.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.DateTime);
  if DBLookupCombobox1.KeyValue = 100 then
  begin
    Form8.ADOQuery10.Parameters.ParamByName('departamento').Value := (100);
    Form8.ADOQuery10.Parameters.ParamByName('departamento2').Value := (110);
    Form8.ADOQuery10.Open;
  end
  else
  begin
      Form8.ADOQuery10.Parameters.ParamByName('departamento').Value := (200);
      Form8.ADOQuery10.Parameters.ParamByName('departamento2').Value := (210);
      Form8.ADOQuery10.Open;
  end;

  a := Form8.ADOQuery8.FieldByName('Total').Value;
  b := Form8.ADOQuery9.FieldByName('Total').Value;
  c := Form8.ADOQuery10.FieldByName('Total').Value;
  f := Form8.ADOQuery1.FieldByName('Total').Value;
  g := Form8.ADOQuery4.FieldByName('Total').Value;
  dbedit1.Text:=inttostr(a);
  dbedit2.Text:=inttostr(b);
  dbedit5.Text:=inttostr(c);
  dbedit4.Text:=inttostr(f);
  dbedit7.Text:=inttostr(g);

  d := (StrToFloat(dbedit2.Text)*100)/StrToFloat(dbedit1.Text);
  dbedit3.Text:= FormatFloat('#,,0.0',d)+'%';

  e := (StrToFloat(dbedit5.Text)*100)/StrToFloat(dbedit1.Text);
  dbedit6.Text:= FormatFloat('#,,0.0',e)+'%';

  if DBLookupCombobox1.KeyValue = 100 then
  begin
    Form8.DataSource6.DataSet := ADOQuery6;
    Form8.ADOQuery6.Close;
    Form8.ADOQuery6.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery6.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery6.Parameters.ParamByName('inicio1').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery6.Parameters.ParamByName('fim1').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery6.Parameters.ParamByName('inicio2').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery6.Parameters.ParamByName('fim2').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery6.Parameters.ParamByName('inicio3').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery6.Parameters.ParamByName('fim3').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery6.Parameters.ParamByName('inicio4').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery6.Parameters.ParamByName('fim4').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery6.Parameters.ParamByName('inicio5').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery6.Parameters.ParamByName('fim5').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery6.Parameters.ParamByName('inicio6').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery6.Parameters.ParamByName('fim6').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery6.Parameters.ParamByName('inicio7').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery6.Parameters.ParamByName('fim7').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery6.Parameters.ParamByName('inicio8').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery6.Parameters.ParamByName('fim8').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery6.Parameters.ParamByName('inicio9').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery6.Parameters.ParamByName('fim9').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery6.Parameters.ParamByName('inicio10').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery6.Parameters.ParamByName('fim10').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery6.Parameters.ParamByName('inicio11').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery6.Parameters.ParamByName('fim11').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery6.Parameters.ParamByName('inicio12').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery6.Parameters.ParamByName('fim12').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery6.Parameters.ParamByName('inicio13').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery6.Parameters.ParamByName('fim13').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery6.Parameters.ParamByName('inicio14').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery6.Parameters.ParamByName('fim14').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery6.Open;
    DBGrid5.Columns[0].FieldName := 'VENDEDOR';
    DBGrid5.Columns[1].FieldName := 'Total Atendimentos'  ;
    DBGrid5.Columns[2].FieldName := 'Ativos' ;
    DBGrid5.Columns[3].FieldName := 'Receptivos' ;
    DBGrid5.Columns[4].FieldName := 'Showroom' ;
    DBGrid5.Columns[5].FieldName := 'Telefone' ;
    DBGrid5.Columns[6].FieldName := 'Lead' ;
    DBGrid5.Columns[7].FieldName := 'Total Vendas' ;
    DBGrid5.Columns[8].FieldName := 'Venda Loja' ;
    DBGrid5.Columns[9].FieldName := 'Venda Direta' ;
    DBGrid5.Columns[10].FieldName := 'Aproveitamento' ;
    DBGrid5.Columns[11].FieldName := 'Vendas Perdidas' ;
    DBGrid5.Columns[12].FieldName := 'Descarte' ;
    DimensionarGrid( DBGrid5 );
    Screen.Cursor := crDefault;
  end
  else
  begin
    Form8.DataSource6.DataSet := ADOQuery14;
    Form8.ADOQuery14.Close;
    Form8.ADOQuery14.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery14.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery14.Parameters.ParamByName('inicio1').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery14.Parameters.ParamByName('fim1').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery14.Parameters.ParamByName('inicio2').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery14.Parameters.ParamByName('fim2').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery14.Parameters.ParamByName('inicio3').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery14.Parameters.ParamByName('fim3').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery14.Parameters.ParamByName('inicio4').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery14.Parameters.ParamByName('fim4').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery14.Parameters.ParamByName('inicio5').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery14.Parameters.ParamByName('fim5').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery14.Parameters.ParamByName('inicio6').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery14.Parameters.ParamByName('fim6').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery14.Parameters.ParamByName('inicio7').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery14.Parameters.ParamByName('fim7').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery14.Parameters.ParamByName('inicio8').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery14.Parameters.ParamByName('fim8').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery14.Parameters.ParamByName('inicio9').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery14.Parameters.ParamByName('fim9').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery14.Parameters.ParamByName('inicio10').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery14.Parameters.ParamByName('fim10').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery14.Parameters.ParamByName('inicio11').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery14.Parameters.ParamByName('fim11').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery14.Parameters.ParamByName('inicio12').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery14.Parameters.ParamByName('fim12').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery14.Parameters.ParamByName('inicio13').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery14.Parameters.ParamByName('fim13').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery14.Parameters.ParamByName('inicio14').Value := DatetoStr(DateTimePicker3.Date);
    Form8.ADOQuery14.Parameters.ParamByName('fim14').Value := DatetoStr(DateTimePicker4.Date);
    Form8.ADOQuery14.Open;
    DBGrid5.Columns[0].FieldName := 'VENDEDOR';
    DBGrid5.Columns[1].FieldName := 'Total Atendimentos'  ;
    DBGrid5.Columns[2].FieldName := 'Ativos' ;
    DBGrid5.Columns[3].FieldName := 'Receptivos' ;
    DBGrid5.Columns[4].FieldName := 'Showroom' ;
    DBGrid5.Columns[5].FieldName := 'Telefone' ;
    DBGrid5.Columns[6].FieldName := 'Lead' ;
    DBGrid5.Columns[7].FieldName := 'Total Vendas' ;
    DBGrid5.Columns[8].FieldName := 'Venda Loja' ;
    DBGrid5.Columns[9].FieldName := 'Venda Direta' ;
    DBGrid5.Columns[10].FieldName := 'Aproveitamento' ;
    DBGrid5.Columns[11].FieldName := 'Vendas Perdidas' ;
    DBGrid5.Columns[12].FieldName := 'Descarte' ;
    DimensionarGrid( DBGrid5 );
    Screen.Cursor := crDefault;


  end;


end;

procedure TForm8.FormResize(Sender: TObject);
begin
  DimensionarGrid( DBGrid2 );
  DimensionarGrid( DBGrid4 );
  DimensionarGrid( DBGrid6 );
end;

procedure TForm8.FormShow(Sender: TObject);
begin
  datetimepicker3.date:= StartOfTheMonth(now);
  datetimepicker4.date:= EndOfTheMonth(now);
  DBLookupComboBox1.KeyValue:= 100;
  Dados.ActivePage := Atendimentos;
end;

procedure TForm8.DBGrid2DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  grid: TDBGrid;
  linha: Integer;
begin
  // obt�m um refer�ncia ao DBGrid
  grid := sender as TDBGrid;

  // obt�m o n�mero da linha atual usando a propriedade
  // RecNo da classe TDataSet
  linha := grid.DataSource.DataSet.RecNo;

  // o n�mero da linha � par?
  if Odd(linha) then
    begin
      grid.Canvas.Brush.Color := $00D8F5DC;
      grid.Canvas.Font.Color := clBlack
    end
  else
    begin
      grid.Canvas.Brush.Color := clSilver;
      grid.Canvas.Font.Color := clBlack
    end;

  if DataCol=1 then
  Begin
    TDBGrid(Sender).Canvas.Brush.Color := clRed;
    TDBGrid(Sender).Canvas.Font.Color := clWhite;
  End;


  // vamos terminar de desenhar a c�lula
  grid.DefaultDrawColumnCell(Rect, DataCol, Column, State);



end;

procedure TForm8.DBGrid4DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  grid: TDBGrid;
  linha: Integer;
begin
  // obt�m um refer�ncia ao DBGrid
  grid := sender as TDBGrid;

  // obt�m o n�mero da linha atual usando a propriedade
  // RecNo da classe TDataSet
  linha := grid.DataSource.DataSet.RecNo;

  // o n�mero da linha � par?
  if Odd(linha) then
    begin
      grid.Canvas.Brush.Color := $00D8F5DC;
      grid.Canvas.Font.Color := clBlack
    end
  else
    begin
      grid.Canvas.Brush.Color := clSilver;
      grid.Canvas.Font.Color := clBlack
    end;

  if DataCol=1 then
  Begin
    TDBGrid(Sender).Canvas.Brush.Color := clRed;
    TDBGrid(Sender).Canvas.Font.Color := clWhite;
  End;


  // vamos terminar de desenhar a c�lula
  grid.DefaultDrawColumnCell(Rect, DataCol, Column, State);



end;

procedure TForm8.DBGrid6DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  grid: TDBGrid;
  linha: Integer;
begin
  // obt�m um refer�ncia ao DBGrid
  grid := sender as TDBGrid;

  // obt�m o n�mero da linha atual usando a propriedade
  // RecNo da classe TDataSet
  linha := grid.DataSource.DataSet.RecNo;

  // o n�mero da linha � par?
  if Odd(linha) then
    begin
      grid.Canvas.Brush.Color := $00D8F5DC;
      grid.Canvas.Font.Color := clBlack
    end
  else
    begin
      grid.Canvas.Brush.Color := clSilver;
      grid.Canvas.Font.Color := clBlack
    end;

  if DataCol=1 then
  Begin
    TDBGrid(Sender).Canvas.Brush.Color := clRed;
    TDBGrid(Sender).Canvas.Font.Color := clWhite;
  End;


  // vamos terminar de desenhar a c�lula
  grid.DefaultDrawColumnCell(Rect, DataCol, Column, State);



end;

procedure TForm8.DimensionarGrid(dbg: TDBGrid);
type
  TArray = Array of Integer;
  procedure AjustarColumns(Swidth, TSize: Integer; Asize: TArray);
  var
    idx: Integer;
  begin
    if TSize = 0 then
    begin
      TSize := dbg.Columns.count;
      for idx := 0 to dbg.Columns.count - 1 do
        dbg.Columns[idx].Width := (dbg.Width - dbg.Canvas.TextWidth('AAAAAA')
          ) div TSize
    end
    else
      for idx := 0 to dbg.Columns.count - 1 do
        dbg.Columns[idx].Width := dbg.Columns[idx].Width +
          (Swidth * Asize[idx] div TSize);
  end;

var
  idx, Twidth, TSize, Swidth: Integer;
  AWidth: TArray;
  Asize: TArray;
  NomeColuna: String;
begin
  SetLength(AWidth, dbg.Columns.count);
  SetLength(Asize, dbg.Columns.count);
  Twidth := 0;
  TSize := 0;
  for idx := 0 to dbg.Columns.count - 1 do
  begin
    NomeColuna := dbg.Columns[idx].Title.Caption;
    dbg.Columns[idx].Width := dbg.Canvas.TextWidth
      (dbg.Columns[idx].Title.Caption + 'A');
    AWidth[idx] := dbg.Columns[idx].Width;
    Twidth := Twidth + AWidth[idx];

    if Assigned(dbg.Columns[idx].Field) then
      Asize[idx] := dbg.Columns[idx].Field.Size
    else
      Asize[idx] := 1;

    TSize := TSize + Asize[idx];
  end;
  if TDBGridOption.dgColLines in dbg.Options then
    Twidth := Twidth + dbg.Columns.count;

  // adiciona a largura da coluna indicada do cursor
  if TDBGridOption.dgIndicator in dbg.Options then
    Twidth := Twidth + IndicatorWidth;

  Swidth := dbg.ClientWidth - Twidth;
  AjustarColumns(Swidth, TSize, Asize);
end;

end.
