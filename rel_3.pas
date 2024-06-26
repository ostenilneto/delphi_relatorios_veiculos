unit rel_3;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, Mask, Grids, ComObj, DBGrids, ComCtrls,
  Vcl.ExtCtrls, DateUtils, Vcl.Buttons, Vcl.DBCtrls, Datasnap.DBClient, SimpleDS,
  Data.SqlExpr, Data.FMTBcd, Datasnap.Provider;

type
  TForm4 = class(TForm)
    DataSource1: TDataSource;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    TabControl2: TTabControl;
    DBGrid3: TDBGrid;
    ADOQuery2: TADOQuery;
    DataSource2: TDataSource;
    DBGrid4: TDBGrid;
    DataSource3: TDataSource;
    ADOQuery3: TADOQuery;
    Panel1: TPanel;
    DateTimePicker3: TDateTimePicker;
    Label3: TLabel;
    Label4: TLabel;
    DateTimePicker4: TDateTimePicker;
    Label1: TLabel;
    DBLookupComboBox1: TDBLookupComboBox;
    BitBtn1: TBitBtn;
    BitBtn4: TBitBtn;
    BitBtn2: TBitBtn;
    Label2: TLabel;
    ComboBox1: TComboBox;
    ADOQuery3USUARIO: TIntegerField;
    ADOQuery3NOME: TStringField;
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DimensionarGrid(dbg: TDBGrid);
    procedure DBGrid3DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid4DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure BitBtn2Click(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);

private

public
    { Public declarations }
end;

var
  Form4: TForm4;

implementation
uses Princ;

{$R *.dfm}

procedure TForm4.BitBtn1Click(Sender: TObject);
begin
  DBGrid3.DataSource.DataSet.Close;
  DBGrid4.DataSource.DataSet.Close;
end;

procedure TForm4.BitBtn2Click(Sender: TObject);
begin
  Form4.Close;
  Form2.Panel12.Width := 150;
  Form2.Caption := 'PROGRAMA DE APOIO';
  DBGrid3.DataSource.DataSet.Close;
  DBGrid4.DataSource.DataSet.Close;
end;

procedure TForm4.BitBtn4Click(Sender: TObject);
begin
  Form4.ADOQuery1.Close;
  Form4.ADOQuery2.Close;

  if ComboBox1.ItemIndex = 0
        then begin
        Form4.ADOQuery2.Parameters.ParamByName('departamento').Value := 100;
        Form4.ADOQuery2.Parameters.ParamByName('departamento2').Value := 110;
        Form4.ADOQuery1.Parameters.ParamByName('departamento').Value := 100;
        Form4.ADOQuery1.Parameters.ParamByName('departamento2').Value := 110;

    end
    else begin
        Form4.ADOQuery2.Parameters.ParamByName('departamento').Value := 200;
        Form4.ADOQuery2.Parameters.ParamByName('departamento2').Value := 210;
        Form4.ADOQuery1.Parameters.ParamByName('departamento').Value := 200;
        Form4.ADOQuery1.Parameters.ParamByName('departamento2').Value := 210;

    end;

  Form4.ADOQuery1.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
  Form4.ADOQuery1.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.DateTime);
  Form4.ADOQuery2.Parameters.ParamByName('vendedor').Value := DBLookupComboBox1.KeyValue;
  Form4.ADOQuery2.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
  Form4.ADOQuery2.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.DateTime);
  Form4.ADOQuery1.Open;
  Form4.ADOQuery2.Open;
  DimensionarGrid( DBGrid3 );
  DimensionarGrid( DBGrid4 );
end;

procedure TForm4.ComboBox1Change(Sender: TObject);
begin
  if ComboBox1.ItemIndex = 0
        then begin
        Form4.ADOQuery1.Close;
        Form4.ADOQuery2.Close;
        Form4.ADOQuery3.Close;
        Form4.ADOQuery1.Parameters.ParamByName('departamento').Value := 100;
        Form4.ADOQuery1.Parameters.ParamByName('departamento2').Value := 110;
        Form4.ADOQuery1.Open;
        Form4.ADOQuery3.Parameters.ParamByName('cargo').Value := 6;
        Form4.DBLookupComboBox1.KeyValue:= 468;
        Form4.ADOQuery3.Open;
    end
    else begin
        Form4.ADOQuery1.Close;
        Form4.ADOQuery2.Close;
        Form4.ADOQuery3.Close;
        Form4.ADOQuery1.Parameters.ParamByName('departamento').Value := 200;
        Form4.ADOQuery1.Parameters.ParamByName('departamento2').Value := 210;
        Form4.ADOQuery1.Open;
        Form4.ADOQuery3.Parameters.ParamByName('cargo').Value := 7;
        Form4.DBLookupComboBox1.KeyValue:= 521;
        Form4.ADOQuery3.Open;
    end;
    DimensionarGrid( DBGrid3 );
end;

procedure TForm4.FormResize(Sender: TObject);
begin
  DimensionarGrid( DBGrid3 );
  DimensionarGrid( DBGrid4 );
end;

procedure TForm4.FormShow(Sender: TObject);
begin
  Form4.ADOQuery1.Close;
  datetimepicker3.date:= StartOfTheMonth(now);
  datetimepicker4.DateTime:= now;
  Form4.ADOQuery1.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
  Form4.ADOQuery1.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.DateTime);
  Form4.ADOQuery1.Parameters.ParamByName('departamento').Value := 100;
  Form4.ADOQuery1.Parameters.ParamByName('departamento2').Value := 110;
  Form4.ADOQuery1.Open;
  DimensionarGrid( DBGrid3 );
  DimensionarGrid( DBGrid4 );
  Form4.ADOQuery3.Close;
  Form4.ADOQuery3.Parameters.ParamByName('cargo').Value := 6;
  Form4.ADOQuery3.Open;
  DBLookupComboBox1.KeyValue:= 468;
  ComboBox1.ItemIndex := 0;
end;

procedure TForm4.DBGrid3DrawColumnCell(Sender: TObject; const Rect: TRect;
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

  if DataCol=3 then
  Begin
    TDBGrid(Sender).Canvas.Brush.Color := clRed;
    TDBGrid(Sender).Canvas.Font.Color := clWhite;
  End;


  // vamos terminar de desenhar a c�lula
  grid.DefaultDrawColumnCell(Rect, DataCol, Column, State);

  if (TStringGrid(DBGrid3).RowCount-1) < 10 then //Se tiver menos de 10 linhas
    ShowScrollBar(DBGrid3.Handle,SB_HORZ,False); //Remove barra Vertical
    ShowScrollBar(DBGrid3.Handle,SB_VERT,False); //Remove barra Vertical


end;


procedure TForm4.DBGrid4DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
    if (TStringGrid(DBGrid4).RowCount-1) < 10 then //Se tiver menos de 10 linhas
    ShowScrollBar(DBGrid4.Handle,SB_HORZ,False); //Remove barra Vertical
    ShowScrollBar(DBGrid4.Handle,SB_VERT,False); //Remove barra Vertical
end;



procedure TForm4.DimensionarGrid(dbg: TDBGrid);
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
