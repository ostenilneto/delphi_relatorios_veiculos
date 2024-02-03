unit rel_5;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, Mask, Grids, ComObj, DBGrids, ComCtrls,
  Vcl.ExtCtrls, DateUtils, Vcl.Buttons, Vcl.DBCtrls, Datasnap.DBClient, SimpleDS,
  Data.SqlExpr, Data.FMTBcd, Datasnap.Provider;

type
  TForm6 = class(TForm)
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
    TabControl2: TTabControl;
    DBGrid3: TDBGrid;
    DBGrid4: TDBGrid;
    DataSource1: TDataSource;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    ADOQuery2: TADOQuery;
    DataSource2: TDataSource;
    DataSource3: TDataSource;
    ADOQuery3: TADOQuery;
    ADOQuery3ORIGEM_TRAFEGO: TIntegerField;
    ADOQuery3DES_ORIGEM_TRAFEGO: TStringField;
    Label2: TLabel;
    ComboBox1: TComboBox;
    DataSource5: TDataSource;
    DBLookupComboBox2: TDBLookupComboBox;
    ADOQuery5: TADOQuery;
    ADOQuery4: TADOQuery;
    procedure FormShow(Sender: TObject);
    procedure DimensionarGrid(dbg: TDBGrid);
    procedure FormResize(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form6: TForm6;

implementation
uses Princ;
{$R *.dfm}

procedure TForm6.FormResize(Sender: TObject);
begin
  DimensionarGrid( DBGrid3 );
  DimensionarGrid( DBGrid4 );
end;

procedure TForm6.FormShow(Sender: TObject);
begin
  Form6.ADOQuery1.Close;
  datetimepicker3.date:= StartOfTheMonth(now);
  datetimepicker4.date:= now;
  Form6.ADOQuery1.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
  Form6.ADOQuery1.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.Date);
  Form6.ADOQuery1.Parameters.ParamByName('inicio2').Value := DatetoStr(DateTimePicker3.Date);
  Form6.ADOQuery1.Parameters.ParamByName('fim2').Value := DatetoStr(DateTimePicker4.Date);
  Form6.ADOQuery1.Parameters.ParamByName('inicio3').Value := DatetoStr(DateTimePicker3.Date);
  Form6.ADOQuery1.Parameters.ParamByName('fim3').Value := DatetoStr(DateTimePicker4.Date);
  Form6.DBLookupComboBox2.Enabled := False;
  Form6.DBLookupComboBox2.Visible := False;
  Form6.ADOQuery1.Open;
  DBLookupComboBox1.KeyValue:= 14;
end;

procedure TForm6.BitBtn1Click(Sender: TObject);
begin
  DBGrid3.DataSource.DataSet.Close;
  DBGrid4.DataSource.DataSet.Close;
end;

procedure TForm6.BitBtn2Click(Sender: TObject);
begin
  Form6.Close;
  Form2.Panel12.Width := 150;
  Form2.Caption := 'PROGRAMA DE APOIO';
  DBGrid3.DataSource.DataSet.Close;
  DBGrid4.DataSource.DataSet.Close;
  DBGrid3.DataSource.DataSet.Close;
  DBGrid4.DataSource.DataSet.Close;
end;

procedure TForm6.BitBtn4Click(Sender: TObject);
begin
  DBGrid3.DataSource.DataSet.Close;
  DBGrid4.DataSource.DataSet.Close;
  Form6.ADOQuery1.Close;
  Form6.ADOQuery2.Close;
  Form6.ADOQuery1.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
  Form6.ADOQuery1.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.Date);
  Form6.ADOQuery1.Parameters.ParamByName('inicio2').Value := DatetoStr(DateTimePicker3.Date);
  Form6.ADOQuery1.Parameters.ParamByName('fim2').Value := DatetoStr(DateTimePicker4.Date);
  Form6.ADOQuery1.Parameters.ParamByName('inicio3').Value := DatetoStr(DateTimePicker3.Date);
  Form6.ADOQuery1.Parameters.ParamByName('fim3').Value := DatetoStr(DateTimePicker4.Date);
  Form6.ADOQuery1.Open;
    DimensionarGrid( DBGrid3 );

  if ComboBox1.ItemIndex = 0
  then begin
    Form6.DataSource2.DataSet := ADOQuery2;
    Form6.ADOQuery2.Parameters.ParamByName('inicio').Value := DateTimetoStr(DateTimePicker3.Date);
    Form6.ADOQuery2.Parameters.ParamByName('fim').Value := DateTimetoStr(DateTimePicker4.Date);
    Form6.ADOQuery2.Parameters.ParamByName('inicio2').Value := DateTimetoStr(DateTimePicker3.Date);
    Form6.ADOQuery2.Parameters.ParamByName('fim2').Value := DateTimetoStr(DateTimePicker4.Date);
    Form6.ADOQuery2.Parameters.ParamByName('origem').Value := DBLookupComboBox1.KeyValue;
    Form6.ADOQuery2.Open;
  end
  else begin
  Form6.DataSource2.DataSet := ADOQuery4;
    Form6.ADOQuery4.Parameters.ParamByName('inicio').Value := DateTimetoStr(DateTimePicker3.Date);
    Form6.ADOQuery4.Parameters.ParamByName('fim').Value := DateTimetoStr(DateTimePicker4.Date);
    Form6.ADOQuery4.Parameters.ParamByName('vendedor').Value := Form6.DBLookupComboBox2.KeyValue;
    Form6.ADOQuery4.Open;
  end;


end;

procedure TForm6.ComboBox1Change(Sender: TObject);
begin
  if ComboBox1.ItemIndex = 0
        then begin

        Form6.ADOQuery2.Close;
        Form6.ADOQuery3.Close;
        Form6.Label1.Caption := 'Origem de Tr�fego:';
        Form6.DBLookupComboBox1.Enabled := True;
        Form6.DBLookupComboBox1.Visible := True;
        Form6.DBLookupComboBox2.Enabled := False;
        Form6.DBLookupComboBox2.Visible := False;
        Form6.DataSource1.DataSet := ADOQuery1;
        Form6.ADOQuery1.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
        Form6.ADOQuery1.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.Date);
        Form6.ADOQuery1.Open;
        Form6.DBLookupComboBox1.ListFieldIndex:= 0;
        Form6.ADOQuery3.Open;
    end
    else begin
        Form6.ADOQuery2.Close;
        Form6.ADOQuery3.Close;
        Form6.ADOQuery1.Close;
        Form6.Label1.Caption := 'Escolha o vendedor:';
        //Form6.DataSource1.DataSet := ADOQuery4;
        //Form6.ADOQuery4.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
        //Form6.ADOQuery4.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.Date);
        Form6.DataSource1.DataSet := ADOQuery1;
        Form6.ADOQuery1.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
        Form6.ADOQuery1.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.Date);
        Form6.DBLookupComboBox1.Enabled := False;
        Form6.DBLookupComboBox1.Visible := False;
        Form6.DBLookupComboBox2.Enabled := True;
        Form6.DBLookupComboBox2.Visible := True;
        Form6.DBLookupComboBox2.KeyValue:= 468;
        Form6.ADOQuery1.Open;
        Form6.ADOQuery5.Open;
    end;
    DimensionarGrid( DBGrid3 );
end;

procedure TForm6.DimensionarGrid(dbg: TDBGrid);
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