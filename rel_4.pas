unit rel_4;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.ExtCtrls, Data.Win.ADODB,
  Vcl.StdCtrls, Vcl.Buttons, ComObj, DateUtils, Vcl.ComCtrls, Vcl.Grids, Vcl.DBGrids,
  Vcl.DBCtrls;

type
  TForm5 = class(TForm)
    Panel1: TPanel;
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
    ComboBox1: TComboBox;
    DBLookupComboBox1: TDBLookupComboBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    DateTimePicker4: TDateTimePicker;
    DateTimePicker3: TDateTimePicker;
    ADOQuery3USUARIO: TIntegerField;
    ADOQuery3NOME: TStringField;
    ADOQuery4: TADOQuery;
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure DimensionarGrid(dbg: TDBGrid);
    procedure FormResize(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form5: TForm5;

implementation
uses Princ;

{$R *.dfm}

procedure TForm5.FormResize(Sender: TObject);
begin
  DimensionarGrid( DBGrid3 );
  DimensionarGrid( DBGrid4 );
end;

procedure TForm5.FormShow(Sender: TObject);
begin
  Form5.ADOQuery1.Close;
  datetimepicker3.date:= StartOfTheMonth(now);;
  datetimepicker4.date:= now;
  Form5.DataSource1.DataSet := ADOQuery1;
  Form5.ADOQuery1.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
  Form5.ADOQuery1.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.Date);
  Form5.ADOQuery1.Open;
  DimensionarGrid( DBGrid3 );
  DimensionarGrid( DBGrid4 );
  Form5.ADOQuery3.Close;
  Form5.ADOQuery3.Parameters.ParamByName('cargo').Value := 6;
  Form5.ADOQuery3.Open;
  DBLookupComboBox1.KeyValue:= 468;
  ComboBox1.ItemIndex := 0;
end;

procedure TForm5.BitBtn1Click(Sender: TObject);
begin
  DBGrid3.DataSource.DataSet.Close;
  DBGrid4.DataSource.DataSet.Close;
end;

procedure TForm5.BitBtn2Click(Sender: TObject);
begin
  Form5.Close;
  Form2.Panel12.Width := 150;
  Form2.Caption := 'PROGRAMA DE APOIO';
  DBGrid3.DataSource.DataSet.Close;
  DBGrid4.DataSource.DataSet.Close;
  DBGrid3.DataSource.DataSet.Close;
  DBGrid4.DataSource.DataSet.Close;
end;

procedure TForm5.BitBtn4Click(Sender: TObject);
begin
  //ADOQuery1
  Form5.ADOQuery1.Close;
  Form5.ADOQuery1.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
  Form5.ADOQuery1.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.Date);
  Form5.ADOQuery1.Open;
  DimensionarGrid( DBGrid3 );

  //ADOQuery2
  Form5.ADOQuery2.Close;
  Form5.ADOQuery2.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
  Form5.ADOQuery2.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.Date);
  Form5.ADOQuery2.Parameters.ParamByName('vendedor').Value := DBLookupComboBox1.KeyValue;
  Form5.ADOQuery2.Open;
end;

procedure TForm5.ComboBox1Change(Sender: TObject);
begin
  if ComboBox1.ItemIndex = 0
        then begin

        Form5.ADOQuery2.Close;
        Form5.ADOQuery3.Close;
        Form5.DataSource1.DataSet := ADOQuery1;
        Form5.ADOQuery1.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
        Form5.ADOQuery1.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.Date);
        Form5.ADOQuery1.Open;
        Form5.ADOQuery3.Parameters.ParamByName('cargo').Value := 6;
        Form5.DBLookupComboBox1.KeyValue:= 468;
        Form5.ADOQuery3.Open;
    end
    else begin
        Form5.ADOQuery2.Close;
        Form5.ADOQuery3.Close;
        Form5.DataSource1.DataSet := ADOQuery4;
        Form5.ADOQuery4.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker3.Date);
        Form5.ADOQuery4.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker4.Date);
        Form5.ADOQuery4.Open;
        Form5.ADOQuery3.Parameters.ParamByName('cargo').Value := 7;
        Form5.DBLookupComboBox1.KeyValue:= 521;
        Form5.ADOQuery3.Open;
    end;
    DimensionarGrid( DBGrid3 );

end;

procedure TForm5.DimensionarGrid(dbg: TDBGrid);
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
