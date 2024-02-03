unit rel_9;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, Mask, Grids, ComObj, DBGrids, ComCtrls,
  Vcl.ExtCtrls, Vcl.Buttons;

type
  TForm10 = class(TForm)
    Label2: TLabel;
    DateTimePicker1: TDateTimePicker;
    Label1: TLabel;
    DateTimePicker2: TDateTimePicker;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    procedure BitBtn4Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form10: TForm10;

implementation
uses Princ;
{$R *.dfm}

procedure TForm10.BitBtn1Click(Sender: TObject);
begin
  DBGrid1.DataSource.DataSet.Close;
end;

procedure TForm10.BitBtn2Click(Sender: TObject);
begin
  Screen.Cursor := crHourglass;
  Form10.ADOQuery1.Close;
  Form10.ADOQuery1.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker1.Date);
  Form10.ADOQuery1.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker2.Date);
  Form10.ADOQuery1.Open;
  Screen.Cursor := crDefault;
end;

procedure TForm10.BitBtn3Click(Sender: TObject);
Var
linha, coluna: integer;
planilha: variant;
valorCampo: string;
begin
   planilha:= CreateOleObject('Excel.Application');
   planilha.Workbooks.add(1);
   planilha.Cells.Select;
   planilha.Selection.NumberFormat := '@';
   planilha.caption:= 'Exportação de dados para o excel';
   planilha.visible:= true;
   AdoQuery1.First;
   for linha:= 0 to AdoQuery1.RecordCount-1 do
   begin
     for coluna:= 1 to AdoQuery1.FieldCount do
      begin
         valorCampo:= AdoQuery1.Fields[coluna-1].AsString;
         planilha.cells[linha+2,coluna]:= valorCampo;
      end;

     AdoQuery1.Next;

   end;
   for coluna:=1 to AdoQuery1.FieldCount do
   begin
      valorCampo:= AdoQuery1.Fields[coluna-1].DisplayLabel;
      planilha.cells[1,coluna]:= valorCampo;
   end;
   planilha.columns.AutoFit;
end;

procedure TForm10.BitBtn4Click(Sender: TObject);
begin
  Form2.Panel12.Width := 150;
  Form2.Caption := 'PROGRAMA DE APOIO';
  Form10.Close;
  DBGrid1.DataSource.DataSet.Close;
end;

procedure TForm10.FormShow(Sender: TObject);
begin
  datetimepicker1.date:= now;
  datetimepicker2.date:= now;
end;

end.
