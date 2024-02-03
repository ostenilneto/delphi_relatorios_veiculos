unit rel_1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, DB, ADODB,ComObj, ComCtrls, Mask,
  OleServer, ExcelXP, DBCtrls, Vcl.ExtCtrls, Vcl.Buttons;

type
  TForm1 = class(TForm)
    DataSource1: TDataSource;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    TabControl1: TTabControl;
    Label2: TLabel;
    DBGrid1: TDBGrid;
    Label1: TLabel;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn5: TBitBtn;
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation
uses Princ;
{$R *.dfm}

procedure TForm1.FormShow(Sender: TObject);
begin
  datetimepicker1.date:= now-1;
  datetimepicker2.date:= now-1;
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
begin
  DBGrid1.DataSource.DataSet.Close;
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
begin
  Form1.ADOQuery1.Close;
  Form1.ADOQuery1.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker1.Date);
  Form1.ADOQuery1.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker2.Date);
  Form1.ADOQuery1.Open;
end;

procedure TForm1.BitBtn3Click(Sender: TObject);
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

procedure TForm1.BitBtn5Click(Sender: TObject);
begin
  Form2.Panel12.Width := 150;
  Form2.Caption := 'PROGRAMA DE APOIO';
  Form1.Close;
  DBGrid1.DataSource.DataSet.Close;
end;

end.
