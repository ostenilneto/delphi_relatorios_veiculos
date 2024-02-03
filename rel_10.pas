unit rel_10;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.StdCtrls, Vcl.Buttons,
  Mask, ComObj, Vcl.ComCtrls, Vcl.Grids, Vcl.DBGrids, Vcl.ExtCtrls, Data.Win.ADODB;

type
  TForm11 = class(TForm)
    TabControl1: TTabControl;
    Panel4: TPanel;
    Label2: TLabel;
    Label1: TLabel;
    DateTimePicker2: TDateTimePicker;
    BitBtn4: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn1: TBitBtn;
    DateTimePicker1: TDateTimePicker;
    Panel3: TPanel;
    DBGrid3: TDBGrid;
    Panel5: TPanel;
    Label4: TLabel;
    Panel1: TPanel;
    DBGrid1: TDBGrid;
    Panel2: TPanel;
    Label3: TLabel;
    Panel6: TPanel;
    DBGrid2: TDBGrid;
    Panel7: TPanel;
    Label5: TLabel;
    DataSource1: TDataSource;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    ADOQuery2: TADOQuery;
    DataSource2: TDataSource;
    ADOQuery3: TADOQuery;
    DataSource3: TDataSource;
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form11: TForm11;

implementation

uses Princ;

{$R *.dfm}


procedure TForm11.BitBtn1Click(Sender: TObject);
begin
  Form11.DBGrid1.DataSource.DataSet.Close;
  Form11.DBGrid2.DataSource.DataSet.Close;
  Form11.DBGrid3.DataSource.DataSet.Close;
end;

procedure TForm11.BitBtn2Click(Sender: TObject);
begin
  Screen.Cursor := crHourglass;
  Form11.ADOQuery1.Close;
  Form11.ADOQuery1.Parameters.ParamByName('inicio').Value := DatetoStr(DateTimePicker1.Date);
  Form11.ADOQuery1.Parameters.ParamByName('fim').Value := DatetoStr(DateTimePicker2.Date);
  Form11.ADOQuery1.Open;
  Form11.ADOQuery2.Parameters.ParamByName('inicio2').Value := DatetoStr(DateTimePicker1.Date);
  Form11.ADOQuery2.Parameters.ParamByName('fim2').Value := DatetoStr(DateTimePicker2.Date);
  Form11.ADOQuery2.Open;
  Form11.ADOQuery3.Parameters.ParamByName('inicio3').Value := DatetoStr(DateTimePicker1.Date);
  Form11.ADOQuery3.Parameters.ParamByName('fim3').Value := DatetoStr(DateTimePicker2.Date);
  Form11.ADOQuery3.Open;
  Screen.Cursor := crDefault;
end;


procedure TForm11.BitBtn3Click(Sender: TObject);
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


procedure TForm11.BitBtn4Click(Sender: TObject);
begin
  Form11.DBGrid1.DataSource.DataSet.Close;
  Form2.Panel12.Width := 150;
  Form11.Close;
  Form2.Caption := 'PROGRAMA DE APOIO';
end;

procedure TForm11.FormShow(Sender: TObject);
begin
  datetimepicker1.date:= now;
  datetimepicker2.date:= now;
end;

end.
