unit relatorios;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, DB, ADODB, ComCtrls;

type
  TForm1 = class(TForm)
    Button2: TButton;
    DataSource1: TDataSource;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    TabControl1: TTabControl;
    date_in: TDateTimePicker;
    Label2: TLabel;
    DBGrid1: TDBGrid;
    date_fin: TDateTimePicker;
    Label1: TLabel;
    Btn1: TButton;
    procedure Btn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Btn1Click(Sender: TObject);
begin
  Form1.ADOQuery1.Active := True;
end;

end.
