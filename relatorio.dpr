program relatorio;

uses
  Forms,
  rel_1 in 'rel_1.pas' {Form1},
  Uconexao in 'Uconexao.pas',
  Princ in 'Princ.pas' {Form2},
  rel_2 in 'rel_2.pas' {Form3},
  rel_3 in 'rel_3.pas' {Form4},
  rel_4 in 'rel_4.pas' {Form5},
  rel_5 in 'rel_5.pas' {Form6},
  rel_6 in 'rel_6.pas' {Form7},
  rel_7 in 'rel_7.pas' {Form8},
  rel_8 in 'rel_8.pas' {Form9},
  rel_9 in 'rel_9.pas' {Form10},
  rel_10 in 'rel_10.pas' {Form11};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm2, Form2);
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm3, Form3);
  Application.CreateForm(TForm4, Form4);
  Application.CreateForm(TForm5, Form5);
  Application.CreateForm(TForm6, Form6);
  Application.CreateForm(TForm7, Form7);
  Application.CreateForm(TForm8, Form8);
  Application.CreateForm(TForm9, Form9);
  Application.CreateForm(TForm10, Form10);
  Application.CreateForm(TForm11, Form11);
  Application.Run;
end.
