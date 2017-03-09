program Project1;

uses
  Vcl.Forms,
  Unit1 in 'Unit1.pas' {MainForm},
  Unit2 in 'Unit2.pas' {MenuSPO},
  Unit3 in 'Unit3.pas' {PriemSPO},
  Unit4 in 'Unit4.pas' {AttestSPO},
  Unit5 in 'Unit5.pas' {AddAttestSPO},
  Unit6 in 'Unit6.pas' {Form6};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TMenuSPO, MenuSPO);
  Application.CreateForm(TPriemSPO, PriemSPO);
  Application.CreateForm(TAttestSPO, AttestSPO);
  Application.CreateForm(TAddAttestSPO, AddAttestSPO);
  Application.CreateForm(TForm6, Form6);
  Application.Run;
end.
