unit Unit2;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls;

type
  TMenuSPO = class(TForm)
    Background: TImage;
    Label1: TLabel;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    Label10: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MenuSPO: TMenuSPO;

implementation

{$R *.dfm}

uses Unit1, Unit3;

procedure TMenuSPO.Button1Click(Sender: TObject);
begin
MenuSPO.Hide;
Form3.show;
end;

procedure TMenuSPO.FormClose(Sender: TObject; var Action: TCloseAction);
begin
MenuSPO.Hide;
MainForm.show;
end;

end.
