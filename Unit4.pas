unit Unit4;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Data.DB,
  Data.Win.ADODB, Vcl.Grids, Vcl.DBGrids;

type
  TAttestSPO = class(TForm)
    Background: TImage;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label10: TLabel;
    DBGrid1: TDBGrid;
    Label4: TLabel;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    ADOQuery1: TADOQuery;
    DataSource1: TDataSource;
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AttestSPO: TAttestSPO;

implementation

{$R *.dfm}

uses Unit1, Unit2, Unit3, Unit5;

procedure TAttestSPO.Button1Click(Sender: TObject);
begin
AttestSPO.Hide;
AddAttestSPO.show;
end;

procedure TAttestSPO.FormActivate(Sender: TObject);
begin
Label3.Caption:=MainForm.DBLookupComboBox1.Text;
ADOQuery1.Close;
ADOQuery1.SQL.Clear;
ADOQuery1.SQL.Add('SELECT * FROM ПедкадрыСПО WHERE [Код ОУ]=:a;');
ADOQuery1.Parameters.ParamByName('a').Value:=MainForm.DataSource1.DataSet.FieldByName('Код').Value;
ADOQuery1.Open;
end;

procedure TAttestSPO.FormClose(Sender: TObject; var Action: TCloseAction);
begin
AttestSPO.hide;
MenuSPO.show;
end;

end.
