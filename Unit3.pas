unit Unit3;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Data.DB,
  Vcl.Grids, Vcl.DBGrids, Data.Win.ADODB;

type
  TPriemSPO = class(TForm)
    Background: TImage;
    Label1: TLabel;
    Label10: TLabel;
    DBGrid1: TDBGrid;
    Label2: TLabel;
    Label3: TLabel;
    ADOQuery1: TADOQuery;
    DataSource1: TDataSource;
    ADOQuery2: TADOQuery;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  PriemSPO: TPriemSPO;
  a:integer;
implementation

{$R *.dfm}

uses Unit2, Unit1;

procedure TPriemSPO.FormActivate(Sender: TObject);
begin
Label3.Caption:=MainForm.DBLookupComboBox1.Text;
ADOQuery2.Close;
ADOQuery2.SQL.Clear;
ADOQuery2.SQL.Add('SELECT *');
ADOQuery2.SQL.Add('FROM ОУ');
ADOQuery2.SQL.Add('WHERE Наименование =:a;');
ADOQuery2.Parameters.ParamByName('a').Value:=Label3.Caption;
ADOQuery2.Open;
a:=ADOQuery2.FieldByName('Код').Value;

ADOQuery1.Close;
ADOQuery1.SQL.Clear;
ADOQuery1.SQL.Add('SELECT Наименование, Квота,Год');
ADOQuery1.SQL.Add('FROM КЦПСПО INNER JOIN СпециальностиСПО ON СпециальностиСПО.ИД_Специальности=КЦПСПО.ИД_Специальности');
ADOQuery1.SQL.Add('WHERE ИД_ОУ =:a;');
ADOQuery1.Parameters.ParamByName('a').Value:=a;
ADOQuery1.Open;
end;

procedure TPriemSPO.FormClose(Sender: TObject; var Action: TCloseAction);
begin
PriemSPO.Hide;
MenuSPO.show;
end;

end.
