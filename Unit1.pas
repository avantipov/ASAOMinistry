unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Data.DB,
  Vcl.DBCtrls, Data.Win.ADODB;

type
  TMainForm = class(TForm)
    Background: TImage;
    Label1: TLabel;
    Edit1: TEdit;
    Label2: TLabel;
    Edit2: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Edit3: TEdit;
    Edit4: TEdit;
    Button1: TButton;
    Button2: TButton;
    Label10: TLabel;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    DBLookupComboBox1: TDBLookupComboBox;
    DBLookupComboBox2: TDBLookupComboBox;
    DataSource1: TDataSource;
    ADOQuery2: TADOQuery;
    DataSource2: TDataSource;
    Auth: TADOQuery;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;
  a:integer;

implementation

{$R *.dfm}

uses Unit2;

procedure TMainForm.Button1Click(Sender: TObject);
begin
a:=DataSource1.DataSet.FieldByName('Тип образовательной организации').Value;
Auth.Close;
Auth.SQL.Clear;
Auth.SQL.Add('SELECT * FROM ОУ');
Auth.SQL.Add('WHERE Наименование =:a AND Логин=:b AND Пароль=:c;');
Auth.Parameters.ParamByName('a').Value:=DBLookupComboBox1.Text;
Auth.Parameters.ParamByName('b').Value:=Edit1.Text;
Auth.Parameters.ParamByName('c').Value:=Edit2.text;
Auth.Open;
Edit1.Text:='';
Edit2.Text:='';
if Auth.RecordCount=0 then
showmessage('Авторизация не удалась, проверьте введенные данные')
  else
  begin
  MainForm.Hide;
    case a of
    1 : showmessage('Ссылка на модуль Меню ШКОЛА');
    2 : MenuSPO.Show;
    3 : showmessage('Ссылка на модуль Меню ВУЗ')
    else
    begin
      showmessage('Вы не выбрали учебное заведение');
      MainForm.Show;
    end;
    end;


  end;

end;

end.
