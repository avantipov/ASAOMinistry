unit Unit5;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.DBCtrls,
  Data.DB, Data.Win.ADODB, Vcl.Grids, Vcl.Samples.Calendar,ComObj;


type
  TAddAttestSPO = class(TForm)
    Background: TImage;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label10: TLabel;
    DBLookupComboBox1: TDBLookupComboBox;
    Label4: TLabel;
    Label5: TLabel;
    Edit1: TEdit;
    Label6: TLabel;
    Label7: TLabel;
    Button5: TButton;
    Label8: TLabel;
    Edit2: TEdit;
    DBConnect: TADOConnection;
    Query_Teacher: TADOQuery;
    DS_Teacher: TDataSource;
    ComboBox1: TComboBox;
    ComboBox2: TComboBox;
    Query_Insert: TADOQuery;
    Calendar1: TCalendar;
    SaveDialog1: TSaveDialog;
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DBLookupComboBox1Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AddAttestSPO: TAddAttestSPO;

implementation


{$R *.dfm}

uses Unit1, Unit4;

procedure TAddAttestSPO.Button5Click(Sender: TObject);
 const
  wdAlignParagraphCenter = 1;
  wdAlignParagraphLeft = 0;
  wdAlignParagraphRight = 2;
var
  wdApp, wdDoc, wdRng : Variant;
  Res : Integer;
  Sd : TSaveDialog;
begin
  Query_Insert.Close;
  Query_Insert.SQL.Clear;
  Query_Insert.SQL.Add('INSERT INTO АттестСПО(Код_ПР,ТипПер,КвКат,Дата,[Статус заявления],[Код ОУ],Основание)');
  Query_Insert.SQL.Add('VALUES (:idT,:per,:kvk,:date,:status,:idOU,:Osn);');
  Query_Insert.Parameters.ParamByName('idT').Value:=DS_Teacher.DataSet.FieldByName('Код').Value;
  Query_Insert.Parameters.ParamByName('per').Value:=Combobox1.Text;
  Query_Insert.Parameters.ParamByName('kvk').Value:=Combobox2.text;
  Query_Insert.Parameters.ParamByName('date').Value:=Calendar1.CalendarDate;
  Query_Insert.Parameters.ParamByName('status').Value:='Направлено в АК';
  Query_Insert.Parameters.ParamByName('idOU').Value:=DS_Teacher.DataSet.FieldByName('Код ОУ').Value;
  Query_Insert.Parameters.ParamByName('osn').Value:='';
  Query_Insert.ExecSQL;
  showmessage('Заявление зарегистрировано!');
    begin
      Sd :=SaveDialog1; //SaveDialog1 уже должен быть на форме.
      //Если начальная папка диалога не задана, то в качестве начальной берём ту папку,
      //в которой расположен исполняемый файл нашей программы.
      if Sd.InitialDir = '' then Sd.InitialDir := ExtractFilePath( ParamStr(0) );
      //Запуск диалога сохранения файла.
      if not Sd.Execute then Exit;
      //Если файл с заданным именем существует, то запускаем диалог с пользователем.
      if FileExists(Sd.FileName) then begin
        Res := MessageBox(0, 'Файл с заданным именем уже существует. Перезаписать?'
          ,'Внимание!', MB_YESNO + MB_ICONQUESTION + MB_APPLMODAL);
        if Res <> IDYES then Exit;
      end;

      //Попытка запустить MS Word.
      try
        wdApp := CreateOleObject('Word.Application');
      except
        MessageBox(0, 'Не удалось запустить MS Word. Действие отменено.'
          ,'Внимание!', MB_OK + MB_ICONERROR + MB_APPLMODAL);
        Exit;
      end;

      //На время отладки делаем видимым окно MS Word.
      wdApp.Visible := True; //После отладки: wdApp.Visible := False;
      //Создаём новый документ.
      wdDoc := wdApp.Documents.Add;
      //Отключение перерисовки окна MS Word, если wdApp.Visible := True.
      //Для ускорения обработки в случае больших текстов.
      wdApp.ScreenUpdating := False;
    end;
      try
      wdRng := wdDoc.Content; //Диапазон, охватывающий всё содержимое документа.
      wdRng.InsertAfter('В высшую аттестационную комиссию');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('Министерства образования и науки РФ');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter(''+DS_Teacher.DataSet.FieldByName('Фамилия').AsString+' '+DS_Teacher.DataSet.FieldByName('Имя').AsString+' '+DS_Teacher.DataSet.FieldByName('Отчество').AsString+'');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('преподаватель '+Label3.Caption+'');
      wdRng.InsertAfter(#13#10);
      wdRng.Font.Name := 'Times New Roman';
      wdRng.Font.Bold := False;
      wdRng.Font.Size := 12;
      wdRng.ParagraphFormat.Alignment := wdAlignParagraphRight;
      wdRng.InsertAfter(#13#10);
      wdRng.Start := wdRng.End;
      wdRng.ParagraphFormat.Reset;
      wdRng.Font.Reset;
      wdRng.InsertAfter('ЗАЯВЛЕНИЕ');
      wdRng.InsertAfter(#13#10);
      wdRng.Font.Name := 'Times New Roman';
      wdRng.Font.Bold := True;
      wdRng.Font.Size := 18;
      wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
      wdRng.Start := wdRng.End;
      wdRng.ParagraphFormat.Reset;
      wdRng.Font.Reset;
      wdRng.InsertAfter('Прошу аттестовать меня в текущем учебном году на ');
      wdRng.InsertAfter(' '+Combobox2.text+' квалификационную категорию по должности преподаватель.');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('С Положением  о порядке аттестации специалистов и руководящих ');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('работников бюджетных  учреждений  сферы  молодежной  политики');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('Российской Федерации ознакомлен(а).');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('Основанием   для   аттестации   на   указанную   в   заявлении ');
      wdRng.InsertAfter('квалификационную категорию считаю следующие результаты работы: ');
      wdRng.InsertAfter('__________________________________________________________________');
      wdRng.InsertAfter('__________________________________________________________________');
      wdRng.InsertAfter('__________________________________________________________________');
      wdRng.InsertAfter('__________________________________________________________________');
      wdRng.InsertAfter('Сообщаю о себе следующие сведения:');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('Образование (какое   образовательное   учреждение   окончил, полученная специальность и квалификация) ');
      wdRng.InsertAfter(''+Edit2.text+'');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('Стаж работы (по специальности) _______ лет, в данной должности _____ лет.');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('"___" ___________ 20__ года            Подпись ________________________');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('                                                               Телефон: дом. __________________');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('                                                               сл. ____________________________');
      wdRng.Font.Name := 'Times New Roman';
      wdRng.Font.Bold := False;
      wdRng.Font.Size := 14;
      wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
       wdRng.Start := wdRng.End;
      wdRng.ParagraphFormat.Reset;
      wdRng.Font.Reset;

      finally
        //Включение перерисовки окна MS Word. В случае, если wdApp.Visible := True.
        wdApp.ScreenUpdating := True;

      //Отключаем режим показа предупреждений.
      wdApp.DisplayAlerts := False;
      try
        //Запись документа в файл.
        wdDoc.SaveAs(FileName:=Sd.FileName);
      finally
        //Включаем режим показа предупреждений.
        wdApp.DisplayAlerts := True;
      end;

      end;
  end;

procedure TAddAttestSPO.DBLookupComboBox1Click(Sender: TObject);
begin
Edit2.Text:=DS_Teacher.DataSet.FieldByName('Образование').AsString;
Edit1.Text:=DS_Teacher.DataSet.FieldByName('Срок действия').AsString;
end;

procedure TAddAttestSPO.FormActivate(Sender: TObject);
begin
Label3.Caption:=MainForm.DBLookupComboBox1.Text;

end;

procedure TAddAttestSPO.FormClose(Sender: TObject; var Action: TCloseAction);
begin
AddAttestSPO.Hide;
AttestSPO.show;
end;

end.
