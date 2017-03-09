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
  Query_Insert.SQL.Add('INSERT INTO ���������(���_��,������,�����,����,[������ ���������],[��� ��],���������)');
  Query_Insert.SQL.Add('VALUES (:idT,:per,:kvk,:date,:status,:idOU,:Osn);');
  Query_Insert.Parameters.ParamByName('idT').Value:=DS_Teacher.DataSet.FieldByName('���').Value;
  Query_Insert.Parameters.ParamByName('per').Value:=Combobox1.Text;
  Query_Insert.Parameters.ParamByName('kvk').Value:=Combobox2.text;
  Query_Insert.Parameters.ParamByName('date').Value:=Calendar1.CalendarDate;
  Query_Insert.Parameters.ParamByName('status').Value:='���������� � ��';
  Query_Insert.Parameters.ParamByName('idOU').Value:=DS_Teacher.DataSet.FieldByName('��� ��').Value;
  Query_Insert.Parameters.ParamByName('osn').Value:='';
  Query_Insert.ExecSQL;
  showmessage('��������� ����������������!');
    begin
      Sd :=SaveDialog1; //SaveDialog1 ��� ������ ���� �� �����.
      //���� ��������� ����� ������� �� ������, �� � �������� ��������� ���� �� �����,
      //� ������� ���������� ����������� ���� ����� ���������.
      if Sd.InitialDir = '' then Sd.InitialDir := ExtractFilePath( ParamStr(0) );
      //������ ������� ���������� �����.
      if not Sd.Execute then Exit;
      //���� ���� � �������� ������ ����������, �� ��������� ������ � �������������.
      if FileExists(Sd.FileName) then begin
        Res := MessageBox(0, '���� � �������� ������ ��� ����������. ������������?'
          ,'��������!', MB_YESNO + MB_ICONQUESTION + MB_APPLMODAL);
        if Res <> IDYES then Exit;
      end;

      //������� ��������� MS Word.
      try
        wdApp := CreateOleObject('Word.Application');
      except
        MessageBox(0, '�� ������� ��������� MS Word. �������� ��������.'
          ,'��������!', MB_OK + MB_ICONERROR + MB_APPLMODAL);
        Exit;
      end;

      //�� ����� ������� ������ ������� ���� MS Word.
      wdApp.Visible := True; //����� �������: wdApp.Visible := False;
      //������ ����� ��������.
      wdDoc := wdApp.Documents.Add;
      //���������� ����������� ���� MS Word, ���� wdApp.Visible := True.
      //��� ��������� ��������� � ������ ������� �������.
      wdApp.ScreenUpdating := False;
    end;
      try
      wdRng := wdDoc.Content; //��������, ������������ �� ���������� ���������.
      wdRng.InsertAfter('� ������ �������������� ��������');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('������������ ����������� � ����� ��');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter(''+DS_Teacher.DataSet.FieldByName('�������').AsString+' '+DS_Teacher.DataSet.FieldByName('���').AsString+' '+DS_Teacher.DataSet.FieldByName('��������').AsString+'');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('������������� '+Label3.Caption+'');
      wdRng.InsertAfter(#13#10);
      wdRng.Font.Name := 'Times New Roman';
      wdRng.Font.Bold := False;
      wdRng.Font.Size := 12;
      wdRng.ParagraphFormat.Alignment := wdAlignParagraphRight;
      wdRng.InsertAfter(#13#10);
      wdRng.Start := wdRng.End;
      wdRng.ParagraphFormat.Reset;
      wdRng.Font.Reset;
      wdRng.InsertAfter('���������');
      wdRng.InsertAfter(#13#10);
      wdRng.Font.Name := 'Times New Roman';
      wdRng.Font.Bold := True;
      wdRng.Font.Size := 18;
      wdRng.ParagraphFormat.Alignment := wdAlignParagraphCenter;
      wdRng.Start := wdRng.End;
      wdRng.ParagraphFormat.Reset;
      wdRng.Font.Reset;
      wdRng.InsertAfter('����� ����������� ���� � ������� ������� ���� �� ');
      wdRng.InsertAfter(' '+Combobox2.text+' ���������������� ��������� �� ��������� �������������.');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('� ����������  � ������� ���������� ������������ � ����������� ');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('���������� ���������  ����������  �����  ����������  ��������');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('���������� ��������� ����������(�).');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('����������   ���   ����������   ��   ���������   �   ��������� ');
      wdRng.InsertAfter('���������������� ��������� ������ ��������� ���������� ������: ');
      wdRng.InsertAfter('__________________________________________________________________');
      wdRng.InsertAfter('__________________________________________________________________');
      wdRng.InsertAfter('__________________________________________________________________');
      wdRng.InsertAfter('__________________________________________________________________');
      wdRng.InsertAfter('������� � ���� ��������� ��������:');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('����������� (�����   ���������������   ����������   �������, ���������� ������������� � ������������) ');
      wdRng.InsertAfter(''+Edit2.text+'');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('���� ������ (�� �������������) _______ ���, � ������ ��������� _____ ���.');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('"___" ___________ 20__ ����            ������� ________________________');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('                                                               �������: ���. __________________');
      wdRng.InsertAfter(#13#10);
      wdRng.InsertAfter('                                                               ��. ____________________________');
      wdRng.Font.Name := 'Times New Roman';
      wdRng.Font.Bold := False;
      wdRng.Font.Size := 14;
      wdRng.ParagraphFormat.Alignment := wdAlignParagraphLeft;
       wdRng.Start := wdRng.End;
      wdRng.ParagraphFormat.Reset;
      wdRng.Font.Reset;

      finally
        //��������� ����������� ���� MS Word. � ������, ���� wdApp.Visible := True.
        wdApp.ScreenUpdating := True;

      //��������� ����� ������ ��������������.
      wdApp.DisplayAlerts := False;
      try
        //������ ��������� � ����.
        wdDoc.SaveAs(FileName:=Sd.FileName);
      finally
        //�������� ����� ������ ��������������.
        wdApp.DisplayAlerts := True;
      end;

      end;
  end;

procedure TAddAttestSPO.DBLookupComboBox1Click(Sender: TObject);
begin
Edit2.Text:=DS_Teacher.DataSet.FieldByName('�����������').AsString;
Edit1.Text:=DS_Teacher.DataSet.FieldByName('���� ��������').AsString;
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
