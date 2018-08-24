unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, StrUtils, ExtCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Memo1: TMemo;
    Button2: TButton;
    Memo2: TMemo;
    Memo3: TMemo;
    Label1: TLabel;
    Memo4: TMemo;
    Label2: TLabel;
    Memo5: TMemo;
    Label3: TLabel;
    Panel1: TPanel;
    Label4: TLabel;
    Edit1: TEdit;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Edit7: TEdit;
    Panel2: TPanel;
    Label11: TLabel;
    Label12: TLabel;
    Panel3: TPanel;
    Label14: TLabel;
    Edit9: TEdit;
    Label13: TLabel;
    Label15: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public

    { Public declarations }
  end;

var
  Form1: TForm1;


implementation

{$R *.dfm}
uses ComObj, ExcelXP;

var

  StFile : AnsiString;
  E, Range :variant;
  kol_users: integer;
  Arr_users: array of string;

//-------------Excel-------------------

Function CreateExcel:boolean;
begin
CreateExcel:=true;
try
E:=CreateOleObject('Excel.Application');
except
CreateExcel:=false;
end;
End;

Function VisibleExcel(visible:boolean):boolean;
begin
VisibleExcel:=true;
try
E.visible:=visible;
except
VisibleExcel:=false;
end;
End;

Function AddWorkBook:boolean;
begin
 AddWorkBook:=true;
 try
  E.Workbooks.Add;
 except
  AddWorkBook:=false;
 end;
End;

Function OpenWorkBook(file_: string):boolean;
begin
 OpenWorkBook:=true;
 try
  E.Workbooks.Open(file_);
 except
  OpenWorkBook:=false;
 end;
End;

Function SaveWorkBookAs(file_:string): boolean;
begin
SaveWorkBookAs:=true;
try
E.DisplayAlerts:=False;
E.ActiveWorkbook.SaveAs(file_);
E.DisplayAlerts:=True;
except
SaveWorkBookAs:=false;
end;
End;

Function CloseWorkBook:boolean;
begin
 CloseWorkBook:=true;
 try
  E.ActiveWorkbook.Close;
 except
  CloseWorkBook:=false;
 end;
End;

Function CloseExcel:boolean;
begin
 CloseExcel:=true;
 try
  E.Quit;
 except
  CloseExcel:=false;
 end;
End;

Function FindText (text_:string):boolean;
begin
 FindText:=true;
 try
  E.Cells.Find(what:=text_, matchcase:=True).Select;
 except
  FindText:=False;
 end;
End;

procedure TForm1.Button1Click(Sender: TObject);
var
  List : TStringList;
  selectedFile : String ;
  i, position_zapetoi, ind_users:Integer;
  St, Naidenie:string;


begin
  Form1.Memo1.Clear;
  List := TStringList.Create;

  if PromptForFileName(selectedFile,        // ����� ������������� �����
                       'CSV (�����������-�������)(*.csv)|*.csv',
                       '',
                       '�������� ������ ����',
                       'C:\',
                       False)  // ��������, ��� ������ ��� ����������
    then
      // ����������� ����� ������� �������� �����/����
      //ShowMessage('��������� ���� = '+selectedFile)
      StFile := selectedFile
    else
      begin
        ShowMessage('������ �� �������');
        exit;
      end;

    List.LoadFromFile(StFile);
    //memo1.Lines.LoadFromFile(StFile);
    SetLength(Arr_users,List.Count);

    For i := 1 to List.Count-1 do
      begin
        //ShowMessage(List[i]);
        St:=List[i];
        //ind_users:=1;
        position_zapetoi:=POS(',', ST);
        If position_zapetoi=0
         then
           begin
             //ShowMessage('����� ��������');
             break;
           end;
        Naidenie:=copy(St,1,position_zapetoi-1);
        //ShowMessage(Naidenie);

        Arr_users[i]:=Naidenie;       // ������ ������������� �� AD.
        //ShowMessage(Arr_users[i]);
        Memo1.Lines.Add(Arr_users[i]);
      end;

    kol_users:=i-1;
    //ShowMessage(inttostr(kol_users));
    List.Free;
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  selectedFile, FirstAddress, addr : String;
  str, str2, kb, cn, office, Tel_num_sity, Tel_num_local, Tel_num, street, city, title, department, company: String;
  i, kol_naidenogo: integer;
begin
str2:='.';
kb:='��. ';
//-----�������� Excel -------------
 if not CreateExcel
   then
     exit;
 //messagebox(handle,'','��������� Excel.',0);
 VisibleExcel(true);
 //messagebox(handle,'','���������� Excel �� ������.',0);
 if PromptForFileName(selectedFile,        // ����� ������������� �����
                       'Excel (*.xls)|*.xls',
                       '',
                       '�������� ������ ����',
                       'C:\',
                       False)
   then
     begin
       //ShowMessage('��������� ���� = '+selectedFile);
       StFile := selectedFile;
     end 
   else
     begin
       ShowMessage('������ �� �������');
       exit;
     end;
 OpenWorkBook(StFile);
 // ----------- ����� �� Excel --------------
 VarClear(Range);
 For i:=1 to kol_users do
  begin
    Range := E.Range['c4:c880'].Find(What:=Arr_users[i], LookIn:=xlValues,  SearchDirection:=xlNext, MatchCase:=False);
    if not VarIsClear(Range)
        then
          begin
            kol_naidenogo:=0;
            FirstAddress := Range.Address;

            //ShowMessage(Range.Value);
            //ShowMessage(FirstAddress);

            //kol_naidenogo:=kol_naidenogo+1;
            repeat
              Range := E.Range['c4:c880'].FindNext(After := Range);
              //ShowMessage(Range.Value);
              //ShowMessage(Range.Address);

              kol_naidenogo:=kol_naidenogo+1;
            until FirstAddress = Range.Address;                   // ������� ������ ���������� �� ���� ���������


            If kol_naidenogo=1
              then
                begin
                  // ���������� ldf-������� ������ ������������

                  // ������� ���
                  //j:=j+1;
                  cn:=Arr_users[i];
                  //ShowMessage('������� ���='+cn);
                  str := StringReplace(Memo3.Lines[0], Memo3.Lines[0], 'dn: cn='+Arr_users[i]+',OU=exkad,OU=FRS,DC=rosregistr,DC=local', [rfReplaceAll, rfIgnoreCase]);
                  Memo3.Lines[0]:=str;
                  //ShowMessage(Memo3.Lines[0]);

                  // �������
                  addr:=Range.Address;
                  addr[2]:='D';
                  office:=E.Range[addr].value;
                  office:= Concat(kb, office);
                  //ShowMessage('�������='+office);
                  str := StringReplace(Memo3.Lines[3], Memo3.Lines[3], 'physicalDeliveryOfficeName: '+office, [rfReplaceAll, rfIgnoreCase]);
                  Memo3.Lines[3]:=str;
                  //ShowMessage(Memo3.Lines[3]);

                  // ���������� �����
                  addr:=Range.Address;
                  addr[2]:='G';
                  Tel_num_sity:=E.Range[addr].value;
                  addr:=Range.Address;
                  addr[2]:='F';
                  Tel_num_local:=E.Range[addr].value;
                  Tel_num := Tel_num_sity + '��. ' + Tel_num_local;
                  //ShowMessage('���������� �����='+Tel_num);
                  str := StringReplace(Memo3.Lines[6], Memo3.Lines[6], 'telephoneNumber: '+Tel_num, [rfReplaceAll, rfIgnoreCase]);
                  Memo3.Lines[6]:=str;
                  //ShowMessage(Memo3.Lines[6]);

                  // ������
                  addr:=Range.Address;
                  addr[2]:='I';
                  street:=E.Range[addr].value;
                  street:= Concat(street, str2);
                  //ShowMessage('������='+street);
                  str := StringReplace(Memo3.Lines[9], Memo3.Lines[9], 'streetAddress: '+street, [rfReplaceAll, rfIgnoreCase]);
                  Memo3.Lines[9]:=str;
                  //ShowMessage(Memo3.Lines[9]);

                  // �����
                  city:='������';
                  //ShowMessage('�����='+city);
                  str := StringReplace(Memo3.Lines[12], Memo3.Lines[12], 'l: '+city, [rfReplaceAll, rfIgnoreCase]);
                  Memo3.Lines[12]:=str;
                  //ShowMessage(Memo3.Lines[12]);

                  // ���������
                  addr:=Range.Address;
                  addr[2]:='B';
                  title:=E.Range[addr].value;
                  title:= Concat(title, str2);
                  //ShowMessage('���������='+title);
                  str := StringReplace(Memo3.Lines[15], Memo3.Lines[15], 'title: '+title, [rfReplaceAll, rfIgnoreCase]);
                  Memo3.Lines[15]:=str;
                  //ShowMessage(Memo3.Lines[15]);

                  // �����������
                  addr:=Range.Address;
                  addr[2]:='J';
                  department:=E.Range[addr].value;
                  department:= Concat(department, str2);
                  //ShowMessage('�����������='+department);
                  str := StringReplace(Memo3.Lines[18], Memo3.Lines[18], 'department: '+department, [rfReplaceAll, rfIgnoreCase]);
                  Memo3.Lines[18]:=str;
                  //ShowMessage(Memo3.Lines[18]);

                  // ��������
                  company:='�� ���������';
                  //ShowMessage('��������='+company);
                  str := StringReplace(Memo3.Lines[21], Memo3.Lines[21], 'company: '+company, [rfReplaceAll, rfIgnoreCase]);
                  Memo3.Lines[21]:=str;
                  //ShowMessage(Memo3.Lines[21]);

                  // ���������� ��������� .ldf �����

                   Memo4.Lines.Add(Memo3.lines.Text);

                end
              else
                begin
                  Memo5.Lines.Add(Arr_users[i]);
                  //ShowMessage('� ����������� ���� ����������');
                end;
          end
        else
          begin
            Memo2.Lines.Add(Arr_users[i]);
            //ShowMessage('� ����������� �� �������');
          end;
  end;
  Memo4.Lines.delete(Memo4.Lines.Count-1);
  Memo4.Lines.delete(Memo4.Lines.Count);
  Memo4.Lines.SaveToFile('c:\i_allusers.ldf'); // ��������� �������� � ���� .ldf
  Memo5.Lines.SaveToFile('c:\����������.txt'); // ��������� � ����������� ���������� ��������� � c:\����������.txt
  Memo2.Lines.SaveToFile('c:\�� �������.txt'); // �� ��������� � ����������� ��������� � c:\�� �������.txt
  ShowMessage('������������� ���� �������� � c:\i_allusers.ldf'+#13#10+'��������� � ����������� ���������� ��������� � c:\����������.txt'+#13#10+'�� ��������� � ����������� ��������� � c:\�� �������.txt');
  CloseWorkBook;
  CloseExcel;
end;

end.
