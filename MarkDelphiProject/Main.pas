unit Main;

interface

uses
Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, ShellAPI, Vcl.ExtCtrls, TLhelp32, JobsApi;
type
  TForm1 = class(TForm)
    ListBox1: TListBox;
    Button1: TButton;
    ListBox2: TListBox;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

Procedure OpenMyProcess();
Procedure GetFormBarHw();
procedure GetNewConnHwAndLogin();
function EnumWindowsProc(MyHWND: HWND; lParam: LPARAM):   Boolean;stdCall;
function EnumWindowsProcTotal(MyHWND: HWND; lParam: LPARAM):   Boolean;stdCall;

function test(MyHWND: HWND; lParam: LPARAM):   Boolean;stdCall;

type
  Hw_record = record                      //�s��HandleID
    Top: HWND;                            //���h����
    Form: HWND;                           //�D����
    Form2: HWND;                          //�Y��Bar�A�}�Ҫ��ݵn�J����
    Bar: HWND;                            //���D�CMenuBar
    Btn: HWND;                            //���s(Ex:�T�{��)
    CkBox: HWND;                          //�ƿ��
    CbBox: HWND;                          //�U�Կ��
    RdBtn: HWND;                          //���s
    Edit: HWND;                           //��J��
    CbBoxEdit: HWND;                      //�U�Կ�檺��J��
    BarCount: Integer;                    //�p��n��ĴX��BarHandleID
    BtnCount: Integer;                    //�p��n��ĴX��BtnHandleID
    EditCount: Integer;                   //�p��n��ĴX��EditHandleID
    CkBoxCount: Integer;                  //�p��n��ĴX��CkBoxHandleID
    CbBoxCount: Integer;                  //�p��n��ĴX��CbBoxHandleID
    RdBtnCount: Integer;                  //�p��n��ĴX��RdBtnHandleID
    CbBoxEditCount: Integer;              //�p��n��ĴX��CbBoxEditHandleID
  end;

type Script_record = record
  Key,All:string;                   //All�G�}�����,Key�G[[]]����
  Value:array of String;            //]]�k���
//  txt:TextFile;                     //��Ӥ�r���
  txt:TStreamReader;                //��Ӥ�r���
  FileLine:Integer;                 //���ε{����}�b�}��Value����l
  FileName:String;                  //���ε{���b�u�@�޲z�����ɦW
  TitleLine:Integer;                //���ε{�����D�b�}��Value����l
  BarLine:Integer;                  //ClassBar�b�}��Value����l
  EditLine:Integer;                 //ClassEdit�b�}��Value����l
  BtnLine:Integer;                  //ClassBtn�b�}��Value����l
  CheckBoxLine:Integer;             //CkBox�b�}��Value����l
  ComboBoxLine:Integer;             //CbBox�b�}��Value����l
  RadioBtnLine:Integer;             //RdBtn�b�}��Value����l
  ComboBoxEditLine:Integer;         //CbBoxEdit�b�}��Value����l
  DelayLine:Integer;
  LastTitleLine:Integer;
  NextTitleLine:array of Integer;   //NextTitle�b�}��Value����l
  Menu:array of Integer;            //���C�n�諸����
  Sequence:array of Integer;        //�n�J���ǱƦC
end;

var
  Form1: TForm1;
  Hw:Hw_record;
  Script:Script_record;
  ClassBuf,TitleBuf: array[Byte] of Char; //�s��Ҧ��l������D&�ݩ�
  CbBoxSum:Integer;                       //�p��CbBox���`��
  EditSum:Integer;                        //�p��Edit���`��
  CbBoxEditSum:Integer;                   //�p��CbBoxEdit���`��
  ProcessExistCount:Integer;              //���ε{���O�_�w�}��
  ProcessId:THandle;                      //���ε{��PID
  OldHook:HHOOK;                          //�_�l

implementation

uses UseScript,Func;

{$R *.dfm}
//��w��L �ޥ�DLL
function SetHook():boolean;stdcall; external 'KeyBoardHook.dll';
//�Ѱ���w��L �ޥ�DLL
function DelHook():boolean;stdcall; external 'KeyBoardHook.dll';

procedure TForm1.Button1Click(Sender: TObject);
var
  i:integer;
begin
  EnumChildWindows(Hw.Form2,@test,0);
  for I := 0 to 10 do begin
    Form1.ListBox2.Items.Add(IntToStr(Script.Sequence[I]));
  end;
end;
procedure TForm1.FormCreate(Sender: TObject);
begin
  Hw.Top:=GetForegroundWindow();
  Loadtxt();
  //�Y���X���~�K���_�D�{��
  if ListBox1.Items.Count=0 then begin
    OpenMyProcess();
  end;
  if ListBox1.Items.Count=0 then begin
    GetFormBarHw();
  end;
  if ListBox1.Items.Count=0 then begin
    GetNewConnHwAndLogin();
  end;
//  DelHook();
//  EnableKeyBoard();
end;
//�}�ҵ{���A�[�J�u�@�s��
Procedure OpenMyProcess();
var
  jLimit: TJobObjectExtendedLimitInformation;
  hJob : Integer;
  hApp : Cardinal;
  ExecInfo:TShellExecuteInfo;
  FormTitle:String;
  tTime,J:Integer;
  TempProcess:Cardinal;
begin
  TempProcess:=0;
  try
    hJob := CreateJobObject(nil, PChar(TimeToStr(Now)));
    if hJob <> 0 then
    begin
      jLimit.BasicLimitInformation.LimitFlags := JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
      SetInformationJobObject(hJob, JobObjectExtendedLimitInformation, @jLimit, SizeOf(TJobObjectExtendedLimitInformation));
    end;
    ZeroMemory(@ExecInfo,SizeOf(ExecInfo));
    //�}�����ε{�����һݸ�T
    With ExecInfo do begin
      cbSize := sizeof(ShellExecuteinfo);
      fMask := SEE_MASK_NOCLOSEPROCESS;
      lpVerb := 'open';
      lpFile := pchar(Script.Value[Script.FileLine]);
      lpParameters := nil;
      Wnd := Application.Handle;
      nShow := SW_SHOWNORMAL;
    end;
    //�}�����ε{��
    ShellExecuteEX(@ExecInfo);
    //���ݵ{���}��
    WaitForInputIdle(ExecInfo.hProcess,INFINITE);
    hApp:=ExecInfo.hProcess;
    //�}��FilePath���~
    if hApp = 0 then begin
      Form1.ListBox1.Items.Add('OpenProcessFail...ShellExecuteFail');
      Abort;
    end;
    sleep(500);
    ProcessId:=GetProcessId(ExecInfo.hProcess);
    //���ݵ{���}�Ҥ����D�����l�������J����
    if FormTitle <> Script.Value[Script.TitleLine] then begin
      tTime:=0;
      J:=0;
      repeat
        Hw.Form:=GetHWndByPID(ProcessId);
        FormTitle:=GetFormTitle(Hw.Form);
        sleep(100);
        //�C0.5���ˬd�@���O�_���u���@�}�ҵ{��
        if (tTime = 5*J) then begin
        J:=J+1;
          //�Y���ε{�����u���@�}�ҥB�w�Q�}�ҡA�ϥΤw�}�Ҫ�PID(�ثe�w��Navicat)
          if (Hw.Form = 0) AND (FormTitle = '') then begin
            //�ˬd���ε{���O�_�s�b�ç��PID
            ProcessId:=CheckProcessNameGetPID(Script.FileName);
            //�ǥ�PID��Process
            TempProcess:=OpenProcess(PROCESS_ALL_ACCESS,False,ProcessId);
            Hw.Form:=GetHWndByPID(ProcessId);
            FormTitle:=GetFormTitle(Hw.Form);
            if ProcessExistCount = 0 then begin
              Form1.ListBox1.Items.Add('���ε{�����}�ҩΤw�Q����');
              break;
            end;
            if (FormTitle <> Script.Value[Script.TitleLine]) AND (FormTitle <> '') then begin
              Form1.ListBox1.Items.Add('[[InitialTitle]]���~�����ε{�����h�������~');
              break;
            end;
          end;
        end;
        tTime:=tTime+1;
        //�]60��W��
        if (tTime=600) AND (FormTitle <> Script.Value[Script.TitleLine]) then begin
          Form1.ListBox1.Items.Add('OpenProcessFail...TimeOut');
          break;
        end;
      until (FormTitle = Script.Value[Script.TitleLine]);
    end;
    //�[�J�u�@�s��
    if hApp <> INVALID_HANDLE_VALUE then begin
      AssignProcessToJobObject(hJob, hApp);
      AssignProcessToJobObject(hJob, TempProcess);
    end;
    SetForegroundWindow(Hw.Form);
  except on E: Exception do
    begin
      Form1.ListBox1.Items.Add('Exception�G'+E.Message);
    end;
  end;
end;
//���o�DForm������HandleID
Procedure GetFormBarHw();
begin
  try
    //�}����Bar�����~����
    if (Script.BarLine > 0) then begin
      Hw.BarCount:=StrToInt(Script.Value[Script.BarLine+1]);
      EnumChildWindows(Hw.Form,@EnumWindowsProc,0);
//      DisableKeyBoard();
//      SetHook();
//      EnableWindow(Hw.Form,False);
      OpenNewConn();
//      EnableWindow(Hw.Form,True);
    end;
  except on E: Exception do
    begin
    Form1.ListBox1.Items.Add('Exception�G'+E.Message);
    end;
  end;
end;
//���oForm2&Form2������HandleID��n�J
procedure GetNewConnHwAndLogin();
var
  I,J,tTime:Integer;
  FormTitle:String;
begin
  try
    J:=0;
    if Script.Menu[0] = 0 then begin
      Hw.Form2:=Hw.Form;
      SetForegroundWindow(Hw.Form2);
    end;
    //��Form2���U�ؼЪ���HandleID��n�J�A�ë��Ӹ}�����Ƕi��
    While(Script.Sequence[J]<>0) do begin
      EditSum:=0;
      CbBoxSum:=0;
      CbBoxEditSum:=0;
      //��Form2���U�ؼЪ����`��
      EnumChildWindows(Hw.Form2,@EnumWindowsProcTotal,0);
      //���s
      if Script.Value[Script.Sequence[J]] = Script.Value[Script.BtnLine] then begin
        Hw.BtnCount:=StrToInt(Script.Value[Script.Sequence[J]+1]);
        EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
        LoginButton();
      end;
      //��J��
      if Script.Value[Script.Sequence[J]] = Script.Value[Script.EditLine] then begin
//        Hw.EditCount:=EditSum;
//        EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
      end;
      //�ƿ��
      if Script.Value[Script.Sequence[J]] = Script.Value[Script.CheckBoxLine] then begin
        Hw.CkBoxCount:=StrToInt(Script.Value[Script.Sequence[J]+1]);
        EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
        LoginCheckBox();
      end;
      //�U�Կ��
      if Script.Value[Script.Sequence[J]] = Script.Value[Script.ComboBoxLine] then begin
//        Hw.CbBoxCount:=CbBoxSum;
//        EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
      end;
      //���s
      if Script.Value[Script.Sequence[J]] = Script.Value[Script.RadioBtnLine] then begin
        Hw.RdBtnCount:=StrToInt(Script.Value[Script.Sequence[J]+1]);
        EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
        LoginRadioButton();
      end;
      //�U�Կ���J��
      if Script.Value[Script.Sequence[J]] = Script.Value[Script.ComboBoxEditLine] then begin
//        Hw.CbBoxEditCount:=CbBoxEditSum;
//        EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
      end;
      //�}��Value���Ʀr�ɡA�n�P�_�w����Ӫ���n�J--ComboBox
      for I := 1 to CbBoxSum do begin
        if (Pos('ComboBoxValue',Script.Value[Script.Sequence[J]+1]) > 0) AND (IntToStr(I) = Script.Value[Script.Sequence[J]]) then begin
          //�]�o��Value�O�}�����A�^���U�Կ��n�諸����
          Script.Value[Script.Sequence[J]+1]:=Copy(Script.Value[Script.Sequence[J]+1],Pos(']]',Script.Value[Script.Sequence[J]+1])+2,Length(Script.Value[Script.Sequence[J]+1]));
          Hw.CbBoxCount := I;
          EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
          LoginComboBox(Script.Sequence[J]+1);
        end;
      end;
      //�}��Value���Ʀr�ɡA�n�P�_�w����Ӫ���n�J--Edit
      for I := 1 to EditSum do begin
        if (Pos('InputEditValue',Script.Value[Script.Sequence[J]+1]) > 0) AND (IntToStr(I) = Script.Value[Script.Sequence[J]]) then begin
          //�]�o��Value�O�}�����A�^����J�حn��J�����
          Script.Value[Script.Sequence[J]+1]:=Copy(Script.Value[Script.Sequence[J]+1],Pos(']]',Script.Value[Script.Sequence[J]+1])+2,Length(Script.Value[Script.Sequence[J]+1]));
          Hw.EditCount := I;
          EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
          LoginEdit(Script.Sequence[J]+1);
        end;
      end;
      //�}��Value���Ʀr�ɡA�n�P�_�w����Ӫ���n�J--ComboBoxEdit
      for I := 1 to CbBoxEditSum do begin
        if (Pos('ComboBoxEditValue',Script.Value[Script.Sequence[J]+1]) > 0) AND (IntToStr(I) = Script.Value[Script.Sequence[J]]) then begin
          //�]�o��Value�O�}�����A�^���U�Կ���J�حn��J�����
          Script.Value[Script.Sequence[J]+1]:=Copy(Script.Value[Script.Sequence[J]+1],Pos(']]',Script.Value[Script.Sequence[J]+1])+2,Length(Script.Value[Script.Sequence[J]+1]));
          Hw.CbBoxEditCount := I;
          EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
          LoginComboBoxEdit(Script.Sequence[J]+1);
        end;
      end;
      //�}��Value���U�@�������D�A�����Form��HandleID
      for I := 0 to Length(Script.NextTitleLine)-1 do begin
        if Script.Value[Script.Sequence[J]] = Script.Value[Script.NextTitleLine[I]] then begin
          tTime:=0;
          repeat
            tTime:=tTime+1;
            sleep(500);
            Hw.Form2:=GetHWndByPID(ProcessId);
            FormTitle:=GetFormTitle(Hw.Form2);
            //�]5��W�ɡA���P�_�������S�����\�}��
            if (Hw.Form = Hw.Form2) And (tTime = 10) then begin
              Form1.ListBox1.Items.Add('TIMEOUT...�N�n�ؼФU�@�h����HWND�������');
              exit;
            end;
            //�A�����D
            if (FormTitle <> Script.Value[Script.NextTitleLine[I]]) And (tTime = 10) then begin
              Form1.ListBox1.Items.Add('TIMEOUT...�N�n�ؼФU�@�h�����}�����D���~');
              exit;
            end;
            SetForegroundWindow(Hw.Form2);
          until (FormTitle = Script.Value[Script.NextTitleLine[I]]);
          Hw.Form:=Hw.Form2;
//          break;
        end;
      end;
      J:=J+1;
    end;
    //�T�w�s�u�O�_���\���\��A�����D
//    if (Form1.ListBox1.Items.Count=0) AND ((GetFormTitle(GetHWndByPID(ProcessId))) = (Script.Value[Script.LastTitleLine])) then begin
//      Form1.ListBox1.Items.Add(Script.Value[Script.LastTitleLine]+' LoginSuccess');
//    end else
//    begin
//      Form1.ListBox1.Items.Add(Script.Value[Script.LastTitleLine]+' LoginFail');
//    end;;
  except on E: Exception do
    begin
    Form1.ListBox1.Items.Add('Exception�G'+E.Message);
    end;
  end;
end;
//�T�|�Ҧ��l����p���`��
function EnumWindowsProcTotal(MyHWND: HWND; lParam: LPARAM):   Boolean;stdCall;
begin
  Result:=True;
  GetClassName(Myhwnd,ClassBuf,SizeOf(ClassBuf));
  //�p���J���`��
  if ClassBuf = Script.Value[Script.EditLine] then begin
    EditSum:=EditSum+1;
  end;
  //�p��U�Կ���`��
  if ClassBuf = Script.Value[Script.ComboBoxLine] then begin
    CbBoxSum:=CbBoxSum+1;
  end;
  //�p��U�Կ���J���`��
  if ClassBuf = Script.Value[Script.ComboBoxEditLine] then begin
    CbBoxEditSum:=CbBoxEditSum+1;
  end;
end;
//�T�|�Ҧ��l������ؼЪ���HandleID
function EnumWindowsProc(MyHWND: HWND; lParam: LPARAM):   Boolean;stdCall;
var
  Hwbuf: HWND;
begin
  Result:=True;
  Hwbuf:=MyHWND;
  //�쪫���ݩ�Class
  GetClassName(Myhwnd,ClassBuf,SizeOf(ClassBuf));
  //���ClassName--MenuBar,��ؼв�N�Ӫ���HandleID
  if ClassBuf = Script.Value[Script.BarLine] then begin
    if Hw.BarCount > 0 then begin
      Hw.Bar:=Hwbuf;
    end;
    Hw.BarCount:=Hw.BarCount-1;
  end;
  //���ClassName--Button,��ؼв�N�Ӫ���HandleID
  if ClassBuf = Script.Value[Script.BtnLine] then begin
    if Hw.BtnCount > 0 then begin
      Hw.Btn:=Hwbuf;
    end;
    Hw.BtnCount:=Hw.BtnCount-1;
  end;
  //��J��HandleId����
  if ClassBuf = Script.Value[Script.EditLine] then begin
    if Hw.EditCount > 0 then begin
//      Hw.Edit[abs(Hw.EditCount-(EditSum+1))]:=Hwbuf;
      Hw.Edit:=Hwbuf;
    end;
    Hw.EditCount:=Hw.EditCount-1;
  end;

  if ClassBuf = Script.Value[Script.CheckBoxLine] then begin
    if Hw.CkBoxCount > 0 then begin
      Hw.CkBox:=Hwbuf;
    end;
    Hw.CkBoxCount:=Hw.CkBoxCount-1;
  end;
  //�U�Կ��HandleId����
  if ClassBuf = Script.Value[Script.ComboBoxLine] then begin
    if Hw.CbBoxCount > 0 then begin
//      Hw.CbBox[abs(Hw.CbBoxCount-(CbBoxSum+1))]:=Hwbuf;
      Hw.CbBox:=Hwbuf;
    end;
    Hw.CbBoxCount:=Hw.CbBoxCount-1;
  end;

  if ClassBuf = Script.Value[Script.RadioBtnLine] then begin
    if Hw.RdBtnCount > 0 then begin
      Hw.RdBtn:=Hwbuf;
    end;
    Hw.RdBtnCount:=Hw.RdBtnCount-1;
  end;
  //�U�Կ���J��HandleId����
  if ClassBuf = Script.Value[Script.ComboBoxEditLine] then begin
    if Hw.CbBoxEditCount > 0 then begin
//      Hw.CbBoxEdit[abs(Hw.CbBoxEditCount-(CbBoxEditSum+1))]:=Hwbuf;
      Hw.CbBoxEdit:=Hwbuf;
    end;
    Hw.CbBoxEditCount:=Hw.CbBoxEditCount-1;
  end;
  sleep(1);
end;
function test(MyHWND: HWND; lParam: LPARAM):   Boolean;stdCall;
begin
  Result:=True;
  GetClassName(Myhwnd,ClassBuf,SizeOf(ClassBuf));
  GetWindowText(Myhwnd,TitleBuf,SizeOf(TitleBuf));
  Form1.ListBox2.Items.Add(ClassBuf);
//  Form1.ListBox2.Items.Add(TitleBuf);
  Form1.ListBox2.Items.Add(IntToStr(MyHWND));
end;
end.
