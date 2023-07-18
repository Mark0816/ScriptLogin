unit Func;

interface
  uses Main, TLhelp32, System.SysUtils, Winapi.Windows, Winapi.Messages, Vcl.Dialogs, JobsApi,
        ShellAPI, Vcl.Forms, System.Variants;

function CheckProcessNameGetPID(FileName:String):THandle;
function GetHWndByPID(const hPID: THandle): THandle;
procedure KeyBoardEvent(KeyNum:Integer);
Procedure OpenNewConn();
function GetFormTitle(FormHwnd : HWND):string;
function GetItemsClass(ItemHwnd : HWND):String;
procedure LoginCheckBox();
procedure LoginComboBox(Line:Integer);
procedure LoginButton();
procedure LoginRadioButton();
procedure LoginEdit(Line:Integer);
procedure LoginComboBoxEdit(Line:Integer);
//function KBHook(code: Integer; wparam: Word; lparam: LongInt): LRESULT; stdcall;
procedure DisableKeyBoard();
procedure EnableKeyBoard();

implementation

//�ˬd�{���O�_�w�Q�}�Ҩè��oPID
function CheckProcessNameGetPID(FileName:String):THandle;
var
  Found: Boolean;
  hPL: THandle;
  ProcessStruct: TProcessEntry32;
begin
  try
    Result:=0;
    ProcessExistCount:=0;
    hPL:= CreateToolHelp32SnapShot(TH32CS_SNAPPROCESS, 0);
    ProcessStruct.dwSize:=SizeOf(TProcessEntry32);
    Found:=Process32First(hPL,ProcessStruct);
    //���u�@�޲z���ˬd
    while Found do begin
      if UpperCase(ProcessStruct.szExeFile) = UpperCase(FileName) then begin
        ProcessExistCount:=ProcessExistCount+1;
        Result:=ProcessStruct.th32ProcessID;
      end;
      Found := Process32Next(hPL,ProcessStruct);
    end;
    CloseHandle(hPL);
  except on E: Exception do
    begin
      Form1.ListBox1.Items.Add('Exception�G' + E.Message);
      Result:=0;
    end;
  end;
end;

//�ǥѵ{��PID���oHandleID
function GetHWndByPID(const hPID: THandle): THandle;
type
  PEnumInfo = ^TEnumInfo ;
  TEnumInfo = record
    ProcessID : DWORD ;
    HWND : THandle ;
  end;
  function EnumWindowsProc3(Wnd: DWORD; var EI: TEnumInfo): bool; stdcall;
    var
      PID: DWORD;
    begin
      GetWindowThreadProcessID(Wnd, @PID);
      Result := (PID <> EI.ProcessID) or (not IsWindowVisible(Wnd)) or (not IsWindowEnabled(Wnd));
      if not Result then
        EI.HWND := Wnd; // break on return FALSE �ҥH�n�ϦV�ˬd
    end;
  function FindMainWindow(PID: DWORD): DWORD;
    var
      EI: TEnumInfo;
    begin
      EI.ProcessID := PID;
      EI.HWND := 0;
      EnumWindows(@EnumWindowsProc3, Integer(@EI));
      Result := EI.HWND;
    end;
begin
  try
    if hPID <> 0 then
      Result := FindMainWindow(hPID)
    else
      Result := 0;
  except on E: Exception do
    begin
      Form1.ListBox1.Items.Add('Exception�G' + E.Message);
      Result:=0;
    end;
  end;
end;
//�̷Ӹ}���������MenuBar���ﶵ
Procedure OpenNewConn();
var
  I:Integer;
  tNum:Integer;
begin
  try
    tNum:=1;
    SetForegroundWindow(Hw.Form);
    if(Hw.Form<>0) then begin
//      SetWindowPos(Hw.Form,HWND_TOP,0,2000,0,0,SWP_SHOWWINDOW);
//      SetForegroundWindow(Hw.Top);
      //VK_F10
      KeyBoardEvent(121);
      //VK_F
      KeyBoardEvent(70);
      for I := 1 to Script.Menu[0]-1 do begin
        //VK_RIGHT
        KeyBoardEvent(39);
      end;
      //VK_UP
      KeyBoardEvent(38);
      //VK_DOWN
      KeyBoardEvent(40);
      //Menu�U�ԫ���
      while Script.Menu[tNum] <> 0 do begin
        for I := 1 to Script.Menu[tNum]-1 do begin
          //VK_DOWN
          KeyBoardEvent(40);
        end;
        //enter VK_RETURN
        KeyBoardEvent(13);
        tNum:=tNum+1;
      end;
//      SetWindowPos(Hw.Top,HWND_NOTOPMOST,0,0,0,0,SWP_NOSIZE);
//      ShowWindow(Hw.Form, SW_MAXIMIZE);
      exit;
    end;
    Form1.ListBox1.Items.Add('�����oFormHandle');
    exit;
  except on E: Exception do
    begin
      Form1.ListBox1.Items.Add('Exception�G' + E.Message);
    end;
  end;
end;
//�w��Bar����L�ƥ�
procedure KeyBoardEvent(KeyNum:Integer);
begin
  try
    PostMessage(Hw.Bar,WM_Keydown,KeyNum,0);
    PostMessage(Hw.Bar,WM_KeyUp,KeyNum,0);
    sleep(100);
  except on E: Exception do
    begin
      Form1.ListBox1.Items.Add('Exception�G' + E.Message);
    end;
  end;
end;
//(�x�s�K�X)CheckBox�Ŀ�
procedure LoginCheckBox();
var
  Ok:Integer;
begin
  Ok:=SendMessage(Hw.CkBox,BM_GETCHECK,0,0);
  //�P�_�O�_�w�Ŀ�
  if Ok = 0 then begin
    //Button�ݩ�
    SendMessage(Hw.CkBox,BM_CLICK,0,0);
    //CheckBox�ݩ�
    SendMessage(Hw.CkBox,BM_SETCHECK,1,0);
    sleep(500);
  end;
end;
//�I���s(�T�{)...
procedure LoginButton();
begin
  SendMessage(Hw.Btn,WM_LBUTTONDOWN,0,0);
  SendMessage(Hw.Btn,WM_LBUTTONUP,0,0);
  sleep(500);
end;
//���s�I��
procedure LoginRadioButton();
begin
  SendMessage(Hw.RdBtn,BM_CLICK,1,0);
  sleep(500);
end;
//�U�Կ��ʧ@
procedure LoginComboBox(Line:Integer);
begin
  //���U�Զ���
  SendMessage(Hw.CbBox,CB_SETCURSEL,StrToInt(Script.Value[Line])-1,0);
  //Ĳ�oOnchange
  SendMessage(Hw.CbBox,WM_COMMAND,MakeLong(0,CBN_SELCHANGE),Hw.CbBox);
  sleep(500);
end;
//��J�n�J���(��J��)
procedure LoginEdit(Line:Integer);
begin
  //�T�w�OEnable�����A
  SendMessage(Hw.Edit,WM_LBUTTONDOWN,0,0);
  SendMessage(Hw.Edit,WM_LBUTTONUP,0,0);
  SendMessage(Hw.Edit,WM_SETTEXT,0,Integer(Script.Value[Line]));
  sleep(500);
end;
//��J�n�J���(�U�Կ���J��)
procedure LoginComboBoxEdit(Line:Integer);
begin
  SendMessage(Hw.CbBoxEdit,WM_SETTEXT,0,Integer(Script.Value[Line]));
  sleep(500);
end;
function KBHook(Code: Integer; wparam: Word; lparam: LongInt): LRESULT; stdcall;
begin
  Result := 0;
//  if Code < 0 then begin
//    Result:=CallNextHookEx(OldHook,Code,wparam,lparam);
//  end;
end;
//��w��L
procedure DisableKeyBoard();
begin
  OldHook := SetWindowsHookEx(WH_KEYBOARD, @KBHook, HINSTANCE, 0);
end;
//������L
procedure EnableKeyBoard();
begin
  if OldHook <> 0 then
  begin
    UnhookWindowsHookEx(OldHook);
    OldHook := 0;
  end;
end;
//���o�ؼЪ����ݩ�
function GetItemsClass(ItemHwnd : HWND):String;
var
  iClass: array [0 .. 128] of char;
begin
  try
    GetClassName(ItemHwnd, iClass, 128);
    Result := StrPas(iClass);
  except on E: Exception do
    begin
      Form1.ListBox1.Items.Add('Exception�G' + E.Message);
      Result:='';
    end;
  end;
end;
//���o�ؼе������D
function GetFormTitle(FormHwnd : HWND):string;
var
  iTitle: array [0 .. 128] of char;
begin
  try
    GetWindowText(FormHwnd, iTitle, 128);
    Result := StrPas(iTitle);
  except on E: Exception do
    begin
      Form1.ListBox1.Items.Add('Exception�G' + E.Message);
      Result:='';
    end;
  end;
end;
end.
