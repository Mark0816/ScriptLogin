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

//檢查程式是否已被開啟並取得PID
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
    //巡工作管理員檢查
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
      Form1.ListBox1.Items.Add('Exception：' + E.Message);
      Result:=0;
    end;
  end;
end;

//藉由程式PID取得HandleID
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
        EI.HWND := Wnd; // break on return FALSE 所以要反向檢查
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
      Form1.ListBox1.Items.Add('Exception：' + E.Message);
      Result:=0;
    end;
  end;
end;
//依照腳本執行按鍵MenuBar的選項
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
      //Menu下拉後選擇
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
    Form1.ListBox1.Items.Add('未取得FormHandle');
    exit;
  except on E: Exception do
    begin
      Form1.ListBox1.Items.Add('Exception：' + E.Message);
    end;
  end;
end;
//針對Bar的鍵盤事件
procedure KeyBoardEvent(KeyNum:Integer);
begin
  try
    PostMessage(Hw.Bar,WM_Keydown,KeyNum,0);
    PostMessage(Hw.Bar,WM_KeyUp,KeyNum,0);
    sleep(100);
  except on E: Exception do
    begin
      Form1.ListBox1.Items.Add('Exception：' + E.Message);
    end;
  end;
end;
//(儲存密碼)CheckBox勾選
procedure LoginCheckBox();
var
  Ok:Integer;
begin
  Ok:=SendMessage(Hw.CkBox,BM_GETCHECK,0,0);
  //判斷是否已勾選
  if Ok = 0 then begin
    //Button屬性
    SendMessage(Hw.CkBox,BM_CLICK,0,0);
    //CheckBox屬性
    SendMessage(Hw.CkBox,BM_SETCHECK,1,0);
    sleep(500);
  end;
end;
//點按鈕(確認)...
procedure LoginButton();
begin
  SendMessage(Hw.Btn,WM_LBUTTONDOWN,0,0);
  SendMessage(Hw.Btn,WM_LBUTTONUP,0,0);
  sleep(500);
end;
//單選鈕點擊
procedure LoginRadioButton();
begin
  SendMessage(Hw.RdBtn,BM_CLICK,1,0);
  sleep(500);
end;
//下拉選單動作
procedure LoginComboBox(Line:Integer);
begin
  //更改下拉項次
  SendMessage(Hw.CbBox,CB_SETCURSEL,StrToInt(Script.Value[Line])-1,0);
  //觸發Onchange
  SendMessage(Hw.CbBox,WM_COMMAND,MakeLong(0,CBN_SELCHANGE),Hw.CbBox);
  sleep(500);
end;
//輸入登入資料(輸入框)
procedure LoginEdit(Line:Integer);
begin
  //確定是Enable的狀態
  SendMessage(Hw.Edit,WM_LBUTTONDOWN,0,0);
  SendMessage(Hw.Edit,WM_LBUTTONUP,0,0);
  SendMessage(Hw.Edit,WM_SETTEXT,0,Integer(Script.Value[Line]));
  sleep(500);
end;
//輸入登入資料(下拉選單輸入框)
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
//鎖定鍵盤
procedure DisableKeyBoard();
begin
  OldHook := SetWindowsHookEx(WH_KEYBOARD, @KBHook, HINSTANCE, 0);
end;
//解鎖鍵盤
procedure EnableKeyBoard();
begin
  if OldHook <> 0 then
  begin
    UnhookWindowsHookEx(OldHook);
    OldHook := 0;
  end;
end;
//取得目標物件屬性
function GetItemsClass(ItemHwnd : HWND):String;
var
  iClass: array [0 .. 128] of char;
begin
  try
    GetClassName(ItemHwnd, iClass, 128);
    Result := StrPas(iClass);
  except on E: Exception do
    begin
      Form1.ListBox1.Items.Add('Exception：' + E.Message);
      Result:='';
    end;
  end;
end;
//取得目標視窗標題
function GetFormTitle(FormHwnd : HWND):string;
var
  iTitle: array [0 .. 128] of char;
begin
  try
    GetWindowText(FormHwnd, iTitle, 128);
    Result := StrPas(iTitle);
  except on E: Exception do
    begin
      Form1.ListBox1.Items.Add('Exception：' + E.Message);
      Result:='';
    end;
  end;
end;
end.
