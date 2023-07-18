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
  Hw_record = record                      //存放HandleID
    Top: HWND;                            //頂層視窗
    Form: HWND;                           //主視窗
    Form2: HWND;                          //若有Bar，開啟的待登入視窗
    Bar: HWND;                            //標題列MenuBar
    Btn: HWND;                            //按鈕(Ex:確認鍵)
    CkBox: HWND;                          //複選框
    CbBox: HWND;                          //下拉選單
    RdBtn: HWND;                          //單選鈕
    Edit: HWND;                           //輸入框
    CbBoxEdit: HWND;                      //下拉選單的輸入框
    BarCount: Integer;                    //計算要找第幾個BarHandleID
    BtnCount: Integer;                    //計算要找第幾個BtnHandleID
    EditCount: Integer;                   //計算要找第幾個EditHandleID
    CkBoxCount: Integer;                  //計算要找第幾個CkBoxHandleID
    CbBoxCount: Integer;                  //計算要找第幾個CbBoxHandleID
    RdBtnCount: Integer;                  //計算要找第幾個RdBtnHandleID
    CbBoxEditCount: Integer;              //計算要找第幾個CbBoxEditHandleID
  end;

type Script_record = record
  Key,All:string;                   //All：腳本整行,Key：[[]]內值
  Value:array of String;            //]]右方值
//  txt:TextFile;                     //整個文字文件
  txt:TStreamReader;                //整個文字文件
  FileLine:Integer;                 //應用程式位址在腳本Value的位子
  FileName:String;                  //應用程式在工作管理員的檔名
  TitleLine:Integer;                //應用程式標題在腳本Value的位子
  BarLine:Integer;                  //ClassBar在腳本Value的位子
  EditLine:Integer;                 //ClassEdit在腳本Value的位子
  BtnLine:Integer;                  //ClassBtn在腳本Value的位子
  CheckBoxLine:Integer;             //CkBox在腳本Value的位子
  ComboBoxLine:Integer;             //CbBox在腳本Value的位子
  RadioBtnLine:Integer;             //RdBtn在腳本Value的位子
  ComboBoxEditLine:Integer;         //CbBoxEdit在腳本Value的位子
  DelayLine:Integer;
  LastTitleLine:Integer;
  NextTitleLine:array of Integer;   //NextTitle在腳本Value的位子
  Menu:array of Integer;            //選單列要選的項次
  Sequence:array of Integer;        //登入順序排列
end;

var
  Form1: TForm1;
  Hw:Hw_record;
  Script:Script_record;
  ClassBuf,TitleBuf: array[Byte] of Char; //存放所有子物件標題&屬性
  CbBoxSum:Integer;                       //計算CbBox的總數
  EditSum:Integer;                        //計算Edit的總數
  CbBoxEditSum:Integer;                   //計算CbBoxEdit的總數
  ProcessExistCount:Integer;              //應用程式是否已開啟
  ProcessId:THandle;                      //應用程式PID
  OldHook:HHOOK;                          //鉤子

implementation

uses UseScript,Func;

{$R *.dfm}
//鎖定鍵盤 引用DLL
function SetHook():boolean;stdcall; external 'KeyBoardHook.dll';
//解除鎖定鍵盤 引用DLL
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
  //若有出錯誤便中斷主程式
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
//開啟程式，加入工作群組
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
    //開啟應用程式之所需資訊
    With ExecInfo do begin
      cbSize := sizeof(ShellExecuteinfo);
      fMask := SEE_MASK_NOCLOSEPROCESS;
      lpVerb := 'open';
      lpFile := pchar(Script.Value[Script.FileLine]);
      lpParameters := nil;
      Wnd := Application.Handle;
      nShow := SW_SHOWNORMAL;
    end;
    //開啟應用程式
    ShellExecuteEX(@ExecInfo);
    //等待程式開啟
    WaitForInputIdle(ExecInfo.hProcess,INFINITE);
    hApp:=ExecInfo.hProcess;
    //腳本FilePath錯誤
    if hApp = 0 then begin
      Form1.ListBox1.Items.Add('OpenProcessFail...ShellExecuteFail');
      Abort;
    end;
    sleep(500);
    ProcessId:=GetProcessId(ExecInfo.hProcess);
    //等待程式開啟比對標題直到初始頁面載入完成
    if FormTitle <> Script.Value[Script.TitleLine] then begin
      tTime:=0;
      J:=0;
      repeat
        Hw.Form:=GetHWndByPID(ProcessId);
        FormTitle:=GetFormTitle(Hw.Form);
        sleep(100);
        //每0.5秒檢查一次是否為只能單一開啟程式
        if (tTime = 5*J) then begin
        J:=J+1;
          //若應用程式為只能單一開啟且已被開啟，使用已開啟的PID(目前針對Navicat)
          if (Hw.Form = 0) AND (FormTitle = '') then begin
            //檢查應用程式是否存在並抓取PID
            ProcessId:=CheckProcessNameGetPID(Script.FileName);
            //藉由PID抓Process
            TempProcess:=OpenProcess(PROCESS_ALL_ACCESS,False,ProcessId);
            Hw.Form:=GetHWndByPID(ProcessId);
            FormTitle:=GetFormTitle(Hw.Form);
            if ProcessExistCount = 0 then begin
              Form1.ListBox1.Items.Add('應用程式未開啟或已被關閉');
              break;
            end;
            if (FormTitle <> Script.Value[Script.TitleLine]) AND (FormTitle <> '') then begin
              Form1.ListBox1.Items.Add('[[InitialTitle]]錯誤或應用程式頂層視窗錯誤');
              break;
            end;
          end;
        end;
        tTime:=tTime+1;
        //設60秒超時
        if (tTime=600) AND (FormTitle <> Script.Value[Script.TitleLine]) then begin
          Form1.ListBox1.Items.Add('OpenProcessFail...TimeOut');
          break;
        end;
      until (FormTitle = Script.Value[Script.TitleLine]);
    end;
    //加入工作群組
    if hApp <> INVALID_HANDLE_VALUE then begin
      AssignProcessToJobObject(hJob, hApp);
      AssignProcessToJobObject(hJob, TempProcess);
    end;
    SetForegroundWindow(Hw.Form);
  except on E: Exception do
    begin
      Form1.ListBox1.Items.Add('Exception：'+E.Message);
    end;
  end;
end;
//取得主Form內物件HandleID
Procedure GetFormBarHw();
begin
  try
    //腳本有Bar項次才執行
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
    Form1.ListBox1.Items.Add('Exception：'+E.Message);
    end;
  end;
end;
//取得Form2&Form2內物件HandleID後登入
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
    //抓Form2底下目標物件HandleID後登入，並按照腳本順序進行
    While(Script.Sequence[J]<>0) do begin
      EditSum:=0;
      CbBoxSum:=0;
      CbBoxEditSum:=0;
      //找Form2底下目標物件總數
      EnumChildWindows(Hw.Form2,@EnumWindowsProcTotal,0);
      //按鈕
      if Script.Value[Script.Sequence[J]] = Script.Value[Script.BtnLine] then begin
        Hw.BtnCount:=StrToInt(Script.Value[Script.Sequence[J]+1]);
        EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
        LoginButton();
      end;
      //輸入框
      if Script.Value[Script.Sequence[J]] = Script.Value[Script.EditLine] then begin
//        Hw.EditCount:=EditSum;
//        EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
      end;
      //複選框
      if Script.Value[Script.Sequence[J]] = Script.Value[Script.CheckBoxLine] then begin
        Hw.CkBoxCount:=StrToInt(Script.Value[Script.Sequence[J]+1]);
        EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
        LoginCheckBox();
      end;
      //下拉選單
      if Script.Value[Script.Sequence[J]] = Script.Value[Script.ComboBoxLine] then begin
//        Hw.CbBoxCount:=CbBoxSum;
//        EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
      end;
      //單選鈕
      if Script.Value[Script.Sequence[J]] = Script.Value[Script.RadioBtnLine] then begin
        Hw.RdBtnCount:=StrToInt(Script.Value[Script.Sequence[J]+1]);
        EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
        LoginRadioButton();
      end;
      //下拉選單輸入框
      if Script.Value[Script.Sequence[J]] = Script.Value[Script.ComboBoxEditLine] then begin
//        Hw.CbBoxEditCount:=CbBoxEditSum;
//        EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
      end;
      //腳本Value為數字時，要判斷針對哪個物件登入--ComboBox
      for I := 1 to CbBoxSum do begin
        if (Pos('ComboBoxValue',Script.Value[Script.Sequence[J]+1]) > 0) AND (IntToStr(I) = Script.Value[Script.Sequence[J]]) then begin
          //因這行Value是腳本整行，擷取下拉選單要選的項次
          Script.Value[Script.Sequence[J]+1]:=Copy(Script.Value[Script.Sequence[J]+1],Pos(']]',Script.Value[Script.Sequence[J]+1])+2,Length(Script.Value[Script.Sequence[J]+1]));
          Hw.CbBoxCount := I;
          EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
          LoginComboBox(Script.Sequence[J]+1);
        end;
      end;
      //腳本Value為數字時，要判斷針對哪個物件登入--Edit
      for I := 1 to EditSum do begin
        if (Pos('InputEditValue',Script.Value[Script.Sequence[J]+1]) > 0) AND (IntToStr(I) = Script.Value[Script.Sequence[J]]) then begin
          //因這行Value是腳本整行，擷取輸入框要輸入的資料
          Script.Value[Script.Sequence[J]+1]:=Copy(Script.Value[Script.Sequence[J]+1],Pos(']]',Script.Value[Script.Sequence[J]+1])+2,Length(Script.Value[Script.Sequence[J]+1]));
          Hw.EditCount := I;
          EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
          LoginEdit(Script.Sequence[J]+1);
        end;
      end;
      //腳本Value為數字時，要判斷針對哪個物件登入--ComboBoxEdit
      for I := 1 to CbBoxEditSum do begin
        if (Pos('ComboBoxEditValue',Script.Value[Script.Sequence[J]+1]) > 0) AND (IntToStr(I) = Script.Value[Script.Sequence[J]]) then begin
          //因這行Value是腳本整行，擷取下拉選單輸入框要輸入的資料
          Script.Value[Script.Sequence[J]+1]:=Copy(Script.Value[Script.Sequence[J]+1],Pos(']]',Script.Value[Script.Sequence[J]+1])+2,Length(Script.Value[Script.Sequence[J]+1]));
          Hw.CbBoxEditCount := I;
          EnumChildWindows(Hw.Form2,@EnumWindowsProc,0);
          LoginComboBoxEdit(Script.Sequence[J]+1);
        end;
      end;
      //腳本Value為下一視窗標題，抓視窗Form的HandleID
      for I := 0 to Length(Script.NextTitleLine)-1 do begin
        if Script.Value[Script.Sequence[J]] = Script.Value[Script.NextTitleLine[I]] then begin
          tTime:=0;
          repeat
            tTime:=tTime+1;
            sleep(500);
            Hw.Form2:=GetHWndByPID(ProcessId);
            FormTitle:=GetFormTitle(Hw.Form2);
            //設5秒超時，先判斷視窗有沒有成功開啟
            if (Hw.Form = Hw.Form2) And (tTime = 10) then begin
              Form1.ListBox1.Items.Add('TIMEOUT...代登目標下一層視窗HWND抓取失敗');
              exit;
            end;
            //再比對標題
            if (FormTitle <> Script.Value[Script.NextTitleLine[I]]) And (tTime = 10) then begin
              Form1.ListBox1.Items.Add('TIMEOUT...代登目標下一層視窗腳本標題錯誤');
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
    //確定連線是否成功的功能，比對標題
//    if (Form1.ListBox1.Items.Count=0) AND ((GetFormTitle(GetHWndByPID(ProcessId))) = (Script.Value[Script.LastTitleLine])) then begin
//      Form1.ListBox1.Items.Add(Script.Value[Script.LastTitleLine]+' LoginSuccess');
//    end else
//    begin
//      Form1.ListBox1.Items.Add(Script.Value[Script.LastTitleLine]+' LoginFail');
//    end;;
  except on E: Exception do
    begin
    Form1.ListBox1.Items.Add('Exception：'+E.Message);
    end;
  end;
end;
//枚舉所有子物件計算總數
function EnumWindowsProcTotal(MyHWND: HWND; lParam: LPARAM):   Boolean;stdCall;
begin
  Result:=True;
  GetClassName(Myhwnd,ClassBuf,SizeOf(ClassBuf));
  //計算輸入框總數
  if ClassBuf = Script.Value[Script.EditLine] then begin
    EditSum:=EditSum+1;
  end;
  //計算下拉選單總數
  if ClassBuf = Script.Value[Script.ComboBoxLine] then begin
    CbBoxSum:=CbBoxSum+1;
  end;
  //計算下拉選單輸入框總數
  if ClassBuf = Script.Value[Script.ComboBoxEditLine] then begin
    CbBoxEditSum:=CbBoxEditSum+1;
  end;
end;
//枚舉所有子物件取目標物件HandleID
function EnumWindowsProc(MyHWND: HWND; lParam: LPARAM):   Boolean;stdCall;
var
  Hwbuf: HWND;
begin
  Result:=True;
  Hwbuf:=MyHWND;
  //抓物件屬性Class
  GetClassName(Myhwnd,ClassBuf,SizeOf(ClassBuf));
  //比對ClassName--MenuBar,抓目標第N個物件HandleID
  if ClassBuf = Script.Value[Script.BarLine] then begin
    if Hw.BarCount > 0 then begin
      Hw.Bar:=Hwbuf;
    end;
    Hw.BarCount:=Hw.BarCount-1;
  end;
  //比對ClassName--Button,抓目標第N個物件HandleID
  if ClassBuf = Script.Value[Script.BtnLine] then begin
    if Hw.BtnCount > 0 then begin
      Hw.Btn:=Hwbuf;
    end;
    Hw.BtnCount:=Hw.BtnCount-1;
  end;
  //輸入框HandleId全抓
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
  //下拉選單HandleId全抓
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
  //下拉選單輸入框HandleId全抓
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
