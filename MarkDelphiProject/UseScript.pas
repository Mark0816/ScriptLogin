unit UseScript;

interface
  uses System.SysUtils, Vcl.Dialogs, System.Classes, Main;
  procedure Loadtxt();
  procedure LoginQueue(ScriptLine:Integer);
  procedure SetArray();

implementation
//設定腳本變數陣列長度
procedure SetArray();
var
  StrList:TStringList;
  I:Integer;
begin
  StrList:=TStringList.Create;
  try
    try
      StrList.LoadFromFile(ExtractFilePath(Paramstr(0))+'/test.txt', TEncoding.UTF8);
      SetLength(Script.Value,Strlist.Count+1);
      SetLength(Script.Sequence,Strlist.Count);
      for I := 0 to StrList.Count-1 do begin
        if Pos('NextTitle',StrList[I]) > 0 then begin
          SetLength(Script.NextTitleLine,Length(Script.NextTitleLine)+1);
        end;
      end;
      SetLength(Script.Menu,1);
    except on E: Exception do
      begin
        Form1.ListBox1.Items.Add('Exception：' + E.Message);
      end;
    end;
  finally
    StrList.Free;
  end;
end;
 //讀腳本相關資訊
procedure Loadtxt();
var
  I,J,tNum1,tNum2:Integer;
  Strs:TStringList;
begin
  SetArray();
  //抓腳本的位址&名稱
  try
    Script.txt:=TStreamReader.Create(ExtractFilePath(Paramstr(0))+'/test.txt');
//    AssignFile(Script.txt,ExtractFilePath(Paramstr(0))+'/vsphere.txt'); //navicat,workstation,vsphere,ssms,plsql9,plsql14
//    Reset(Script.txt);
  except on E: Exception do
    begin
      Form1.ListBox1.Items.Add('Exception：' + E.Message);
    end;
  end;
  //開始處理腳本內容
  try
    try
      I:=1;
      //掃描到文件字尾
      while not Script.txt.EndOfStream do begin
        //讀取每行放入All
  //      readln(Script.txt,Script.All);
        Script.All:=Script.txt.ReadLine;
        //確認格式
        tNum1:=Pos('[[',Script.All);
        tNum2:=Pos(']]',Script.All);
        if tNum1 <= 0 then begin
          Form1.ListBox1.Items.Add('腳本Key格式錯誤');
          Exit;
        end;
        tNum1:=tNum1+2;
        if tNum2 <= 0 then begin
          Form1.ListBox1.Items.Add('腳本Key格式錯誤');
          Exit;
        end;
        //存放每行Value
        Script.Value[I]:=Copy(Script.All,tNum2+2,Length(Script.All));
        //空白填0
        if (Length(Script.All)-tNum2) = 1 then begin
          Script.Value[I]:='0';
        end;
        //存放每行Key
        Script.Key:=Copy(Script.All,tNum1,tNum2-tNum1);
        //用Key判斷腳本Value的位子：檔案位址
        if (Script.Key = 'FilePath') then begin
          Script.FileLine := I;
          Strs:=TStringList.Create;
          Strs.Delimiter:='\';
          Strs.DelimitedText:=Script.Value[I];
          //取最後\右邊的值
          for J := 0 to Strs.Count-1 do begin
            Script.FileName:=Strs.Strings[J];
          end;
          Strs.Free;
        end;
        //用Key判斷腳本Value的位子
        if (Script.Key = 'InitialTitle') then begin
          Script.TitleLine := I;
        end;

        if (Script.Key = 'MenuClass') then begin
          Script.BarLine := I;
        end;

        if (Script.Key = 'InputClass') then begin
          Script.EditLine := I;
          LoginQueue(Script.EditLine);
        end;

        if (Script.Key = 'ButtonClass') then begin
          Script.BtnLine := I;
          LoginQueue(Script.BtnLine);
        end;

        if (Script.Key = 'CheckBoxClass') then begin
          Script.CheckBoxLine := I;
          LoginQueue(Script.CheckBoxLine);
        end;

        if (Script.Key = 'ComboBoxClass') then begin
          Script.ComboBoxLine := I;
          LoginQueue(Script.ComboBoxLine);
        end;

        if (Script.Key = 'RadioBtnClass') then begin
          Script.RadioBtnLine := I;
          LoginQueue(Script.RadioBtnLine);
        end;

        if (Script.Key = 'ComboBoxEditClass') then begin
          Script.ComboBoxEditLine := I;
          LoginQueue(Script.ComboBoxEditLine);
        end;

        if (Script.Key = 'ComboBoxValue') and (Script.Value[I] <> '0') then begin
          LoginQueue(I-1);
          Script.Value[I]:=Script.All;
        end;

        if (Script.Key = 'InputEditValue') and (Script.Value[I] <> '0') then begin
          LoginQueue(I-1);
          Script.Value[I]:=Script.All;
        end;

        if (Script.Key = 'ComboBoxEditValue') and (Script.Value[I] <> '0') then begin
          LoginQueue(I-1);
          Script.Value[I]:=Script.All;
        end;
        //用Key判斷腳本Value的位子：選單列選取項次
        if (Script.Key = 'MenuValue') and (Script.Value[I] <> '0') then begin
          Strs:=TStringList.Create;
          Strs.Delimiter:='-';
          Strs.DelimitedText:=Script.Value[I];
          SetLength(Script.Menu,Strs.Count+1);
          //以-分隔儲存
          for J := 0 to Strs.Count-1 do begin
            Script.Menu[J]:=StrToInt(Strs.Strings[J]);
          end;
          Strs.Free;
        end;
        //用Key判斷腳本Value的位子：下一個視窗標題
        if (Script.Key = 'NextTitle') then begin
          for J := 0 to Length(Script.NextTitleLine) do begin
            if Script.NextTitleLine[J] = 0 then begin
              Script.NextTitleLine[J]:=I;
              LoginQueue(I);
              break;
            end;
          end;
        end;

        if (Script.Key = 'LastTitle') then begin
          Script.LastTitleLine:=I;
        end;
      I:=I+1;
      end;
    except on E: Exception do
      begin
        Form1.ListBox1.Items.Add('Exception：' + E.Message);
      end;
    end;
  finally
    Script.txt.Free;
  end;
end;
//將腳本先後順序排列
procedure LoginQueue(ScriptLine:Integer);
var
  I: Integer;
begin
  for I := 0 to 30 do begin
    if Script.Sequence[I] =0 then begin
      Script.Sequence[I]:=ScriptLine;
      Exit;
    end;
  end;
end;
end.
