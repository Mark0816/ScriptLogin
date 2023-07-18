unit UseScript;

interface
  uses System.SysUtils, Vcl.Dialogs, System.Classes, Main;
  procedure Loadtxt();
  procedure LoginQueue(ScriptLine:Integer);
  procedure SetArray();

implementation
//�]�w�}���ܼư}�C����
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
        Form1.ListBox1.Items.Add('Exception�G' + E.Message);
      end;
    end;
  finally
    StrList.Free;
  end;
end;
 //Ū�}��������T
procedure Loadtxt();
var
  I,J,tNum1,tNum2:Integer;
  Strs:TStringList;
begin
  SetArray();
  //��}������}&�W��
  try
    Script.txt:=TStreamReader.Create(ExtractFilePath(Paramstr(0))+'/test.txt');
//    AssignFile(Script.txt,ExtractFilePath(Paramstr(0))+'/vsphere.txt'); //navicat,workstation,vsphere,ssms,plsql9,plsql14
//    Reset(Script.txt);
  except on E: Exception do
    begin
      Form1.ListBox1.Items.Add('Exception�G' + E.Message);
    end;
  end;
  //�}�l�B�z�}�����e
  try
    try
      I:=1;
      //���y����r��
      while not Script.txt.EndOfStream do begin
        //Ū���C���JAll
  //      readln(Script.txt,Script.All);
        Script.All:=Script.txt.ReadLine;
        //�T�{�榡
        tNum1:=Pos('[[',Script.All);
        tNum2:=Pos(']]',Script.All);
        if tNum1 <= 0 then begin
          Form1.ListBox1.Items.Add('�}��Key�榡���~');
          Exit;
        end;
        tNum1:=tNum1+2;
        if tNum2 <= 0 then begin
          Form1.ListBox1.Items.Add('�}��Key�榡���~');
          Exit;
        end;
        //�s��C��Value
        Script.Value[I]:=Copy(Script.All,tNum2+2,Length(Script.All));
        //�ťն�0
        if (Length(Script.All)-tNum2) = 1 then begin
          Script.Value[I]:='0';
        end;
        //�s��C��Key
        Script.Key:=Copy(Script.All,tNum1,tNum2-tNum1);
        //��Key�P�_�}��Value����l�G�ɮצ�}
        if (Script.Key = 'FilePath') then begin
          Script.FileLine := I;
          Strs:=TStringList.Create;
          Strs.Delimiter:='\';
          Strs.DelimitedText:=Script.Value[I];
          //���̫�\�k�䪺��
          for J := 0 to Strs.Count-1 do begin
            Script.FileName:=Strs.Strings[J];
          end;
          Strs.Free;
        end;
        //��Key�P�_�}��Value����l
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
        //��Key�P�_�}��Value����l�G���C�������
        if (Script.Key = 'MenuValue') and (Script.Value[I] <> '0') then begin
          Strs:=TStringList.Create;
          Strs.Delimiter:='-';
          Strs.DelimitedText:=Script.Value[I];
          SetLength(Script.Menu,Strs.Count+1);
          //�H-���j�x�s
          for J := 0 to Strs.Count-1 do begin
            Script.Menu[J]:=StrToInt(Strs.Strings[J]);
          end;
          Strs.Free;
        end;
        //��Key�P�_�}��Value����l�G�U�@�ӵ������D
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
        Form1.ListBox1.Items.Add('Exception�G' + E.Message);
      end;
    end;
  finally
    Script.txt.Free;
  end;
end;
//�N�}�����ᶶ�ǱƦC
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
