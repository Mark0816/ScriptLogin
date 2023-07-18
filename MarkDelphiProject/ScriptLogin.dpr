program ScriptLogin;

uses
  Vcl.Forms,
  Func in 'Func.pas',
  JobsApi in 'JobsApi.pas',
  Main in 'Main.pas' {Form1},
  UseScript in 'UseScript.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
