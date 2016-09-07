program checkdw;

uses
  Forms,
  check in 'check.pas' {Form1},
  Checklist in 'Checklist.pas',
  Udebug in '..\lxyshare\Udebug.pas',
  uselectfield in 'uselectfield.pas' {fmselectfield},
  communit in '..\ShareUnit\communit.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(Tfmselectfield, fmselectfield);
  Application.Run;
end.
