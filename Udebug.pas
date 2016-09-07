unit Udebug;

interface

uses
  Windows, Classes, SysUtils, Variants, adxAddIn, forms, FileCtrl, excel2000,
  StrUtils, ZcGridClasses, ZJGrid, ZcDataGrid, ZcDBGrids, ZcUniClass, Dialogs,
  Controls, IniFiles, ShellAPI, ExtCtrls, DB, ADODB;

const
  WilltoDebug = true;

var
  mainpath: string;

procedure NewTxt(FileName: string);

procedure OpenTxt(FileName: string);

procedure ReadTxt(FileName: string);

procedure AppendTxt(Str: string; FileName: string);

procedure debugto(str: string);

procedure DebugReset();

procedure DebugList;

implementation

procedure NewTxt(FileName: string);
var
  F: Textfile;
begin
  if not WilltoDebug then
    exit;
  if fileExists(FileName) then
    DeleteFile(FileName); {���ļ��Ƿ����,�ھ̈́h��}
  AssignFile(F, FileName); {���ļ�������� F ����}
  ReWrite(F); {����һ���µ��ļ�������Ϊ ek.txt}
  Writeln(F, 'test:');
  Closefile(F); {�ر��ļ� F}
end;

procedure OpenTxt(FileName: string);
var
  F: Textfile;
begin
  if not WilltoDebug then
    exit;
  AssignFile(F, FileName); {���ļ�������� F ����}
  Append(F); {�Ա༭��ʽ���ļ� F }
  Writeln(F, '����Ҫд����ı�д�뵽һ�� .txt �ļ�');
  Closefile(F); {�ر��ļ� F}
end;

procedure ReadTxt(FileName: string);
var
  F: Textfile;
  str: string;
begin
  if not WilltoDebug then
    exit;
  AssignFile(F, FileName); {���ļ�������� F ����}
  Reset(F); {�򿪲���ȡ�ļ� F }
  Readln(F, str);
  ShowMessage('�ļ���:' + str + '�С�');
  Closefile(F); {�ر��ļ� F}
end;

procedure AppendTxt(Str: string; FileName: string);
var
  F: Textfile;
begin
  if not WilltoDebug then
    exit;
  AssignFile(F, FileName);
  Append(F);
  Writeln(F, Str);
  Closefile(F);
end;

procedure DebugReset();
begin
  if not WilltoDebug then
    exit;
  NewTxt(mainpath + 'test.txt');
end;

procedure debugto(str: string);
var
  afile: string;
begin
  if not WilltoDebug then
    exit;
  afile := mainpath + 'test.txt';

  if not FileExists(afile) then
    NewTxt(afile);

  AppendTxt(Str, afile);

end;

procedure DebugList;
begin
  if not WilltoDebug then
    exit;
  ShellExecute(0, 'open', PChar(mainpath + 'TEST.txt'), '', nil, 1);
end;

end.

