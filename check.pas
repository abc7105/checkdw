unit check;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  OleCtl, OleCtrls, ComCtrls, ComObj, OleServer, ShlObj, EXCEL2000,
  Dialogs, Checklist, StdCtrls, Udebug;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    function ExePath(): string;
    procedure testone();
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  DOThisDB: lxyMdb;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  aXlsToMdb: xlstomdb;
  EXCELAPP: Variant;
  zdname, zdlength: string;
  zd1, zd2: string;
begin

  //测试建库
  debugreset;
  try
    excelapp := createoleobject('excel.application');
  except
    Exit;
  end;

  mainpath := ExtractFilePath(Application.ExeName);
  excelapp.WorkBooks.Open(FileName := ExtractFilePath(Application.ExeName) +
    '单位往来交易模块.xlsx', UpdateLinks := 0);
  excelapp.visible := TRUE;

  aXlsToMdb := XlsToMdb.create(excelapp, mainpath + 'checkzw.mdb');

  zd1 := ' 单位名称,	对方单位编码 ,对方单位名称 ,	应收票据,	 应收帐款';
  zd2 := ' 单位名称,	对方单位编码 ,对方单位名称 ,	应收票据,	 应收帐款';

  try
    axlstomdb.XlsSheetdata_into_Mdbtable('内部往来', '内部往来', zd1, zd2);
  finally
    excelapp.WorkBooks.close;
    excelapp.quit;
    excelapp := Unassigned;
  end;
  DebugList;
  showmessage('字段到字段完成！');
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  aXlsToMdb: xlstomdb;
  EXCELAPP: Variant;
  zdname, zdlength: string;
  zd1, zd2: string;
begin

  //测试建库
  debugreset;
  try
    excelapp := createoleobject('excel.application');
  except
    Exit;
  end;
  mainpath := ExtractFilePath(Application.ExeName);
  excelapp.WorkBooks.Open(FileName := ExtractFilePath(Application.ExeName) +
    '单位往来交易模块.xlsx', UpdateLinks := 0);
  excelapp.visible := TRUE;

  aXlsToMdb := XlsToMdb.create(excelapp, mainpath + 'checkzw.mdb');
  axlstomdb.XlsSheet_into_Mdbtable('内部往来', '内部往来');
  excelapp.WorkBooks.close;
  excelapp.quit;
  excelapp := Unassigned;
  DebugList;
  showmessage('表到表完成');
end;

procedure TForm1.Button3Click(Sender: TObject);
var
  aXlsToMdb: xlstomdb;
  EXCELAPP: Variant;
  zdname, zdlength: string;
  zd1, zd2: string;
  // mainpath:string;
begin
  //测试建库
  debugreset;
  try
    excelapp := createoleobject('excel.application');
  except
    Exit;
  end;
  mainpath := ExtractFilePath(Application.ExeName);
  excelapp.WorkBooks.Open(FileName := ExtractFilePath(Application.ExeName) +
    '单位往来交易模块.xlsx', UpdateLinks := 0);
  excelapp.visible := TRUE;

  aXlsToMdb := XlsToMdb.create(excelapp, mainpath + 'checkzw.mdb');

  zdname := '姓名,年龄';
  zdlength := '30,20';
  aXlsToMdb.Createxlsxheet('lxy', zdname, zdlength);

  axlstomdb.XlssSheet_TOcreate_MdbTable('内部往来', '内部往来');

  excelapp.WorkBooks.close;
  excelapp.quit;
  excelapp := Unassigned;

  showmessage('建库完成！');

end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  mainpath := exepath;
  DOThisDB := lxyMdb.create(mainpath + 'CHECKZW.MDB');

end;

function TForm1.ExePath: string;
begin
  result := ExtractFilePath(Application.ExeName);
end;

procedure TForm1.testone;
begin
  // //
 //  DOThisDB.sheetname := 'testA';
 //  DOThisDB.fieldlength.Add('id');
  ;
end;

end.

