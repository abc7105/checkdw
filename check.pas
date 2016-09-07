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
  xx, yy, zz: tstringlist;
begin
  xx := tstringlist.create();
  yy := tstringlist.create();
  zz := tstringlist.create();

  //  DOThisDB.sheetname := 'lxytest';
  //
  //  xx.Add('id');
  //  xx.Add('lxyname');
  //  xx.Add('address');
  //  xx.Add('MONEYdd');
  //  xx.Add('是否重点');
  //  dothisdb.fieldname := xx;
  //
  //  yy.clear;
  //  yy.Add('N');
  //  yy.Add('C');
  //  yy.Add('C');
  //  yy.Add('F');
  //  yy.Add('B');
  //  dothisdb.fieldtype := yy;
  //
  //  zz.clear;
  //  zz.Add('0');
  //  zz.Add('20');
  //  zz.Add('50');
  //  zz.Add('0');
  //  zz.Add('0');
  //  dothisdb.fieldlength := zz;

   // DOThisDB.createTable;
  DebugList;

end;

procedure TForm1.Button2Click(Sender: TObject);
//var
//  EXCELAPP: Variant;
//  aLXYdoexcel: LXYdoexcel;
begin
  // DebugReset;
 //  try
 //    excelapp := createoleobject('excel.application');
 //    //   showbar(0, ' 建立EXCEL进程ok！  ');
 //  except
 //    //   showbar(0, ' ERROR:建立EXCEL进程失败！  ');
 //    Exit;
 //  end;
 //  excelapp.WorkBooks.Open(FileName := ExtractFilePath(Application.ExeName) +
 //    '单位往来交易模块.xlsx', UpdateLinks := 0);
 //  //  EXCElAPP.WORKBOOKS.OPEN(ExtractFilePath(Application.ExeName) +
 //  //    '单位往来交易模块.xlsx');
 //  aLXYdoexcel := LXYdoexcel.create(EXCELAPP, mainpath + 'checkzw.MDB');
 //  aLXYdoexcel.aworkbook := EXCElAPP.activeworkbook;
 //  aLXYdoexcel.sheetname := '内部单位';
 //  aLXYdoexcel.CreateMdbTable();
 //  DEBUGLIST;
 //
 //  aLXYdoexcel.aworkbook := EXCElAPP.activeworkbook;
 //  aLXYdoexcel.sheetname := '内部往来';
 //  aLXYdoexcel.CreateMdbTable();
 //
 //  excelapp.activeworkbook.close(false);
 //  ExcelApp.quit;
 //  excelapp := unassigned;
 //  ShowMessage('OK');

end;

procedure TForm1.Button3Click(Sender: TObject);
var
  aXlsToMdb: xlstomdb;
  EXCELAPP: Variant;
  zdname, zdlength: string;
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

  axlstomdb.XlsSheet_into_Mdbtable('内部往来', '内部往来');
  DebugList;
  showmessage('建库成功');

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

