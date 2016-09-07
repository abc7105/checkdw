unit Checklist;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  communit, DB,
  ComObj, Dialogs, ADODB, Udebug;

type
  LxyMdb = class
  private
    fMdbName: string;
    sheetfiledstring: string;
    Fmdbfieldlength: TStringList;
    Fmdbfieldname: TStringList;
    Fmdbfieldtype: TStringList;
    con1: tadoconnection;
    FmdbTableName: string;
    sheetfieldlist: string;
    sqlstr: Integer;
  private
    function mergefieldINFO: Boolean;
    function ConvertFieldtypeString(fieldtype, fieldlen: string): string;
  public
    constructor create(MdbName: string);
    procedure openDB(MDBfilename: string);
    function createTable(tablename: string; mdbfieldname, mdbfieldtype,
      mdbfieldlength: TStringList): Boolean;
    procedure createMDB(MDBfilename: string);
    property con: tadoconnection read con1 write con1;
    //  property mdbTableName: string read FmdbTableName write FmdbTableName;
  end;

  XlsToMdb = class(tobject)
  private
    alxydb: LxyMdb;
    FMdbName: string;
    atable: TADOTable;

    Fsheetname: string;
    EXCELAPP: VARIANT;
    Faworkbook: Variant;
    curSheet: Variant;

    ffieldlength: tstringlist;
    fMDBfieldname: tstringlist;
    ffieldname: tstringlist;
    ffieldtype: tstringlist;
    ffieldpos: TStringList;
    fkeyfield: string;

    Ffmanyfieldstr: string;
    Ffmanyfieldlength: string;

    procedure Setsheetname(const Value: string);

  private
    procedure format_sheet_column(asheet: Variant; COLUMNID: Integer;
      COLUMNTYPE:
      string);
    function filldate(asheet: Variant; ncolumn: integer): boolean;
    procedure fillzero(asheet: Variant; ncolumn: integer);
    procedure xlsfieldinfotolist;
    procedure listclear;
    function INmdbtable(zdname: string; zdlists: TStringList): Integer;
    function sheetexists(excelapp: Variant; aname: string): boolean;
    function zdlens: tstringlist;
    function zdtypes: tstringlist;
    function zdtypestring(zdtype: TFieldType): string;
    procedure sheetdata_tomdb;
  published
    property MdbName: string read FMdbName write FMdbName;
    property aworkbook: Variant read Faworkbook write Faworkbook;
    property sheetname: string read Fsheetname write Setsheetname;
  public

    //建立 XLS模板表
    procedure Createxlsxheet(sheetname, manyfieldnamestr, manyfieldlengthstr:
      string);

    //SHEET数据字段建立 MDB数据库表
    procedure XlssSheet_TOcreate_MdbTable(sheetname, tablename: string);
    procedure XlsSheetfieldstr_Tocreate_Mdbtable(tablename, manyfieldnamestr,
      manyfieldlengthstr, manyfieldtypestr: string);

    //SHEET数据导入到MDB数据库中去
    procedure XlsSheetdata_into_Mdbtable(XlsSheetname, MdbTablename,
      manyfieldnamestr_inxls, manyfieldnamestr_inmdb: string);
    procedure XlsSheet_into_Mdbtable(XlsSheetname, MdbTablename: string);
    function zdnames: tstringlist;

    constructor create(AEXCELAPP: VARIANT; MdbName: string);
    destructor destroy;
  end;

implementation

uses
  uselectfield;

{ LxyMdb }

constructor LxyMdb.create(MdbName: string);
begin
  con1 := TADOConnection.Create(nil);
  con1.LoginPrompt := false;
  fMdbName := MdbName;
  openDB(fMdbName);
end;

procedure LxyMdb.createMDB(MDBfilename: string);
var
  Dbnew: OleVariant;
begin

  if FileExists(fMdbName) then
  begin
    if MessageBox(Application.Handle, PChar('数据库 ' + fMdbName + ' 已存在！' +
      #13#10 + '是否覆盖？'), '警告', MB_YESNO + MB_ICONWARNING) = idNo then
      exit;
    if not DeleteFile(fMdbName) then
    begin
      MessageBox(Application.Handle, PChar('不能删除数据库：' + fMdbName),
        '错误', MB_OK + MB_ICONERROR);
      exit;
    end;
  end;
  try
    dbnew := CreateOleObject('ADOX.Catalog');
    fMdbName := MDBfilename;
    dbnew.Create('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + fMdbName);
  finally
    DBNEW := Null;
  end;
end;

function LxyMdb.createTable(tablename: string; mdbfieldname, mdbfieldtype,
  mdbfieldlength: TStringList): Boolean;
var
  qrytmp: TADOQuery;
  sqlstr: string;
begin
  result := false;
  openDB(Fmdbname);
  FmdbTableName := tablename;

  Fmdbfieldname := mdbfieldname;
  Fmdbfieldlength := mdbfieldlength;
  Fmdbfieldtype := mdbfieldtype;

  if mergefieldINFO = False then
    exit;

  try
    qrytmp := TADOQuery.Create(nil);
    qrytmp.Connection := con1;
    qrytmp.Close;
    qrytmp.SQL.Clear;
    sqlstr := 'drop table ' + FmdbTableName;
    DEBUGTO('LxyMdb.createTable:' + sqlstr);
    qrytmp.SQL.Add(sqlstr);
    qrytmp.ExecSQL;
    result := True;
    DEBUGTO('LxyMdb.createTable: Boolean  删除 :' + sqlstr);
  except
  end;

  try

    qrytmp := TADOQuery.Create(nil);
    qrytmp.Connection := con1;
    qrytmp.Close;
    qrytmp.SQL.Clear;
    sqlstr := 'create table ' + FmdbTableName + '(' + sheetfiledstring + ')';
    DEBUGTO(sqlstr);
    qrytmp.SQL.Add(sqlstr);
    qrytmp.ExecSQL;
    result := True;
    DEBUGTO('LxyMdb.createTable: Boolean   :' + sqlstr);

  except
  end;
end;

function LxyMdb.ConvertFieldtypeString(fieldtype, fieldlen: string): string;
begin
  result := '';
  if lowercase(fieldtype) = 'd' then
  begin
    result := ' datetime ';
  end
  else if lowercase(fieldtype) = 'b' then
  begin
    result := ' bit ';
  end
  else if lowercase(fieldtype) = 'c' then
  begin
    result := ' char(' + trim(fieldlen) + ')';
  end
  else if lowercase(fieldtype) = 'f' then
  begin
    result := ' double ';
  end
  else if lowercase(fieldtype) = 'n' then
  begin
    result := ' integer ';
  end;
end;

function LxyMdb.mergefieldINFO: Boolean;
var
  i: Integer;
begin
  //
  result := false;
  sheetfiledstring := '';
  if Fmdbfieldlength.Count <> Fmdbfieldname.Count then
    Exit;

  if Fmdbfieldlength.Count <> Fmdbfieldtype.Count then
    Exit;

  //  DEBUGTO(INTTOSTR(Fmdbfieldname.count));
  for i := 0 to Fmdbfieldname.count - 1 do
  begin
    if (Trim(Fmdbfieldname[i]) = '') or (Trim(Fmdbfieldtype[i]) = '') then
      Break;

    if i > 0 then
      sheetfiledstring := sheetfiledstring + ',';

    sheetfiledstring := sheetfiledstring + Fmdbfieldname[i] + ' ' +
      ConvertFieldtypeString(Fmdbfieldtype[i], Fmdbfieldlength[i]);

  end;
  DEBUGTO('LxyMdb.mergefieldINFO  ' + sheetfiledstring);
  result := True;
end;

procedure LxyMdb.openDB(MDBfilename: string);
var
  ausername, apassword: string;
begin
  ausername := 'admin';
  apassword := '';
  fMdbName := MDBfilename;

  if not FileExists(MDBfilename) then
    createMDB(MDBfilename);

  if con1.Connected then
    con1.Connected := false;

  con1.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;' + 'User ID=' +
    AUserName + ';' + 'Jet OLEDB:Database Password=' + APassword + ';' +
    'Data Source=' + MDBfilename + ';' + 'Mode=ReadWrite;' +
    'Extended Properties="";';

  con1.Connected := true;
end;

constructor XlsToMdb.create(AEXCELAPP: VARIANT; MdbName: string);
begin
  ffieldname := TStringList.Create;
  ffieldTYPE := TStringList.Create;
  ffieldLENGTH := TStringList.Create;
  ffieldpos := tstringlist.create;

  debugto('XlsToMdb.create  =BEGIN OK');
  EXCELAPP := AEXCELAPP;
  FMdbName := MdbName;
  alxydb := LxyMdb.create(fMdbName);
  Faworkbook := EXCELAPP.activeworkbook;
  debugto('XlsToMdb.create  =END OK');

  atable := TADOTable.Create(nil);
  atable.Connection := alxydb.con;
end;

procedure XlsToMdb.Createxlsxheet(sheetname, manyfieldnamestr,
  manyfieldlengthstr: string);
var
  I: INTEGER;
  cols: integer;
  fdnamelist, fdlengthlist: TStringList;
begin
  Faworkbook := EXCELAPP.activeworkbook;
  // ShowMessage(Faworkbook.NAME);
  try
    if not sheetexists(EXCELAPP, SHEETNAME) then
    begin
      curSheet := Faworkbook.SHEETS.ADD;
      curSheet.NAME := sheetname;
    end
    else
      curSheet := Faworkbook.WORKSHEETS.ITEM[SHEETNAME];
    //   ShowMessage(curSheet.NAME);
  except
    ShowMessage('ERROR: EXCEL文件未打开！');
    exit;
  end;

  fdnamelist := TStringList.Create;
  fdlengthlist := TStringList.Create;

  fdnamelist := splitstring(manyfieldnamestr, ',');
  fdlengthlist := splitstring(manyfieldlengthstr, ',');

  if fdnamelist.Count < 1 then
  begin
    ShowMessage('ERROR: 无相应的字段，请检查！');
    exit;
  end;

  if fdnamelist.Count <> fdlengthlist.Count then
  begin
    ShowMessage('ERROR: 字段数与字段宽度的数量不一致，请检查');
    exit;
  end;

  cols := fdnamelist.Count;
  for i := 0 to cols - 1 do
  begin
    curSheet.cells[1, i + 1].value := fdnamelist[i];
    curSheet.Columns[i + 1].ColumnWidth := strtoint(fdlengthlist[i]);
  end;

end;

destructor XlsToMdb.destroy;
begin
  //
  EXCELAPP := null;
  CURsheet := null;
  faworkbook := null;

  try
    alxydb.Free;
    alxydb := nil;
  except
  end;
end;

procedure XlsToMdb.listclear;
begin
  ffieldname.Clear;
  ffieldlength.Clear;
  ffieldtype.Clear;
  ffieldpos.Clear;
  // Fsheetname := '';
end;

procedure XlsToMdb.Setsheetname(const Value: string);
begin
  debugto(' XlsToMdb.Setsheetname  =BEGIN OK');

  Fsheetname := Value;
  CURsheet := NULL;
  try
    if not VarIsNull(faworkbook) then
    begin
      CURsheet := faworkbook.SHEETS.ITEM[Value];
      CURsheet.activate;
    end;
  except
  end;
  debugto(' XlsToMdb.Setsheetname  =END OK');
end;

function XlsToMdb.sheetexists(excelapp: Variant; aname: string): boolean;
var
  xstr: string;
  i: Integer;
begin
  //  表名是否存在
  result := False;
  try
    for i := 1 to excelapp.ActiveWorkbook.Sheets.count do
    begin
      if Trim(excelapp.ActiveWorkbook.Sheets.Item[i].name) =
        Trim(aname) then
      begin
        result := true;
        exit;
      end;
    end;

  except
  end;
end;

procedure XlsToMdb.XlsFieldInfotoLIST;
var
  i: integer;
  AFORM: Tfmselectfield;
begin
  debugto(' XlsToMdb.XlsFieldInfotoLIST  =BEGIN OK');

  listclear;

  CURsheet.activate;
  for i := 1 to CURsheet.USEDRANGE.columns.count do
  begin
    ffieldname.Add(CURsheet.CELLS.ITEM[1, i].text);
    ffieldTYPE.Add('C');
    ffieldLENGTH.Add('20');
  end;

  try
    AFORM := Tfmselectfield.Create(nil);
    AFORM.fieldlength := ffieldlength;
    AFORM.fieldname := ffieldname;
    AFORM.fieldtype := fFIELDTYPE;
    AFORM.ShowMODAL;

  finally
    ffieldlength := AFORM.fieldlength;
    ffieldname := AFORM.fieldname;
    fFIELDTYPE := AFORM.fieldtype;
    AFORM.Close;
    AFORM.Free;
    AFORM := nil;
  end;
  debugto(' XlsToMdb.XlsFieldInfotoLIST  =END OK');
end;

procedure XlsToMdb.XlsSheetfieldstr_Tocreate_Mdbtable(tablename,
  manyfieldnamestr, manyfieldlengthstr, manyfieldtypestr: string);
begin
  try
    if not sheetexists(EXCELAPP, SHEETNAME) then
    begin
      curSheet := Faworkbook.SHEETS.ADD;
      curSheet.NAME := sheetname;
    end
    else
      curSheet := Faworkbook.WORKSHEETS(SHEETNAME);

    curSheet.activate;
  except
    ShowMessage('表文件不存在或建表失败！');
    exit;
  end;

  listclear;
  ffieldname := splitstring(manyfieldnamestr, ',');
  ffieldtype := splitstring(manyfieldtypestr, ',');
  ffieldlength := splitstring(manyfieldlengthstr, ',');

  if (ffieldname.Count <> ffieldtype.Count)
    or (ffieldname.Count <> ffieldlength.Count) then
  begin
    ShowMessage('字段名，字段类型，字段长度的数量不匹配，请检查！');
    EXIT;
  end;

  alxydb.createTable(tablename, ffieldname, ffieldTYPE, ffieldlength);
end;

procedure XlsToMdb.XlssSheet_TOcreate_MdbTable(sheetname, tablename: string);
begin
  debugto('XlsToMdb.XlsSheet_TO_MdbTable  =BEGIN OK');

  try
    if not sheetexists(EXCELAPP, SHEETNAME) then
    begin
      curSheet := Faworkbook.SHEETS.ADD;
      curSheet.NAME := sheetname;
    end
    else
      curSheet := Faworkbook.WORKSHEETS.ITEM[SHEETNAME];

    curSheet.activate;
  except
    ShowMessage('表文件不存在或建表失败！');
    exit;
  end;

  XlsFieldInfotoLIST;

  alxydb.createTable(tablename, ffieldname, ffieldTYPE, ffieldlength);
  debugto('XlsToMdb.XlsSheet_TO_MdbTable  =END OK');

  //
end;

procedure XlsToMdb.XlsSheetdata_into_Mdbtable(XlsSheetname, MdbTablename,
  manyfieldnamestr_inxls, manyfieldnamestr_inmdb: string);
var
  I, iindex: INTEGER;
  cols: integer;
  MDBzdNAMEs: TStringList;
  MDBzdTYPEs: TStringList;
  MDBzddlens: TStringList;
  xlszdnames: TStringList;
  xlszdnames_para: TStringList;
  strx: string;

begin
  //要加入改变
  xlszdnames := TStringList.Create;
  Faworkbook := excelapp.activeworkbook;
  try
    if not sheetexists(EXCELAPP, XlsSheetname) then
    begin
      curSheet := Faworkbook.SHEETS.ADD;
      curSheet.NAME := XlsSheetname;
    end
    else
      curSheet := Faworkbook.WORKSHEETS.item[XlsSheetname];
  except
    ShowMessage('ERROR: EXCEL文件未打开！');
    exit;
  end;

  listclear;

  try
    atable.TableName := MdbTablename;
    atable.Open;
  except
    ShowMessage('数据库中无对应的表！请检查。');
    exit;
  end;

  cols := curSheet.usedrange.columns.count;
  for i := 1 to cols do
    xlszdnames.Add(Trim(curSheet.cells[1, i].text));

  MDBzdNAMEs := zdnames();
  MDBzdTYPEs := zdtypes();
  MDBzddlens := zdlens();

  strx := manyfieldnamestr_inmdb;
  ffieldname := splitstring(strx, ',');
  debugstringlist(ffieldname);
  strx := manyfieldnamestr_inxls;
  xlszdnames_para := splitstring(manyfieldnamestr_inxls, ',');
  debugstringlist(xlszdnames_para);

  for i := 0 to ffieldname.count - 1 do
  begin
    iindex := INmdbtable(ffieldname[i], MDBzdNAMEs);
    if iindex = -1 then
    begin
      showmessage('有字段不匹配数据库，请检查 ！' + ffieldname[i] + '==');
      exit;
    end;

    ffieldtype.Add(MDBzdTYPEs[iindex]);
    format_sheet_column(curSheet, i, MDBzdTYPEs[iindex]);
    ffieldlength.Add(MDBzddlens[iindex]);
    ffieldpos.Add(IntToStr(INmdbtable(xlszdnames_para[i], xlszdnames) + 1));
  end;
  sheetdata_tomdb;

  xlszdnames.Free;
  xlszdnames := nil;
end;

procedure XlsToMdb.XlsSheet_into_Mdbtable(XlsSheetname,
  MdbTablename: string);
var
  I, iindex: INTEGER;
  cols: integer;
  MDBzdNAMEs: TStringList;
  MDBzdTYPEs: TStringList;
  MDBzddlens: TStringList;

begin
  Faworkbook := excelapp.activeworkbook;
  try
    if not sheetexists(EXCELAPP, XlsSheetname) then
    begin
      curSheet := Faworkbook.SHEETS.ADD;
      curSheet.NAME := XlsSheetname;
    end
    else
      curSheet := Faworkbook.WORKSHEETS.item[XlsSheetname];
  except
    ShowMessage('ERROR: EXCEL文件未打开！');
    exit;
  end;

  listclear;

  try
    atable.TableName := MdbTablename;
    atable.Open;
  except
    ShowMessage('数据库中无对应的表！请检查。');
    exit;
  end;

  cols := curSheet.usedrange.columns.count;
  MDBzdNAMEs := zdnames();
  MDBzdTYPEs := zdtypes();
  MDBzddlens := zdlens();

  for i := 1 to cols do
  begin
    iindex := INmdbtable(trim(curSheet.cells[1, i].Text), MDBzdNAMEs);
    if iindex <> -1 then
    begin
      ffieldname.Add(trim(MDBzdNAMEs[iindex]));
      ffieldtype.Add(MDBzdTYPEs[iindex]);
      format_sheet_column(curSheet, i, MDBzdTYPEs[iindex]);
      ffieldlength.Add(MDBzddlens[iindex]);
      ffieldpos.Add(IntToStr(i));
    end;
  end;
  sheetdata_tomdb;
end;

procedure XlsToMdb.sheetdata_tomdb;
var
  rows, i, j: Integer;
begin
  rows := curSheet.usedrange.rows.count;

  for i := 2 to rows do
  begin
    atable.append;

    for j := 1 to ffieldname.count do
    begin
      try
        if LowerCase(ffieldtype[j - 1]) = 'd' then
          atable.fieldbyname(ffieldname[j - 1]).asdatetime := curSheet.cells[i,
            strtoint(ffieldpos[j - 1])].value
        else if LowerCase(ffieldtype[j - 1]) = 'n' then
          atable.fieldbyname(ffieldname[j - 1]).ASINTEGER := curSheet.cells[i,
            strtoint(ffieldpos[j - 1])].VALUE
        else if LowerCase(ffieldtype[j - 1]) = 'f' then
          atable.fieldbyname(ffieldname[j - 1]).asfloat := curSheet.cells[i,
            strtoint(ffieldpos[j - 1])].VALUE
        else if LowerCase(ffieldtype[j - 1]) = 'c' then
          atable.fieldbyname(ffieldname[j - 1]).asstring := curSheet.cells[i,
            strtoint(ffieldpos[j - 1])].text;
      except
      end;
    end;
    atable.post;
  end;

end;

function XlsToMdb.zdnames: tstringlist;
var
  tbnames: tstringlist;
  i: integer;
begin
  //
  debugto('');
  debugto('====zdnames');
  tbnames := TStringList.Create;
  for i := 0 to atable.FieldCount - 1 do
  begin
    tbnames.Add(atable.Fields[i].FieldName);
    debugto(atable.Fields[i].FieldName);
  end;

  result := tbnames;
end;

function XlsToMdb.zdtypes: tstringlist;
var
  tbnames: tstringlist;
  i: integer;
begin
  //
  debugto('');
  debugto('====zdtypes');
  tbnames := TStringList.Create;
  for i := 0 to atable.FieldCount - 1 do
  begin
    tbnames.Add(zdtypestring(atable.Fields[i].DataType));
    debugto(zdtypestring(atable.Fields[i].DataType));
  end;

  result := tbnames;
end;

function XlsToMdb.zdlens: tstringlist;
var
  tbnames: tstringlist;
  i: integer;
begin
  //
  debugto('');
  debugto('====zdlens');
  tbnames := TStringList.Create;
  for i := 0 to atable.FieldCount - 1 do
  begin
    tbnames.Add(IntToStr(atable.Fields[i].DataSize));
    debugto(IntToStr(atable.Fields[i].DataSize));
  end;

  result := tbnames;

end;

function XlsToMdb.INmdbtable(zdname: string; zdlists: TStringList): Integer;
var
  i: integer;
begin
  //
  result := -1;
  for i := 0 to zdlists.Count - 1 do
  begin
    if Trim(zdlists[i]) = Trim(zdname) then
    begin
      result := i;
      Exit;
    end;
  end;

end;

function XlsToMdb.zdtypestring(zdtype: TFieldType): string;
begin
  //
  result := '-';
  case zdtype of
    ftSmallint, ftInteger, ftWord, ftLargeint:
      result := 'n';
    ftFloat, ftCurrency:
      result := 'f';
    ftString, ftWideString:
      result := 'c';
    ftBoolean:
      result := 'l';
    ftDate, ftTime, ftDateTime:
      RESULT := 'd';
  end

  { TFieldType = (ftUnknown, ftString,
   ftBoolean,, ftDate, ftTime, ftDateTime,
   ftBytes, ftVarBytes, ftAutoInc, ftBlob, ftMemo, ftGraphic, ftFmtMemo,
   ftParadoxOle, ftDBaseOle, ftTypedBinary, ftCursor, ftFixedChar,

   ftADT, ftArray, ftReference, ftDataSet, ftOraBlob, ftOraClob,
   ftVariant, ftInterface, ftIDispatch, ftGuid, ftTimeStamp, ftFMTBcd);   }
end;

procedure XlsToMdb.fillzero(asheet: Variant; ncolumn: integer);
var
  i, ncount: integer;
  acol: Variant;
begin
  //
  ncount := asheet.usedrange.rows.count;

  acol := asheet.Range[asheet.cells.Item[2, ncolumn],
    asheet.cells.Item[ncount, ncolumn]].Value;

  for i := 1 to ncount - 1 do
  begin
    try
      if VarIsNull(acol[i, 1]) then
        acol[i, 1] := 0
      else if VarIsEmpty(acol[i, 1]) then
        acol[i, 1] := 0
      else if VarIsStr(acol[i, 1]) then
        if (acol[i, 1] = '-') or (Trim(acol[i, 1]) = '') then
          acol[i, 1] := 0;

    except
    end;
  end;
  asheet.Range[asheet.cells.Item[2, ncolumn],
    asheet.cells.Item[ncount, ncolumn]].Value := acol;

end;

function XlsToMdb.filldate(asheet: Variant; ncolumn: integer): boolean;
var
  kk, ncount: integer;
  acol: Variant;
begin
  //
  result := False;
  ncount := asheet.usedrange.rows.count;

  acol := asheet.Range[asheet.cells.Item[2, ncolumn],
    asheet.cells.Item[ncount, ncolumn]].value;

  kk := 1;
  while kk <= ncount - 1 do
  begin

    if VarIsNumeric(acol[kk, 1]) then
      acol[kk, 1] := str8todate(IntToStr(acol[kk, 1]))
    else if VarIsStr(acol[kk, 1]) then
      acol[kk, 1] := str8todate(acol[kk, 1]);
    kk := kk + 1;
  end;
  asheet.Range[asheet.cells.Item[2, ncolumn],
    asheet.cells.Item[ncount, ncolumn]].Value := acol;
  result := True;
end;

procedure XlsToMdb.format_sheet_column(asheet: Variant; COLUMNID: Integer;
  COLUMNTYPE: string);
begin
  //
  if LowerCase(COLUMNTYPE) = 'n' then
    fillzero(asheet, COLUMNID)
  else if LowerCase(COLUMNTYPE) = 'f' then
    fillzero(asheet, COLUMNID)
  else if LowerCase(COLUMNTYPE) = 'D' then
    filldate(asheet, COLUMNID);

end;

end.

