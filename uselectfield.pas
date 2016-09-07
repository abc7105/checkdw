unit uselectfield;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ZcGridStyle, ZcCalcExpress, ZcFormulas, StdCtrls, ExtCtrls,
  ZcGridClasses, ZcUniClass, ZJGrid, ZcDataGrid;

type
  Tfmselectfield = class(TForm)
    EjunDataGrid1: TEjunDataGrid;
    EjunLicense1: TEjunLicense;
    Panel1: TPanel;
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);

  private
    FDBname: Integer;
    Ffieldlength: TStringList;
    Ffieldname: TStringList;
    Ffieldtype: TStringList;
  published
    property fieldlength: TStringList read Ffieldlength write Ffieldlength;
    property fieldname: TStringList read Ffieldname write Ffieldname;
    property fieldtype: TStringList read Ffieldtype write Ffieldtype;
    property DBname: Integer read FDBname write FDBname;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmselectfield: Tfmselectfield;

implementation

{$R *.dfm}

procedure Tfmselectfield.Button1Click(Sender: TObject);
var
  i: Integer;
begin
  for i := 1 to ffieldname.count do
  begin
    ffieldname[i - 1] := EjunDataGrid1.Cells[1, I].Text;
    ffieldtype[i - 1] := EjunDataGrid1.Cells[2, I].text;
    ffieldlength[i - 1] := EjunDataGrid1.Cells[3, I].Text;
  end;

  CLOSE;

end;

procedure Tfmselectfield.FormShow(Sender: TObject);
var
  i: Integer;
begin
  EjunDataGrid1.RowCount := ffieldname.count + 1;
  EjunDataGrid1.ColCount := 4;
  EjunDataGrid1.Cells[1, 1].Value := '字段名';
  EjunDataGrid1.Cells[2, 1].Value := '字段类型';
  EjunDataGrid1.Cells[3, 1].Value := '字段长度';

  for i := 1 to ffieldname.count do
  begin
    EjunDataGrid1.Cells[1, I].Value := ffieldname[i - 1];
    EjunDataGrid1.Cells[2, I].Value := ffieldtype[i - 1];
    EjunDataGrid1.Cells[3, I].Value := ffieldlength[i - 1];
  end;

end;

end.

