unit Cor_shtvec;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, Buttons, OleCtrls, ExtCtrls, StdCtrls, FileCtrl, Grids,
  DBGrids, DB, DBTables, TTF160_TLB, AxCtrls;

type
  TfmCorVecSheet = class(TForm)
    Panel1: TPanel;
    sbClose: TSpeedButton;
    sbSheet: TStatusBar;
    SaveDialogSprdSheet: TSaveDialog;
    gb3: TGroupBox;
    bbSaveSheet: TBitBtn;
    SprdSheet: TF1Book6;
    Table1: TTable;
    ds1: TDataSource;
    procedure sbCloseClick(Sender: TObject);
    procedure bbSaveSheetClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
    MaxIsData : Integer;
    FileSavePath : string;
    FileSaveName : string;
    FileOpenPathAndName : string;
    IsData : array[1..200] of boolean;
    procedure FillSheet;
  end;

var
  fmCorVecSheet: TfmCorVecSheet;

implementation

{$R *.DFM}

procedure TfmCorVecSheet.sbCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TfmCorVecSheet.bbSaveSheetClick(Sender: TObject);
const
  Excel5Type = 4;
  Excel97Type = 11;
  VisualComponentType = 5;
  FormulaOne6Type = 12;
var
  pFileType : smallint;
  pBuf      : string;
  pTitle    : string;
  tmpStr    : string[3];
begin
  SaveDialogSprdSheet.InitialDir := FileSavePath;
  SaveDialogSprdSheet.FileName := FileSaveName;
  if SaveDialogSprdSheet.Execute then
  begin
    FileSavePath := ExtractFilePath(SaveDialogSprdSheet.FileName);
    pFileType := Excel97Type;
    case SaveDialogSprdSheet.FilterIndex of
      1 : pFileType := Excel97Type;
      2 : pFileType := Excel5Type;
    end;
    pBuf := SaveDialogSprdSheet.FileName;
    SprdSheet.Write(pBuf,pFileType);
  end;
end;

procedure TfmCorVecSheet.FillSheet;
var
  i, j : integer;
begin
  ds1.DataSet.DisableControls;
  i := 1;
  ds1.DataSet.First;
  SprdSheet.Row := i;
  for j := 0 to ds1.DataSet.FieldCount - 1 do
  begin
    SprdSheet.Col := j+1;
    SprdSheet.Text := ds1.DataSet.Fields[j].FieldName;
  end;
  for i := 1 to ds1.DataSet.RecordCount do
  begin
    SprdSheet.Row := i+1;
    for j := 0 to ds1.DataSet.FieldCount - 1 do
    begin
      SprdSheet.Col := j+1;
      if ((j+1 <= MaxIsData) and (IsData[j+1] = true)) then
      begin
        SprdSheet.Number := ds1.DataSet.Fields[j].AsVariant;
      end else
      begin
        SprdSheet.Text := ds1.DataSet.Fields[j].AsString;
      end;
    end;
    ds1.DataSet.Next;
  end;
  SprdSheet.Row := 1;
  ds1.DataSet.First;
  ds1.DataSet.EnableControls;
  Table1.Close;
end;

procedure TfmCorVecSheet.FormCreate(Sender: TObject);
begin
  FileSavePath := '';
  FileOpenPathAndName := '';
  FileSaveName := '';
  MaxIsData := 200;
end;

procedure TfmCorVecSheet.FormShow(Sender: TObject);
begin
  gb3.enabled := true;
  bbSaveSheet.Enabled := true;
  Table1.TableName := FileOpenPathAndName;
  Table1.Open;
  FillSheet;
end;


procedure TfmCorVecSheet.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Table1.Close;
end;

end.
