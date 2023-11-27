unit Cor_ShtImEigVal;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, Buttons, OleCtrls, ExtCtrls, StdCtrls, Mask, IniFiles,
  Grids, DBGrids, DBCtrls, AxCtrls,
  Vcl.Tabs,
  VCL.FlexCel.Core, FlexCel.XlsAdapter, FlexCel.Render, FlexCel.Preview,
  Data.DB;

type
  TfmSheetImportEigVal = class(TForm)
    pControl: TPanel;
    sbSheet: TStatusBar;
    bbOpenSheet: TBitBtn;
    OpenDialogSprdSheet: TOpenDialog;
    gbDefineFields: TGroupBox;
    bbCancel: TBitBtn;
    gbDefineRows: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    eFromRow: TEdit;
    eToRow: TEdit;
    bbImport: TBitBtn;
    Label5: TLabel;
    eImportSpecNameCol: TEdit;
    Label10: TLabel;
    sbFindLastRow: TSpeedButton;
    Panel3: TPanel;
    gbDefineTabSheet: TGroupBox;
    cbSheetName: TComboBox;
    Label21: TLabel;
    Label22: TLabel;
    Splitter1: TSplitter;
    pDefinitions: TPanel;
    Splitter2: TSplitter;
    pData: TPanel;
    SheetData: TStringGrid;
    Tabs: TTabSet;
    procedure bbOpenSheetClick(Sender: TObject);
    procedure bbImportClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bbCancelClick(Sender: TObject);
    procedure sbFindLastRowClick(Sender: TObject);
    procedure cbSheetNameChange(Sender: TObject);
  private
    { Private declarations }
    Xls : TXlsFile;
    procedure FillTabs;
    procedure ClearGrid;
    procedure FillGrid(const Formatted: boolean);
    function GetStringFromCell(iRow,iCol : integer) : string;
  public
    { Public declarations }
  end;

var
  fmSheetImportEigVal: TfmSheetImportEigVal;

implementation

uses Allsorts, Cor_varb, Cor_dm_acs;

{$R *.DFM}

var
  iRec, iRecCount      : integer;

function TfmSheetImportEigVal.GetStringFromCell(iRow,iCol : integer) : string;
begin
  Result := Xls.GetStringFromCell(iRow,iCol);
end;

procedure TfmSheetImportEigVal.FillTabs;
var
  s: integer;
begin
  Tabs.Tabs.Clear;
  cbSheetname.Items.Clear;
  for s := 1 to Xls.SheetCount do
  begin
    Tabs.Tabs.Add(Xls.GetSheetName(s));
    cbSheetname.Items.Add(Xls.GetSheetName(s));
  end;
end;

procedure TfmSheetImportEigVal.ClearGrid;
var
  r: integer;
begin
  for r := 1 to SheetData.RowCount do SheetData.Rows[r].Clear;
end;

procedure TfmSheetImportEigVal.FillGrid(const Formatted: boolean);
var
  r, c, cIndex: Integer;
  v: TCellValue;
begin
  if Xls = nil then exit;
  if (Tabs.TabIndex + 1 <= Xls.SheetCount) and (Tabs.TabIndex >= 0) then Xls.ActiveSheet := Tabs.TabIndex + 1 else Xls.ActiveSheet := 1;
  //Clear data in previous grid
  ClearGrid;
  SheetData.RowCount := 1;
  SheetData.ColCount := 1;
  //FmtBox.Text := '';
  SheetData.RowCount := Xls.RowCount + 1; //Include fixed row
  SheetData.ColCount := Xls.ColCount + 1; //Include fixed col. NOTE THAT COLCOUNT IS SLOW. We use it here because we really need it. See the Performance.pdf doc.
  if (SheetData.ColCount > 1) then SheetData.FixedCols := 1; //it is deleted when we set the width to 1.
  if (SheetData.RowCount > 1) then SheetData.FixedRows := 1;

  for r := 1 to Xls.RowCount do
  begin
    //Instead of looping in all the columns, we will just loop in the ones that have data. This is much faster.
    for cIndex := 1 to Xls.ColCountInRow(r) do
    begin
      c := Xls.ColFromIndex(r, cIndex); //The real column.
      if Formatted then
      begin
        SheetData.Cells[c, r] := Xls.GetStringFromCell(r, c);
      end
      else
      begin
        v := Xls.GetCellValue(r, c);
        SheetData.Cells[c, r] := v.ToString;
      end;
    end;
  end;
  //Fill the row headers
  for r := 1 to SheetData.RowCount - 1 do
  begin
    SheetData.Cells[0, r] := IntToStr(r);
    SheetData.RowHeights[r] := Round(Xls.GetRowHeight(r) / TExcelMetrics.RowMultDisplay(Xls));
  end;
  //Fill the column headers
  for c := 1 to SheetData.ColCount - 1 do
  begin
    SheetData.Cells[c, 0] := TCellAddress.EncodeColumn(c);
    SheetData.ColWidths[c] := Round(Xls.GetColWidth(c) / TExcelMetrics.ColMult(Xls));
  end;
  //SelectedCell(1,1);
end;

procedure TfmSheetImportEigVal.bbOpenSheetClick(Sender: TObject);
var
  i : integer;
begin
  //TabControl.Tabs.Clear;
  cbSheetname.Items.Clear;
  OpenDialogSprdSheet.InitialDir := DataPath;
  if OpenDialogSprdSheet.Execute then
  begin
    DataPath := ExtractFilePath(OpenDialogSprdSheet.FileName);
    //Open the Excel file.
    if Xls = nil then Xls := TXlsFile.Create(false);
    xls.Open(OpenDialogSprdSheet.FileName);
    FillTabs;
    Tabs.TabIndex := Xls.ActiveSheet - 1;
    //FlexCelImport1.OpenFile(OpenDialogSprdSheet.FileName);
    //for i := 1 to FlexCelImport1.SheetCount do
    //begin
    //  FlexCelImport1.ActiveSheet:=i;
    //  TabControl.Tabs.Add(FlexCelImport1.ActiveSheetName);
    //  cbSheetname.Items.Add(FlexCelImport1.ActiveSheetName);
    //end;
    //FlexCelImport1.ActiveSheet:=1;
    //TabControl.TabIndex:=FlexCelImport1.ActiveSheet-1;
    //cbSheetName.ItemIndex := FlexCelImport1.ActiveSheet-1;
    //Data.LoadSheet;
    //Data.Zoom := 70;
    FillGrid(true);
    pDefinitions.Visible := true;
    Splitter1.Visible := true;
    //TabControl.Visible := true;

    //cbSheetName.Items.Clear;
    sbFindLastRowClick(Sender);
  end;
end;


procedure TfmSheetImportEigVal.bbImportClick(Sender: TObject);
var
  j      : integer;
  iCode  : integer;
  i      : integer;
  ii, jj : integer;
  iii, jjj : integer;
  tmpStr : string;
begin
  ImportSheetNumber := cbSheetName.ItemIndex + 1;
  FromRowValueString := UpperCase(eFromRow.Text);
  ToRowValueString := UpperCase(eToRow.Text);
  eImportSpecNameCol.Text := UpperCase(eImportSpecNameCol.Text);
  {check row variables}
  iCode := 1;
  repeat
    {From Row}
    tmpStr := eFromRow.Text;
    Val(tmpStr, FromRow, iCode);
    {To Row}
    if (iCode = 0) then
    begin
      tmpStr := eToRow.Text;
      Val(tmpStr, ToRow, iCode);
    end else
    begin
      ShowMessage('Incorrect value entered for From row');
      Exit;
    end;
    if (iCode = 0) then
    begin
      if (ToRow >= FromRow) then iCode := 0
                            else iCode := -1;
    end else
    begin
      ShowMessage('Incorrect value entered for To row');
      Exit;
    end;
    if (iCode <> 0)
      then begin
        ShowMessage('Incorrect values entered for rows to import');
        Exit;
      end;
  until (iCode = 0);
  {convert input columns for variables to numeric}
  ImportSpecNameCol := ConvertCol2Int(eImportSpecNameCol.Text);
  dmCor.EigenVal.Open;
  dmCor.EigenVal.Last;
  if not (dmCor.EigenVal.Bof  and dmCor.EigenVal.Eof) then
  begin
    sbSheet.SimpleText := 'Clearing existing eigen values';
    repeat
      dmCor.EigenVal.Delete;
    until dmCor.EigenVal.Bof;
  end;
  sbSheet.SimpleText := 'Appending new eigen values';
  ii := 0;
  for i := FromRow to ToRow do
  begin
    ii := ii + 1;
    jj := 0;
    dmCor.EigenVal.Append;
    dmCor.EigenValVector.AsInteger := ii;
    iii := i+ImportSpecNameCol-1;
    for j := FromRow to FromRow + 2 do
    begin
      jj := jj + 1;
      try
        jjj := j+ImportSpecNameCol-1;
        //tmpStr := FlexCelImport1.CellValue[i,jjj];
        tmpStr := Xls.GetStringFromCell(i,jjj);
        dmCor.EigenVal.Fields[jj].AsString := tmpStr;
      except
      end;
    end;
    dmCor.EigenVal.Post;
  end;
  dmCor.EigenVal.Close;
  if (iCode = 0) then
  begin
    ModalResult := mrOK;
  end else
  begin
    ModalResult := mrNone;
  end;
end;

procedure TfmSheetImportEigVal.FormShow(Sender: TObject);
var
  i, j : integer;
begin
  //TabControl.Visible := false;
  Splitter1.Visible := false;
  eFromRow.Text := FromRowValueString;
  eToRow.Text := ToRowValueString;
  eImportSpecNameCol.Text := 'A';
  pDefinitions.Visible := false;
  bbOpenSheetClick(Sender);
end;


procedure TfmSheetImportEigVal.bbCancelClick(Sender: TObject);
begin
  ModalResult := mrNone;
  Close;
end;


procedure TfmSheetImportEigVal.sbFindLastRowClick(Sender: TObject);
var
  iCode : integer;
  tmpStr : string;
  i,j : integer;
begin
  eToRow.Text := '';
  ToRow := 0;
  i := 1;
  j := 1;
  try
    i := 2;
    j := ConvertCol2Int(eImportSpecNameCol.Text);
    //tmpStr := FlexCelImport1.CellValue[i,j];
    tmpStr := Xls.GetStringFromCell(i,j);
    while (tmpStr <> '') do
    begin
      if (i > ToRow) then ToRow := i;
      i := i + 1;
      //tmpStr := FlexCelImport1.CellValue[i,j];
      tmpStr := Xls.GetStringFromCell(i,j);
    end;
    eToRow.Text := IntToStr(ToRow);
    RowCount[ImportSheetNumber] := ToRow + 1;
  except
    //MessageDlg('Error reading data for main variable',mtwarning,[mbOK],0);
  end;
end;

procedure TfmSheetImportEigVal.cbSheetNameChange(Sender: TObject);
begin
  //ImportSheetNumber := cbSheetName.ItemIndex+1;
  //TabControl.TabIndex := cbSheetname.ItemIndex;
  //FlexCelImport1.ActiveSheet:= TabControl.TabIndex+1;
  //Data.ApplySheet;
  //Data.Zoom := 70;
  //Data.LoadSheet;
  Tabs.TabIndex := cbSheetname.ItemIndex;
  sbFindLastRowClick(Sender);
end;

end.
