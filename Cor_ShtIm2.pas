unit Cor_ShtIm2;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, Buttons, OleCtrls, ExtCtrls, StdCtrls, Mask, IniFiles,
  Grids, DBGrids, DBCtrls, AxCtrls,
  Vcl.Tabs,
  VCL.FlexCel.Core, FlexCel.XlsAdapter, FlexCel.Render, FlexCel.Preview,
  Data.DB;

type
  TfmSheetImport2 = class(TForm)
    pControl: TPanel;
    sbSheet: TStatusBar;
    bbOpenSheet: TBitBtn;
    gbDefineFields: TGroupBox;
    bbCancel: TBitBtn;
    gbDefineRows: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    eFromRow: TEdit;
    eToRow: TEdit;
    bbImport: TBitBtn;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    eImportSpecNameCol: TEdit;
    ePositionCol: TEdit;
    eCalledCol: TEdit;
    Label10: TLabel;
    sbFindLastRow: TSpeedButton;
    pSpreadSheet: TPanel;
    Panel3: TPanel;
    gbDefineTabSheet: TGroupBox;
    cbSheetName: TComboBox;
    Label21: TLabel;
    Label22: TLabel;
    Label12: TLabel;
    eColumnCol: TEdit;
    Label13: TLabel;
    eTakeLogCol: TEdit;
    Splitter1: TSplitter;
    pDefinitions: TPanel;
    gbDefaults: TGroupBox;
    Splitter2: TSplitter;
    Label1: TLabel;
    eDefaultMinimum: TEdit;
    Label4: TLabel;
    Label8: TLabel;
    lMaxVar: TLabel;
    lMaxSmp: TLabel;
    Memo1: TMemo;
    Label9: TLabel;
    eWSumFacCol: TEdit;
    pData: TPanel;
    SheetData: TStringGrid;
    Tabs: TTabSet;
    OpenDialogSprdSheet: TOpenDialog;
    procedure bbOpenSheetClick(Sender: TObject);
    procedure bbImportClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bbCancelClick(Sender: TObject);
    procedure sbFindLastRowClick(Sender: TObject);
    procedure cbSheetNameChange(Sender: TObject);
  private
    { Private declarations }
    Xls : TXlsFile;
    function ConvertCol2Int(AnyString : string) : integer;
    procedure FillTabs;
    procedure ClearGrid;
    procedure FillGrid(const Formatted: boolean);
    function GetStringFromCell(iRow,iCol : integer) : string;
  public
    { Public declarations }
  end;

var
  fmSheetImport2: TfmSheetImport2;

implementation

uses Allsorts, Cor_varb, Cor_dm_acs;

{$R *.DFM}

var
  iRec, iRecCount      : integer;

function TfmSheetImport2.GetStringFromCell(iRow,iCol : integer) : string;
begin
  Result := Xls.GetStringFromCell(iRow,iCol);
end;

procedure TfmSheetImport2.FillTabs;
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

procedure TfmSheetImport2.ClearGrid;
var
  r: integer;
begin
  for r := 1 to SheetData.RowCount do SheetData.Rows[r].Clear;
end;

procedure TfmSheetImport2.FillGrid(const Formatted: boolean);
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

function TfmSheetImport2.ConvertCol2Int(AnyString : string) : integer;
var
  itmp    : integer;
  tmpStr  : string;
  tmpChar : char;
begin
    AnyString := UpperCase(AnyString);
    tmpStr := AnyString;
    ClearNull(tmpStr);
    Result := 0;
    if (length(tmpStr) = 2) then
    begin
      tmpChar := tmpStr[1];
      itmp := (ord(tmpChar)-64)*26;
      tmpChar := tmpStr[2];
      Result := itmp+(ord(tmpChar)-64);
    end else
    begin
      tmpChar := tmpStr[1];
      Result := (ord(tmpChar)-64);
    end;
end;

procedure TfmSheetImport2.bbOpenSheetClick(Sender: TObject);
var
  tmpStr    : string;
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
    pDefinitions.Visible := true;
    Splitter1.Visible := true;
    //TabControl.Visible := true;

    bbImport.Visible := true;
    //Data.Row := 1;
    //Data.Col := 1;
    FillGrid(true);
    pDefinitions.Visible := true;
    sbFindLastRowClick(Sender);
  end;
end;


procedure TfmSheetImport2.bbImportClick(Sender: TObject);
var
  j      : integer;
  iCode  : integer;
  i      : integer;
  tmpStr : string;
begin
  ImportSheetNumber := cbSheetName.ItemIndex + 1;
  //Data.Row := 1;
  //Data.Col := 1;
  FromRowValueString := UpperCase(eFromRow.Text);
  ToRowValueString := UpperCase(eToRow.Text);
  eImportSpecNameCol.Text := UpperCase(eImportSpecNameCol.Text);
  ePositionCol.Text := UpperCase(ePositionCol.Text);
  eCalledCol.Text := UpperCase(eCalledCol.Text);
  eColumnCol.Text := UpperCase(eColumnCol.Text);
  eTakeLogCol.Text := UpperCase(eTakeLogCol.Text);
  eWSumFacCol.Text := UpperCase(eWSumFacCol.Text);
  ImportSpecNameColStr := eImportSpecNameCol.Text;
  PositionColStr := ePositionCol.Text;
  CalledColStr := eCalledCol.Text;
  ColumnColStr := eColumnCol.Text;
  TakeLogColStr := eTakeLogCol.Text;
  WSumFacColStr := eWSumFacCol.Text;
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
  {check Default values}
  iCode := 1;
  repeat
    tmpStr := eDefaultMinimum.Text;
    Val(tmpStr, DefaultMinimum, iCode);
    if (iCode = 0) then
    begin
      eDefaultMinimum.Text := FormatFloat('##0.0000e-00',DefaultMinimum);
    end else
    begin
      ShowMessage('Incorrect value entered for Default Minimum');
      Exit;
    end;
  until (iCode = 0);
  {convert input columns for variables to numeric}
  ImportSpecNameCol := ConvertCol2Int(eImportSpecNameCol.Text);
  PositionCol := ConvertCol2Int(ePositionCol.Text);
  CalledCol := ConvertCol2Int(eCalledCol.Text);
  ColumnCol := ConvertCol2Int(eColumnCol.Text);
  TakeLogCol := ConvertCol2Int(eTakeLogCol.Text);
  WSumFacCol := ConvertCol2Int(eWSumFacCol.Text);
  dmCor.CoranFacAll.Open;
  dmCor.CoranFacAll.Last;
  if not (dmCor.CoranFacAll.Bof  and dmCor.CoranFacAll.Eof) then
  begin
    sbSheet.SimpleText := 'Clearing existing definitions';
    repeat
      dmCor.CoranFacAll.Delete;
    until dmCor.CoranFacAll.Bof;
  end;
  sbSheet.SimpleText := 'Appending new definitions';
  //ShowMessage('5');
  //ShowMessage('FromRow = '+IntToStr(FromRow)+'   ToRow = '+IntToStr(ToRow));
  for i := FromRow to ToRow do
  begin
    try
      dmCor.CoranFacAll.Append;
      //Data.Row := i;
      //Data.Col := ImportSpecNameCol;
      j := ImportSpecNameCol;
      //tmpStr := FlexCelImport1.CellValue[i,j];
      tmpStr := Xls.GetStringFromCell(i,j);
      //ShowMessage('i = '+IntToStr(i)+' '+tmpStr+'   j = '+IntToStr(j));
      dmCor.CoranFacAllImportGroup.AsString := tmpStr;
      //ShowMessage(IntToStr(i)+'**'+tmpStr);
      j := PositionCol;
      //tmpStr := FlexCelImport1.CellValue[i,j];
      tmpStr := Xls.GetStringFromCell(i,j);
      dmCor.CoranFacAllPOS.AsString := tmpStr;
      j := CalledCol;
      //tmpStr := FlexCelImport1.CellValue[i,j];
      tmpStr := Xls.GetStringFromCell(i,j);
      dmCor.CoranFacAllCalled.AsString := tmpStr;
      j := ColumnCol;
      //tmpStr := FlexCelImport1.CellValue[i,j];
      tmpStr := Xls.GetStringFromCell(i,j);
      dmCor.CoranFacAllCOLUMN.AsString := tmpStr;
      j := TakeLogCol;
      //tmpStr := FlexCelImport1.CellValue[i,j];
      tmpStr := Xls.GetStringFromCell(i,j);
      dmCor.CoranFacAllTakeLog.AsString := tmpStr;
      j := WSumFacCol;
      //tmpStr := FlexCelImport1.CellValue[i,j];
      tmpStr := Xls.GetStringFromCell(i,j);
      dmCor.CoranFacAllWSumFac.AsString := tmpStr;
      dmCor.CoranFacAll.Post;
    except
    end;
  end;
  dmCor.CoranFacAll.First;
  //ShowMessage('5b');
  repeat
    if (dmCor.CoranFacAllPOS.AsInteger > MM) then dmCor.CoranFacAll.Delete
                                             else dmCor.CoranFacAll.Next;
  until dmCor.CoranFacAll.Eof;
  dmCor.CoranFacAll.Close;
  if (iCode = 0) then
  begin
    ModalResult := mrOK;
  end else
  begin
    ModalResult := mrNone;
  end;
  //ShowMessage('6');
end;

procedure TfmSheetImport2.FormShow(Sender: TObject);
var
  i, j : integer;
begin
  lMaxVar.Caption := IntToStr(MM);
  lMaxSmp.Caption := IntToStr(NN);
  bbImport.Visible := false;
  //TabControl.Visible := false;
  Splitter1.Visible := false;
  if (FromRowValueString = '') then FromRowValueString := '2';
  if (ToRowValueString = '') then ToRowValueString := '2';
  eFromRow.Text := FromRowValueString;
  eToRow.Text := ToRowValueString;
  eImportSpecNameCol.Text := ImportSpecNameColStr;
  ePositionCol.Text := PositionColStr;
  eCalledCol.Text := CalledColStr;
  eColumnCol.Text := ColumnColStr;
  eTakeLogCol.Text := TakeLogColStr;
  eWSumFacCol.Text := WSumFacColStr;
  eDefaultMinimum.Text := FormatFloat('##0.0000e-00',DefaultMinimum);
  pDefinitions.Visible := false;
  bbOpenSheetClick(Sender);
end;


procedure TfmSheetImport2.bbCancelClick(Sender: TObject);
begin
  ModalResult := mrNone;
  Close;
end;


procedure TfmSheetImport2.sbFindLastRowClick(Sender: TObject);
var
  iCode : integer;
  tmpStr : string;
  i,j : integer;
begin
  ImportSheetNumber := cbSheetName.ItemIndex+1;
  eToRow.Text := '';
  ToRow := 0;
  iCode := 1;
  repeat
    tmpStr := eFromRow.Text;
    Val(tmpStr, FromRow, iCode);
    if (iCode = 0) then
    begin
    end else
    begin
      ShowMessage('Incorrect value entered for From row');
      Exit;
    end;
  until (iCode = 0);
  try
    i := FromRow;
    j := ConvertCol2Int(eImportSpecNameCol.Text);
    //Data.Row := i;
    //Data.Col := j;
    ToRow := 0;
    repeat
      //if (Data.Row > 48) then ShowMessage('repeat '+IntToStr(Data.Row)+'   '+FlexCelImport1.CellValue[Data.Row,Data.Col]);
      i := i + 1;
      //Data.Row := i;
      //Data.Col := j;
      if (i > ToRow) then ToRow := i-1;
      eToRow.Text := IntToStr(ToRow);
      //tmpStr := FlexCelImport1.CellValue[i,j];
      tmpStr := Xls.GetStringFromCell(i,j);
    until (tmpStr = '');
    eToRow.Text := IntToStr(ToRow);
    RowCount[ImportSheetNumber] := ToRow + 1;
  except
    //MessageDlg('Error reading data for main variable',mtwarning,[mbOK],0);
  end;
end;

procedure TfmSheetImport2.cbSheetNameChange(Sender: TObject);
begin
  //ImportSheetNumber := cbSheetName.ItemIndex+1;
  //TabControl.TabIndex := cbSheetname.ItemIndex;
  //FlexCelImport1.ActiveSheet:= TabControl.TabIndex+1;
  //Data.ApplySheet;
  //Data.Zoom := 70;
  //Data.LoadSheet;
  //sbFindLastRowClick(Sender);
  Tabs.TabIndex := cbSheetname.ItemIndex;
  sbFindLastRowClick(Sender);
end;

end.
