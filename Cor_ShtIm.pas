unit Cor_ShtIm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, Buttons, OleCtrls, ExtCtrls, StdCtrls, Mask, IniFiles,
  Grids, DBGrids, DBCtrls, AxCtrls,
  Vcl.Tabs,
  VCL.FlexCel.Core, FlexCel.XlsAdapter, FlexCel.Render, FlexCel.Preview,
  Data.DB;

type
  TfmSheetImport = class(TForm)
    pControl: TPanel;
    sbSheet: TStatusBar;
    bbOpenSheet: TBitBtn;
    OpenDialogSprdSheet: TOpenDialog;
    gbDefineFields: TGroupBox;
    bbCancel: TBitBtn;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    gbDefineRows: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    meFromRow: TEdit;
    meToRow: TEdit;
    bbImport: TBitBtn;
    Memo1: TMemo;
    Label4: TLabel;
    dblcbImportSpec: TDBLookupComboBox;
    Label5: TLabel;
    Panel2: TPanel;
    Splitter1: TSplitter;
    gbDefineTabSheet: TGroupBox;
    cbSheetName: TComboBox;
    sbFindLastRow: TSpeedButton;
    cbCheckZero: TCheckBox;
    cbDoAitchison: TCheckBox;
    pDefinitions: TPanel;
    Panel1: TPanel;
    cbPositiveOnly: TCheckBox;
    pData: TPanel;
    SheetData: TStringGrid;
    Tabs: TTabSet;
    procedure bbOpenSheetClick(Sender: TObject);
    procedure bbImportClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bbCancelClick(Sender: TObject);
    procedure sbFindLastRowClick(Sender: TObject);
    procedure cbSheetNameChange(Sender: TObject);
    procedure dblcbImportSpecCloseUp(Sender: TObject);
    procedure TabsClick(Sender: TObject);
  private
    { Private declarations }
    Xls : TXlsFile;
    function ConvertCol2Int(AnyString : string) : integer;
    procedure FillTabs;
    procedure ClearGrid;
    procedure FillGrid(const Formatted: boolean);
    function GetStringFromCell(iRow,iCol : integer) : string;
    procedure GetElementOrder;
    procedure MatchElementsInFile;
    procedure CheckForZeros(Nox : integer);
    procedure CentredLogRatio(Nox : integer);
    procedure CentredLogRatioA(Nox : integer);
    procedure CalculateLogs(Nox : integer);
    procedure CalculateStatistics(Nox : integer);
    procedure SetStatsValue(j : integer; AValue : double);
  public
    { Public declarations }
  end;

var
  fmSheetImport: TfmSheetImport;

implementation

{$R *.DFM}

uses
  AllSorts, cor_varb, cor_dm_acs, mathproc, MATRIX;

var
  iRec, iRecCount      : integer;

function TfmSheetImport.GetStringFromCell(iRow,iCol : integer) : string;
begin
  Result := Xls.GetStringFromCell(iRow,iCol);
end;

procedure TfmSheetImport.FillTabs;
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

procedure TfmSheetImport.ClearGrid;
var
  r: integer;
begin
  for r := 1 to SheetData.RowCount do SheetData.Rows[r].Clear;
end;

procedure TfmSheetImport.FillGrid(const Formatted: boolean);
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

procedure TfmSheetImport.bbOpenSheetClick(Sender: TObject);
var
  tmpStr    : string;
  i         : integer;
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
    //dblcbImportSpec.KeyValue := dmCor.CoranFacImportGroup.AsString;
    try
      sbFindLastRowClick(Sender);
    except
    end;
  end;
end;

function TfmSheetImport.ConvertCol2Int(AnyString : string) : integer;
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

procedure TfmSheetImport.bbImportClick(Sender: TObject);
var
  j     : integer;
  iCode : integer;
  i : integer;
  FromRow, ToRow : integer;
  tmpStr : string;
begin
  ImportSheetNumber := cbSheetName.ItemIndex + 1;
  iCode := 1;
  repeat
    tmpStr := meFromRow.Text;
    Val(tmpStr, FromRow, iCode);
    if (iCode = 0) then
    begin
      tmpStr := meToRow.Text;
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
  with dmCor do
  begin
    cdsCoranChem.Close;
    cdsCoranRaw.Close;
    CoranFac.Close;
    SmpLoc.Close;
    ElemNames.Close;
    //cdsCoranChem.DisableControls;
    //cdsCoranRaw.DisableControls;
    //ElemNames.DisableControls;
    cdsCoranChem.Open;
    cdsCoranRaw.Open;
    CoranFac.Open;
    SmpLoc.Open;
    ElemNames.Open;
    CoranFac.First;
    CoranStats.Open;
  end;
  MatchElementsInFile;
  GetElementOrder;
  with dmCor do
  begin
    //cdsFacLoadingsSmp.DisableControls;
    //FacLoadingsVar.DisableControls;
    //GroupedSmp.DisableControls;
    //QGroupedSmp.DisableControls;
    //SmpLoc.DisableControls;
    //GroupedSmpLoc.DisableControls;
    {import groups, sample numbers and chemical data into CoranRaw}
    sbSheet.SimpleText := 'Importing data for '+IntToStr(Nox)+' variables';
    sbSheet.Refresh;
    j := 1;
    for i := FromRow to ToRow do
    begin
      cdsCoranRaw.Append;
      cdsCoranRawSequence.AsInteger := i;
      for j:= 1 to Nox+6 do
      begin
        //Data.Row := i;
        if (ElementPos[j] > 0) then
        begin
          //Data.Col := ElementPos[j];
          try
            //tmpStr := FlexCelImport1.CellValue[i,ElementPos[j]];
            tmpStr := Xls.GetStringFromCell(i,ElementPos[j]);
          except
            tmpStr := '0.0';
          end;
          if (j = 1) then   //Group
            try
              cdsCoranRawGroupName.AsString:= tmpStr;
            except
            end;
          if (j = 2) then   //PlotGroup
            try
              cdsCoranRawPlotGroupName.AsString := tmpStr;
            except
            end;
          if (j = 3) then    //Sample
            try
              cdsCoranRawSampleNum.AsString := tmpStr;
            except
            end;
          if (j > 6) then    //since do not want to read locality data at this stage
          begin
              if (tmpStr <> '') then
              begin
                try
                  cdsCoranRaw.Fields[j-4].AsString := tmpStr;
                except
                  cdsCoranRaw.Fields[j-4].AsString := '';
                end;
              end else
              begin
                cdsCoranRaw.Fields[j-4].AsFloat := 0.0;
              end;
              if ((cdsCoranRaw.Fields[j-4].AsFloat < 0.0) and (cbPositiveOnly.Checked)) then
              begin
                cdsCoranRaw.Fields[j-4].AsFloat := (Abs(cdsCoranRaw.Fields[j-4].AsFloat))/2.0;
              end;
          end;
        end else
        begin
          cdsCoranRaw.Fields[j-4].AsVariant := 0.0;
        end;
      end;
      cdsCoranRaw.Post;
    end;
    sbSheet.SimpleText := 'Saving untransformed data to database';
    sbSheet.Refresh;
    cdsCoranRaw.ApplyUpdates(-1);
    cdsCoranChem.First;
    cdsCoranRaw.First;
    {transfer to CoranChem}
    sbSheet.SimpleText := 'Transfer untransformed data to CoranChem';
    sbSheet.Refresh;
    repeat
      cdsCoranChem.Append;
      cdsCoranChemGroupName.AsString := cdsCoranRawGroupName.AsString;
      cdsCoranChemPlotGroupName.AsString := cdsCoranRawPlotGroupName.AsString;
      cdsCoranChemSampleNum.AsString := cdsCoranRawSampleNum.AsString;
      cdsCoranChemSequence.AsInteger := cdsCoranRawSequence.AsInteger;
      ElemNames.First;
      for j:= 1 to Nox do
      begin
        cdsCoranChem.Fields[j+2].AsVariant := cdsCoranRaw.Fields[j+2].AsVariant;
        if (dmCor.ElemNamesTakeLog.AsString = 'Y') then
        begin
          if (cdsCoranRaw.Fields[j+2].AsFloat <> 0.0) then
          begin
            cdsCoranChem.Fields[j+2].AsFloat := LogX(cdsCoranRaw.Fields[j+2].AsFloat);
          end else
          begin
            cdsCoranChem.Fields[j+2].AsFloat := -19.0;
          end;
        end else
        begin
          cdsCoranChem.Fields[j+2].AsFloat := cdsCoranRaw.Fields[j+2].AsFloat;
        end;
        ElemNames.Next;
      end;
      cdsCoranChem.Post;
      cdsCoranRaw.Next;
    until cdsCoranRaw.Eof;
    cdsCoranRaw.First;
    cdsCoranChem.First;
  end;
  with dmCor do
  begin
    {import latitude, longitude and elevation data}
    sbSheet.SimpleText := 'Importing latitude, longitude and elevation data';
    sbSheet.Refresh;
    j := 1;
    for i := FromRow to ToRow do
    begin
      SmpLoc.Append;
      for j:= 1 to 6 do
      begin
        //Data.Row := i;
        if (ElementPos[j] > 0) then
        begin
          //Data.Col := ElementPos[j];
          try
            //tmpStr := FlexCelImport1.CellValue[i,ElementPos[j]];
            tmpStr := Xls.GetStringFromCell(i,ElementPos[j]);
          except
            tmpStr := '0.0';
          end;
          if (j = 1) then
            try
              SmpLocGroupName.AsString := tmpStr;
            except
            end;
          if (j = 2) then
            try
              SmpLocPlotGroupName.AsString := tmpStr;
            except
            end;
          if (j = 3) then
            try
              SmpLocSampleNum.AsString := tmpStr;
            except
            end;
          if ((j = 4) or (j = 5) or (j = 6)) then
          begin
              if (tmpStr <> '') then
              begin
                try
                  SmpLoc.Fields[j-1].AsVariant := tmpStr;
                except
                  SmpLoc.Fields[j-1].AsString := '';
                end;
              end else
              begin
                SmpLoc.Fields[j-1].AsVariant := 0.0;
              end;
          end;
        end else
        begin
          SmpLoc.Fields[j-1].AsVariant := 0.0;
        end;
      end;
      SmpLoc.Post;
    end;
    SmpLoc.First;
  end;
  if cbCheckZero.Checked then
  begin
    sbSheet.SimpleText := 'Checking for zeros';
    sbSheet.Refresh;
    CheckForZeros(Nox);    //deletes zero records in CoranRaw, CoranChem and SmpLoc
  end;
  sbSheet.SimpleText := 'Re-saving untransformed data to database';
  sbSheet.Refresh;
  dmCor.cdsCoranRaw.ApplyUpdates(-1);
  if cbDoAitchison.Checked then CentredLogRatio(Nox)
                           else CentredLogRatioA(Nox);
  sbSheet.SimpleText := 'Re-saving transformed data to database';
  sbSheet.Refresh;
  dmCor.cdsCoranChem.ApplyUpdates(-1);
  CalculateStatistics(Nox);
  dmCor.cdsCoranChem.EnableControls;
  dmCor.cdsCoranRaw.EnableControls;
  dmCor.ElemNames.EnableControls;
  sbSheet.SimpleText := 'Finished importing data for '+IntToStr(Nox)+' variables';
  sbSheet.Refresh;
end;

procedure TfmSheetImport.FormShow(Sender: TObject);
var
  i, j : integer;
begin
  //TabControl.Visible := false;
  Splitter1.Visible := false;
  pDefinitions.Visible := false;
  meFromRow.Text := '2';
  meToRow.Text := '3';
  dmCor.CoranFac.Open;
  bbOpenSheetClick(Sender);
end;

procedure TfmSheetImport.bbCancelClick(Sender: TObject);
begin
  ModalResult := mrNone;
  Close;
end;

procedure TfmSheetImport.GetElementOrder;
var
  Sum : double;
  i : integer;
begin
   with dmCor do
   begin
     ElemNames.Last;
     if not (ElemNames.Bof and ElemNames.Eof) then
     begin
       repeat
         ElemNames.Delete;
       until ElemNames.Bof;
     end;
     CoranFac.First;
     repeat
       ElementPos[CoranFacPos.AsInteger+6] := CoranFacColumnNo.AsInteger;
       if ((CoranFacPos.AsInteger > 0) and (CoranfacPos.AsInteger <= MM)) then
       begin
         if (CoranFacColumnNo.AsInteger > 0) then
         begin
           OxideName[CoranFacPos.AsInteger] := CoranFacCalled.AsString;
           TakeLogs[CoranFacPos.AsInteger] := UpperCase(CoranFacTakeLog.AsString);
           if ((TakeLogs[CoranFacPos.AsInteger] <> 'Y') and (TakeLogs[CoranFacPos.AsInteger] <> 'A'))
             then TakeLogs[CoranFacPos.AsInteger] := 'N';
           ElemNames.Append;
           ElemNamesPos.AsInteger := CoranFacPos.AsInteger;
           ElemNamesCalled.AsString := CoranFacCalled.AsString;
           ElemNamesTakeLog.AsString := TakeLogs[CoranFacPos.AsInteger];
           ElemNamesWSumFac.AsString := CoranFacWSumFac.AsVariant;
           ElemNamesWSumFacAdj.AsFloat := 0.0;
           ElemNames.Post;
         end;
       end;
       CoranFac.Next;
     until CoranFac.EOF;
     CoranFac.First;
     Nox := dmCor.ElemNames.RecordCount;
     ElemNames.First;
     Sum := 0.0;
     for i := 1 to Nox do
     begin
       Sum := Sum + Abs(ElemNamesWSumFac.AsFloat);
       ElemNames.Next;
     end;
     ElemNames.First;
     for i := 1 to Nox do
     begin
       ElemNames.Edit;
       dmCor.ElemNamesWSumFacAdj.AsFloat := dmCor.ElemNamesWSumFac.AsFloat/Sum;
       ElemNames.Post;
       ElemNames.Next;
     end;
     ElemNames.First;
   end;
end;{proc GetElementOrder}

procedure TfmSheetImport.MatchElementsInFile;
var
  tmpStr : string;
begin
  with dmCor do
  begin
    CoranFac.First;
    repeat
      tmpStr := UpperCase(CoranFacColumn.AsString);
      ClearNull(tmpStr);
      CoranFac.Edit;
      CoranFacColumn.AsString := tmpStr;
      CoranFacTakeLog.AsString := UpperCase(CoranFacTakeLog.AsString);
      CoranFac.Post;
      CoranFac.Edit;
      if (CoranFacColumn.AsString >= 'A') then
      begin
        CoranFacColumnNo.AsInteger := ConvertCol2Int(tmpStr);
      end else
      begin
        CoranFacColumnNo.AsInteger := 0;
      end;
      CoranFac.Post;
      CoranFac.Next;
    until CoranFac.EOF;
    CoranFac.First;
  end;
end;

procedure TfmSheetImport.CheckForZeros(Nox : integer);
var
  Total : double;
  i, j : integer;
begin
  with dmCor do
  begin
    sbSheet.SimpleText := 'Checking CoranRaw for zeros';
    sbSheet.Refresh;
    cdsCoranRaw.First;
    i := 1;
    repeat
      Total := 0.0;
      for j := 1 to Nox do
      begin
        Total := Total + Abs(cdsCoranRaw.Fields[j+2].AsFloat);
      end;
      if (Total < 0.00001) then
      begin
        cdsCoranRaw.Delete;
        cdsCoranChem.Delete;
      end else
      begin
        cdsCoranRaw.Next;
        cdsCoranChem.Next;
      end;
    until cdsCoranRaw.EOF;
    cdsCoranRaw.First;
    sbSheet.SimpleText := 'Checking SmpLoc for zeros';
    sbSheet.Refresh;
    SmpLoc.First;
    repeat
      Total := 0.0;
      for j := 3 to 5 do
      begin
        Total := Total + SmpLoc.Fields[j].AsFloat
      end;
      if (Total = 0.0) then SmpLoc.Delete
                       else SmpLoc.Next;
    until SmpLoc.EOF;
    SmpLoc.First;
  end;
  sbSheet.SimpleText := 'Finished checking for zeros';
  sbSheet.Refresh;
end;

procedure TfmSheetImport.CalculateLogs(Nox : integer);
var
  Total : double;
  i, j : integer;
begin
  with dmCor do
  begin
    sbSheet.SimpleText := 'Calculating logs for specified variables';
    sbSheet.Refresh;
    for j := 1 to Nox do
    begin
      if (TakeLogs[j] = 'Y') then
      begin
        cdsCoranChem.First;
        repeat
          cdsCoranChem.Edit;
          if (cdsCoranChem.Fields[j+2].AsFloat > 0.0) then
          begin
            cdsCoranChem.Fields[j+2].AsFloat := LogX(cdsCoranChem.Fields[j+2].AsFloat);
          end else
          begin
            cdsCoranChem.Fields[j+2].AsFloat := LogX(DefaultMinimum);
          end;
          cdsCoranChem.Post;
          cdsCoranChem.Next;
        until cdsCoranChem.EOF;
      end;
    end;
  end;
  sbSheet.SimpleText := 'Finished calculating logs for specified variables';
  sbSheet.Refresh;
end;

procedure TfmSheetImport.sbFindLastRowClick(Sender: TObject);
var
  iCode : integer;
  tmpStr : string;
  i,j : integer;
begin
  ImportSheetNumber := cbSheetName.ItemIndex+1;
  meToRow.Text := '';
  ToRow := 0;
  dmCor.CoranFac.First;
  iCode := 1;
  repeat
    tmpStr := meFromRow.Text;
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
    j := ConvertCol2Int(dmCor.CoranFacCOLUMN.AsString);
    //i := Data.Row;
    //j := Data.Col;
    ToRow := 0;
    repeat
      //if (Data.Row > 48) then ShowMessage('repeat '+IntToStr(Data.Row)+'   '+FlexCelImport1.CellValue[Data.Row,Data.Col]);
      i := i + 1;
      if (i > ToRow) then ToRow := i-1;
      meToRow.Text := IntToStr(ToRow);
      try
        //tmpStr := FlexCelImport1.CellValue[i,j];
        tmpStr := Xls.GetStringFromCell(i,j);
      except
        tmpStr := '0.0';
      end;
    until (tmpStr = '');
  except
    //MessageDlg('Error reading data in column '+IntToStr(Data.Col),mtwarning,[mbOK],0);
  end;
  meToRow.Text := IntToStr(ToRow);
  RowCount[ImportSheetNumber] := ToRow + 1;
  //Data.Row := 1;
end;

procedure TfmSheetImport.cbSheetNameChange(Sender: TObject);
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

procedure TfmSheetImport.dblcbImportSpecCloseUp(Sender: TObject);
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

procedure TfmSheetImport.TabsClick(Sender: TObject);
begin
  FillGrid(true);
end;

procedure TfmSheetImport.CentredLogRatio(Nox : integer);
var
  i, j : integer;
  RowSum, ColSum : double;
  tmp : double;
begin
  with dmCor do
  begin
    sbSheet.SimpleText := 'Logratio centering all records';
    sbSheet.Refresh;
    cdsCoranChem.First;
    i := 0;
    repeat
      i := i+1;
      DR[i] := 0.0;
      cdsCoranChem.Edit;
      for j := 1 to Nox do
      begin
        if ((TakeLogs[j] = 'N') or (TakeLogs[j] = 'A')) then
        begin
          if (cdsCoranChem.Fields[j+2].AsFloat > 0.0) then
          begin
            cdsCoranChem.Fields[j+2].AsFloat := LogX(cdsCoranChem.Fields[j+2].AsFloat);
          end else
          begin
            cdsCoranChem.Fields[j+2].AsFloat := LogX(DefaultMinimum);
          end;
        end;
        DR[i] := DR[i] + cdsCoranChem.Fields[j+2].AsFloat;
      end;
      DR[i] := DR[i] / (1.0*Nox);
      cdsCoranChem.Next;
    until cdsCoranChem.EOF;
    cdsCoranChem.First;
    i := 0;
    repeat
      i := i+1;
      for j := 1 to Nox do
      begin
      end;
      cdsCoranChem.Next;
    until cdsCoranChem.EOF;
    cdsCoranChem.First;
    i := 0;
    repeat
      i := i+1;
      cdsCoranChem.Edit;
      for j := 1 to Nox do
      begin
        tmp := cdsCoranChem.Fields[j+2].AsFloat;
        cdsCoranChem.Fields[j+2].AsFloat := tmp - DR[i];
      end;
      cdsCoranChem.Next;
    until cdsCoranChem.EOF;
    cdsCoranChem.First;
  end;
  sbSheet.SimpleText := 'Finished logratio centering all records';
  sbSheet.Refresh;
end;

procedure TfmSheetImport.CentredLogRatioA(Nox : integer);
var
  i, j, n : integer;
  RowSum, ColSum : double;
begin
  with dmCor do
  begin
    sbSheet.SimpleText := 'Logratio centering user-specified records';
    sbSheet.Refresh;
    n := 0;
    CoranFac.First;
    repeat
      if (CoranFacTakeLog.AsString = 'A') then n := n + 1;
      CoranFac.Next;
    until CoranFac.Eof;
    if (n > 0) then
    begin
      sbSheet.SimpleText := 'Logratio centering '+IntToStr(n)+' user-specified fields ';
      sbSheet.Refresh;
      cdsCoranChem.First;
      i := 0;
      repeat
        i := i+1;
        DR[i] := 0.0;
        cdsCoranChem.Edit;
        for j := 1 to Nox do
        begin
          if (TakeLogs[j] = 'A') then
          begin
            if (cdsCoranChem.Fields[j+2].AsFloat > 0) then
            begin
              cdsCoranChem.Fields[j+2].AsFloat := LogX(cdsCoranChem.Fields[j+2].AsFloat);
            end else
            begin
              cdsCoranChem.Fields[j+2].AsFloat := LogX(DefaultMinimum);
            end;
            DR[i] := DR[i] + cdsCoranChem.Fields[j+2].AsFloat;
          end;
        end;
        DR[i] := DR[i] / (1.0*n);
        cdsCoranChem.Next;
      until cdsCoranChem.EOF;
      cdsCoranChem.First;
      i := 0;
      repeat
        i := i+1;
        cdsCoranChem.Edit;
        for j := 1 to Nox do
        begin
          if (TakeLogs[j] = 'A') then
          begin
            cdsCoranChem.Fields[j+2].AsFloat := cdsCoranChem.Fields[j+2].AsFloat - DR[i];
          end;
        end;
        cdsCoranChem.Next;
      until cdsCoranChem.EOF;
      cdsCoranChem.First;
    end;
  end;
  sbSheet.SimpleText := 'Finished logratio centering user-specified records for '+IntToStr(n)+' fields';
  sbSheet.Refresh;
end;

procedure TfmSheetImport.CalculateStatistics(Nox : integer);
var
  i, j, n, nn : integer;
  RowSum, ColSum, t : double;
  XMed : double;
  ml : integer;
begin
  with dmCor do
  begin
    sbSheet.SimpleText := 'Calculating Statistics';
    sbSheet.Refresh;
    CoranStats.First;
    repeat
      CoranStats.Edit;
      for j := 1 to MMaxFields do
      begin
        SetStatsValue(j,zero);
      end;
      CoranStats.Next;
    until CoranStats.Eof;
    CoranStats.First;
    cdsCoranChem.First;
    for j := 1 to Nox do
    begin
      cdsCoranChem.First;
      cdsCoranRaw.First;
      ColSum := 0.0;
      XMin := 9.99e20;
      XMax := -9.99e20;
      n := 0;
      nn := 0;
      DD[j] := 0.0;
      repeat
        n := n + 1;
        if (cdsCoranRaw.Fields[j+2].AsFloat > zero) then nn := nn + 1;
        if (XMin > cdsCoranChem.Fields[j+2].AsFloat) then XMin := cdsCoranChem.Fields[j+2].AsFloat;
        if (XMax < cdsCoranChem.Fields[j+2].AsFloat) then XMax := cdsCoranChem.Fields[j+2].AsFloat;
        ColSum := ColSum +cdsCoranChem.Fields[j+2].AsFloat;
        cdsCoranChem.Next;
        cdsCoranRaw.Next;
      until cdsCoranChem.EOF;
      CoranStats.First;
      CoranStats.Locate('StatsID',1,[]);
      CoranStats.Edit;
      SetStatsValue(j,1.0*n);
      CoranStats.Next;
      CoranStats.Locate('StatsID',2,[]);
      CoranStats.Edit;
      SetStatsValue(j,1.0*nn);
      CoranStats.Next;
      CoranStats.Locate('StatsID',15,[]);
      CoranStats.Edit;
      SetStatsValue(j,XMin);
      CoranStats.Next;
      CoranStats.Locate('StatsID',9,[]);
      CoranStats.Edit;
      SetStatsValue(j,XMax);
      CoranStats.Next;
      CoranStats.Locate('StatsID',8,[]);
      CoranStats.Edit;
      SetStatsValue(j,XMax-XMin);
      CoranStats.Next;
      CoranStats.Locate('StatsID',3,[]);
      CoranStats.Edit;
      DD[j] := ColSum/(1.0*n);
      SetStatsValue(j,DD[j]);
      CoranStats.Post;
      Ndig[j] := n;
    end;
    cdsCoranChem.First;
    cdsCoranRaw.First;
    CoranStats.Next;
    for j := 1 to Nox do
    begin
      cdsCoranChem.First;
      ColSum := 0.0;
      repeat
        t := cdsCoranChem.Fields[j+2].AsFloat - DD[j];
        ColSum := ColSum + t*t;
        cdsCoranChem.Next;
      until cdsCoranChem.EOF;
      ColSum := Sqrt(ColSum/(1.0*(Ndig[j]-1)));
      CoranStats.Locate('StatsID',4,[]);
      CoranStats.Edit;
      SetStatsValue(j,ColSum);
      CoranStats.Post;
    end;
    CoranStats.Next;
    for j := 1 to Nox do
    begin
      cdsCoranChem.First;
      i := 1;
      repeat
        DR[i] :=  cdsCoranChem.Fields[j+2].AsFloat;
        i := i + 1;
        cdsCoranChem.Next;
      until cdsCoranChem.Eof;
      Sort(DR,n,1);
      ml := (n+1) div 2;
      {calculate and store the median}
      n := Ndig[j];
      XMed := 0.5*(DR[ml] + DR[n-ml+1]);
      DC[1,j] := XMed;
      CoranStats.Locate('StatsID',12,[]);
      CoranStats.Edit;
      SetStatsValue(j,DC[1,j]);
      CoranStats.Post;
    end;
    CoranStats.Next;
    for j := 1 to Nox do
    begin
      cdsCoranChem.First;
      i := 1;
      repeat
        DR[i] :=  Abs(cdsCoranChem.Fields[j+2].AsFloat - DD[j]);
        i := i + 1;
        cdsCoranChem.Next;
      until cdsCoranChem.Eof;
      n := Ndig[j];
      Sort(DR,n,1);
      {calculate and store the median deviation from the median}
      XMed := 0.5*(DR[ml] + DR[n-ml+1]);
      CoranStats.Locate('StatsID',6,[]);
      CoranStats.Edit;
      SetStatsValue(j,XMed);
      CoranStats.Post;
    end;
    CoranStats.Next;
    for j := 1 to Nox do
    begin
      cdsCoranChem.First;
      i := 1;
      repeat
        DR[i] :=  cdsCoranChem.Fields[j+2].AsFloat;
        i := i + 1;
        cdsCoranChem.Next;
      until cdsCoranChem.Eof;
      n := Ndig[j];
      Sort(DR,n,1);
      {calculate upper and lower hinge values}
      Hinges(DR,XMin,XMax,n);
      DC[2,j] := XMin;
      DC[3,j] := XMax;
      DC[4,j] := 0.5*(DC[3,j]-DC[2,j]);
      DC[5,j] := 0.5*DC[1,j] + 0.25*DC[2,j] + 0.25*DC[3,j];
      CoranStats.Locate('StatsID',13,[]);
      CoranStats.Edit;
      {store lower hinge value}
      SetStatsValue(j,DC[2,j]);
      CoranStats.Post;
    end;
    CoranStats.Next;
    CoranStats.Locate('StatsID',11,[]);
    for j := 1 to Nox do
    begin
      {store upper hinge value}
      CoranStats.Edit;
      SetStatsValue(j,DC[3,j]);
      CoranStats.Post;
    end;
    CoranStats.Next;
    CoranStats.Locate('StatsID',7,[]);
    for j := 1 to Nox do
    begin
      {store half hinge width}
      CoranStats.Edit;
      SetStatsValue(j,DC[4,j]);
      CoranStats.Post;
    end;
    CoranStats.Next;
    CoranStats.Locate('StatsID',10,[]);
    for j := 1 to Nox do
    begin
      {store upper fence value}
      CoranStats.Edit;
      SetStatsValue(j,DC[3,j]+1.5*DC[4,j]);
      CoranStats.Post;
    end;
    CoranStats.Next;
    CoranStats.Locate('StatsID',14,[]);
    for j := 1 to Nox do
    begin
      {store lower fence value}
      CoranStats.Edit;
      SetStatsValue(j,DC[2,j]-1.5*DC[4,j]);
      CoranStats.Post;
    end;
    CoranStats.Next;
    CoranStats.Locate('StatsID',5,[]);
    for j := 1 to Nox do
    begin
      {store trimean}
      CoranStats.Edit;
      SetStatsValue(j,DC[5,j]);
      CoranStats.Post;
    end;
    CoranStats.First;
  end;
  sbSheet.SimpleText := 'Finished Calculating Statistics';
  sbSheet.Refresh;
end;

procedure TfmSheetImport.SetStatsValue(j : integer; AValue : double);
begin
  dmCor.CoranStats.Fields[j+1].AsVariant := AValue;
end;

end.
