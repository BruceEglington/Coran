unit Cor_mn_test;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, StdCtrls, Buttons, ExtCtrls, ToolWin, ComCtrls,
  Printers, Menus, Mask, DBCtrls, Db, DBTables, IniFiles, TeEngine, Series,
  TeeProcs, Chart, DbChart, AxCtrls, OleCtrls, TTF160_TLB, ActnList,
  ActnMan, TeeSurfa, TeePoin3;

type
  TfmCoranMain = class(TForm)
    ToolBar1: TToolBar;
    sbMain: TStatusBar;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Exit1: TMenuItem;
    pc1: TPageControl;
    tsControl: TTabSheet;
    N1: TMenuItem;
    Import1: TMenuItem;
    Help: TMenuItem;
    About1: TMenuItem;
    Printersetup1: TMenuItem;
    PrinterSetupDialog1: TPrinterSetupDialog;
    N2: TMenuItem;
    PrintDialog1: TPrintDialog;
    Options1: TMenuItem;
    DataLinkFile1: TMenuItem;
    tsResults: TTabSheet;
    Panel4: TPanel;
    Label1: TLabel;
    eTitle: TEdit;
    MemoResults: TMemo;
    tsLoadings: TTabSheet;
    pSmp: TPanel;
    DBGrid4: TDBGrid;
    pVar: TPanel;
    bbEmptyCoranVecVar: TBitBtn;
    DBGrid5: TDBGrid;
    DBNavigator2: TDBNavigator;
    tsGraph1: TTabSheet;
    DBChart1: TDBChart;
    Series1: TPointSeries;
    Series2: TPointSeries;
    tsGraph2: TTabSheet;
    tsGraph3: TTabSheet;
    DBChart2: TDBChart;
    PointSeries1: TPointSeries;
    PointSeries2: TPointSeries;
    DBChart3: TDBChart;
    PointSeries3: TPointSeries;
    PointSeries4: TPointSeries;
    tsSpreadSheets: TTabSheet;
    pSaveVar: TPanel;
    F1Var: TF1Book6;
    pSaveSmp: TPanel;
    F1Smp: TF1Book6;
    SaveDialogSprdSheet: TSaveDialog;
    Series3: TPointSeries;
    Series4: TPointSeries;
    Series5: TPointSeries;
    tsCheck: TTabSheet;
    DBNavigator8: TDBNavigator;
    DBGrid10: TDBGrid;
    Button1: TButton;
    DBGrid11: TDBGrid;
    DBGrid9: TDBGrid;
    Series6: TPointSeries;
    DBNavigator7: TDBNavigator;
    Series7: TPointSeries;
    Series8: TPointSeries;
    tsLocalities: TTabSheet;
    DBChart4: TDBChart;
    PointSeries6: TPointSeries;
    PointSeries7: TPointSeries;
    PointSeries8: TPointSeries;
    DBNavigator13: TDBNavigator;
    DBGrid19: TDBGrid;
    Series10: TPointSeries;
    Series9: TPointSeries;
    Series11: TPointSeries;
    PrintGraph1: TMenuItem;
    DBGrid15: TDBGrid;
    lIgnoreLoadingsVar: TLabel;
    tsScores: TTabSheet;
    DBChart5: TDBChart;
    PointSeries12: TPointSeries;
    PointSeries5: TPointSeries;
    DBChart7: TDBChart;
    PointSeries10: TPointSeries;
    PointSeries11: TPointSeries;
    DBChart6: TDBChart;
    PointSeries9: TPointSeries;
    PointSeries13: TPointSeries;
    DBGrid22: TDBGrid;
    DBGrid23: TDBGrid;
    Panel6: TPanel;
    Panel3: TPanel;
    DBGrid2: TDBGrid;
    Panel11: TPanel;
    pGraph1Var: TPanel;
    DBGrid17: TDBGrid;
    dbnVar: TDBNavigator;
    Panel7: TPanel;
    DBGrid7: TDBGrid;
    dbnGroups: TDBNavigator;
    DBGrid24: TDBGrid;
    DBNavigator11: TDBNavigator;
    Panel9: TPanel;
    DBGrid12: TDBGrid;
    dbnGroupedSmp: TDBNavigator;
    Panel8: TPanel;
    pGraph2Var: TPanel;
    DBGrid6: TDBGrid;
    DBNavigator3: TDBNavigator;
    Panel13: TPanel;
    DBGrid8: TDBGrid;
    DBNavigator4: TDBNavigator;
    DBGrid13: TDBGrid;
    DBNavigator9: TDBNavigator;
    Panel17: TPanel;
    DBGrid16: TDBGrid;
    DBNavigator12: TDBNavigator;
    Panel12: TPanel;
    pGraph3Var: TPanel;
    DBGrid14: TDBGrid;
    DBNavigator10: TDBNavigator;
    Panel18: TPanel;
    DBGrid18: TDBGrid;
    DBNavigator14: TDBNavigator;
    DBGrid21: TDBGrid;
    DBNavigator15: TDBNavigator;
    Panel19: TPanel;
    DBGrid25: TDBGrid;
    DBNavigator16: TDBNavigator;
    Panel20: TPanel;
    bbSaveVar: TBitBtn;
    Panel21: TPanel;
    bbSaveSmp: TBitBtn;
    Panel15: TPanel;
    Panel22: TPanel;
    DBGrid26: TDBGrid;
    DBNavigator18: TDBNavigator;
    DBGrid27: TDBGrid;
    DBNavigator19: TDBNavigator;
    Panel23: TPanel;
    DBGrid28: TDBGrid;
    DBNavigator20: TDBNavigator;
    Panel16: TPanel;
    Panel1: TPanel;
    rbCorrespondence: TRadioButton;
    rbRQModeVariance: TRadioButton;
    rbRQModeCorrelation: TRadioButton;
    bbCalculate: TBitBtn;
    cbPrint: TCheckBox;
    rbPCACorrelation: TRadioButton;
    rbPCAVariance: TRadioButton;
    cbPrintData: TCheckBox;
    rbDiscrim: TRadioButton;
    Panel2: TPanel;
    Panel25: TPanel;
    DBGrid3: TDBGrid;
    Panel26: TPanel;
    DBGrid1: TDBGrid;
    Panel24: TPanel;
    bbEmptyNormChem: TBitBtn;
    DBNavigator5: TDBNavigator;
    Panel27: TPanel;
    DBNavigator6: TDBNavigator;
    bbEmptyCoranMin: TBitBtn;
    Panel10: TPanel;
    BitBtn1: TBitBtn;
    lIgnoreLoadingsSmp: TLabel;
    DBNavigator1: TDBNavigator;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    ActionManager1: TActionManager;
    Importdatadefinitions1: TMenuItem;
    bbExit: TBitBtn;
    tsEigenvalues: TTabSheet;
    ChartEigenvalue: TChart;
    Series12: TLineSeries;
    Splitter3: TSplitter;
    Splitter4: TSplitter;
    Panel5: TPanel;
    Splitter5: TSplitter;
    Splitter6: TSplitter;
    Panel14: TPanel;
    Splitter7: TSplitter;
    Panel28: TPanel;
    DBGrid20: TDBGrid;
    Panel29: TPanel;
    DBNavigator17: TDBNavigator;
    Label2: TLabel;
    Splitter8: TSplitter;
    Splitter9: TSplitter;
    Splitter10: TSplitter;
    Splitter11: TSplitter;
    rbDiscrimMulti: TRadioButton;
    rbCluster: TRadioButton;
    ts3D: TTabSheet;
    Panel30: TPanel;
    Splitter12: TSplitter;
    DBChart8: TDBChart;
    Panel31: TPanel;
    DBGrid29: TDBGrid;
    DBNavigator21: TDBNavigator;
    Panel32: TPanel;
    DBGrid30: TDBGrid;
    DBNavigator22: TDBNavigator;
    DBGrid31: TDBGrid;
    DBNavigator23: TDBNavigator;
    Series13: TPoint3DSeries;
    Splitter13: TSplitter;
    Panel33: TPanel;
    udRotation: TUpDown;
    Panel34: TPanel;
    udElevation: TUpDown;
    udPerspective: TUpDown;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    eRotation: TEdit;
    eElevation: TEdit;
    ePerspective: TEdit;
    Series14: TPoint3DSeries;
    Series15: TPoint3DSeries;
    udZoom: TUpDown;
    Label6: TLabel;
    eZoom: TEdit;
    tsSummary: TTabSheet;
    Panel35: TPanel;
    Panel36: TPanel;
    DBGridStats: TDBGrid;
    Panel37: TPanel;
    DBNavigator24: TDBNavigator;
    procedure bbCalculateClick(Sender: TObject);
    procedure bbEmptyCoranVecClick(Sender: TObject);
    procedure Import1Click(Sender: TObject);
    procedure bbEmptyCoranChemClick(Sender: TObject);
    procedure About1Click(Sender: TObject);
    procedure Printersetup1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DataLinkFile1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure bbSaveVarClick(Sender: TObject);
    procedure bbSaveSmpClick(Sender: TObject);
    procedure dbnGroupedSmpClick(Sender: TObject; Button: TNavigateBtn);
    procedure dbnGroupsClick(Sender: TObject; Button: TNavigateBtn);
    procedure dbnVarClick(Sender: TObject; Button: TNavigateBtn);
    procedure PrintGraph1Click(Sender: TObject);
    procedure bbEmptyCoranVecVarClick(Sender: TObject);
    procedure Importdatadefinitions1Click(Sender: TObject);
    procedure bbExitClick(Sender: TObject);
    procedure udRotationClick(Sender: TObject; Button: TUDBtnType);
    procedure udElevationClick(Sender: TObject; Button: TUDBtnType);
    procedure udPerspectiveClick(Sender: TObject; Button: TUDBtnType);
    procedure pc1Change(Sender: TObject);
    procedure udZoomClick(Sender: TObject; Button: TUDBtnType);
  private
    { Private declarations }
    procedure DoProcess;
    procedure CalcComponentScores;
    procedure PrepareGraphs;
    procedure Discrm;
    procedure CalcDiscrmScores;
    procedure DiscrimMulti;
    procedure Cluster;
  public
    { Public declarations }
    procedure GetIniFile;
    procedure SetIniFile;
  end;

var
  fmCoranMain: TfmCoranMain;

implementation

uses cor_varb, Cor_ShtIm, About, Cor_dm_acs, Matrix, AllSorts, Cor_ShtIm2;

{$R *.DFM}
{$D+}
var
  ImportForm : TfmSheetImport;
  ImportForm2 : TfmSheetImport2;

procedure TfmCoranMain.bbCalculateClick(Sender: TObject);
begin
  TooMuch := false;
   if rbCorrespondence.Checked then SimilarityChoice := 1;
   if rbRQModeVariance.Checked then SimilarityChoice := 2;
   if rbRQModeCorrelation.Checked then SimilarityChoice := 3;
   if rbPCAVariance.Checked then SimilarityChoice := 4;
   if rbPCACorrelation.Checked then SimilarityChoice := 5;
   if rbDiscrim.Checked then SimilarityChoice := 6;
   if rbDiscrimMulti.Checked then SimilarityChoice := 7;
   if rbCluster.Checked then SimilarityChoice := 8;
   try
     dmCor.QGroups.Open;
     dmCor.QPlotGroups.Open;
   except
   end;
   DBChart7.Visible := true;
   if (SimilarityChoice < 6) then DoProcess;
   if (SimilarityChoice = 6) then Discrm;
   if (SimilarityChoice = 7) then DiscrimMulti;
   if (SimilarityChoice = 8) then Cluster;
   dmCor.FacLoadingsSmp.First;
   dmCor.CoranChem.First;
   sbMain.Panels[1].Text := 'Completed';
end;

procedure TfmCoranMain.DoProcess;
var
  ii, jj : integer;
  I, J, K, KP : integer;
  tmpStr : string;
  blankstr : string;
begin
  blankstr := '   ';
  MemoResults.Clear;
  MemoResults.Font.Name := 'Courier';
  MemoResults.Font.Pitch := fpFixed;
  MemoResults.Font.Style := [fsBold];
  MemoResults.Font.Size := 12;
  MemoResults.Lines.Add(eTitle.Text+'                '+DateToStr(Now));
  MemoResults.Font.Style := [];
  MemoResults.Font.Size := 10;
  MemoResults.Lines.Add('  ');
  MemoResults.Lines.Add('  ');
  if (cbPrint.Checked=true) then IPrn := 'Y'
                            else IPrn := 'N';
  if IPrn = 'Y' then
  begin
    AssignPrn(Lst);
    Rewrite(Lst);
    Printer.Canvas.Font.Name := 'Courier';
    Printer.Canvas.Font.Style := [fsBold];
    Printer.Canvas.Font.Size := 7;
    Writeln(Lst,' ');
    Writeln(Lst,'');
    Writeln(Lst,'');
    Write(Lst,' ':10,eTitle.Text);
    Writeln(Lst,'  ');
    Printer.Canvas.Font.Style := [];
  end;
  Nox := dmCor.ElemNames.RecordCount;
  for ii := Nox+3 to MM+2 do
  begin
    DBGrid1.Columns[ii].Visible := false;
    DBGrid2.Columns[ii].Visible := false;
  end;
  for ii := Nox+2 to MM+1 do
  begin
    DBGrid5.Columns[ii].Visible := false;
    DBGridStats.Columns[ii].Visible := false;
  end;
  for ii := Nox+3 to MM+2 do
  begin
    DBGrid4.Columns[ii].Visible := false;
  end;
  for ii := Nox+1 to MM do
  begin
    DBGrid20.Columns[ii].Visible := false;
  end;
  I:=0;
  N:=1;
  TooMuch:=false;
  TotalRecs := dmCor.CoranChem.RecordCount;
  sbmain.Panels[1].Text :='Reading data for '+IntToStr(Nox)+' elements';
  sbMain.Refresh;
  dmCor.Coranchem.First;
  repeat
    if ((I<(NN-10))) then
    begin
      I:=I+1;
      Component[I]:=dmCor.CoranChemSampleNum.AsString;
      for J:=1 to Nox do begin
        case J of
          1 : X[I,J] := dmCor.coranchemParam1.AsFloat;
          2 : X[I,J] := dmCor.coranchemParam2.AsFloat;
          3 : X[I,J] := dmCor.coranchemParam3.AsFloat;
          4 : X[I,J] := dmCor.coranchemParam4.AsFloat;
          5 : X[I,J] := dmCor.coranchemParam5.AsFloat;
          6 : X[I,J] := dmCor.coranchemParam6.AsFloat;
          7 : X[I,J] := dmCor.coranchemParam7.AsFloat;
          8 : X[I,J] := dmCor.coranchemParam8.AsFloat;
          9 : X[I,J] := dmCor.coranchemParam9.AsFloat;
          10 : X[I,J] := dmCor.coranchemParam10.AsFloat;
          11 : X[I,J] := dmCor.coranchemParam11.AsFloat;
          12 : X[I,J] := dmCor.coranchemParam12.AsFloat;
          13 : X[I,J] := dmCor.coranchemParam13.AsFloat;
          14 : X[I,J] := dmCor.coranchemParam14.AsFloat;
          15 : X[I,J] := dmCor.coranchemParam15.AsFloat;
          16 : X[I,J] := dmCor.coranchemParam16.AsFloat;
          17 : X[I,J] := dmCor.coranchemParam17.AsFloat;
          18 : X[I,J] := dmCor.coranchemParam18.AsFloat;
          19 : X[I,J] := dmCor.coranchemParam19.AsFloat;
          20 : X[I,J] := dmCor.coranchemParam20.AsFloat;
          21 : X[I,J] := dmCor.coranchemParam21.AsFloat;
          22 : X[I,J] := dmCor.coranchemParam22.AsFloat;
          23 : X[I,J] := dmCor.coranchemParam23.AsFloat;
          24 : X[I,J] := dmCor.coranchemParam24.AsFloat;
          25 : X[I,J] := dmCor.coranchemParam25.AsFloat;
          26 : X[I,J] := dmCor.coranchemParam26.AsFloat;
          27 : X[I,J] := dmCor.coranchemParam27.AsFloat;
          28 : X[I,J] := dmCor.coranchemParam28.AsFloat;
          29 : X[I,J] := dmCor.coranchemParam29.AsFloat;
          30 : X[I,J] := dmCor.coranchemParam30.AsFloat;
        end;
      end;
    end;
    if ((I>=(NN-10))) then TooMuch:=true;
    sbmain.Panels[1].Text :='Processing record '+IntToStr(N)+' of '+IntToStr(TotalRecs)+' Total included '+IntToStr(I);
    sbMain.Refresh;
    N:=N+1;
    dmCor.coranchem.Next;
  until ((N>TotalRecs) or (I>=NN-10) or dmCor.Coranchem.eof);
  if TooMuch then
  begin
    MessageDlg('Data overflow. Truncating!!',mtWarning,[mbOK],0);
  end;
  N:=I;
  M:=Nox;
  for I:=1 to Nox-1 do begin
    VarbSymbol[I]:=48+I;
  end;
  VarbSymbol[Nox]:=48;
  if (N<M) then begin
    M:=N;
    MessageDlg('Insufficient data in file. Decreasing variables',mtWarning,[mbOK],0);
  end;
  {
  MemoResults.Lines.Add('Input variables are:  ');
  }
  i := 1;
  dmCor.ElemNames.First;
  repeat
    if (dmCor.ElemNamesPos.AsInteger > 0) then
    begin
      OxideName[dmCor.ElemNamesPos.AsInteger] := dmCor.ElemNamesCalled.AsString;
      {
      MemoResults.Lines.Add(dmCor.ElemNamesPos.AsString+'   '+OxideName[dmCor.ElemNamesPos.AsInteger]);
      }
    end;
    i := i + 1;
    dmCor.ElemNames.Next;
  until ((dmCor.ElemNames.Eof) or (i > M));
  dmCor.ElemNames.First;
  {
  MemoResults.Lines.Add('  ');
  MemoResults.Lines.Add('  ');
  }
  for I:=1 to N do
  begin
    for J:=1 to M do
    begin
      if (X[I,J]=0.0) then X[I,J]:=0.001;
      X1[i,j] := X[i,j];
    end;
  end;
  if (IPrn='Y') then
  begin
    Write(Lst,'                                                                      CORAN version '+CoranVersion);
    Writeln(Lst,'       '+'Printed on '+DateToStr(Now));
    Writeln(Lst);
    Writeln(Lst,'    ',eTitle.Text);
    Writeln(Lst);
    Writeln(Lst);
    Writeln(Lst,'Data Matrix');
    if cbPrintData.Checked then PrintM(Component,X,N,M);
  end;
  for I:=1 to M do
  begin
    Str(I:10,tempstr);
    EigName[I]:=tempstr;
  end;
  sbmain.Panels[1].Text :='Calculating similarity matrix';
  sbMain.Refresh;
  case SimilarityChoice of
    1 : begin
      ScaleCorAnal;
    end;
    2 : begin
      ScaleRQAnalVar;
    end;
    3 : begin
      ScaleRQAnalStd;
    end;
    4 : begin
      Cov(X,A3,N,M);
    end;
    5 : begin
      Stand(X,N,M);
      RCoef(X,A3,N,M);
    end;
  end;
  case SimilarityChoice of
    1,2,3 : begin
      for i:= 1 to N do
      begin
        for j:= 1 to M do
        begin
          X1[i,j] := W[i,j];
        end;
      end;
      Transp(W,WP,N,M);
      {derive similarity matrix}
      for I:=1 to M do
      begin
        for J:=1 to M do
        begin
          A3[I,J]:=0.0;
          for K:=1 to N do
          begin
            A3[I,J]:=A3[I,J]+WP[I,K]*W[K,J];
          end;
        end;
      end;
    end;
    4,5 : begin
    end;
  end;
  if (IPrn='Y') then
  begin
    Writeln(Lst);
    Writeln(Lst);
    case SimilarityChoice of
      1 : Writeln(Lst,'Similarity Matrix - Correspondence Analysis');
      2 : Writeln(Lst,'Similarity Matrix - R- Q-mode variance');
      3 : Writeln(Lst,'Similarity Matrix - R- Q-mode correlation');
      4 : Writeln(Lst,'Similarity Matrix - Principal components - covariance');
      5 : Writeln(Lst,'Similarity Matrix - Principal components - correlation');
    end;
    PrintMC(OxideName,A3,M,M);
  end;
  MemoResults.Lines.Add(' ');
  case SimilarityChoice of
    1 : Memoresults.Lines.Add('Similarity Matrix - Correspondence Analysis');
    2 : Memoresults.Lines.Add('Similarity Matrix - R- Q-mode variance');
    3 : Memoresults.Lines.Add('Similarity Matrix - R- Q-mode correlation');
    4 : Memoresults.Lines.Add('Similarity Matrix - Principal components - covariance');
    5 : Memoresults.Lines.Add('Similarity Matrix - Principal components - correlation');
  end;
  Memoresults.Lines.Add(' ');
  {try writing all in separate columns}
  sbmain.Panels[1].Text :='after case similarity matrix';
  sbMain.Refresh;
  for i := 1 to M do
  begin
    tmpstr := '';
    tmpstr := tmpstr + FormatFloat('00',i);
    ResultsArray[i,1] := tmpstr + '   ';
    tmpstr := OxideName[i];
    ResultsArray[i,2] := tmpstr + CharStream(10-Length(tmpstr),32) + blankstr;
    tmpStr := TakeLogs[i];
    ResultsArray[i,3] := ' ' + tmpstr + '  ';
    for J:=1 to M do
    begin
      if (A3[i,J] <= 9000000.0) then tmpStr := FormatFloat('###0.000',A3[i,J]);
      if (A3[i,J] > 9000000.0) then tmpStr := FormatFloat('###0.00',A3[i,J]);
      ResultsArray[i,j+3] := CharStream(12-Length(tmpstr),32)+tmpstr + blankstr;
    end;
  end;
  tmpstr := '                     ';
  for i := 1 to M do
  begin
    tmpstr := tmpstr + FormatFloat('    00      ',i);
  end;
  MemoResults.Lines.Add(tmpstr);
  Memoresults.Lines.Add(' ');
  for i := 1 to M do
  begin
    tmpstr := '';
    for j := 1 to M+3 do
    begin
      tmpstr := tmpstr + ResultsArray[i,j];
    end;
    MemoResults.Lines.Add(tmpstr);
  end;
  MemoResults.Lines.Add('  ');
  MemoResults.Lines.Add('  ');
  {
  for I:=1 to M do
  begin
    tmpStr := OxideName[I]+'    ';
    for J:=1 to M do
    begin
      tmpStr := tmpStr + FormatFloat('######0.0000',A3[I,J])+'    ';
    end;
    MemoResults.Lines.Add(tmpStr);
  end;
  Memoresults.Lines.Add(' ');
  Memoresults.Lines.Add(' ');
  }
  sbmain.Panels[1].Text :='Calculating eigenvectors and eigenvalues';
  sbMain.Refresh;
  EigenJ(A3,A2,M);
  A1[1,1]:=A3[1,1];
  if (SimilarityChoice=1) then A3[1,1]:=0.0;
  SumE:=0.0;
  for I:=1 to M do
  begin
    A1[I,1]:=A3[I,I];
    SumE:=SumE+Abs(A1[I,1]);
  end;
  A1[1,2]:=0.0;
  A1[1,3]:=0.0;
  SumEE:=0.0;
  ChartEigenvalue.Series[0].Clear;
  ChartEigenvalue.Series[0].XValues.Order := loNone;
  ChartEigenvalue.Series[0].YValues.Order := loNone;
  for I:=1 to M do
  begin
    SumEE:=SumEE+Abs(A1[I,1]);
    A1[I,2]:=Abs(A1[I,1])*100.0/SumE;
    A1[I,3]:=SumEE*100.0/SumE;
    if not ((SimilarityChoice=1) and (I=1)) then
    begin
      ChartEigenvalue.Series[0].AddXY(1.0*I,A1[I,2]);
    end;
  end;
  if (SimilarityChoice=1) then
  begin
    A3[1,1]:=1.0;
    A1[1,1]:=1.0;
  end;
  if (SimilarityChoice=1) then
  begin
    Memoresults.Lines.Add('Eigenvalue 1 and Eigenvector 1 are artifices of the method. Ignore them!');
    Memoresults.Lines.Add(' ');
  end;
  {
  Memoresults.Lines.Add('Column 1 = Eigenvalues,  Column 2 = % of trace,  Column 3 = cum. % of trace');
  Memoresults.Lines.Add(' ');
  }
  MemoResults.Lines.Add(' ');
  MemoResults.Lines.Add('Component   Eigen       % of      Cumulative');
  MemoResults.Lines.Add('            value       trace     % of trace');
  MemoResults.Lines.Add(' ');
  {try writing all in separate columns}
  for i := 1 to M do
  begin
    tmpstr := '';
    tmpstr := FormatFloat('    00',i);
    ResultsArray[i,1] := tmpstr + CharStream(6-Length(tmpstr),32) + blankstr;
    for J:=1 to 3 do
    begin
      if (A1[i,j] <= 9000000.0) then tmpStr := FormatFloat('###0.0000',A1[i,j]);
      if (A1[i,j] > 9000000.0) then tmpStr := FormatFloat('###0.00',A1[i,j]);
      ResultsArray[i,j+1] := CharStream(12-Length(tmpstr),32)+tmpstr + blankstr;
    end;
  end;
  for i := 1 to M do
  begin
    tmpstr := '';
    for j := 1 to 4 do
    begin
      tmpstr := tmpstr + ResultsArray[i,j];
    end;
    MemoResults.Lines.Add(tmpstr);
  end;
  MemoResults.Lines.Add('  ');
  MemoResults.Lines.Add('  ');
  {
  for I:=1 to M do
  begin
    tmpStr := EigName[I]+'    ';
    for J:=1 to 3 do
    begin
      tmpStr := tmpStr + FormatFloat('######0.0000',A1[I,J])+'    ';
    end;
    MemoResults.Lines.Add(tmpStr);
  end;
  }
  if (IPrn='Y') then
  begin
    Writeln(Lst);
    if (SimilarityChoice=1) then
    begin
      Writeln(Lst,'Eigenvalue 1 and Eigenvector 1 are artifices of the method. Ignore them!');
      Writeln(Lst);
    end;
    Writeln(Lst,'Column 1 = Eigenvalues,  Column 2 = % of trace,  Column 3 = cum. % of trace');
    PrintMC(EigName,A1,M,3);
    Writeln(Lst);
    Writeln(Lst);
    Writeln(Lst,'Principal axis matrix  -  columns = eigenvectors, rows = variables');
    PrintMC(OxideName,A2,M,M);
  end;
  dmCor.EigenVec.Open;
  dmCor.EigenVec.Last;
  if not (dmCor.EigenVec.Bof and dmCor.EigenVec.Eof) then
  begin
    dmCor.EigenVec.Last;
    repeat
     dmCor.EigenVec.Delete;
    until dmCor.EigenVec.Bof;
  end;
  for ii := 1 to Nox do
  begin
    dmCor.EigenVec.Append;
    dmCor.EigenVecPos.AsInteger := ii;
    for jj := 1 to Nox do
    begin
      case jj of
        1 : dmCor.EigenVecVector1.AsFloat := A2[ii,jj];
        2 : dmCor.EigenVecVector2.AsFloat := A2[ii,jj];
        3 : dmCor.EigenVecVector3.AsFloat := A2[ii,jj];
        4 : dmCor.EigenVecVector4.AsFloat := A2[ii,jj];
        5 : dmCor.EigenVecVector5.AsFloat := A2[ii,jj];
        6 : dmCor.EigenVecVector6.AsFloat := A2[ii,jj];
        7 : dmCor.EigenVecVector7.AsFloat := A2[ii,jj];
        8 : dmCor.EigenVecVector8.AsFloat := A2[ii,jj];
        9 : dmCor.EigenVecVector9.AsFloat := A2[ii,jj];
        10 : dmCor.EigenVecVector10.AsFloat := A2[ii,jj];
        11 : dmCor.EigenVecVector11.AsFloat := A2[ii,jj];
        12 : dmCor.EigenVecVector12.AsFloat := A2[ii,jj];
        13 : dmCor.EigenVecVector13.AsFloat := A2[ii,jj];
        14 : dmCor.EigenVecVector14.AsFloat := A2[ii,jj];
        15 : dmCor.EigenVecVector15.AsFloat := A2[ii,jj];
        16 : dmCor.EigenVecVector16.AsFloat := A2[ii,jj];
        17 : dmCor.EigenVecVector17.AsFloat := A2[ii,jj];
        18 : dmCor.EigenVecVector18.AsFloat := A2[ii,jj];
        19 : dmCor.EigenVecVector19.AsFloat := A2[ii,jj];
        20 : dmCor.EigenVecVector20.AsFloat := A2[ii,jj];
        21 : dmCor.EigenVecVector21.AsFloat := A2[ii,jj];
        22 : dmCor.EigenVecVector22.AsFloat := A2[ii,jj];
        23 : dmCor.EigenVecVector23.AsFloat := A2[ii,jj];
        24 : dmCor.EigenVecVector24.AsFloat := A2[ii,jj];
        25 : dmCor.EigenVecVector25.AsFloat := A2[ii,jj];
        26 : dmCor.EigenVecVector26.AsFloat := A2[ii,jj];
        27 : dmCor.EigenVecVector27.AsFloat := A2[ii,jj];
        28 : dmCor.EigenVecVector28.AsFloat := A2[ii,jj];
        29 : dmCor.EigenVecVector29.AsFloat := A2[ii,jj];
        30 : dmCor.EigenVecVector30.AsFloat := A2[ii,jj];
      end;
    end;
    dmCor.EigenVec.Post;
  end;
  sbmain.Panels[1].Text :='Calculating factor loadings';
  sbMain.Refresh;
  case SimilarityChoice of
    1 : begin
      for J:=1 to M do
      begin
        A3[J,J]:=Sqrt(Abs(A3[J,J]));
      end;
      MmultC(A2,A3,TempC,M,M,M);  {factor loadings for variables}
      for J:=1 to M do
      begin
        A3[J,J]:=1.0/A3[J,J];
      end;
      MmultC(DC,TempC,A1,M,M,M);  {scaled factor loadings for variables}
      Mmult(W,A2,B,N,M,M);
      Mmult(B,A3,W,N,M,M);
      MmultR(DR,W,TempM,N,M);     {factor loadings for samples}
      for J:=1 to M do
      begin
        A3[J,J]:=1.0/A3[J,J];
      end;
      Mmult(TempM,A3,B,N,M,M);    {scaled factor loadings for samples}
      if (IPrn='Y') then
      begin
        Writeln(Lst,'Variable factors used for loadings -  columns = factors, rows = variables');
      end;
      (*
      MInv(X,TempM,N,M);
      for I:=1 to M do
      begin
        for J:=1 to N do
        begin
          W[I,J]:=0.0;
          for K:=1 to M do
          begin
            W[I,J]:=W[I,J]+TempM[I,K]*B[K,J];
          end;
        end;
      end;
      if (IPrn='Y') then
      begin
        PrintMX(OxideName,W,M,M);
      end;
      MmultX(X,W,B,N,M,M);
      if (IPrn='Y') then
      begin
        Writeln(Lst,'Factor loadings  -  columns = factors, rows = samples');
        if cbPrintData.Checked then PrintM(Component,B,N,M);
        Writeln(Lst);
      end;
      *)
    end;
    2,3 : begin
      for J:=1 to M do
      begin
        A3[J,J]:=Sqrt(Abs(A3[J,J]));
      end;
      MmultC(A2,A3,A1,M,M,M);  {factor loadings for variables}
      Mmult(W,A2,B,N,M,M);     {factor loadings for samples}
    end;
    4,5 : begin
      Mmult(X,A2,B,N,M,M);     {factor scores for samples}
    end;
  end;
  CalcComponentScores;
  i := 1;
  dmCor.QGroups.Open;
  dmCor.QPlotGroups.Open;
  if (IPrn='Y') then
  begin
    if (SimilarityChoice in [1..3]) then
    begin
      Writeln(Lst,'Factor loadings  -  columns = factors, rows = variables');
      PrintMC(OxideName,A1,M,M);
    end;
    Writeln(Lst,'Factor loadings  -  columns = factors, rows = samples');
    if cbPrintData.Checked then PrintM(Component,B,N,M);
  end;
  if ((SimilarityChoice=1) and (M>4)) then
  begin
    for I:=1 to N do
    begin
     DR[I]:=1.0/DR[I];
     DR[I]:=DR[I]*DR[I];
    end;
    for J:=1 to M do
    begin
     DC[J,J]:=1.0/DC[J,J];
     DC[J,J]:=DC[J,J]*DC[J,J];
     A3[J,J]:=A3[J,J]*A3[J,J];
    end;
    if (IPrn='Y') then
    begin
     Writeln(Lst,'Absolute and relative contributions for variables');
     Writeln(Lst);
     Write(Lst,'                                  2         ');
     Write(Lst,'            3         ');
     Write(Lst,'            4         ');
     Write(Lst,'            5         ');
     Writeln(Lst);
     Write(Lst,'                Weight      Abs       Rel   ');
     Write(Lst,'      Abs       Rel   ');
     Write(Lst,'      Abs       Rel   ');
     Write(Lst,'      Abs       Rel   ');
     Writeln(Lst);
     for J:=1 to M do
     begin
       D:=0.0;
       for K:=2 to M do
       begin
         D:=D+A1[J,K]*A1[J,K];
       end;
       for K:=2 to 5 do
       begin
         C:=100.0*A1[J,K]*A1[J,K];
         CRR[K-1]:=C/D;
         CA[K-1]:=DC[J,J]*C/A3[K,K];
       end;
       if (IPrn='Y') then
       begin
         Write(Lst,OxideName[J]:11,DC[J,J]:11:4);
         for K:=1 to 4 do
         begin
           Write(Lst,CA[K]:11:4,CRR[K]:11:4);
         end;
         Writeln(Lst);
       end;
     end;
     if (IPrn='Y') then
     begin
       Writeln(Lst);
       Writeln(Lst,'Absolute and relative contributions for samples');
       Writeln(Lst);
       Write(Lst,'                                  2         ');
       Write(Lst,'            3         ');
       Write(Lst,'            4         ');
       Write(Lst,'            5         ');
       Writeln(Lst);
       Write(Lst,'                Weight      Abs       Rel   ');
       Write(Lst,'      Abs       Rel   ');
       Write(Lst,'      Abs       Rel   ');
       Write(Lst,'      Abs       Rel   ');
       Writeln(Lst);
     end;
     for I:=1 to N do
     begin
       D:=0.0;
       for K:=2 to M do
       begin
         D:=D+B[I,K]*B[I,K];
       end;
       for K:=2 to 5 do
       begin
         C:=100.0*B[I,K]*B[I,K];
         CRR[K-1]:=C/D;
         CA[K-1]:=DR[I]*C/A3[K,K];
       end;
       if (IPrn='Y') then
       begin
         Write(Lst,Component[I]:11,DR[I]:11:4);
         for K:=1 to 4 do
         begin
           Write(Lst,CA[K]:11:4,CRR[K]:11:4);
         end;
         Writeln(Lst);
       end;
     end;
    end;
  end;
  pVar.Visible := true;
  pGraph1Var.Visible := true;
  pGraph2Var.Visible := true;
  pGraph3Var.Visible := true;
  case SimilarityChoice of
    1 : begin
       lIgnoreLoadingsVar.Visible := true;
       lIgnoreLoadingsSmp.Visible := true;
       NX:=2;
    end;
    2..3 : begin
       lIgnoreLoadingsVar.Visible := false;
       lIgnoreLoadingsSmp.Visible := false;
       NX:=1;
    end;
    4..5 : begin
       lIgnoreLoadingsVar.Visible := false;
       lIgnoreLoadingsSmp.Visible := false;
       pVar.Visible := false;
       pGraph1Var.Visible := false;
       pGraph2Var.Visible := false;
       pGraph3Var.Visible := false;
       NX:=1;
    end;
  end;
  case SimilarityChoice of
    1..3 : begin
       pVar.Visible := true;
       Xmin:=A1[1,NX];
       Xmax:=Xmin;
       Ymin:=A1[1,NX+1];
       Ymax:=Ymin;
       Zmin:=A1[1,NX+2];
       Zmax:=Zmin;
       for J:=1 to M do
       begin
         if (Xmin > A1[J,NX]) then Xmin:=A1[J,NX];
         if (Xmax < A1[J,NX]) then Xmax:=A1[J,NX];
         if (Ymin > A1[J,NX+1]) then Ymin:=A1[J,NX+1];
         if (Ymax < A1[J,NX+1]) then Ymax:=A1[J,NX+1];
         if (Zmin > A1[J,NX+2]) then Zmin:=A1[J,NX+2];
         if (Zmax < A1[J,NX+2]) then Zmax:=A1[J,NX+2];
         for K:=1 to 5 do
         begin
           X[J,K]:=A1[J,K];
         end;
       end;
    end;
    4,5 : begin
       pVar.Visible := false;
       Xmin:=B[1,NX];
       Xmax:=Xmin;
       Ymin:=B[1,NX+1];
       Ymax:=Ymin;
       Zmin:=B[1,NX+2];
       Zmax:=Zmin;
    end;
  end;
  for I:=1 to N do
  begin
    if (Xmin > B[I,NX]) then Xmin:=B[I,NX];
    if (Xmax < B[I,NX]) then Xmax:=B[I,NX];
    if (Ymin > B[I,NX+1]) then Ymin:=B[I,NX+1];
    if (Ymax < B[I,NX+1]) then Ymax:=B[I,NX+1];
    if (Zmin > B[I,NX+2]) then Zmin:=B[I,NX+2];
    if (Zmax < B[I,NX+2]) then Zmax:=B[I,NX+2];
    for K:=1 to 5 do
    begin
      X[I+M,K]:=B[I,K];
    end;
  end;
  {
  PlotValRec[1].XMini:=XMin;
  PlotValRec[1].XMaxi:=XMax;
  PlotValRec[2].XMini:=XMin;
  PlotValRec[2].XMaxi:=XMax;
  PlotValRec[3].XMini:=YMin;
  PlotValRec[3].XMaxi:=YMax;
  PlotValRec[1].YMini:=YMin;
  PlotValRec[1].YMaxi:=YMax;
  PlotValRec[2].YMini:=ZMin;
  PlotValRec[2].YMaxi:=ZMax;
  PlotValRec[3].YMini:=ZMin;
  PlotValRec[3].YMaxi:=ZMax;
  for KP:=1 to 3 do
  begin
    PlotValRec[KP].XTics:=(PlotValRec[KP].XMaxi-PlotValRec[KP].XMini)/10.0;
    PlotValRec[KP].YTics:=(PlotValRec[KP].YMaxi-PlotValRec[KP].YMini)/10.0;
  end;
  }
  tsGraph1.TabVisible := false;
  tsGraph2.TabVisible := false;
  tsGraph3.TabVisible := false;
  tsLocalities.TabVisible := false;
  tsSpreadsheets.TabVisible := false;
  tsScores.TabVisible := false;
  pVar.Visible := true;
  pGraph1Var.Visible := true;
  case SimilarityChoice of
    1 : begin
      tsGraph1.TabVisible := true;
      tsGraph2.TabVisible := true;
      tsGraph3.TabVisible := true;
      tsLocalities.TabVisible := true;
      tsSpreadsheets.TabVisible := true;
      tsScores.TabVisible := true;
      pSaveVar.Visible := true;
      PrepareGraphs;
      {
      PlotCorAnal;
      }
    end;
    2..3 : begin
      tsGraph1.TabVisible := true;
      tsGraph2.TabVisible := true;
      tsGraph3.TabVisible := true;
      tsScores.TabVisible := true;
      tsLocalities.TabVisible := true;
      tsSpreadsheets.TabVisible := true;
      pSaveVar.Visible := true;
      PrepareGraphs;
      {
      PlotRQAnal;
      }
    end;
    4..5 : begin
      tsGraph1.TabVisible := true;
      tsGraph2.TabVisible := true;
      tsGraph3.TabVisible := true;
      tsScores.TabVisible := false;
      tsLocalities.TabVisible := true;
      tsSpreadsheets.TabVisible := true;
      pSaveVar.Visible := false;
      pVar.Visible := false;
      pGraph1Var.Visible := false;
      PrepareGraphs;
      {
      PlotRQAnal;
      }
    end;
  end;
  if (IPrn='Y') then
  begin
   Writeln(Lst,char(13));
   System.CloseFile(Lst);
  end;
end;

procedure TfmCoranMain.CalcComponentScores;
var
  I, J, K, KP : integer;
begin
  {clear contents of spreadsheets}
  F1Var.ClearRange(1,1,50,50,F1ClearAll);
  F1Smp.ClearRange(1,1,3100,50,F1ClearAll);
  {clear factor loadings table for variables}
  dmCor.FacLoadingsVar.Last;
  if not (dmCor.FacLoadingsVar.Bof and dmCor.FacLoadingsVar.Eof) then
  begin
    dmCor.FacLoadingsVar.Last;
    repeat
      dmCor.FacLoadingsVar.Delete;
    until dmCor.FacLoadingsVar.Bof;
  end;
  {clear factor loadings table for samples}
  dmCor.FacLoadingsSmp.Last;
  if not (dmCor.FacLoadingsSmp.Bof and dmCor.FacLoadingsSmp.Eof) then
  begin
    dmCor.FacLoadingsSmp.Last;
    repeat
      dmCor.FacLoadingsSmp.Delete;
    until dmCor.FacLoadingsSmp.Bof;
  end;
  if (SimilarityChoice < 4) then
  begin
    {fill factor loadings table for variables}
    dmCor.FacLoadingsVar.First;
    F1Var.Row := 1;
    F1Var.Col := 1;
    F1Var.Text := 'Variable';
    for j := 1 to M do
    begin
      F1Var.Col := j+1;
      F1Var.text := 'Component ' + IntToStr(j);
    end;
    ii := 0;
    for i := 1 to M do
    begin
      ii := ii + 1;
      dmCor.FacLoadingsVar.Append;
      dmCor.FacLoadingsVarPos.AsInteger := i;
      dmCor.FacLoadingsVarCalled.AsString := OxideName[i];
      F1Var.Row := ii+1;
      F1Var.Col := 1;
      F1Var.Text := OxideName[i];
      for j := 1 to M do
      begin
        F1Var.Col := j+1;
        case j of
          1 : dmCor.FacLoadingsVarVector1.AsFloat := A1[i,j];
          2 : dmCor.FacLoadingsVarVector2.AsFloat := A1[i,j];
          3 : dmCor.FacLoadingsVarVector3.AsFloat := A1[i,j];
          4 : dmCor.FacLoadingsVarVector4.AsFloat := A1[i,j];
          5 : dmCor.FacLoadingsVarVector5.AsFloat := A1[i,j];
          6 : dmCor.FacLoadingsVarVector6.AsFloat := A1[i,j];
          7 : dmCor.FacLoadingsVarVector7.AsFloat := A1[i,j];
          8 : dmCor.FacLoadingsVarVector8.AsFloat := A1[i,j];
          9 : dmCor.FacLoadingsVarVector9.AsFloat := A1[i,j];
          10 : dmCor.FacLoadingsVarVector10.AsFloat := A1[i,j];
          11 : dmCor.FacLoadingsVarVector11.AsFloat := A1[i,j];
          12 : dmCor.FacLoadingsVarVector12.AsFloat := A1[i,j];
          13 : dmCor.FacLoadingsVarVector13.AsFloat := A1[i,j];
          14 : dmCor.FacLoadingsVarVector14.AsFloat := A1[i,j];
          15 : dmCor.FacLoadingsVarVector15.AsFloat := A1[i,j];
          16 : dmCor.FacLoadingsVarVector16.AsFloat := A1[i,j];
          17 : dmCor.FacLoadingsVarVector17.AsFloat := A1[i,j];
          18 : dmCor.FacLoadingsVarVector18.AsFloat := A1[i,j];
          19 : dmCor.FacLoadingsVarVector19.AsFloat := A1[i,j];
          20 : dmCor.FacLoadingsVarVector20.AsFloat := A1[i,j];
          21 : dmCor.FacLoadingsVarVector21.AsFloat := A1[i,j];
          22 : dmCor.FacLoadingsVarVector22.AsFloat := A1[i,j];
          23 : dmCor.FacLoadingsVarVector23.AsFloat := A1[i,j];
          24 : dmCor.FacLoadingsVarVector24.AsFloat := A1[i,j];
          25 : dmCor.FacLoadingsVarVector25.AsFloat := A1[i,j];
          26 : dmCor.FacLoadingsVarVector26.AsFloat := A1[i,j];
          27 : dmCor.FacLoadingsVarVector27.AsFloat := A1[i,j];
          28 : dmCor.FacLoadingsVarVector28.AsFloat := A1[i,j];
          29 : dmCor.FacLoadingsVarVector29.AsFloat := A1[i,j];
          30 : dmCor.FacLoadingsVarVector30.AsFloat := A1[i,j];
        end;
        F1Var.Number := A1[i,j];
      end;
      dmCor.FacLoadingsVar.Post;
      {insert a row of zero values between each
      variable in spreadsheet to define origin for Grapher plotting
      }
      ii := ii + 1;
      F1Var.Row := ii+1;
      for j := 1 to M do
      begin
        F1Var.Col := j+1;
        F1Var.Number := 0.0;
      end;
    end;
    dmCor.FacLoadingsVar.First;
  end;
  {fill factor loadings table for samples}
  F1Smp.Row := 1;
  F1Smp.Col := 1;
  F1Smp.Text := 'Group';
  F1Smp.Col := 2;
  F1Smp.Text := 'Sample';
  for j := 1 to M do
  begin
    F1Smp.Col := j+2;
    F1Smp.Text := 'Component ' + IntToStr(j);
  end;
  dmCor.CoranChem.First;
  i := 1;
  dmCor.CoranChem.First;
  repeat
    F1Smp.Row := i+1;
    dmCor.CoranVecLinked.Append;
    dmCor.CoranVecLinkedGROUPNAME.AsString := dmCor.CoranChemGroupName.AsString;
    dmCor.CoranVecLinkedSampleNum.AsString := dmCor.CoranChemSampleNum.AsString;
    F1Smp.Col := 1;
    F1Smp.Text := dmCor.CoranVecLinkedGROUPNAME.AsString;
    F1Smp.Col := 2;
    F1Smp.Text := dmCor.CoranVecLinkedSampleNum.AsString;
    for j := 1 to M do
    begin
      F1Smp.Col := j+2;
      case j of
        1 : dmCor.CoranVecLinkedVector1.AsFloat := B[i,j];
        2 : dmCor.CoranVecLinkedVector2.AsFloat := B[i,j];
        3 : dmCor.CoranVecLinkedVector3.AsFloat := B[i,j];
        4 : dmCor.CoranVecLinkedVector4.AsFloat := B[i,j];
        5 : dmCor.CoranVecLinkedVector5.AsFloat := B[i,j];
        6 : dmCor.CoranVecLinkedVector6.AsFloat := B[i,j];
        7 : dmCor.CoranVecLinkedVector7.AsFloat := B[i,j];
        8 : dmCor.CoranVecLinkedVector8.AsFloat := B[i,j];
        9 : dmCor.CoranVecLinkedVector9.AsFloat := B[i,j];
        10 : dmCor.CoranVecLinkedVector10.AsFloat := B[i,j];
        11 : dmCor.CoranVecLinkedVector11.AsFloat := B[i,j];
        12 : dmCor.CoranVecLinkedVector12.AsFloat := B[i,j];
        13 : dmCor.CoranVecLinkedVector13.AsFloat := B[i,j];
        14 : dmCor.CoranVecLinkedVector14.AsFloat := B[i,j];
        15 : dmCor.CoranVecLinkedVector15.AsFloat := B[i,j];
        16 : dmCor.CoranVecLinkedVector16.AsFloat := B[i,j];
        17 : dmCor.CoranVecLinkedVector17.AsFloat := B[i,j];
        18 : dmCor.CoranVecLinkedVector18.AsFloat := B[i,j];
        19 : dmCor.CoranVecLinkedVector19.AsFloat := B[i,j];
        20 : dmCor.CoranVecLinkedVector20.AsFloat := B[i,j];
        21 : dmCor.CoranVecLinkedVector21.AsFloat := B[i,j];
        22 : dmCor.CoranVecLinkedVector22.AsFloat := B[i,j];
        23 : dmCor.CoranVecLinkedVector23.AsFloat := B[i,j];
        24 : dmCor.CoranVecLinkedVector24.AsFloat := B[i,j];
        25 : dmCor.CoranVecLinkedVector25.AsFloat := B[i,j];
        26 : dmCor.CoranVecLinkedVector26.AsFloat := B[i,j];
        27 : dmCor.CoranVecLinkedVector27.AsFloat := B[i,j];
        28 : dmCor.CoranVecLinkedVector28.AsFloat := B[i,j];
        29 : dmCor.CoranVecLinkedVector29.AsFloat := B[i,j];
        30 : dmCor.CoranVecLinkedVector30.AsFloat := B[i,j];
      end;
      F1Smp.Number := B[i,j];
    end;
    dmCor.CoranVecLinked.Post;
    dmCor.CoranChem.Next;
    i := i + 1;
  until dmCor.CoranChem.Eof;
  dmCor.CoranChem.First;
  dmCor.FacLoadingsSmp.First;
  dmCor.FacLoadingsSmp.Close;
  dmCor.FacLoadingsSmp.Open;
end;

procedure TfmCoranMain.bbEmptyCoranVecClick(Sender: TObject);
begin
  dmCor.FacLoadingsSmp.Last;
  if not(dmCor.FacLoadingsSmp.BOF and dmCor.FacLoadingsSmp.EOF) then
  begin
    dmCor.FacLoadingsSmp.Last;
    repeat
      dmCor.FacLoadingsSmp.Delete;
    until dmCor.FacLoadingsSmp.BOF;
  end;
end;

procedure TfmCoranMain.Import1Click(Sender: TObject);
var
  i, ii : integer;
begin
  for i := 0 to MM+2 do
  begin
    DBGrid1.Columns[i].Visible := true;
    DBGrid2.Columns[i].Visible := true;
  end;
  for i := 1 to MM do
  begin
    DBGridStats.Columns[i+1].Visible := true;
  end;
  try
    try
      dmCor.ImportGroup.Open;
      dmCor.ImportGroup.First;
      dmCor.CoranFac.Open;
    except
    end;
    ImportForm := TfmSheetImport.Create(Self);
    ImportForm.OpenDialogSprdSheet.FileName := 'CoranCHEM';
    ImportForm.ShowModal;
  finally
    ImportForm.Free;
    try
      dmCor.ImportGroup.Close;
      dmCor.CoranFac.Close;
      dmCor.QGroups.Close;
      dmCor.QGroups.Open;
      dmCor.QPlotGroups.Close;
      dmCor.QPlotGroups.Open;
    except
    end;
  end;
  dmCor.CoranChem.First;
  for ii := Nox+2 to MM+1 do
  begin
    DBGridStats.Columns[ii].Visible := false;
  end;
  dmCor.ElemNames.First;
  for ii := 1 to Nox do
  begin
    DBGrid1.Columns[ii+2].Title.Caption := dmCor.ElemNamesCalled.AsString;
    DBGridStats.Columns[ii+1].Title.Caption := dmCor.ElemNamesCalled.AsString;
    dmCor.ElemNames.Next;
  end;
end;

procedure TfmCoranMain.bbEmptyCoranChemClick(Sender: TObject);
begin
  dmCor.CoranChem.Last;
  if not(dmCor.CoranChem.Bof and dmCor.CoranChem.Eof) then
  begin
    dmCor.CoranChem.Last;
    repeat
      dmCor.CoranChem.Delete;
      dmCor.CoranChem.Next;
    until dmCor.CoranChem.BOF;
  end;
  dmCor.SmpLoc.Last;
  if not(dmCor.SmpLoc.Bof and dmCor.SmpLoc.Eof) then
  begin
    dmCor.SmpLoc.Last;
    repeat
      dmCor.SmpLoc.Delete;
      dmCor.SmpLoc.Next;
    until dmCor.SmpLoc.BOF;
  end;
  dmCor.FacLoadingsVar.Last;
  if not(dmCor.FacLoadingsVar.Bof and dmCor.FacLoadingsVar.Eof) then
  begin
    dmCor.FacLoadingsVar.Last;
    repeat
      dmCor.FacLoadingsVar.Delete;
      dmCor.FacLoadingsVar.Next;
    until dmCor.FacLoadingsVar.BOF;
  end;
end;

procedure TfmCoranMain.About1Click(Sender: TObject);
begin
  AboutBox.ShowModal;
end;

procedure TfmCoranMain.Printersetup1Click(Sender: TObject);
begin
  PrinterSetupDialog1.Execute;
end;

procedure TfmCoranMain.GetIniFile;
var
  AppIni   : TIniFile;
  tmpStr   : string;
  iCode    : integer;
begin
  AppIni := TIniFile.Create('Coran.INI');
  try
    ADODataLinkFile := AppIni.ReadString('ADO','Data link file','C:\CoranData\Coran.udl');
    DataPath := AppIni.ReadString('Spreadsheets','Results path','C:\CoranData');
    ImportSpecNameColStr := AppIni.ReadString('ColumnDefinitions','ImportSpecNameColStr','A');
    PositionColStr := AppIni.ReadString('ColumnDefinitions','PositionColStr','B');
    CalledColStr := AppIni.ReadString('ColumnDefinitions','CalledColStr','C');
    ColumnColStr := AppIni.ReadString('ColumnDefinitions','ColumnColStr','D');
    TakeLogColStr := AppIni.ReadString('ColumnDefinitions','TakeLogColStr','E');
    tmpStr := AppIni.ReadString('Defaults','DefaultMinimum','1.0e-6');
    Val(tmpStr,DefaultMinimum,iCode);
    if (iCode > 0) then DefaultMinimum := 1.0e-6;
  finally
    AppIni.Free;
  end;
end;

procedure TfmCoranMain.SetIniFile;
var
  AppIni   : TIniFile;
begin
  AppIni := TIniFile.Create('Coran.INI');
  try
    AppIni.WriteString('ADO','Data link file',ADODataLinkFile);
    AppIni.WriteString('Spreadsheets','Results path',DataPath);
    AppIni.WriteString('ColumnDefinitions','ImportSpecNameColStr',ImportSpecNameColStr);
    AppIni.WriteString('ColumnDefinitions','PositionColStr',PositionColStr);
    AppIni.WriteString('ColumnDefinitions','CalledColStr',CalledColStr);
    AppIni.WriteString('ColumnDefinitions','ColumnColStr',ColumnColStr);
    AppIni.WriteString('ColumnDefinitions','TakeLogColStr',TakeLogColStr);
    AppIni.WriteString('Defaults','DefaultMinimum',FormatFloat('##0.0000e-00',DefaultMinimum));
  finally
    AppIni.Free;
  end;
end;

procedure TfmCoranMain.FormShow(Sender: TObject);
var
  ii : integer;
begin
  tsCheck.TabVisible := false;
  pc1.ActivePage := tsControl;
  DefaultMinimum := 1.0e-6;
  GetIniFile;
  with dmCor do
  begin
    try
      Coran.Connected := false;
    except
    end;
    {
    Provider=MSDASQL.1;Persist Security Info=False;User ID=admin;Mode=ReadWrite;
    Extended Properties="DSN=MS Access Database;
    DBQ=C:\CoranDataT\coran.mdb;
    DefaultDir=C:\CoranDataT;DriverId=281;FIL=MS Access;
    FILEDSN=C:\Program Files\Common Files\ODBC\Data Sources\MS Access Database (not sharable).dsn;
    MaxBufferSize=2048;PageTimeout=5;UID=admin;"
    }
    Coran.ConnectionString := 'FILE NAME='+ADODataLinkFile;
    Coran.Provider := ADODataLinkFile;
    Coran.Open('admin','');
  end;
  tsSpreadsheets.TabVisible := false;
  with dmCor do
  begin
    CoranChem.Open;
    ElemNames.Open;
    QGroups.Open;
    QPlotGroups.Open;
    GroupedSmp.Open;
    QGroupedSmp.Open;
    SmpLoc.Open;
    GroupedSmpLoc.Open;
    QGroupedSmpLoc.Open;
    CoranStats.Open;
    try
      FacLoadingsSmp.Open;
      FacLoadingsVar.Open;
      CoranVecLinked.Open;
      FacLoadingsVarLinked.Open;
    except
    end;
    ElemNames.First;
    ii := 0;
    try
      repeat
        ii :=ii + 1;
        Oxidename[ElemNamesPos.AsInteger] := ElemNamesCalled.AsString;
        TakeLogs[ElemNamesPos.AsInteger] := ElemNamesTakeLog.AsString;
        DBGrid1.Columns[ii+2].Title.Caption := dmCor.ElemNamesCalled.AsString;
        DBGridStats.Columns[ii+1].Title.Caption := dmCor.ElemNamesCalled.AsString;
        ElemNames.Next;
      until ElemNames.Eof;
    except
    end;
    Nox := ii;
  end;
  try
    PrepareGraphs;
  except
  end;
  FromRowValueString := '2';
  ToRowValueString := '2';
end;

procedure TfmCoranMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  SetIniFile;
end;

procedure TfmCoranMain.DataLinkFile1Click(Sender: TObject);
begin
  ADODataLinkFile := InputBox('ADO','Data link file ',ADODataLinkFile);
  ADODataLinkFile := UpperCase(ADODataLinkFile);
end;

procedure TfmCoranMain.PrepareGraphs;
begin
  dmCor.FacLoadingsVar.Close;
  dmCor.FacLoadingsVarLinked.Close;
  dmCor.FacLoadingsSmp.Close;
  dmCor.SmpLoc.Close;
  dmCor.QGroups.Close;
  dmCor.GroupedSmp.Close;
  dmCor.QGroupedSmp.Close;
  dmCor.GroupedSmpLoc.Close;
  dmCor.QGroupedSmpLoc.Close;
  {loadings}
  try
    DBChart5.Series[0].DataSource := dmCor.FacLoadingsVar;
    DBChart6.Series[0].DataSource := dmCor.FacLoadingsVar;
    DBChart7.Series[0].DataSource := dmCor.FacLoadingsVar;
    DBChart5.Series[0].Active := true;
    DBChart6.Series[0].Active := true;
    DBChart7.Series[0].Active := true;
    DBChart5.BottomAxis.Title.Caption := 'Component 1';
    DBChart6.BottomAxis.Title.Caption := 'Component 2';
    DBChart7.BottomAxis.Title.Caption := 'Component 3';
    DBChart5.Series[1].Active := false;
    DBChart6.Series[1].Active := false;
    DBChart7.Series[1].Active := false;
    case SimilarityChoice of
      1 : begin
        DBChart5.BottomAxis.Title.Caption := 'Component 2';
        DBChart6.BottomAxis.Title.Caption := 'Component 3';
        DBChart7.BottomAxis.Title.Caption := 'Component 4';
        DBChart5.Series[0].XValues.ValueSource := 'Vector2';
        DBChart5.Series[0].YValues.ValueSource := 'Vector2';
        DBChart6.Series[0].XValues.ValueSource := 'Vector3';
        DBChart6.Series[0].YValues.ValueSource := 'Vector3';
        DBChart7.Series[0].XValues.ValueSource := 'Vector4';
        DBChart7.Series[0].YValues.ValueSource := 'Vector4';
      end;
      2..3 : begin
        DBChart5.Series[0].XValues.ValueSource := 'Vector1';
        DBChart5.Series[0].YValues.ValueSource := 'Vector1';
        DBChart6.Series[0].XValues.ValueSource := 'Vector2';
        DBChart6.Series[0].YValues.ValueSource := 'Vector2';
        DBChart7.Series[0].XValues.ValueSource := 'Vector3';
        DBChart7.Series[0].YValues.ValueSource := 'Vector3';
      end;
      4..5 : begin
        DBChart5.Series[0].Active := false;
        DBChart6.Series[0].Active := false;
        DBChart7.Series[0].Active := false;
      end;
      6 : begin
        DBChart5.Series[0].Active := false;
        DBChart6.Series[0].Active := false;
        DBChart7.Series[0].Active := false;
      end;
    end;
    DBChart5.Series[0].XValues.Order := loNone;
    DBChart5.Series[0].YValues.Order := loNone;
    DBChart6.Series[0].XValues.Order := loNone;
    DBChart6.Series[0].YValues.Order := loNone;
    DBChart7.Series[0].XValues.Order := loNone;
    DBChart7.Series[0].YValues.Order := loNone;
  except
  end;
  {all variables}
  try
    DBChart1.Series[0].DataSource := dmCor.FacLoadingsVar;
    DBChart2.Series[0].DataSource := dmCor.FacLoadingsVar;
    DBChart3.Series[0].DataSource := dmCor.FacLoadingsVar;
    DBChart1.Series[0].Active := true;
    DBChart2.Series[0].Active := true;
    DBChart3.Series[0].Active := true;
    DBChart1.BottomAxis.Title.Caption := 'Component 1';
    DBChart1.LeftAxis.Title.Caption := 'Component 2';
    DBChart2.BottomAxis.Title.Caption := 'Component 1';
    DBChart2.LeftAxis.Title.Caption := 'Component 3';
    DBChart3.BottomAxis.Title.Caption := 'Component 2';
    DBChart3.LeftAxis.Title.Caption := 'Component 3';
    case SimilarityChoice of
      1 : begin
        DBChart1.BottomAxis.Title.Caption := 'Component 2';
        DBChart1.LeftAxis.Title.Caption := 'Component 3';
        DBChart2.BottomAxis.Title.Caption := 'Component 2';
        DBChart2.LeftAxis.Title.Caption := 'Component 4';
        DBChart3.BottomAxis.Title.Caption := 'Component 3';
        DBChart3.LeftAxis.Title.Caption := 'Component 4';
        DBChart1.Series[0].XValues.ValueSource := 'Vector2';
        DBChart1.Series[0].YValues.ValueSource := 'Vector3';
        DBChart2.Series[0].XValues.ValueSource := 'Vector2';
        DBChart2.Series[0].YValues.ValueSource := 'Vector4';
        DBChart3.Series[0].XValues.ValueSource := 'Vector3';
        DBChart3.Series[0].YValues.ValueSource := 'Vector4';
      end;
      2..3 : begin
        DBChart1.Series[0].XValues.ValueSource := 'Vector1';
        DBChart1.Series[0].YValues.ValueSource := 'Vector2';
        DBChart2.Series[0].XValues.ValueSource := 'Vector1';
        DBChart2.Series[0].YValues.ValueSource := 'Vector3';
        DBChart3.Series[0].XValues.ValueSource := 'Vector2';
        DBChart3.Series[0].YValues.ValueSource := 'Vector3';
      end;
      4..5 : begin
        DBChart1.Series[0].Active := false;
        DBChart2.Series[0].Active := false;
        DBChart3.Series[0].Active := false;
      end;
      6 : begin
        DBChart1.Series[0].Active := false;
        DBChart2.Series[0].Active := false;
        DBChart3.Series[0].Active := false;
      end;
    end;
    DBChart1.Series[0].XValues.Order := loNone;
    DBChart1.Series[0].YValues.Order := loNone;
    DBChart2.Series[0].XValues.Order := loNone;
    DBChart2.Series[0].YValues.Order := loNone;
    DBChart3.Series[0].XValues.Order := loNone;
    DBChart3.Series[0].YValues.Order := loNone;
  except
  end;
  {all samples}
  try
    DBChart1.Series[1].DataSource := dmCor.FacLoadingsSmp;
    DBChart2.Series[1].DataSource := dmCor.FacLoadingsSmp;
    DBChart2.Series[1].DataSource := dmCor.FacLoadingsSmp;
    DBChart1.Series[1].Active := true;
    DBChart2.Series[1].Active := true;
    DBChart3.Series[1].Active := true;
    case SimilarityChoice of
      1 : begin
        DBChart1.Series[1].XValues.ValueSource := 'Vector2';
        DBChart1.Series[1].YValues.ValueSource := 'Vector3';
        DBChart2.Series[1].XValues.ValueSource := 'Vector2';
        DBChart2.Series[1].YValues.ValueSource := 'Vector4';
        DBChart3.Series[1].XValues.ValueSource := 'Vector3';
        DBChart3.Series[1].YValues.ValueSource := 'Vector4';
      end;
      2..3 : begin
        DBChart1.Series[1].XValues.ValueSource := 'Vector1';
        DBChart1.Series[1].YValues.ValueSource := 'Vector2';
        DBChart2.Series[1].XValues.ValueSource := 'Vector1';
        DBChart2.Series[1].YValues.ValueSource := 'Vector3';
        DBChart3.Series[1].XValues.ValueSource := 'Vector2';
        DBChart3.Series[1].YValues.ValueSource := 'Vector3';
      end;
      4..5 : begin
        DBChart1.Series[1].XValues.ValueSource := 'Vector1';
        DBChart1.Series[1].YValues.ValueSource := 'Vector2';
        DBChart2.Series[1].XValues.ValueSource := 'Vector1';
        DBChart2.Series[1].YValues.ValueSource := 'Vector3';
        DBChart3.Series[1].XValues.ValueSource := 'Vector2';
        DBChart3.Series[1].YValues.ValueSource := 'Vector3';
      end;
      6 : begin
        DBChart1.Series[1].XValues.ValueSource := 'Vector1';
        DBChart1.Series[1].YValues.ValueSource := 'Vector2';
        DBChart2.Series[1].Active := false;
        DBChart3.Series[1].Active := false;
      end;
    end;
    DBChart1.Series[1].XValues.Order := loNone;
    DBChart1.Series[1].YValues.Order := loNone;
    DBChart2.Series[1].XValues.Order := loNone;
    DBChart2.Series[1].YValues.Order := loNone;
    DBChart3.Series[1].XValues.Order := loNone;
    DBChart3.Series[1].YValues.Order := loNone;
  except
  end;
  {all samples in group}
  try
    DBChart1.Series[2].DataSource := dmCor.GroupedSmp;
    DBChart2.Series[2].DataSource := dmCor.GroupedSmp;
    DBChart3.Series[2].DataSource := dmCor.GroupedSmp;
    DBChart1.Series[2].Active := true;
    DBChart2.Series[2].Active := true;
    DBChart3.Series[2].Active := true;
    case SimilarityChoice of
      1 : begin
        DBChart1.Series[2].XValues.ValueSource := 'Vector2';
        DBChart1.Series[2].YValues.ValueSource := 'Vector3';
        DBChart2.Series[2].XValues.ValueSource := 'Vector2';
        DBChart2.Series[2].YValues.ValueSource := 'Vector4';
        DBChart3.Series[2].XValues.ValueSource := 'Vector3';
        DBChart3.Series[2].YValues.ValueSource := 'Vector4';
      end;
      2..3 : begin
        DBChart1.Series[2].XValues.ValueSource := 'Vector1';
        DBChart1.Series[2].YValues.ValueSource := 'Vector2';
        DBChart2.Series[2].XValues.ValueSource := 'Vector1';
        DBChart2.Series[2].YValues.ValueSource := 'Vector3';
        DBChart3.Series[2].XValues.ValueSource := 'Vector2';
        DBChart3.Series[2].YValues.ValueSource := 'Vector3';
      end;
      4..5 : begin
        DBChart1.Series[2].XValues.ValueSource := 'Vector1';
        DBChart1.Series[2].YValues.ValueSource := 'Vector2';
        DBChart2.Series[2].XValues.ValueSource := 'Vector1';
        DBChart2.Series[2].YValues.ValueSource := 'Vector3';
        DBChart3.Series[2].XValues.ValueSource := 'Vector2';
        DBChart3.Series[2].YValues.ValueSource := 'Vector3';
      end;
      6 : begin
        DBChart1.Series[2].XValues.ValueSource := 'Vector1';
        DBChart1.Series[2].YValues.ValueSource := 'Vector2';
        DBChart2.Series[2].Active := false;
        DBChart3.Series[2].Active := false;
      end;
    end;
    DBChart1.Series[2].XValues.Order := loNone;
    DBChart1.Series[2].YValues.Order := loNone;
    DBChart2.Series[2].XValues.Order := loNone;
    DBChart2.Series[2].YValues.Order := loNone;
    DBChart3.Series[2].XValues.Order := loNone;
    DBChart3.Series[2].YValues.Order := loNone;
  except
  end;
  {current sample}
  try
    DBChart1.Series[3].DataSource := dmCor.QGroupedSmp;
    DBChart2.Series[3].DataSource := dmCor.QGroupedSmp;
    DBChart3.Series[3].DataSource := dmCor.QGroupedSmp;
    DBChart1.Series[3].Active := true;
    DBChart2.Series[3].Active := true;
    DBChart3.Series[3].Active := true;
    case SimilarityChoice of
      1 : begin
        DBChart1.Series[3].XValues.ValueSource := 'Vector2';
        DBChart1.Series[3].YValues.ValueSource := 'Vector3';
        DBChart2.Series[3].XValues.ValueSource := 'Vector2';
        DBChart2.Series[3].YValues.ValueSource := 'Vector4';
        DBChart3.Series[3].XValues.ValueSource := 'Vector3';
        DBChart3.Series[3].YValues.ValueSource := 'Vector4';
      end;
      2..3 : begin
        DBChart1.Series[3].XValues.ValueSource := 'Vector1';
        DBChart1.Series[3].YValues.ValueSource := 'Vector2';
        DBChart2.Series[3].XValues.ValueSource := 'Vector1';
        DBChart2.Series[3].YValues.ValueSource := 'Vector3';
        DBChart3.Series[3].XValues.ValueSource := 'Vector2';
        DBChart3.Series[3].YValues.ValueSource := 'Vector3';
      end;
      4..5 : begin
        DBChart1.Series[3].XValues.ValueSource := 'Vector1';
        DBChart1.Series[3].YValues.ValueSource := 'Vector2';
        DBChart2.Series[3].XValues.ValueSource := 'Vector1';
        DBChart2.Series[3].YValues.ValueSource := 'Vector3';
        DBChart3.Series[3].XValues.ValueSource := 'Vector2';
        DBChart3.Series[3].YValues.ValueSource := 'Vector3';
      end;
      6 : begin
        DBChart1.Series[3].XValues.ValueSource := 'Vector1';
        DBChart1.Series[3].YValues.ValueSource := 'Vector2';
        DBChart2.Series[3].Active := false;
        DBChart3.Series[3].Active := false;
      end;
    end;
    DBChart1.Series[3].XValues.Order := loNone;
    DBChart1.Series[3].YValues.Order := loNone;
    DBChart2.Series[3].XValues.Order := loNone;
    DBChart2.Series[3].YValues.Order := loNone;
    DBChart3.Series[3].XValues.Order := loNone;
    DBChart3.Series[3].YValues.Order := loNone;
  except
  end;
  {current variable}
  try
    DBChart1.Series[4].DataSource := dmCor.FacLoadingsVarLinked;
    DBChart2.Series[4].DataSource := dmCor.FacLoadingsVarLinked;
    DBChart3.Series[4].DataSource := dmCor.FacLoadingsVarLinked;
    DBChart1.Series[4].Active := true;
    DBChart2.Series[4].Active := true;
    DBChart3.Series[4].Active := true;
    case SimilarityChoice of
      1 : begin
        DBChart1.Series[4].XValues.ValueSource := 'Vector2';
        DBChart1.Series[4].YValues.ValueSource := 'Vector3';
        DBChart2.Series[4].XValues.ValueSource := 'Vector2';
        DBChart2.Series[4].YValues.ValueSource := 'Vector4';
        DBChart3.Series[4].XValues.ValueSource := 'Vector3';
        DBChart3.Series[4].YValues.ValueSource := 'Vector4';
      end;
      2..3 : begin
        DBChart1.Series[4].XValues.ValueSource := 'Vector1';
        DBChart1.Series[4].YValues.ValueSource := 'Vector2';
        DBChart2.Series[4].XValues.ValueSource := 'Vector1';
        DBChart2.Series[4].YValues.ValueSource := 'Vector3';
        DBChart3.Series[4].XValues.ValueSource := 'Vector2';
        DBChart3.Series[4].YValues.ValueSource := 'Vector3';
      end;
      4..5 : begin
        DBChart1.Series[4].Active := false;
        DBChart2.Series[4].Active := false;
        DBChart3.Series[4].Active := false;
      end;
      6 : begin
        DBChart1.Series[4].Active := false;
        DBChart2.Series[4].Active := false;
        DBChart3.Series[4].Active := false;
      end;
    end;
    DBChart1.Series[4].XValues.Order := loNone;
    DBChart1.Series[4].YValues.Order := loNone;
    DBChart2.Series[4].XValues.Order := loNone;
    DBChart2.Series[4].YValues.Order := loNone;
    DBChart3.Series[4].XValues.Order := loNone;
    DBChart3.Series[4].YValues.Order := loNone;
  except
  end;
  {all samples - localities}
  try
    DBChart4.Series[0].DataSource := dmCor.SmpLoc;
    DBChart4.Series[0].XValues.Order := loNone;
    DBChart4.Series[0].YValues.Order := loNone;
    DBChart8.Series[0].DataSource := dmCor.SmpLoc;
    DBChart8.Series[0].XValues.Order := loNone;
    DBChart8.Series[0].YValues.Order := loNone;
  except
  end;
  {all samples in group - localities}
  try
    DBChart4.Series[1].DataSource := dmCor.GroupedSmpLoc;
    DBChart4.Series[1].XValues.Order := loNone;
    DBChart4.Series[1].YValues.Order := loNone;
    DBChart8.Series[1].DataSource := dmCor.GroupedSmpLoc;
    DBChart8.Series[1].XValues.Order := loNone;
    DBChart8.Series[1].YValues.Order := loNone;
  except
  end;
  {current sample - locality}
  try
    DBChart4.Series[2].DataSource := dmCor.QGroupedSmpLoc;
    DBChart4.Series[2].XValues.Order := loNone;
    DBChart4.Series[2].YValues.Order := loNone;
    DBChart8.Series[2].DataSource := dmCor.QGroupedSmpLoc;
    DBChart8.Series[2].XValues.Order := loNone;
    DBChart8.Series[2].YValues.Order := loNone;
  except
  end;
  try
  dmCor.FacLoadingsVar.Open;
  dmCor.FacLoadingsVarLinked.Open;
  dmCor.FacLoadingsSmp.Open;
  dmCor.SmpLoc.Open;
  dmCor.QGroups.Open;
  dmCor.QPlotGroups.Open;
  dmCor.GroupedSmp.Open;
  dmCor.QGroupedSmp.Open;
  dmCor.GroupedSmpLoc.Open;
  dmCor.QGroupedSmpLoc.Open;
  except
  end;
  DBChart8.View3DOptions.Rotation := 315;
  eRotation.Text := IntToStr(DBChart8.View3DOptions.Rotation);
  DBChart8.View3DOptions.Elevation := 350;
  eElevation.Text := IntToStr(DBChart8.View3DOptions.Elevation);
  DBChart8.View3DOptions.Perspective := 0;
  ePerspective.Text := IntToStr(DBChart8.View3DOptions.Perspective);
  DBChart8.View3DOptions.Zoom := 50;
  eZoom.Text := IntToStr(DBChart8.View3DOptions.Zoom);
end;

procedure TfmCoranMain.Button1Click(Sender: TObject);
begin
  PrepareGraphs;
end;

procedure TfmCoranMain.bbSaveVarClick(Sender: TObject);
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
  if SaveDialogSprdSheet.Execute then
  begin
    pFileType := Excel97Type;
    case SaveDialogSprdSheet.FilterIndex of
      1 : pFileType := Excel97Type;
      2 : pFileType := Excel5Type;
    end;
    pBuf := SaveDialogSprdSheet.FileName;
    F1Var.Write(pBuf,pFileType);
  end;
end;

procedure TfmCoranMain.bbSaveSmpClick(Sender: TObject);
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
  if SaveDialogSprdSheet.Execute then
  begin
    pFileType := Excel97Type;
    case SaveDialogSprdSheet.FilterIndex of
      1 : pFileType := Excel97Type;
      2 : pFileType := Excel5Type;
    end;
    pBuf := SaveDialogSprdSheet.FileName;
    F1Smp.Write(pBuf,pFileType);
  end;
end;

procedure TfmCoranMain.dbnGroupedSmpClick(Sender: TObject;
  Button: TNavigateBtn);
begin
  try
    dmCor.QGroupedSmp.Close;
    dmCor.QGroupedSmpLoc.Close;
  except
  end;
  {
  try
    DBChart1.Series[3].DataSource := dmCor.QGroupedSmp;
    DBChart1.Series[3].XValues.Order := loNone;
    DBChart1.Series[3].YValues.Order := loNone;
  except
  end;
  try
    DBChart2.Series[3].DataSource := dmCor.QGroupedSmp;
    DBChart2.Series[3].XValues.Order := loNone;
    DBChart2.Series[3].YValues.Order := loNone;
  except
  end;
  try
    DBChart3.Series[3].DataSource := dmCor.QGroupedSmp;
    DBChart3.Series[3].XValues.Order := loNone;
    DBChart3.Series[3].YValues.Order := loNone;
  except
  end;
  try
    DBChart4.Series[2].DataSource := dmCor.QGroupedSmpLoc;
    DBChart4.Series[2].XValues.Order := loNone;
    DBChart4.Series[2].YValues.Order := loNone;
  except
  end;
  }
  try
    dmCor.QGroupedSmp.Open;
    dmCor.QGroupedSmpLoc.Open;
  except
  end;
end;

procedure TfmCoranMain.dbnGroupsClick(Sender: TObject;
  Button: TNavigateBtn);
begin
  try
    dmCor.GroupedSmp.Close;
    dmCor.GroupedSmpLoc.Close;
  except
  end;
  {
  try
    DBChart1.Series[2].DataSource := dmCor.GroupedSmp;
    DBChart1.Series[2].XValues.Order := loNone;
    DBChart1.Series[2].YValues.Order := loNone;
  except
  end;
  try
    DBChart2.Series[2].DataSource := dmCor.GroupedSmp;
    DBChart2.Series[2].XValues.Order := loNone;
    DBChart2.Series[2].YValues.Order := loNone;
  except
  end;
  try
    DBChart3.Series[2].DataSource := dmCor.GroupedSmp;
    DBChart3.Series[2].XValues.Order := loNone;
    DBChart3.Series[2].YValues.Order := loNone;
  except
  end;
  try
    DBChart4.Series[1].DataSource := dmCor.GroupedSmpLoc;
    DBChart4.Series[1].XValues.Order := loNone;
    DBChart4.Series[1].YValues.Order := loNone;
  except
  end;
  }
  try
    dmCor.GroupedSmp.Open;
    dmCor.GroupedSmpLoc.Open;
    dbnGroupedSmpClick(Sender,nbFirst);
  except
  end;
end;

procedure TfmCoranMain.dbnVarClick(Sender: TObject; Button: TNavigateBtn);
begin
  try
    dmCor.FacLoadingsVarLinked.Close;
  except
  end;
  {
  try
    DBChart1.Series[4].DataSource := dmCor.FacLoadingsVarLinked;
    DBChart1.Series[4].XValues.Order := loNone;
    DBChart1.Series[4].YValues.Order := loNone;
  except
  end;
  try
    DBChart2.Series[4].DataSource := dmCor.FacLoadingsVarLinked;
    DBChart2.Series[4].XValues.Order := loNone;
    DBChart2.Series[4].YValues.Order := loNone;
  except
  end;
  try
    DBChart3.Series[4].DataSource := dmCor.FacLoadingsVarLinked;
    DBChart3.Series[4].XValues.Order := loNone;
    DBChart3.Series[4].YValues.Order := loNone;
  except
  end;
  }
  try
    dmCor.FacLoadingsVarLinked.Open;
  except
  end;
end;

procedure TfmCoranMain.PrintGraph1Click(Sender: TObject);
begin
  if (pc1.ActivePage = tsGraph1) then DBChart1.Print;
  if (pc1.ActivePage = tsGraph2) then DBChart2.Print;
  if (pc1.ActivePage = tsGraph3) then DBChart3.Print;
end;

procedure TfmCoranMain.bbEmptyCoranVecVarClick(Sender: TObject);
begin
  dmCor.FacLoadingsVar.Last;
  if not(dmCor.FacLoadingsVar.BOF and dmCor.FacLoadingsVar.EOF) then
  begin
    dmCor.FacLoadingsVar.Last;
    repeat
      dmCor.FacLoadingsVar.Delete;
    until dmCor.FacLoadingsVar.BOF;
  end;
end;

procedure TfmCoranMain.Discrm;
var
  I, J, K, L : integer;
  tmpStr : string;
  tmpstr10 : string[10];
  blankstr : string;
  {A is MxM, X is NxMx2, A2 is MxM, C is 2xM, NS is 2 matrix}
begin
  blankstr := '  ';
  DBChart7.Visible := false;
  MemoResults.Clear;
  MemoResults.Font.Name := 'Courier';
  MemoResults.Font.Pitch := fpFixed;
  MemoResults.Font.Style := [fsBold];
  MemoResults.Font.Size := 12;
  MemoResults.Lines.Add(eTitle.Text+'                '+DateToStr(Now));
  MemoResults.Font.Style := [];
  MemoResults.Font.Size := 8;
  MemoResults.Lines.Add('  ');
  MemoResults.Lines.Add('  ');
  if (cbPrint.Checked=true) then IPrn := 'Y'
                            else IPrn := 'N';
  if IPrn = 'Y' then
  begin
    AssignPrn(Lst);
    Rewrite(Lst);
    Printer.Canvas.Font.Name := 'CourierNew';
    Printer.Canvas.Font.Style := [fsBold];
    Printer.Canvas.Font.Size := 7;
    Writeln(Lst,' ');
    Writeln(Lst,'');
    Writeln(Lst,'');
    Write(Lst,' ':10,eTitle.Text);
    Writeln(Lst,'  ');
    Writeln(Lst,'Discriminant function analysis  ');
    Writeln(Lst,'  ');
    Printer.Canvas.Font.Style := [];
  end;
  FillChar(A3,SizeOf(A3),0);
  FillChar(A1,SizeOf(A1),0);
  FillChar(A2,SizeOf(A2),0);
  Nox := dmCor.ElemNames.RecordCount;
  TotalRecs := dmCor.CoranChem.RecordCount;
  M := Nox;
  i := 1;
  dmCor.ElemNames.First;
  repeat
    if (dmCor.ElemNamesPos.AsInteger > 0) then
    begin
      OxideName[dmCor.ElemNamesPos.AsInteger] := dmCor.ElemNamesCalled.AsString;
      {
      MemoResults.Lines.Add(dmCor.ElemNamesPos.AsString+'   '+OxideName[dmCor.ElemNamesPos.AsInteger]);
      }
    end;
    i := i + 1;
    dmCor.ElemNames.Next;
  until ((dmCor.ElemNames.Eof) or (i > Nox));
  {
  MemoResults.Lines.Add('  ');
  MemoResults.Lines.Add('  ');
  }
  try
    dmCor.QGroups.Open;
    dmCor.QPlotGroups.Open;
  except
  end;
  dmCor.QGroups.First;
  dmCor.GroupedChem.Open;
  {read group 1 data}
  i:=0;
  N:=1;
  sbmain.Panels[1].Text :='Reading data for '+IntToStr(Nox)+' elements for first group';
  sbMain.Refresh;
  dmCor.GroupedChem.First;
  repeat
    i:=i+1;
    Component[I]:=dmCor.GroupedChemSAMPLENUM.AsString;
    for J:=1 to Nox do begin
      case J of
        1 : X[i,J] := dmCor.GroupedChemParam1.AsFloat;
        2 : X[I,J] := dmCor.GroupedChemParam2.AsFloat;
        3 : X[I,J] := dmCor.GroupedChemParam3.AsFloat;
        4 : X[I,J] := dmCor.GroupedChemParam4.AsFloat;
        5 : X[I,J] := dmCor.GroupedChemParam5.AsFloat;
        6 : X[I,J] := dmCor.GroupedChemParam6.AsFloat;
        7 : X[I,J] := dmCor.GroupedChemParam7.AsFloat;
        8 : X[I,J] := dmCor.GroupedChemParam8.AsFloat;
        9 : X[I,J] := dmCor.GroupedChemParam9.AsFloat;
        10 : X[I,J] := dmCor.GroupedChemParam10.AsFloat;
        11 : X[I,J] := dmCor.GroupedChemParam11.AsFloat;
        12 : X[I,J] := dmCor.GroupedChemParam12.AsFloat;
        13 : X[I,J] := dmCor.GroupedChemParam13.AsFloat;
        14 : X[I,J] := dmCor.GroupedChemParam14.AsFloat;
        15 : X[I,J] := dmCor.GroupedChemParam15.AsFloat;
        16 : X[I,J] := dmCor.GroupedChemParam16.AsFloat;
        17 : X[I,J] := dmCor.GroupedChemParam17.AsFloat;
        18 : X[I,J] := dmCor.GroupedChemParam18.AsFloat;
        19 : X[I,J] := dmCor.GroupedChemParam19.AsFloat;
        20 : X[I,J] := dmCor.GroupedChemParam20.AsFloat;
        21 : X[I,J] := dmCor.GroupedChemParam21.AsFloat;
        22 : X[I,J] := dmCor.GroupedChemParam22.AsFloat;
        23 : X[I,J] := dmCor.GroupedChemParam23.AsFloat;
        24 : X[I,J] := dmCor.GroupedChemParam24.AsFloat;
        25 : X[I,J] := dmCor.GroupedChemParam25.AsFloat;
        26 : X[I,J] := dmCor.GroupedChemParam26.AsFloat;
        27 : X[I,J] := dmCor.GroupedChemParam27.AsFloat;
        28 : X[I,J] := dmCor.GroupedChemParam28.AsFloat;
        29 : X[I,J] := dmCor.GroupedChemParam29.AsFloat;
        30 : X[I,J] := dmCor.GroupedChemParam30.AsFloat;
      end;
    end;
    sbmain.Panels[1].Text :='Processing record '+IntToStr(N)+' of '+IntToStr(TotalRecs)+' Group 1 ';
    sbMain.Refresh;
    N:=N+1;
    dmCor.GroupedChem.Next;
  until (dmCor.GroupedChem.Eof);
  Ndig[1] := i;
  for j := 1 to Ndig[1] do
  begin
    for k := 1 to Nox do
    begin
      A3[1,k] := A3[1,k] + X[j,k];
      for l := 1 to Nox do
      begin
        A1[k,l] := A1[k,l] + X[j,k] * X[j,l];
      end;
    end;
  end;
  {
  MemoResults.Lines.Add('Vector mean for group 1.    Number of samples = '+IntToStr(Ndig[1]));
  }
  MemoResults.Lines.Add('Group 1.  n = '+IntToStr(Ndig[1]));
  tmpStr := '';
  if IPrn = 'Y' then
  begin
    Writeln(Lst,'Vector mean for group 1.    Number of samples = ',IntToStr(Ndig[1]));
    Writeln(Lst,'  ');
  end;
  for k := 1 to Nox do
  begin
    tmpStr := tmpStr + FormatFloat('#####0.0###',A3[1,k]/Ndig[1])+'   ';
    if IPrn = 'Y' then Write(Lst,(A3[1,k]/Ndig[1]):10:4);
  end;
  if IPrn = 'Y' then Writeln(Lst,'  ');
  {
  MemoResults.Lines.Add(tmpStr);
  MemoResults.Lines.Add('  ');
  }
  dmCor.QGroups.Next;
  {read group 2 data}
  N:=1;
  sbmain.Panels[1].Text :='Reading data for '+IntToStr(Nox)+' elements for second group';
  sbMain.Refresh;
  dmCor.GroupedChem.First;
  repeat
    i:=i+1;
    Component[I]:=dmCor.GroupedChemSAMPLENUM.AsString;
    for J:=1 to Nox do begin
      case J of
        1 : X[I,J] := dmCor.GroupedChemParam1.AsFloat;
        2 : X[I,J] := dmCor.GroupedChemParam2.AsFloat;
        3 : X[I,J] := dmCor.GroupedChemParam3.AsFloat;
        4 : X[I,J] := dmCor.GroupedChemParam4.AsFloat;
        5 : X[I,J] := dmCor.GroupedChemParam5.AsFloat;
        6 : X[I,J] := dmCor.GroupedChemParam6.AsFloat;
        7 : X[I,J] := dmCor.GroupedChemParam7.AsFloat;
        8 : X[I,J] := dmCor.GroupedChemParam8.AsFloat;
        9 : X[I,J] := dmCor.GroupedChemParam9.AsFloat;
        10 : X[I,J] := dmCor.GroupedChemParam10.AsFloat;
        11 : X[I,J] := dmCor.GroupedChemParam11.AsFloat;
        12 : X[I,J] := dmCor.GroupedChemParam12.AsFloat;
        13 : X[I,J] := dmCor.GroupedChemParam13.AsFloat;
        14 : X[I,J] := dmCor.GroupedChemParam14.AsFloat;
        15 : X[I,J] := dmCor.GroupedChemParam15.AsFloat;
        16 : X[I,J] := dmCor.GroupedChemParam16.AsFloat;
        17 : X[I,J] := dmCor.GroupedChemParam17.AsFloat;
        18 : X[I,J] := dmCor.GroupedChemParam18.AsFloat;
        19 : X[I,J] := dmCor.GroupedChemParam19.AsFloat;
        20 : X[I,J] := dmCor.GroupedChemParam20.AsFloat;
        21 : X[I,J] := dmCor.GroupedChemParam21.AsFloat;
        22 : X[I,J] := dmCor.GroupedChemParam22.AsFloat;
        23 : X[I,J] := dmCor.GroupedChemParam23.AsFloat;
        24 : X[I,J] := dmCor.GroupedChemParam24.AsFloat;
        25 : X[I,J] := dmCor.GroupedChemParam25.AsFloat;
        26 : X[I,J] := dmCor.GroupedChemParam26.AsFloat;
        27 : X[I,J] := dmCor.GroupedChemParam27.AsFloat;
        28 : X[I,J] := dmCor.GroupedChemParam28.AsFloat;
        29 : X[I,J] := dmCor.GroupedChemParam29.AsFloat;
        30 : X[I,J] := dmCor.GroupedChemParam30.AsFloat;
      end;
    end;
    sbmain.Panels[1].Text :='Processing record '+IntToStr(N)+' of '+IntToStr(TotalRecs)+' Group 2';
    sbMain.Refresh;
    N:=N+1;
    dmCor.GroupedChem.Next;
  until (dmCor.GroupedChem.Eof);
  Ndig[2] := i - Ndig[1];
  for j := Ndig[1]+1 to Ndig[1]+Ndig[2] do
  begin
    for k := 1 to Nox do
    begin
      A3[2,k] := A3[2,k] + X[j,k];
      for l := 1 to Nox do
      begin
        A1[k,l] := A1[k,l] + X[j,k] * X[j,l];
      end;
    end;
  end;
  {
  MemoResults.Lines.Add('Vector mean for group 2.    Number of samples = '+IntToStr(Ndig[2]));
  }
  MemoResults.Lines.Add('Group 2.  n = '+IntToStr(Ndig[2]));
  if IPrn = 'Y' then
  begin
    Writeln(Lst,'Vector mean for group 2.    Number of samples = ',IntToStr(Ndig[2]));
    Writeln(Lst,'  ');
  end;
  tmpStr := '';
  for k := 1 to Nox do
  begin
    tmpStr := tmpStr + FormatFloat('#####0.0###',A3[2,k]/Ndig[2])+'   ';
    if IPrn = 'Y' then Write(Lst,(A3[2,k]/Ndig[2]):10:4);
  end;
  if IPrn = 'Y' then Writeln(Lst,'  ');
  {
  MemoResults.Lines.Add(tmpStr);
  MemoResults.Lines.Add('  ');
  }

  AN1 := 1.0*Ndig[1];
  AN2 := 1.0*Ndig[2];
  AN3 := 1.0*AN1 + 1.0*AN2 - 2.0;
  for i := 1 to M do
  begin
    A2[i,i] := A3[1,i]/AN1 - A3[2,i]/AN2;
    for j := 1 to Nox do
    begin
      A1[i,j] := (A1[i,j] - A3[1,i]*A3[1,j]/AN1-A3[2,i]*A3[2,j]/AN2)/AN3;
    end;
  end;
  {
  MemoResults.Lines.Add('Vector of mean differences');
  }
  if IPrn = 'Y' then
  begin
    Writeln(Lst,'Vector of mean differences');
    Writeln(Lst,'  ');
  end;
  tmpStr := '';
  for k := 1 to Nox do
  begin
    tmpStr := tmpStr + FormatFloat('#####0.0###',A2[k,k])+'   ';
    if IPrn = 'Y' then Write(Lst,A2[k,k]:10:4);
  end;
  if IPrn = 'Y' then Writeln(Lst,'  ');
  {
  MemoResults.Lines.Add(tmpStr);
  MemoResults.Lines.Add('  ');
  }
  {try writing all in separate columns}
  for k := 1 to Nox do
  begin
    tmpstr := '';
    tmpstr := tmpstr + FormatFloat('000',k);
    ResultsArray[k,1] := FormatFloat('000',k) + '   ';
    tmpstr := OxideName[k];
    ResultsArray[k,2] := tmpstr + CharStream(10-Length(tmpstr),32) + blankstr;
    tmpStr := FormatFloat('#####0.0000',A3[1,k]/Ndig[1]);
    ResultsArray[k,3] := CharStream(10-Length(tmpstr),32)+tmpstr + blankstr;
    tmpStr := FormatFloat('#####0.0000',A3[2,k]/Ndig[2]);
    ResultsArray[k,4] := CharStream(10-Length(tmpstr),32)+tmpstr + blankstr;
    tmpStr := FormatFloat('#####0.0000',A2[k,k]);
    ResultsArray[k,5] := CharStream(10-Length(tmpstr),32)+tmpstr + blankstr;
  end;
  SLE(A1,A2,M,MM,1.0e-7);
  {
  MemoResults.Lines.Add('Discriminant function vector parameters');
  }
  if IPrn = 'Y' then
  begin
    Writeln(Lst,'Discriminant function vector parameters');
    Writeln(Lst,'  ');
  end;
  tmpStr := '';
  for k := 1 to Nox do
  begin
    tmpStr := tmpStr + FormatFloat('#####0.0###',A2[k,k])+'   ';
    if IPrn = 'Y' then Write(Lst,A2[k,k]:10:4);
  end;
  if IPrn = 'Y' then Writeln(Lst,'  ');
  {
  MemoResults.Lines.Add(tmpStr);
  MemoResults.Lines.Add('  ');
  }
  {add data for another column}
  for k := 1 to Nox do
  begin
    tmpStr := FormatFloat('#####0.0000',A2[k,k]);
    ResultsArray[k,6] := CharStream(10-Length(tmpstr),32)+tmpstr + blankstr;
  end;
  R0 := 0.0;
  R1 := 0.0;
  R2 := 0.0;
  D2 := 0.0;
  for i := 1 to Nox do
  begin
    R0 := R0 + A2[i,i]*(A3[1,i]/AN1 + A3[2,i]/AN2)/2.0;
    R1 := R1 + A2[i,i]*A3[1,i]/AN1;
    R2 := R2 + A2[i,i]*A3[2,i]/AN2;
    D2 := D2 + A2[i,i]*(A3[1,i]/AN1 - A3[2,i]/AN2);
  end;
  AM := 1.0*M;
  F := (((AN1 + AN2 - AM - 1.0)*AN1*AN2)/(AN3*AM*(AN1 + AN2)))*D2;
  ND1 := M;
  ND2 := Ndig[1] + Ndig[2] - M - 1;
  MemoResults.Lines.Add('  ');
  MemoResults.Lines.Add('  ');
  MemoResults.Lines.Add('F = '+FormatFloat('######0.0#',F)+' with '+IntToStr(ND1)+' and '+IntToStr(ND2)+' degrees of freedom');
  MemoResults.Lines.Add('  ');
  MemoResults.Lines.Add('Mahalanobis D2 = '+FormatFloat('########0.0#',D2));
  MemoResults.Lines.Add('  ');
  MemoResults.Lines.Add('R1 = '+FormatFloat('########0.0###',R1));
  MemoResults.Lines.Add('R0 = '+FormatFloat('########0.0###',R0));
  MemoResults.Lines.Add('R2 = '+FormatFloat('########0.0###',R2));
  MemoResults.Lines.Add('  ');
  if IPrn = 'Y' then
  begin
    Writeln(Lst,'F = ',F:10:2,' with ',ND1,' and ',ND2,'degrees of freedom');
    Writeln(Lst,'  ');
    Writeln(Lst,'Mahalanobis D2= ',D2:10:2);
    Writeln(Lst,'  ');
    Writeln(Lst,'R1= ',R1:14:4);
    Writeln(Lst,'R0= ',R0:14:4);
    Writeln(Lst,'R2= ',R2:14:4);
    Writeln(Lst,'  ');
  end;
  {
  MemoResults.Lines.Add('Variable          Constant         % added');
  }
  if IPrn = 'Y' then Writeln(Lst,'Variable          Constant         % added');
  for i := 1 to M do
  begin
    E := (A2[i,i]*(A3[1,i]/AN1 - A3[2,i]/AN2)/D2)*100.0;
    {
    MemoResults.Lines.Add(' '+FormatFloat('000',i)+' '+OxideName[i]+'          '
       +FormatFloat('#######0.0000',A2[i,i])+'          '+FormatFloat('#######0.0###',E));
    }
    if IPrn = 'Y' then Writeln(Lst,OxideName[i]:15,A2[i,i]:14:4,E:12:2);
  end;
  {add data for another two columns}
  for i := 1 to Nox do
  begin
    E := (A2[i,i]*(A3[1,i]/AN1 - A3[2,i]/AN2)/D2)*100.0;
    tmpStr := FormatFloat('#####0.0000',E);
    ResultsArray[i,7] := CharStream(10-Length(tmpstr),32)+tmpstr;
    tmpStr := '     ' + TakeLogs[i];
    ResultsArray[i,8] := tmpstr;
  end;
  {send formatted matrix to Memo}
  MemoResults.Lines.Add('  ');
  MemoResults.Lines.Add(' #    Variable       Mean 1      Mean 2    Difference    Vector         %      Log');
  MemoResults.Lines.Add('                                                        parameter     added    taken');
  MemoResults.Lines.Add('  ');
  for i := 1 to Nox do
  begin
    tmpstr := '';
    for j := 1 to 8 do
    begin
      tmpstr := tmpstr + ResultsArray[i,j];
    end;
    MemoResults.Lines.Add(tmpstr);
  end;
  MemoResults.Lines.Add('  ');
  MemoResults.Lines.Add('  ');
  if IPrn = 'Y' then Writeln(Lst,'   ');
  CalcDiscrmScores;
  if (IPrn='Y') then
  begin
   Writeln(Lst,char(13));
   System.CloseFile(Lst);
  end;
  tsGraph1.TabVisible := true;
  tsGraph2.TabVisible := false;
  tsGraph3.TabVisible := false;
  tsScores.TabVisible := true;
  tsLocalities.TabVisible := true;
  tsSpreadsheets.TabVisible := true;
  pVar.Visible := false;
  pGraph1Var.Visible := false;
  pSaveVar.Visible := false;
  PrepareGraphs;
end;

procedure TfmCoranMain.CalcDiscrmScores;
var
  I, J, K, L : integer;
  E : double;
begin
  {clear factor loadings table for variables}
  dmCor.FacLoadingsVar.Last;
  if not (dmCor.FacLoadingsVar.Bof and dmCor.FacLoadingsVar.Eof) then
  begin
    dmCor.FacLoadingsVar.Last;
    repeat
      dmCor.FacLoadingsVar.Delete;
    until dmCor.FacLoadingsVar.Bof;
  end;
  {fill factor loadings table for variables}
  dmCor.FacLoadingsVar.First;
  F1Var.Row := 1;
  F1Var.Col := 1;
  F1Var.Text := 'Variable';
  j := 1;
  F1Var.Col := j+1;
  F1Var.text := 'Constant ';
  j := 2;
  F1Var.Col := j+1;
  F1Var.text := '% added ';
  ii := 0;
  for i := 1 to M do
  begin
    ii := ii + 1;
    dmCor.FacLoadingsVar.Append;
    dmCor.FacLoadingsVarPos.AsInteger := i;
    dmCor.FacLoadingsVarCalled.AsString := OxideName[i];
    F1Var.Row := ii+1;
    F1Var.Col := 1;
    F1Var.Text := OxideName[i];
    E := (A2[i,i]*(A3[1,i]/AN1 - A3[2,i]/AN2)/D2)*100.0;
    for j := 1 to 2 do
    begin
      F1Var.Col := j+1;
      case j of
        1 : begin
          dmCor.FacLoadingsVarVector1.AsFloat := A2[i,i];
          F1Var.Number := A1[i,i];
        end;
        2 : begin
          dmCor.FacLoadingsVarVector2.AsFloat := E;
          F1Var.Number := E;
        end;
      end;
    end;
    dmCor.FacLoadingsVar.Post;
  end;
  dmCor.FacLoadingsVar.First;
  {clear factor loadings table for samples}
  dmCor.FacLoadingsSmp.Last;
  if not (dmCor.FacLoadingsSmp.Bof and dmCor.FacLoadingsSmp.Eof) then
  begin
    dmCor.FacLoadingsSmp.Last;
    repeat
      dmCor.FacLoadingsSmp.Delete;
    until dmCor.FacLoadingsSmp.Bof;
  end;
  sbmain.Panels[1].Text :='Calculating scores for group 1';
  sbMain.Refresh;
  {fill factor loadings table for samples for first group}
  F1Smp.Row := 1;
  F1Smp.Col := 1;
  F1Smp.Text := 'Group';
  F1Smp.Col := 2;
  F1Smp.Text := 'Plot Group';
  F1Smp.Col := 3;
  F1Smp.Text := 'Sample';
  j := 1;
  F1Smp.Col := j+3;
  F1Smp.Text := 'Component ' + IntToStr(j);
  j := 2;
  F1Smp.Col := j+3;
  F1Smp.Text := 'Group # ';
  dmCor.QGroups.First;
  dmCor.GroupedChem.First;
  i := 1;
  k := 1;
  dmCor.GroupedChem.First;
  repeat
    sbmain.Panels[1].Text :='Calculating scores for group 1'+dmCor.GroupedChemSAMPLENUM.AsString;
    sbMain.Refresh;
    D := 0.0;
    for l := 1 to M do
    begin
      D := D + A2[l,l] * X[k,l];
    end;
    F1Smp.Row := i+1;
    dmCor.QGroupedSmp.Append;
    dmCor.QGroupedSmpGroupName.AsString := dmCor.GroupedChemGroupName.AsString;
    dmCor.QGroupedSmpPlotGroupName.AsString := dmCor.GroupedChemPlotGroupName.AsString;
    dmCor.QGroupedSmpSampleNum.AsString := dmCor.GroupedChemSAMPLENUM.AsString;
    F1Smp.Col := 1;
    F1Smp.Text := dmCor.QGroupedSmpGroupName.AsString;
    F1Smp.Col := 2;
    F1Smp.Text := dmCor.QGroupedSmpPlotGroupName.AsString;
    F1Smp.Col := 3;
    F1Smp.Text := dmCor.QGroupedSmpSampleNum.AsString;
    j := 1;
    F1Smp.Col := j+3;
    dmCor.QGroupedSmpVector1.AsFloat := D;
    F1Smp.Number := D;
    j := 2;
    F1Smp.Col := j+3;
    dmCor.QGroupedSmpVector2.AsFloat := 1.0;
    F1Smp.Number := 1.0;
    dmCor.QGroupedSmp.Post;
    dmCor.GroupedChem.Next;
    i := i + 1;
    k := k + 1;
  until dmCor.GroupedChem.Eof;
  dmCor.GroupedChem.First;
  dmCor.FacLoadingsSmp.First;
  dmCor.FacLoadingsSmp.Close;
  dmCor.FacLoadingsSmp.Open;
  {fill factor loadings table for samples for second group}
  sbmain.Panels[1].Text :='Calculating scores for group 2';
  sbMain.Refresh;
  dmCor.QGroups.Next;
  dmCor.GroupedChem.First;
  dmCor.GroupedChem.First;
  repeat
    sbmain.Panels[1].Text :='Calculating scores for group 2'+dmCor.GroupedChemSAMPLENUM.AsString;
    sbMain.Refresh;
    D := 0.0;
    for l := 1 to M do
    begin
      D := D + A2[l,l] * X[k,l];
    end;
    F1Smp.Row := i+1;
    dmCor.QGroupedSmp.Append;
    dmCor.QGroupedSmpGroupName.AsString := dmCor.GroupedChemGroupName.AsString;
    dmCor.QGroupedSmpPlotGroupName.AsString := dmCor.GroupedChemPlotGroupName.AsString;
    dmCor.QGroupedSmpSampleNum.AsString := dmCor.GroupedChemSAMPLENUM.AsString;
    F1Smp.Col := 1;
    F1Smp.Text := dmCor.QGroupedSmpGroupName.AsString;
    F1Smp.Col := 2;
    F1Smp.Text := dmCor.QGroupedSmpPlotGroupName.AsString;
    F1Smp.Col := 3;
    F1Smp.Text := dmCor.QGroupedSmpSampleNum.AsString;
    j := 1;
    F1Smp.Col := j+3;
    dmCor.QGroupedSmpVector1.AsFloat := D;
    F1Smp.Number := D;
    j := 2;
    F1Smp.Col := j+3;
    dmCor.QGroupedSmpVector2.AsFloat := 2.0;
    F1Smp.Number := 2.0;
    dmCor.QGroupedSmp.Post;
    dmCor.GroupedChem.Next;
    i := i + 1;
    k := k + 1;
  until dmCor.GroupedChem.Eof;
  dmCor.GroupedChem.First;
  dmCor.FacLoadingsSmp.First;
  dmCor.FacLoadingsSmp.Close;
  dmCor.FacLoadingsSmp.Open;
  dmCor.GroupedChem.Close;
end;

procedure TfmCoranMain.Importdatadefinitions1Click(Sender: TObject);
var
  i : integer;
begin
  try
    ImportForm2 := TfmSheetImport2.Create(Self);
    ImportForm2.OpenDialogSprdSheet.FileName := 'CoranDefinitions';
    ImportForm2.ShowModal;
  finally
    ImportForm2.Free;
  end;
end;

procedure TfmCoranMain.bbExitClick(Sender: TObject);
begin
  dmCor.Coran.Connected := false;
  Close;
end;

procedure TfmCoranMain.DiscrimMulti;
var
  ii, jj : integer;
  I, J, K, KP : integer;
  tmpStr : string;
  blankstr : string;
begin
{
  Original data in array X                                      X (n . m)
  Calculate mean for each variable in each group                XGroupMean (m . g)
  Calculate grand mean for each variable                        XGrandMean (1 . m)
  Calculate covariance between variables for all observations   TotalSumProducts (m . m)
  Calculate within-group covariances                            WithinGroupSumProducts (m . m)
  Calculate between-group covariances                           BetweenGroupSumProducts (m . m)

  TotalSumProducts = WithinGroupSumProducts + BetweenGroupSumProducts;

  InvWithinGroupSumProducts = inverse of WithinGroupSumProducts
  Derive eigenvalues and vectors for InvWithingroupSumProducts * BetweenGroupSumProducts
            may be difficult to solve this since not symetric
  Call the array of eigenvectors EigenVec (m . m)
  Invert the array of eigenvectors = InvEigenVec

  Project observations into space defined by the discriminant axes
    ScoreData = InvEigenVec * X  (n . m)
  Centroid of each group is projected into space defined by the discriminant axes
    ScoreCentroid = InveigenVec * XGroupMean (m . m  x  m . g)

  Plot scores for each of first three disciminant axes against each other
}
end;

procedure TfmCoranMain.Cluster;
var
  ii, jj : integer;
  I, J, K, KP : integer;
  tmpStr : string;
  blankstr : string;
begin
{}
end;

procedure TfmCoranMain.udRotationClick(Sender: TObject;
  Button: TUDBtnType);
begin
  case Button of
    btNext : begin
      DBChart8.View3DOptions.Rotation := DBChart8.View3DOptions.Rotation + 1;
    end;
    btPrev : begin
      DBChart8.View3DOptions.Rotation := DBChart8.View3DOptions.Rotation - 1;
    end;
  end;
  if (DBChart8.View3DOptions.Rotation > 360) then DBChart8.View3DOptions.Rotation := DBChart8.View3DOptions.Rotation - 360;
  if (DBChart8.View3DOptions.Rotation < 0) then DBChart8.View3DOptions.Rotation := DBChart8.View3DOptions.Rotation + 360;
  eRotation.Text := IntToStr(DBChart8.View3DOptions.Rotation);
end;

procedure TfmCoranMain.udElevationClick(Sender: TObject;
  Button: TUDBtnType);
begin
  case Button of
    btNext : begin
      DBChart8.View3DOptions.Elevation := DBChart8.View3DOptions.Elevation + 1;
    end;
    btPrev : begin
      DBChart8.View3DOptions.Elevation := DBChart8.View3DOptions.Elevation - 1;
    end;
  end;
  if (DBChart8.View3DOptions.Elevation > 360) then DBChart8.View3DOptions.Elevation := DBChart8.View3DOptions.Elevation - 360;
  if (DBChart8.View3DOptions.Elevation < 0) then DBChart8.View3DOptions.Elevation := DBChart8.View3DOptions.Elevation + 360;
  eElevation.Text := IntToStr(DBChart8.View3DOptions.Elevation);
end;

procedure TfmCoranMain.udPerspectiveClick(Sender: TObject;
  Button: TUDBtnType);
begin
  case Button of
    btNext : begin
      DBChart8.View3DOptions.Perspective := DBChart8.View3DOptions.Perspective + 1;
    end;
    btPrev : begin
      DBChart8.View3DOptions.Perspective := DBChart8.View3DOptions.Perspective - 1;
    end;
  end;
  if (DBChart8.View3DOptions.Perspective > 100) then DBChart8.View3DOptions.Perspective := DBChart8.View3DOptions.Perspective - 100;
  if (DBChart8.View3DOptions.Perspective < 0) then DBChart8.View3DOptions.Perspective := DBChart8.View3DOptions.Perspective + 100;
  ePerspective.Text := IntToStr(DBChart8.View3DOptions.Perspective);
end;

procedure TfmCoranMain.pc1Change(Sender: TObject);
begin
  if (pc1.ActivePageIndex in [2,4,5,6,7,9,11])
    then sbMain.Panels[1].Text := 'Graphs : left click and drag to zoom;    right click and drag to move'
    else sbMain.Panels[1].Text := '';
end;

procedure TfmCoranMain.udZoomClick(Sender: TObject; Button: TUDBtnType);
begin
  case Button of
    btNext : begin
      DBChart8.View3DOptions.Zoom := DBChart8.View3DOptions.Zoom + 1;
    end;
    btPrev : begin
      DBChart8.View3DOptions.Zoom := DBChart8.View3DOptions.Zoom - 1;
    end;
  end;
  if (DBChart8.View3DOptions.Zoom < 0) then DBChart8.View3DOptions.Zoom := 0;
  eZoom.Text := IntToStr(DBChart8.View3DOptions.Zoom);
end;

end.
