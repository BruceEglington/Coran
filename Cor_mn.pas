unit Cor_mn;

interface

uses
  Windows, Messages, System.SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, StdCtrls, Buttons, ExtCtrls, ToolWin, ComCtrls,
  Printers, Menus, Mask, DBCtrls, Db, IniFiles, VCLTee.TeEngine, Series,
  TeeProcs, VCLTee.Chart, VCLTee.DbChart, AxCtrls, OleCtrls, ActnList,
  System.IOUtils, System.UITypes,
  Generics.Collections, Generics.Defaults,
  Cor_Varb,
  ActnMan, VCLTee.TeeFunci, Math, VCLTee.TeeBoxPlot,
  XPStyleActnCtrls, VCLTee.TeeMapSeries, VCLTee.TeeURL, VCLTee.TeeSeriesTextEd,
  VCLTee.TeeComma, VCLTee.TeeEdit, VclTee.TeeGDIPlus, System.Actions,
  VCLTee.TeeSurfa, VCLTee.TeePoin3, VCLTee.TeeTools, VCLTee.StatChar,
  Flexcel.Core,Flexcel.Report, Vcl.StdStyleActnCtrls;

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
    ImportData1: TMenuItem;
    Help: TMenuItem;
    About1: TMenuItem;
    Printersetup1: TMenuItem;
    PrinterSetupDialog1: TPrinterSetupDialog;
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
    dbgFacLoadingsSmp: TDBGrid;
    pVar: TPanel;
    dbgFacLoadingsVar: TDBGrid;
    DBNavigator2: TDBNavigator;
    tsGraph1: TTabSheet;
    DBChart1: TDBChart;
    SeriesG11: TPointSeries;
    tsGraph2: TTabSheet;
    tsGraph3: TTabSheet;
    DBChart2: TDBChart;
    SeriesG21: TPointSeries;
    DBChart3: TDBChart;
    SeriesG31: TPointSeries;
    SaveDialogSprdSheet: TSaveDialog;
    SeriesG12: TPointSeries;
    SeriesG32: TPointSeries;
    tsCheck: TTabSheet;
    DBNavigator8: TDBNavigator;
    DBGrid10: TDBGrid;
    Button1: TButton;
    DBGrid11: TDBGrid;
    DBGrid9: TDBGrid;
    SeriesG13: TPointSeries;
    DBNavigator7: TDBNavigator;
    SeriesG23: TPointSeries;
    SeriesG33: TPointSeries;
    tsLocalities: TTabSheet;
    DBChart4: TDBChart;
    SeriesG41: TPointSeries;
    SeriesG42: TPointSeries;
    SeriesG43: TPointSeries;
    DBNavigator13: TDBNavigator;
    DBGrid19: TDBGrid;
    SeriesG14: TPointSeries;
    SeriesG24: TPointSeries;
    SeriesG34: TPointSeries;
    PrintGraph1: TMenuItem;
    DBGrid15: TDBGrid;
    lIgnoreLoadingsVar: TLabel;
    tsScores: TTabSheet;
    DBChart5: TDBChart;
    DBChart7: TDBChart;
    DBChart6: TDBChart;
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
    Panel2: TPanel;
    Panel25: TPanel;
    DBGrid3: TDBGrid;
    Panel26: TPanel;
    DBGrid1: TDBGrid;
    Panel24: TPanel;
    bbEmptyCoranChem: TBitBtn;
    DBNavigator5: TDBNavigator;
    Panel27: TPanel;
    DBNavigator6: TDBNavigator;
    Panel10: TPanel;
    lIgnoreLoadingsSmp: TLabel;
    DBNavigator1: TDBNavigator;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    ActionManager1: TActionManager;
    ImportDataDefinitions1: TMenuItem;
    bbExit: TBitBtn;
    tsEigenvalues: TTabSheet;
    ChartEigenvalue: TChart;
    Series12: TLineSeries;
    Splitter4: TSplitter;
    Panel5: TPanel;
    Splitter5: TSplitter;
    Splitter6: TSplitter;
    Panel14: TPanel;
    Splitter7: TSplitter;
    Panel28: TPanel;
    dbgEigenVec: TDBGrid;
    Panel29: TPanel;
    DBNavigator17: TDBNavigator;
    Label2: TLabel;
    Splitter8: TSplitter;
    Splitter9: TSplitter;
    Splitter10: TSplitter;
    Splitter11: TSplitter;
    ts3D: TTabSheet;
    Splitter12: TSplitter;
    DBChart8: TDBChart;
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
    udZoom: TUpDown;
    Label6: TLabel;
    eZoom: TEdit;
    tsSummary: TTabSheet;
    Panel35: TPanel;
    Panel36: TPanel;
    DBGridStats: TDBGrid;
    Panel37: TPanel;
    DBNavigator24: TDBNavigator;
    Series16: THorizBarSeries;
    Series17: THorizBarSeries;
    Series18: THorizBarSeries;
    Label7: TLabel;
    cbX: TComboBox;
    cbY: TComboBox;
    cbZ: TComboBox;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    sbShow3D: TSpeedButton;
    Panel30: TPanel;
    p3DVar: TPanel;
    DBGrid29: TDBGrid;
    DBNavigator21: TDBNavigator;
    Panel32: TPanel;
    DBGrid30: TDBGrid;
    DBNavigator22: TDBNavigator;
    DBGrid31: TDBGrid;
    DBNavigator23: TDBNavigator;
    Panel38: TPanel;
    DBGrid32: TDBGrid;
    DBNavigator25: TDBNavigator;
    N3: TMenuItem;
    tsLoc3D: TTabSheet;
    Panel39: TPanel;
    Panel41: TPanel;
    DBGrid34: TDBGrid;
    DBNavigator27: TDBNavigator;
    DBGrid35: TDBGrid;
    DBNavigator28: TDBNavigator;
    Panel42: TPanel;
    DBGrid36: TDBGrid;
    DBNavigator29: TDBNavigator;
    Splitter14: TSplitter;
    Panel43: TPanel;
    Splitter15: TSplitter;
    DBChart9: TDBChart;
    Panel44: TPanel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    sbShow3DLoc: TSpeedButton;
    udRotationLoc: TUpDown;
    udElevationLoc: TUpDown;
    udPerspectiveLoc: TUpDown;
    eRotationLoc: TEdit;
    eElevationLoc: TEdit;
    ePerspectiveLoc: TEdit;
    udZoomLoc: TUpDown;
    eZoomLoc: TEdit;
    cbXLoc: TComboBox;
    cbYLoc: TComboBox;
    cbZLoc: TComboBox;
    Panel45: TPanel;
    Splitter16: TSplitter;
    Panel46: TPanel;
    Panel47: TPanel;
    Label18: TLabel;
    DBNavigator30: TDBNavigator;
    dbgEigenVal: TDBGrid;
    Panel48: TPanel;
    lIgnoreLoadings: TLabel;
    tsSimilarity: TTabSheet;
    Panel49: TPanel;
    Panel50: TPanel;
    dbgSimilarity: TDBGrid;
    Panel51: TPanel;
    Label19: TLabel;
    DBNavigator31: TDBNavigator;
    Calculate1: TMenuItem;
    Correspondenceanalysis1: TMenuItem;
    Clusteranalysis1: TMenuItem;
    Project1: TMenuItem;
    ProjectCorrespondenceanalysis2: TMenuItem;
    DBChart10: TDBChart;
    HorizBarSeries1: THorizBarSeries;
    Splitter19: TSplitter;
    Panel40: TPanel;
    Panel52: TPanel;
    Label24: TLabel;
    cbComponent: TComboBox;
    PrintResults1: TMenuItem;
    PrintData1: TMenuItem;
    N4: TMenuItem;
    SaveDialogJPEG: TSaveDialog;
    CombinedRandQmode1: TMenuItem;
    RQmodevariance1: TMenuItem;
    RQmodePearsonscorrelation1: TMenuItem;
    RQmodeSpearmanscorrelation1: TMenuItem;
    RQmodeKendallscorrelation1: TMenuItem;
    PrincipalComponentsAnalysis1: TMenuItem;
    PCAvariance1: TMenuItem;
    PCAPearsonscorrelation1: TMenuItem;
    DiscriminantFunctionAnalysis1: TMenuItem;
    Discriminantanalysis2group1: TMenuItem;
    Discriminantanalysisngroup1: TMenuItem;
    PCASpearmanscorrelation1: TMenuItem;
    PCAKendallscorrelation1: TMenuItem;
    SeriesG81: TPoint3DSeries;
    SeriesG82: TPoint3DSeries;
    SeriesG83: TPoint3DSeries;
    SeriesG84: TPoint3DSeries;
    SeriesG91: TPoint3DSeries;
    SeriesG92: TPoint3DSeries;
    SeriesG93: TPoint3DSeries;
    Panel1: TPanel;
    DBNavigator26: TDBNavigator;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    SeriesG20: TLineSeries;
    SeriesG30: TLineSeries;
    SeriesG80: TPoint3DSeries;
    PrincipalComponentAnalysis1: TMenuItem;
    ProjectPCAvariance2: TMenuItem;
    ProjectPCAPearsonscorrelation2: TMenuItem;
    ProjectPCASpearmanscorrelation2: TMenuItem;
    ProjectPCAKendallscorrelation2: TMenuItem;
    SimultaneousRandQmode1: TMenuItem;
    ProjectRQmodevariance2: TMenuItem;
    ProjectRQmodePearsonscorrelation2: TMenuItem;
    ProjectRQmodeSpearmanscorrelation2: TMenuItem;
    ProjectRQmodeKendallscorrelation2: TMenuItem;
    DiscriminantFunctionAnalysis2: TMenuItem;
    tsLoc4D: TTabSheet;
    Panel54: TPanel;
    Panel57: TPanel;
    DBGrid4: TDBGrid;
    DBNavigator32: TDBNavigator;
    DBGrid5: TDBGrid;
    DBNavigator33: TDBNavigator;
    Panel58: TPanel;
    DBGrid20: TDBGrid;
    DBNavigator34: TDBNavigator;
    Splitter21: TSplitter;
    Panel59: TPanel;
    Splitter22: TSplitter;
    DBChart11: TDBChart;
    SeriesG1103: TPoint3DSeries;
    Panel60: TPanel;
    Label29: TLabel;
    Label30: TLabel;
    Label31: TLabel;
    Label32: TLabel;
    Label33: TLabel;
    Label34: TLabel;
    Label35: TLabel;
    sbShow4D: TSpeedButton;
    udRotation4D: TUpDown;
    udElevation4D: TUpDown;
    udPerspective4D: TUpDown;
    eRotation4D: TEdit;
    eElevation4D: TEdit;
    ePerspective4D: TEdit;
    udZoom4D: TUpDown;
    eZoom4D: TEdit;
    cbX4D: TComboBox;
    cbY4D: TComboBox;
    cbZ4D: TComboBox;
    SeriesG1105: TPoint3DSeries;
    SeriesG1106: TPoint3DSeries;
    SeriesG1107: TPoint3DSeries;
    SeriesG1108: TPoint3DSeries;
    SeriesG1110: TPoint3DSeries;
    SeriesG1109: TPoint3DSeries;
    Label36: TLabel;
    cbVariableID: TComboBox;
    Showvariablelabelsingraphs1: TMenuItem;
    SeriesG10: TLineSeries;
    Defaultvalues1: TMenuItem;
    SeriesG22: TPointSeries;
    DBGrid33: TDBGrid;
    DBNavigator35: TDBNavigator;
    tsOutlierMap: TTabSheet;
    Panel61: TPanel;
    pOutlierMapVar: TPanel;
    Panel63: TPanel;
    DBGrid38: TDBGrid;
    DBNavigator37: TDBNavigator;
    DBGrid39: TDBGrid;
    DBNavigator38: TDBNavigator;
    Panel64: TPanel;
    DBGrid40: TDBGrid;
    DBNavigator39: TDBNavigator;
    Splitter23: TSplitter;
    DBChart12: TDBChart;
    SeriesG1201: TPointSeries;
    SeriesG1202: TPointSeries;
    SeriesG1203: TPointSeries;
    Panel65: TPanel;
    pLocalitiesVar: TPanel;
    SeriesG1205: TLineSeries;
    SeriesG1206: TLineSeries;
    DBGrid37: TDBGrid;
    Button2: TButton;
    bCloseAll: TButton;
    boOpenAll: TButton;
    Button3: TButton;
    ScaleAxesEqually1: TMenuItem;
    Splitter24: TSplitter;
    Panel31: TPanel;
    Panel62: TPanel;
    Splitter25: TSplitter;
    DBChart14: TDBChart;
    SeriesG1400: TBarSeries;
    Label37: TLabel;
    Label38: TLabel;
    Panel66: TPanel;
    Splitter26: TSplitter;
    DBChartHist: TDBChart;
    Panel67: TPanel;
    DBChartQuantile: TDBChart;
    Panel68: TPanel;
    DBChartBox: TDBChart;
    GridBandTool1: TGridBandTool;
    GridBandTool2: TGridBandTool;
    Splitter27: TSplitter;
    Series1501: THistogramSeries;
    Panel69: TPanel;
    Panel70: TPanel;
    cbRawGraphVar: TComboBox;
    Series2: TPointSeries;
    Series3: TLineSeries;
    Series1601: THorizBoxSeries;
    Label39: TLabel;
    Panel71: TPanel;
    Panel72: TPanel;
    Splitter28: TSplitter;
    Splitter29: TSplitter;
    Panel73: TPanel;
    DBChartHistScore: TDBChart;
    HistogramSeries1: THistogramSeries;
    Panel74: TPanel;
    DBChartQuantileScore: TDBChart;
    PointSeries1: TPointSeries;
    LineSeries1: TLineSeries;
    Panel75: TPanel;
    DBChartBoxScore: TDBChart;
    HorizBoxSeries1: THorizBoxSeries;
    GridBandTool3: TGridBandTool;
    GridBandTool4: TGridBandTool;
    Panel76: TPanel;
    Label40: TLabel;
    cbScoreGraphVar: TComboBox;
    Splitter30: TSplitter;
    Calculatequantiles1: TMenuItem;
    N5: TMenuItem;
    EmptyDataTables1: TMenuItem;
    Include4DVariableData1: TMenuItem;
    ProjectDiscriminantanalysis2group2: TMenuItem;
    Discriminantanalysisngroup2: TMenuItem;
    Panel77: TPanel;
    DBChart13: TDBChart;
    PointSeries2: TPointSeries;
    Splitter31: TSplitter;
    Panel78: TPanel;
    DBChart15: TDBChart;
    PointSeries3: TPointSeries;
    Splitter32: TSplitter;
    Button4: TButton;
    bDelete: TButton;
    bDim4Smp: TButton;
    TeeFunction1: TVarianceFunction;
    N2: TMenuItem;
    Export1: TMenuItem;
    Transformeddata1: TMenuItem;
    Summarydata1: TMenuItem;
    N4Ddata1: TMenuItem;
    Samplescores1: TMenuItem;
    VariableLoadings1: TMenuItem;
    Eigenvalues1: TMenuItem;
    Eigenvectors1: TMenuItem;
    Similaritymatrix1: TMenuItem;
    Importexternal1: TMenuItem;
    ImportEigenValues1: TMenuItem;
    ImportEigenvectors1: TMenuItem;
    Importdiscriminantfactors1: TMenuItem;
    ImportMeans1: TMenuItem;
    Discretisedata1: TMenuItem;
    Series1: TMapSeries;
    SeriesTextSource1: TSeriesTextSource;
    TeeCommander1: TTeeCommander;
    ChartEditor1: TChartEditor;
    Button5: TButton;
    N6: TMenuItem;
    Connecttodatabase1: TMenuItem;
    procedure bbEmptyCoranVecClick(Sender: TObject);
    procedure ImportData1Click(Sender: TObject);
    procedure bbEmptyCoranChemClick(Sender: TObject);
    procedure About1Click(Sender: TObject);
    procedure Printersetup1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DataLinkFile1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure dbnGroupedSmpClick(Sender: TObject; Button: TNavigateBtn);
    procedure dbnGroupsClick(Sender: TObject; Button: TNavigateBtn);
    procedure dbnVarClick(Sender: TObject; Button: TNavigateBtn);
    procedure PrintGraph1Click(Sender: TObject);
    procedure bbEmptyCoranVecVarClick(Sender: TObject);
    procedure ImportDataDefinitions1Click(Sender: TObject);
    procedure bbExitClick(Sender: TObject);
    procedure udRotationClick(Sender: TObject; Button: TUDBtnType);
    procedure udElevationClick(Sender: TObject; Button: TUDBtnType);
    procedure udPerspectiveClick(Sender: TObject; Button: TUDBtnType);
    procedure pc1Change(Sender: TObject);
    procedure udZoomClick(Sender: TObject; Button: TUDBtnType);
    procedure sbShow3DClick(Sender: TObject);
    procedure ImportEigenValues1Click(Sender: TObject);
    procedure ImportEigenvectors1Click(Sender: TObject);
    procedure udRotationLocClick(Sender: TObject; Button: TUDBtnType);
    procedure udElevationLocClick(Sender: TObject; Button: TUDBtnType);
    procedure udPerspectiveLocClick(Sender: TObject; Button: TUDBtnType);
    procedure udZoomLocClick(Sender: TObject; Button: TUDBtnType);
    procedure sbShow3DLocClick(Sender: TObject);
    procedure SelectProcess(Sender: TObject);
    procedure ProjectSelectProcess(Sender: TObject);
    procedure cbComponentChange(Sender: TObject);
    procedure PrintResults1Click(Sender: TObject);
    procedure PrintData1Click(Sender: TObject);
    procedure ExportGraph1Click(Sender: TObject);
    procedure sbShow4DClick(Sender: TObject);
    procedure udRotation4DClick(Sender: TObject; Button: TUDBtnType);
    procedure udElevation4DClick(Sender: TObject; Button: TUDBtnType);
    procedure udPerspective4DClick(Sender: TObject; Button: TUDBtnType);
    procedure udZoom4DClick(Sender: TObject; Button: TUDBtnType);
    procedure cbVariableIDChange(Sender: TObject);
    procedure Showvariablelabelsingraphs1Click(Sender: TObject);
    procedure Defaultvalues1Click(Sender: TObject);
    procedure cbVariableIDOutlierMapChange(Sender: TObject);
    procedure sbShowOutlierMapClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure bCloseAllClick(Sender: TObject);
    procedure boOpenAllClick(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure ScaleAxesEqually1Click(Sender: TObject);
    procedure cbRawGraphVarChange(Sender: TObject);
    procedure cbScoreGraphVarChange(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Calculatequantiles1Click(Sender: TObject);
    procedure ExportGraph2Click(Sender: TObject);
    procedure PrintGraph2Click(Sender: TObject);
    procedure TablesDblClick(Sender: TObject);
    procedure DBChartHistDblClick(Sender: TObject);
    procedure DBChartQuantileDblClick(Sender: TObject);
    procedure DBChartBoxDblClick(Sender: TObject);
    procedure ChartEigenvalueDblClick(Sender: TObject);
    procedure DBChartHistScoreDblClick(Sender: TObject);
    procedure DBChartQuantileScoreDblClick(Sender: TObject);
    procedure DBChartBoxScoreDblClick(Sender: TObject);
    procedure DBChart5DblClick(Sender: TObject);
    procedure DBChart12DblClick(Sender: TObject);
    procedure Include4DVariableData1Click(Sender: TObject);
    procedure ImportDiscriminantFactors1Click(Sender: TObject);
    procedure ProjectDiscriminantanalysis2group2Click(Sender: TObject);
    procedure bDeleteClick(Sender: TObject);
    procedure bDim4SmpClick(Sender: TObject);
    procedure Transformeddata1Click(Sender: TObject);
    procedure Summarydata1Click(Sender: TObject);
    procedure N4Ddata1Click(Sender: TObject);
    procedure Samplescores1Click(Sender: TObject);
    procedure Eigenvalues1Click(Sender: TObject);
    procedure Eigenvectors1Click(Sender: TObject);
    procedure Similaritymatrix1Click(Sender: TObject);
    procedure VariableLoadings1Click(Sender: TObject);
    procedure Discretisedata1Click(Sender: TObject);
    procedure ImportMeans1Click(Sender: TObject);
    procedure Connecttodatabase1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    procedure DoProcess;
    procedure CalcComponentScores;
    procedure PrepareComboBoxes(NumAttributes : integer);
    procedure PrepareGraphs;
    procedure Discrm;
    procedure CalcDiscrmScores;
    procedure DiscrimProject;
    procedure CalcProjectedDiscrmScores;
    procedure DiscrimMulti;
    procedure Cluster;
    procedure AddToComboBoxXYZ(i : integer; Choice : string);
    procedure AddToComboBoxVariableID(A : SingleArrayM; i : integer);
    procedure AddToComboBoxVar(Nox : integer);
    procedure AddToComboBoxScoreVar(Nox : integer);
    procedure CalculateProjections;
    procedure CreateEigSprdShts;
    procedure FillSimilaritySpreadSheet;
    procedure CalculateHingeStatistics(NumAttributes : integer);
    procedure CalculateHingeStatisticsForOriginalvariables(NumAttributes : integer);
    procedure Discretise(NumAttributes : integer);
    procedure SetSymbolDefaults;
    procedure AssignChartDataSources(OpenClose : string);
    procedure DisableEnableControls(DisableEnable : string);
    procedure GetMinMax(AxisChosen : string;
                    var AxisMin : extended;
                    var AxisMax : extended);
    procedure CalculateOrthogonalDistances(NumEigenVectorsSelected : integer);
    procedure CalculateScoreOrthogonalCutoffs(NumEigenVectorsSelected : integer);
    procedure ScaleAxesEqually;
    procedure ScaleAxesEquallyGraph1;
    procedure ScaleAxesEquallyGraph2;
    procedure ScaleAxesEquallyGraph3;
    procedure ScaleAxesEquallyGraph3D;
    procedure CheckColumnTotals(N, M : integer; var AllOK : boolean);
    procedure CheckColumnTotalsStats(N, M : integer; var AllOK : boolean);
    procedure ConstructHistogram(iVariable : integer);
    procedure CalculateQuantiles(Nox : integer);
    procedure CalculateQuantilesA(Nox : integer);
    procedure CalculateQuantile(iField, N : integer);
  public
    { Public declarations }
    procedure GetIniFile;
    procedure SetIniFile;
  end;

var
  fmCoranMain: TfmCoranMain;

implementation

uses
  shlobj,
  Cor_ShtIm, Matrix, AllSorts, Cor_ShtIm2,
  Cor_ShtImEigVal, Cor_ShtImEigVec, JPEG, TeeJPEG, mathproc, Cor_def,
  Cor_SelLV, NumRecipes_varb, NumRecipes, Cor_About, icnorm,
  Cor_ShtImdiscrimFac, Cor_dm_flex, Cor_ShtImMean, Cor_dm_acs;

{$R *.DFM}
{$D+}
var
  ImportForm : TfmSheetImport;
  ImportForm2 : TfmSheetImport2;
  ImportFormEigVal : TfmSheetImportEigVal;
  ImportFormEigVec : TfmSheetImportEigVec;
  ImportFormDiscrimFac : TfmSheetImportDiscrimFac;
  ImportFormMean : TfmSheetImportMean;
  AboutForm : TAboutBox;
  DefaultForm : TfmDefaults;
  SelectNumvectorsForm : TfmSelectNumVectors;

procedure TfmCoranMain.DoProcess;
var
  ii, jj, Noxtemp : integer;
  I, J, K, KP : integer;
  tmpStr : string;
  blankstr : string;
  n : integer;
  tmp : double;
  ml : integer;
  XMed, TriMean, HingeL, HingeU : double;
begin
  Noxtemp := Nox;
  if (Nox > 10) then Noxtemp := 10;
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
  if (Printresults1.Checked=true) then IPrn := 'Y'
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
  for ii := Nox+1 to MMaxFields do
  begin
    DBGrid1.Columns[ii+2].Visible := false;
    DBGrid2.Columns[ii+2].Visible := false;
  end;
  for ii := Nox+1 to MMaxFields do
  begin
    dbgFacLoadingsVar.Columns[ii+1].Visible := false;
    DBGridStats.Columns[ii+1].Visible := false;
  end;
  for ii := Nox+1 to MMaxFields do
  begin
    dbgFacLoadingsSmp.Columns[ii+2].Visible := false;
  end;
  for ii := Nox+1 to MMaxFields do
  begin
    dbgEigenVec.Columns[ii].Visible := false;
  end;
  for ii := Nox+1 to MMaxFields do
  begin
    dbgSimilarity.Columns[ii+1].Visible := false;
  end;
  I:=0;
  n := 1;
  TooMuch:=false;
  TotalRecs := dmCor.cdsCoranChem.RecordCount;
  sbmain.Panels[1].Text :='Reading data for '+IntToStr(Nox)+' elements';
  sbMain.Refresh;
  dmCor.cdsCoranChem.First;
  repeat
    if (I<NN) then
    begin
      I:=I+1;
      ComponentStr[I]:=dmCor.cdsCoranChemSampleNum.AsString;
      for J:=1 to Nox do
      begin
        X[I,J] := dmCor.cdsCoranChem.Fields[J+2].AsVariant;
      end;
    end;
    if ((I>=NN)) then TooMuch:=true;
    sbmain.Panels[1].Text :='Processing record '+IntToStr(N)+' of '+IntToStr(TotalRecs)+' Total included '+IntToStr(I);
    sbMain.Refresh;
    n := n+1;
    dmCor.cdsCoranChem.Next;
  until ((n>TotalRecs) or (I>=NN) or dmCor.cdsCoranChem.eof);
  if TooMuch then
  begin
    MessageDlg('Data overflow. Truncating!!',mtWarning,[mbOK],0);
  end;
  n := I;
  NumSamples := I;
  //M:=Nox;
  for I:=1 to Nox-1 do begin
    VarbSymbol[I]:=48+I;
  end;
  VarbSymbol[Nox]:=48;
  if (NumSamples<Nox) then begin
    Nox := NumSamples;
    MessageDlg('Insufficient data in file. Decreasing variables',mtWarning,[mbOK],0);
  end;
  {
  MemoResults.Lines.Add('Input variables are:  ');
  }
  {
  //if (scSimilarityChoice in [scPCAPearsonR,scRQPearsonR]) then Stand(X,NumSamples,Nox);
  }
  i := 1;
  dmCor.ElemNames.Open;
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
  dmCor.ElemNames.First;
  {
  MemoResults.Lines.Add('  ');
  MemoResults.Lines.Add('  ');
  }
  {ensure that none of the values are zero}
  {
  for I:=1 to N do
  begin
    for J:=1 to M do
    begin
      if (X[I,J]=zero) then X[I,J]:=DefaultMinimum;
      X1[i,j] := X[i,j];
    end;
  end;
  }
  if (IPrn='Y') then
  begin
    Write(Lst,'                                                                      CORAN version '+CoranVersion);
    Writeln(Lst,'       '+'Printed on '+DateToStr(Now));
    Writeln(Lst);
    Writeln(Lst,'    ',eTitle.Text);
    Writeln(Lst);
    Writeln(Lst);
    if PrintData1.Checked then
    begin
      Writeln(Lst,'Data Matrix');
      PrintM(ComponentStr,X,NumSamples,Nox);
    end;
  end;
  for I:=1 to Nox do
  begin
    Str(I:10,tempstr);
    EigName[I]:=tempstr;
  end;
  sbmain.Panels[1].Text :='Calculating similarity matrix';
  sbMain.Refresh;
  case scSimilarityChoice of
    scCorrespondence : begin
      ScaleCorAnal(NumSamples,Nox);
    end;
    scPCAVariance : begin
      Cov2(X,A3,NumSamples,Nox);
      {
      ShowMessage('A3 '+IntToStr(N)+' '+IntToStr(M));
      for i := 1 to M do
      begin
        for j := 1 to M do
        begin
          ShowMessage('A3 '+IntToStr(i)+IntToStr(j)+'  '+FormatFloat('#####0.00000',A3[i,j]));
        end;
      end;
      }
    end;
    scPCAPearsonR : begin
      Stand(X,NumSamples,Nox);
      RCoef2(X,A3,NumSamples,Nox);
    end;
    scPCASpearmanR : begin
      SpearmanRho(X,A3,NumSamples,Nox);
    end;
    scPCAKendallR : begin
      KendallTau(X,A3,NumSamples,Nox);
    end;
    scRQVariance : begin
      ScaleRQAnalVar(NumSamples,Nox);
    end;
    scRQPearsonR : begin
      ScaleRQAnalStd(NumSamples,Nox);
    end;
    scRQSpearmanR : begin
      ScaleRQAnalSpearman(NumSamples,Nox);
    end;
    scRQKendallR : begin
    end;
    scDiscrim2Grp : begin
    end;
    scDiscrimnGrp : begin
    end;
    scCluster : begin
    end;
  end;
  case scSimilarityChoice of
    scCorrespondence,scRQVariance,scRQPearsonR : begin
      for i:= 1 to NumSamples do
      begin
        for j:= 1 to Nox do
        begin
          X1[i,j] := W[i,j];
        end;
      end;
      Transp(W,WP,NumSamples,Nox);
      {derive similarity matrix}
      for I:=1 to Nox do
      begin
        for J:=1 to Nox do
        begin
          A3[I,J]:=0.0;
          for K:=1 to NumSamples do
          begin
            A3[I,J]:=A3[I,J]+WP[I,K]*W[K,J];
          end;
        end;
      end;
    end;
    scPCAVariance,scPCAPearsonR : begin
    end;
    scPCASpearmanR,scPCAKendallR : begin
    end;
    scRQSpearmanR,scRQKendallR : begin
      for i:= 1 to NumSamples do
      begin
        for j:= 1 to Nox do
        begin
          X1[i,j] := W[i,j];
        end;
      end;
    end;
  end;
  if (dmCor.CoranSimilarity.RecordCount > 0) then
  begin
    dmCor.CoranSimilarity.Last;
    repeat
      dmCor.CoranSimilarity.Delete;
    until dmCor.CoranSimilarity.Eof;
  end;
  for i := 1 to Nox do
  begin
    dmCor.CoranSimilarity.Append;
    dmCor.CoranSimilarityParamNum.AsInteger := i;
    dmCor.CoranSimilarityParamCalled.AsString := OxideName[i];
    for j := 1 to Nox do
    begin
      dmCor.CoranSimilarity.Fields[j+1].AsVariant := A3[i,j];
    end;
    dmCor.CoranSimilarity.Post;
    dbgSimilarity.Columns[i+1].Title.Caption := Oxidename[i];
  end;
  FillSimilaritySpreadSheet;
  if (IPrn='Y') then
  begin
    Writeln(Lst);
    Writeln(Lst);
    case scSimilarityChoice of
      scCorrespondence : Writeln(Lst,'Similarity Matrix - Correspondence Analysis');
      scRQVariance : Writeln(Lst,'Similarity Matrix - R- Q-mode variance');
      scRQPearsonR : Writeln(Lst,'Similarity Matrix - R- Q-mode Pearsons correlation');
      scPCAVariance : Writeln(Lst,'Similarity Matrix - Principal components - covariance');
      scPCAPearsonR : Writeln(Lst,'Similarity Matrix - Principal components - Pearsons correlation');
      scPCASpearmanR : Writeln(Lst,'Similarity Matrix - Principal components Spearmans correlation');
      scPCAKendallR : Writeln(Lst,'Similarity Matrix - Principal components Kendalls correlation');
      scRQSpearmanR : Writeln(Lst,'Similarity Matrix - R- Q-mode Spearmans correlation');
      scRQKendallR : Writeln(Lst,'Similarity Matrix - R- Q-mode Kendalls correlation');
    end;
    PrintMC(OxideName,A3,Nox,Nox);
  end;
  MemoResults.Lines.Add(' ');
  case scSimilarityChoice of
    scCorrespondence : Memoresults.Lines.Add('Similarity Matrix - Correspondence Analysis');
    scRQVariance : Memoresults.Lines.Add('Similarity Matrix - R- Q-mode variance');
    scRQPearsonR : Memoresults.Lines.Add('Similarity Matrix - R- Q-mode Pearsons correlation');
    scPCAVariance : Memoresults.Lines.Add('Similarity Matrix - Principal components - covariance');
    scPCAPearsonR : Memoresults.Lines.Add('Similarity Matrix - Principal components - Pearsons correlation');
    scPCASpearmanR : Memoresults.Lines.Add('Similarity Matrix - Principal components Spearmans correlation');
    scPCAKEndallR : Memoresults.Lines.Add('Similarity Matrix - Principal components Kendalls correlation');
    scRQSpearmanR : Memoresults.Lines.Add('Similarity Matrix - R- Q-mode Spearmans correlation');
    scRQKendallR : Memoresults.Lines.Add('Similarity Matrix - R- Q-mode Kendalls correlation');
  end;
  Memoresults.Lines.Add(' ');
  {try writing all in separate columns}
  for i := 1 to Nox do
  begin
    tmpstr := '';
    tmpstr := tmpstr + FormatFloat('00',i);
    ResultsArray[i,1] := tmpstr + '   ';
    tmpstr := OxideName[i];
    ResultsArray[i,2] := tmpstr + CharStream(10-Length(tmpstr),32) + blankstr;
    tmpStr := TakeLogs[i];
    ResultsArray[i,3] := ' ' + tmpstr + '  ';
    for J:=1 to Nox do
    begin
      if (A3[i,J] <= 9000000.0) then tmpStr := FormatFloat('###0.000',A3[i,J]);
      if (A3[i,J] > 9000000.0) then tmpStr := FormatFloat('###0.00',A3[i,J]);
      ResultsArray[i,j+3] := CharStream(12-Length(tmpstr),32)+tmpstr + blankstr;
    end;
  end;
  tmpstr := '                     ';
  for i := 1 to Nox do
  begin
    tmpstr := tmpstr + FormatFloat('    00      ',i);
  end;
  MemoResults.Lines.Add(tmpstr);
  Memoresults.Lines.Add(' ');
  for i := 1 to Nox do
  begin
    tmpstr := '';
    for j := 1 to Nox+3 do
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

  // procedure from Davis (1973)
  EigenJ(A3,A2,Nox);   //eigenvalues returned in A3, eigenvectors in A2

  {
  ShowMessage('before VVV = A3');
  for i := 1 to Nox do
  begin
    for j := 1 to Nox do
    begin
      VVV[i,j] := A3[i,j];
    end;
  end;
  ShowMessage('before eigsrt');
  // procedure from Numerical Recipes
  eigsrt(DDD, VVV, Nox);
  ShowMessage('before FillChar');
  FillChar(A3,sizeof(A3),0);
  ShowMessage('before A3 = DDD');
  for j := 1 to Nox do
  begin
    ShowMessage(IntToStr(j)+' '+FormatFloat('####0.0000',DDD[j]));
    A3[j,j] := DDD[j];
  end;
  for i := 1 to Nox do
  begin
    for j := 1 to Nox do
    begin
      //ShowMessage(IntToStr(i)+IntToStr(j)+' '+FormatFloat('####0.0000',VVV[i,j]));
      A2[i,j] := VVV[i,j];
    end;
  end;
  ShowMessage('after A2 = VVV');
  }
  A1[1,1]:=A3[1,1];
  if (scSimilarityChoice=scCorrespondence) then A3[1,1]:=0.0;
  SumE:=0.0;
  for I:=1 to Nox do
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
  for I:=1 to Nox do
  begin
    SumEE:=SumEE+Abs(A1[I,1]);
    A1[I,2]:=Abs(A1[I,1])*100.0/SumE;
    A1[I,3]:=SumEE*100.0/SumE;
    if not ((scSimilarityChoice=scCorrespondence) and (I=1)) then
    begin
      ChartEigenvalue.Series[0].AddXY(1.0*I,A1[I,2]);
    end;
  end;
  if (scSimilarityChoice=scCorrespondence) then
  begin
    A3[1,1]:=1.0;
    A1[1,1]:=1.0;
  end;
  if (scSimilarityChoice=scCorrespondence) then
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
  dmCor.EigenVal.Open;
  if (dmCor.EigenVal.RecordCount > 0) then
  begin
    dmCor.EigenVal.Last;
    repeat
      dmCor.EigenVal.Delete;
    until dmCor.EigenVal.Bof;
  end;
  for i := 1 to Nox do
  begin
    tmpstr := '';
    tmpstr := FormatFloat('    00',i);
    ResultsArray[i,1] := tmpstr + CharStream(6-Length(tmpstr),32) + blankstr;
    dmCor.EigenVal.Append;
    dmCor.EigenValVector.AsInteger := i;
    dmCor.EigenValEigenValue.AsFloat := A1[i,1];
    dmCor.EigenValEigenValuePct.AsFloat := A1[i,2];
    dmCor.EigenValEigenValueCumPct.AsFloat := A1[i,3];
    for J:=1 to 3 do
    begin
      if (A1[i,j] <= 9000000.0) then tmpStr := FormatFloat('###0.0000',A1[i,j]);
      if (A1[i,j] > 9000000.0) then tmpStr := FormatFloat('###0.00',A1[i,j]);
      ResultsArray[i,j+1] := CharStream(12-Length(tmpstr),32)+tmpstr + blankstr;
    end;
    dmCor.EigenVal.Post;
  end;
  for i := 1 to Nox do
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
    if (scSimilarityChoice=scCorrespondence) then
    begin
      Writeln(Lst,'Eigenvalue 1 and Eigenvector 1 are artifices of the method. Ignore them!');
      Writeln(Lst);
    end;
    Writeln(Lst,'Column 1 = Eigenvalues,  Column 2 = % of trace,  Column 3 = cum. % of trace');
    PrintMC(EigName,A1,Nox,3);
    Writeln(Lst);
    Writeln(Lst);
    Writeln(Lst,'Principal axis matrix  -  columns = eigenvectors, rows = variables');
    PrintMC(OxideName,A2,Nox,Nox);
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
      dmCor.EigenVec.Fields[jj].AsVariant := A2[ii,jj];
    end;
    dmCor.EigenVec.Post;
  end;
  sbmain.Panels[1].Text :='Calculating factor loadings';
  sbMain.Refresh;
  case scSimilarityChoice of
    scCorrespondence : begin
      for J:=1 to Nox do
      begin
        A3[J,J]:=Sqrt(Abs(A3[J,J]));
      end;
      MmultC(A2,A3,TempC,Nox,Nox,Nox);  {factor loadings for variables}
      for J:=1 to Nox do
      begin
        A3[J,J]:=1.0/A3[J,J];
      end;
      MmultC(DC,TempC,A1,Nox,Nox,Nox);  {scaled factor loadings for variables}
      Mmult(W,A2,B,NumSamples,Nox,Nox);
      Mmult(B,A3,W,NumSamples,Nox,Nox);
      MmultR(DR,W,TempM,NumSamples,Nox);     {factor loadings for samples}
      for J:=1 to Nox do
      begin
        A3[J,J]:=1.0/A3[J,J];
      end;
      Mmult(TempM,A3,B,NumSamples,Nox,Nox);    {scaled factor loadings for samples}
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
    scPCAVariance, scPCAPearsonR : begin
      {
      //solution 1 approach of Koreskog et al 1976
      //this solution scales the scores and loadings
      for J:=1 to M do
      begin
        for I:=1 to M do
        begin
          A3[I,J] := 0.0;
        end;
        A3[J,J]:=Sqrt(Abs(A1[J,1]));
      end;
      MmultC(A2,A3,A1,M,M,M);  //factor loadings for variables
      for J:=1 to M do
      begin
        DD[J] := 0.0;
        for I:=1 to N do
        begin
          DD[J]:=DD[J] + X[I,J];
        end;
        DD[J] := DD[J] / (1.0*N);
        for I:=1 to N do
        begin
          X[I,J]:=X[I,J]-DD[J];
        end;
      end;
      Mmult(X,A1,B,N,M,M);     //factor scores for samples
      //now scale columns by reciprocals of eigenvalues
      for J:=1 to M do
      begin
        tmp := A3[J,J]*A3[J,J];
        if (tmp <= 0.0) then tmp := 1.0;
        for I:=1 to N do
        begin
          B[I,J] := B[I,J]/tmp;
        end;
      end;
      }
      //solution 2 approach of Koreskog et al 1976
      //this is the approach used by ROBPCA classical
      //factor loadings for variables
      for J:=1 to Nox do
      begin
        for I:=1 to Nox do
        begin
          A1[I,J]:=A2[I,J];
        end;
      end;
      for J:=1 to Nox do
      begin
        DD[J] := 0.0;
        for I:=1 to NumSamples do
        begin
          DD[J]:=DD[J] + X[I,J];
        end;
        DD[J] := DD[J] / (1.0*N);
        for I:=1 to NumSamples do
        begin
          X[I,J]:=X[I,J]-DD[J];
        end;
      end;
      Mmult(X,A2,B,NumSamples,Nox,Nox);     //factor scores for samples

    end;
    {
    scPCAPearsonR : begin
      for J:=1 to M do
      begin
        for I:=1 to M do
        begin
          A3[I,J] := 0.0;
        end;
        A3[J,J]:=Sqrt(Abs(A1[J,1]));
      end;
      MmultC(A2,A3,A1,M,M,M);  //factor loadings for variables
      Mmult(X,A1,B,N,M,M);     //factor scores for samples
      //now scale columns by reciprocals of eigenvalues
      for J:=1 to M do
      begin
        tmp := A3[J,J]*A3[J,J];
        if (tmp <= 0.0) then tmp := 1.0;
        for I:=1 to N do
        begin
          B[I,J] := B[I,J]/tmp;
        end;
      end;
    end;
    }
    scPCASpearmanR,scPCAKendallR : begin
      //solution 1 approach of Koreskog et al 1976
      for J:=1 to Nox do
      begin
        for I:=1 to Nox do
        begin
          A3[I,J] := 0.0;
        end;
        A3[J,J]:=Sqrt(Abs(A1[J,1]));
      end;
      MmultC(A2,A3,A1,Nox,Nox,Nox);  //factor loadings for variables

      for j := 1 to Nox do
      begin
        //dmCor.qDim4Smp.Close;
        dmCor.cdsqDim4Smp.Filtered := false;
        dmCor.cdsqDim4Smp.Filter := 'VariableID = '+''''+'Vector'+IntToStr(j)+'''';
        dmCor.cdsqDim4Smp.Filtered := true;
        for i := 1 to NumSamples do
        begin
          DR[i] :=  dmCor.cdsqDim4SmpSmpValue.AsFloat;
          dmCor.cdsqDim4Smp.Next;
        end;
        Sort(DR,NumSamples,1);
        ml := (NumSamples+1) div 2;
        XMed := 0.5*(DR[ml] + DR[NumSamples-ml+1]);
        Hinges(DR,HingeL,HingeU,NumSamples);
        TriMean := 0.5*XMed + 0.25*HingeL + 0.25*HingeU;
        DD[j] := TriMean;
      end;
      for J:=1 to Nox do
      begin
        for I:=1 to NumSamples do
        begin
          X[I,J]:=X[I,J]-DD[J];       //use TriMean here as a robust estimate
        end;
      end;
      Mmult(X,A1,B,NumSamples,Nox,Nox);     //factor scores for samples
      //now scale columns by reciprocals of eigenvalues
      for J:=1 to Nox do
      begin
        tmp := A3[J,J]*A3[J,J];
        if (tmp <= 0.0) then tmp := 1.0;
        for I:=1 to NumSamples do
        begin
          B[I,J] := B[I,J]/tmp;
        end;
      end;
      {
      for J:=1 to M do
      begin
        A3[J,J]:=Sqrt(Abs(A3[J,J]));
      end;
      MmultC(A2,A3,A1,M,M,M);  //actor loadings for variables
      for J:=1 to M do
      begin
        TempC[J,J]:=1.0/(A3[J,J]*A3[J,J]);
      end;
      Mmult(X,A2,B,N,M,M);     //factor scores for samples
      //now scale columns by reciprocals of eigenvalues
      for J:=1 to M do
      begin
        tmp := A3[J,J]*A3[J,J];
        if (tmp <= 0.0) then tmp := 1.0;
        for I:=1 to N do
        begin
          B[I,J] := B[I,J]/tmp;
        end;
      end;
    }
    end;
    scRQVariance,scRQPearsonR : begin
      for J:=1 to Nox do
      begin
        A3[J,J]:=Sqrt(Abs(A3[J,J]));
      end;
      MmultC(A2,A3,A1,Nox,Nox,Nox);  {factor loadings for variables}
      Mmult(W,A2,B,NumSamples,Nox,Nox);     {factor loadings for samples}
    end;
    scRQSpearmanR,scRQKendallR : begin
      for J:=1 to Nox do
      begin
        A3[J,J]:=Sqrt(Abs(A3[J,J]));
      end;
      MmultC(A2,A3,A1,Nox,Nox,Nox);  {factor loadings for variables}
      Mmult(W,A2,B,NumSamples,Nox,Nox);     {factor loadings for samples}
    end;
  end;
  sbmain.Panels[1].Text :='Calculating component scores';
  sbMain.Refresh;
  CalcComponentScores;
  dmCor.QGroups.Open;
  dmCor.QPlotGroups.Open;
  if (IPrn='Y') then
  begin
    if (scSimilarityChoice in [scCorrespondence,scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR]) then
    begin
      Writeln(Lst,'Factor loadings  -  columns = factors, rows = variables');
      PrintMC(OxideName,A1,Nox,Nox);
    end;
    if PrintData1.Checked then
    begin
      Writeln(Lst,'Factor loadings  -  columns = factors, rows = samples');
      PrintM(ComponentStr,B,NumSamples,Nox);
    end;
  end;
  if ((scSimilarityChoice=scCorrespondence) and (Nox>4)) then
  begin
    for I:=1 to NumSamples do
    begin
     DR[I]:=1.0/DR[I];
     DR[I]:=DR[I]*DR[I];
    end;
    for J:=1 to Nox do
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
     for J:=1 to Nox do
     begin
       D:=0.0;
       for K:=2 to Nox do
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
     for I:=1 to NumSamples do
     begin
       D:=0.0;
       for K:=2 to Nox do
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
         Write(Lst,ComponentStr[I]:11,DR[I]:11:4);
         for K:=1 to 4 do
         begin
           Write(Lst,CA[K]:11:4,CRR[K]:11:4);
         end;
         Writeln(Lst);
       end;
     end;
    end;
  end;
  pc1.ActivePage := tsEigenValues;
  NumEigenVectorsSelected := 2;  // need to make this variable ***********
  try
    SelectNumvectorsForm := TfmSelectNumvectors.Create(Self);
    SelectNumvectorsForm.ShowModal;
  finally
    SelectNumvectorsForm.Free;
  end;
  pc1.ActivePage := tsSummary;
  //NumAttributes := NumEigenVectorsSelected;
  CalculateOrthogonalDistances(NumEigenVectorsSelected);
  pVar.Visible := true;
  pGraph1Var.Visible := true;
  pGraph2Var.Visible := true;
  pGraph3Var.Visible := true;
  pOutlierMapVar.Visible := true;
  p3DVar.Visible := true;
  pLocalitiesVar.Visible := true;
  case scSimilarityChoice of
    scCorrespondence : begin
       lIgnoreLoadings.Visible := true;
       lIgnoreLoadingsVar.Visible := true;
       lIgnoreLoadingsSmp.Visible := true;
       NX:=2;
    end;
    scRQVariance,scRQPearsonR : begin
       lIgnoreLoadings.Visible := false;
       lIgnoreLoadingsVar.Visible := false;
       lIgnoreLoadingsSmp.Visible := false;
       NX:=1;
    end;
    scPCAVariance,scPCAPearsonR : begin
       lIgnoreLoadings.Visible := false;
       lIgnoreLoadingsVar.Visible := false;
       lIgnoreLoadingsSmp.Visible := false;
       pVar.Visible := true;
       pGraph1Var.Visible := false;
       pGraph2Var.Visible := false;
       pGraph3Var.Visible := false;
       pOutlierMapVar.Visible := false;
       p3DVar.Visible := false;
       pLocalitiesVar.Visible := false;
       NX:=1;
    end;
    scDiscrim2Grp,scDiscrimnGrp,scCluster : begin
       lIgnoreLoadings.Visible := false;
       lIgnoreLoadingsVar.Visible := false;
       lIgnoreLoadingsSmp.Visible := false;
       NX:=1;
    end;
    scPCASpearmanR,scPCAKendallR : begin
       lIgnoreLoadings.Visible := false;
       lIgnoreLoadingsVar.Visible := false;
       lIgnoreLoadingsSmp.Visible := false;
       pVar.Visible := true;
       pGraph1Var.Visible := false;
       pGraph2Var.Visible := false;
       pGraph3Var.Visible := false;
       pOutlierMapVar.Visible := false;
       p3DVar.Visible := false;
       pLocalitiesVar.Visible := false;
       NX:=1;
    end;
    scRQSpearmanR,scRQKendallR : begin
       lIgnoreLoadings.Visible := false;
       lIgnoreLoadingsVar.Visible := false;
       lIgnoreLoadingsSmp.Visible := false;
       NX:=1;
    end;
  end;
  case scSimilarityChoice of
    scCorrespondence,scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
       pVar.Visible := true;
       Xmin:=A1[1,NX];
       Xmax:=Xmin;
       Ymin:=A1[1,NX+1];
       Ymax:=Ymin;
       Zmin:=A1[1,NX+2];
       Zmax:=Zmin;
       for J:=1 to Nox do
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
    scPCAVariance,scPCAPearsonR : begin
       pVar.Visible := true;
       Xmin:=B[1,NX];
       Xmax:=Xmin;
       Ymin:=B[1,NX+1];
       Ymax:=Ymin;
       Zmin:=B[1,NX+2];
       Zmax:=Zmin;
    end;
    scPCASpearmanR,scPCAKendallR : begin
       pVar.Visible := true;
       Xmin:=B[1,NX];
       Xmax:=Xmin;
       Ymin:=B[1,NX+1];
       Ymax:=Ymin;
       Zmin:=B[1,NX+2];
       Zmax:=Zmin;
    end;
  end;
  for I:=1 to NumSamples do
  begin
    if (Xmin > B[I,NX]) then Xmin:=B[I,NX];
    if (Xmax < B[I,NX]) then Xmax:=B[I,NX];
    if (Ymin > B[I,NX+1]) then Ymin:=B[I,NX+1];
    if (Ymax < B[I,NX+1]) then Ymax:=B[I,NX+1];
    if (Zmin > B[I,NX+2]) then Zmin:=B[I,NX+2];
    if (Zmax < B[I,NX+2]) then Zmax:=B[I,NX+2];
    for K:=1 to 5 do
    begin
      //X[I+Nox,K]:=B[I,K];  //as it used to be Oct 2006. Not sure why I+Nox
      X[I,K]:=B[I,K];
    end;
  end;
  tsGraph1.TabVisible := false;
  sbmain.Panels[1].Text :='Preparing graphs';
  sbMain.Refresh;
  tsGraph2.TabVisible := false;
  tsGraph3.TabVisible := false;
  tsLocalities.TabVisible := false;
  tsScores.TabVisible := false;
  pVar.Visible := true;
  pGraph1Var.Visible := true;
  pOutlierMapVar.Visible := true;
  p3DVar.Visible := true;
  pLocalitiesVar.Visible := true;
  case scSimilarityChoice of
    scCorrespondence : begin
      tsGraph1.TabVisible := true;
      tsGraph2.TabVisible := true;
      tsGraph3.TabVisible := true;
      tsLocalities.TabVisible := true;
      tsScores.TabVisible := true;
    end;
    scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
      tsGraph1.TabVisible := true;
      tsGraph2.TabVisible := true;
      tsGraph3.TabVisible := true;
      tsScores.TabVisible := true;
      tsLocalities.TabVisible := true;
    end;
    scPCAVariance,scPCAPearsonR,scPCASpearmanR, scPCAKendallR : begin
      tsGraph1.TabVisible := true;
      tsGraph2.TabVisible := true;
      tsGraph3.TabVisible := true;
      tsScores.TabVisible := false;
      tsLocalities.TabVisible := true;
      pVar.Visible := true;
      pGraph1Var.Visible := false;
      pGraph2Var.Visible := false;
      pGraph3Var.Visible := false;
      pOutlierMapVar.Visible := false;
      p3DVar.Visible := false;
      pLocalitiesVar.Visible := false;
    end;
  end;
  PrepareGraphs;
  sbmain.Panels[1].Text :='Done preparing graphs';
  sbMain.Refresh;
  if (IPrn='Y') then
  begin
   Writeln(Lst,char(13));
   System.CloseFile(Lst);
  end;
end;

procedure TfmCoranMain.CalcComponentScores;
var
  ii, I, J : integer;
  {
  A1t : RealArrayC;
  }
  SumE, SumEE : double;
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
  {clear factor loadings table for samples}
  dmCor.cdsFacLoadingsSmp.Last;
  if not (dmCor.cdsFacLoadingsSmp.Bof and dmCor.cdsFacLoadingsSmp.Eof) then
  begin
    dmCor.cdsFacLoadingsSmp.Last;
    repeat
      dmCor.cdsFacLoadingsSmp.Delete;
    until dmCor.cdsFacLoadingsSmp.Bof;
  end;
  dmCor.cdsFacLoadingsSmp.ApplyUpdates(-1);
  if (scSimilarityChoice in [scCorrespondence,scRQVariance,scRQPearsonR,
                   scPCAVariance,scPCAPearsonR,scPCASpearmanR,
                   scPCAKendallR,scRQSpearmanR,scRQKendallR]) then
  begin
    {fill factor loadings table for variables}
    dmCor.FacLoadingsVar.First;
    ii := 0;
    for i := 1 to Nox do
    begin
      ii := ii + 1;
      dmCor.FacLoadingsVar.Append;
      dmCor.FacLoadingsVarPos.AsInteger := (i*2) - 1;
      dmCor.FacLoadingsVarCalled.AsString := OxideName[i];
      for j := 1 to Nox do
      begin
        dmCor.FacLoadingsVar.Fields[j+1].AsVariant:= A1[i,j];
      end;
      dmCor.FacLoadingsVar.Post;
      {insert a record of zero values between each
      record to define origin for plotting
      }
      dmCor.FacLoadingsVar.Append;
      dmCor.FacLoadingsVarPos.AsInteger := (i*2);
      dmCor.FacLoadingsVarCalled.AsString := ' ';
      for j := 1 to Nox do
      begin
        dmCor.FacLoadingsVar.Fields[j+1].AsVariant:= 0.0;
      end;
      dmCor.FacLoadingsVar.Post;
      {insert a row of zero values between each
      variable in spreadsheet to define origin for Grapher plotting
      }
      ii := ii + 1;
    end;
    dmCor.FacLoadingsVar.First;
  end;
  {fill factor loadings table for samples}
  dmCor.cdsCoranChem.First;
  i := 1;
  dmCor.CoranVecLinked.Open;
  dmCor.cdsCoranChem.First;
  repeat
    dmCor.CoranVecLinked.Append;
    dmCor.CoranVecLinkedGROUPNAME.AsString := dmCor.cdsCoranChemGroupName.AsString;
    dmCor.CoranVecLinkedSampleNum.AsString := dmCor.cdsCoranChemSampleNum.AsString;
    dmCor.CoranVecLinkedSequence.AsInteger := i;
    for j := 1 to Nox do
    begin
      dmCor.CoranVecLinked.Fields[j+2].AsVariant := B[i,j];
    end;
    dmCor.CoranVecLinked.Post;
    dmCor.cdsCoranChem.Next;
    i := i + 1;
  until dmCor.cdsCoranChem.Eof;
  dmCor.cdsCoranChem.First;
  dmCor.cdsFacLoadingsSmp.First;
  dmCor.cdsFacLoadingsSmp.Close;
  dmCor.cdsFacLoadingsSmp.Open;
  if (scSimilarityChoice in [scCorrespondence,scRQVariance,scRQPearsonR,
                   scPCAVariance,scPCAPearsonR,scPCASpearmanR,
                   scPCAKendallR,scRQSpearmanR,scRQKendallR]) then
  begin
    dmCor.EigenVal.First;
    i := 0;
    repeat
      i := i+1;
      TempC[i,1] := dmCor.EigenValEigenValue.AsFloat;   //was A1t
      TempC[i,2] := dmCor.EigenValEigenValuePct.AsFloat;
      TempC[i,3] := dmCor.EigenValEigenValueCumPct.AsFloat;
      dmCor.EigenVal.Next;
    until dmCor.EigenVal.Eof;
    dmCor.EigenVal.First;
    if (scSimilarityChoice=scCorrespondence) then TempC[1,1]:=0.0;
    SumE:=0.0;
    for I:=1 to Nox do
    begin
      SumE:=SumE+Abs(TempC[I,1]);
    end;
    TempC[1,2]:=0.0;
    TempC[1,3]:=0.0;
    SumEE:=0.0;
    for I:=1 to Nox do
    begin
      SumEE:=SumEE+Abs(TempC[I,1]);
      TempC[I,2]:=Abs(TempC[I,1])*100.0/SumE;
      TempC[I,3]:=SumEE*100.0/SumE;
    end;
    if (scSimilarityChoice=scCorrespondence) then
    begin
      TempC[1,1]:=1.0;
    end;
    {fill eigen values in spreadsheet}
    ii := 0;
  end;
  if (scSimilarityChoice in [scCorrespondence,scRQVariance,scRQPearsonR,
                   scPCAVariance,scPCAPearsonR,scPCASpearmanR,
                   scPCAKendallR,scRQSpearmanR,scRQKendallR]) then
  begin
    dmCor.EigenVec.First;
    for i := 1 to Nox do
    begin
      for j := 1 to Nox do
      begin
        TempC[i,j] := dmCor.EigenVec.Fields[j+1].AsFloat;
      end;
      dmCor.EigenVec.Next;
    end;
    dmCor.EigenVec.First;
    if (scSimilarityChoice=scCorrespondence) then
    begin
      for i := 1 to Nox do TempC[i,1]:=0.0;
    end;
    {fill eigen vectors in spreadsheet}
    dmCor.EigenVec.First;
  end;
end;

procedure TfmCoranMain.bbEmptyCoranVecClick(Sender: TObject);
begin
  //dmCor.cdsFacLoadingsSmp.DisableControls;
  dmCor.cdsFacLoadingsSmp.Last;
  if not(dmCor.cdsFacLoadingsSmp.BOF and dmCor.cdsFacLoadingsSmp.EOF) then
  begin
    dmCor.cdsFacLoadingsSmp.Last;
    repeat
      dmCor.cdsFacLoadingsSmp.Delete;
    until dmCor.cdsFacLoadingsSmp.BOF;
  end;
  dmCor.cdsFacLoadingsSmp.EnableControls;
end;

procedure TfmCoranMain.ImportData1Click(Sender: TObject);
var
  i, ii : integer;
  DataImported : boolean;
begin
  sbMain.Panels[1].Text := 'Importing';
  sbMain.Refresh;
  tsResults.TabVisible := false;
  tsSimilarity.TabVisible := false;
  tsEigenValues.TabVisible := false;
  tsLoadings.TabVisible := false;
  tsScores.TabVisible := false;
  tsGraph1.TabVisible := false;
  tsGraph2.TabVisible := false;
  tsGraph3.TabVisible := false;
  ts3D.TabVisible := false;
  tsLoc4D.TabVisible := Include4DVarData;

  AssignChartDataSources('Close');

  dmCor.DeleteDim4Smp.SQL.Clear;
  dmCor.DeleteDim4Smp.SQL.Add('delete * from Dimension4Smp');
  dmCor.DeleteDim4Smp.ExecSQL;
  dmCor.cdsqDim4Smp.Close;
  dmCor.cdsqDim4Smp.Open;
  for i := 0 to MM+2 do
  begin
    try
      DBGrid1.Columns[i].Visible := true;
      DBGrid2.Columns[i].Visible := true;
    except
      ShowMessage('DBGrid1 or 2 problem');
    end;
  end;
  for i := 1 to MM+1 do
  begin
    try
      DBGridStats.Columns[i].Visible := true;
    except
      ShowMessage('DBGridStats problem');
    end;
  end;
  try
    try
      dmCor.ImportGroup.Open;
      dmCor.ImportGroup.First;
      dmCor.CoranFac.Open;
    except
    end;
    try
      ImportForm := TfmSheetImport.Create(Self);
      ImportForm.OpenDialogSprdSheet.FileName := 'CoranChem';
      //ImportForm.FillData;
      if (ImportForm.ShowModal = mrOK) then DataImported := true
                                       else DataImported := false;
    finally
      //ImportForm.FlexCelImport1.CloseFile;
    end; //finally
  finally
    ImportForm.Free;
    try
      dmCor.ImportGroup.Close;
      dmCor.CoranFac.Close;
      dmCor.QGroups.Close;
      //dmCor.QGroups.Open;
      dmCor.QPlotGroups.Close;
      //dmCor.QPlotGroups.Open;
      dmCor.GroupedSmpLoc.Close;
      //dmCor.GroupedSmpLoc.Open;
      dmCor.FacLoadingsVar.Close;
      dmCor.FacLoadingsVarNoZero.Close;
      dmCor.FacLoadingsVarLinked.Close;
      //dmCor.FacLoadingsVarLinked2.Close;
      dmCor.cdsFacLoadingsSmp.Close;
    except
    end;
  end;
  if DataImported then
  begin
    fmCoranMain.Refresh;
    dmCor.ElemNames.Close;
    dmCor.ElemNames.Open;
    Nox := dmCor.ElemNames.RecordCount;
    dmCor.cdsCoranChem.Open;
    NumSamples := dmCor.cdsCoranChem.RecordCount;
    sbMain.Panels[1].Text := 'Finished';
    sbMain.Refresh;
    for i := 1 to Nox do
    begin
      DBGrid1.Columns[i+2].Visible := true;
      DBGrid2.Columns[i+2].Visible := true;
      DBGridStats.Columns[i+1].Visible := true;
      dbgSimilarity.Columns[i+1].Visible := true;
      dbgEigenVec.Columns[i].Visible := true;
      dbgFacLoadingsVar.Columns[i+1].Visible := true;
      dbgFacLoadingsSmp.Columns[i+2].Visible := true;
    end;
    sbMain.Panels[1].Text := 'Finished adjusting column visibility';
    sbMain.Refresh;
    sbMain.Panels[1].Text := 'Adjust DBGridStats column visibility';
    sbMain.Refresh;
    for ii := Nox+1 to MMaxFields do
    begin
      try
        DBGridStats.Columns[ii+1].Visible := false;
      except
        ShowMessage('DBGridStats problem second');
      end;
    end;
    sbMain.Panels[1].Text := 'Finished adjusting DBGridStats column visibility';
    sbMain.Refresh;
    sbMain.Panels[1].Text := 'Adjusting other grids column visibility';
    sbMain.Refresh;
    dmCor.ElemNames.First;
    for ii := 1 to Nox do
    begin
      try
        DBGrid1.Columns[ii+2].Title.Caption := dmCor.ElemNamesCalled.AsString;
        DBGrid2.Columns[ii+2].Title.Caption := dmCor.ElemNamesCalled.AsString;
      except
        ShowMessage('DBGrid1 column title problem');
      end;
      try
        DBGridStats.Columns[ii+1].Title.Caption := dmCor.ElemNamesCalled.AsString;
      except
        ShowMessage('DBGridStats column title problem');
      end;
      try
        dbgSimilarity.Columns[ii+1].Title.Caption := dmCor.ElemNamesCalled.AsString;
      except
        ShowMessage('dbgSimilarity column title problem');
      end;
      dmCor.ElemNames.Next;
    end;
    AddToComboBoxVar(Nox);
    sbMain.Panels[1].Text := 'Check column totals';
    sbMain.Refresh;
    //CheckColumnTotals;
    AssignChartDataSources('Close');
    sbMain.Panels[1].Text := 'Calculate histogram values';
    sbMain.Refresh;
    //dmCor.cdsCoranChem.DisableControls;
    //AssignChartDataSources('Open');
    for i := 1 to Nox do
    begin
      ConstructHistogram(i);
    end;
    sbMain.Panels[1].Text := 'Calculate quantiles';
    sbMain.Refresh;
    if MakeQQPlot then CalculateQuantiles(Nox);
    sbMain.Panels[1].Text := 'Calculate values for 4D original variables';
    sbMain.Refresh;
    CalculateHingeStatisticsForOriginalVariables(Nox);
    //AssignChartDataSources('Close');
    dmCor.cdsCoranChem.EnableControls;
    try
      dmCor.cdsCoranChem.EnableControls;
    except
      ShowMessage('Problem enablecontrols CoranChem');
    end;
    try
      dmCor.ElemNames.EnableControls;
    except
      ShowMessage('Problem enablecontrols ElemNames');
    end;
    try
      dmCor.cdsFacLoadingsSmp.EnableControls;
    except
      ShowMessage('Problem enablecontrols FacLoadingsSmp');
    end;
    try
      dmCor.FacLoadingsVar.EnableControls;
    except
      ShowMessage('Problem enablecontrols FacLoadingsVar');
    end;
    try
      dmCor.GroupedSmp.EnableControls;
    except
      ShowMessage('Problem enablecontrols GroupedSmp');
    end;
    try
      dmCor.QGroupedSmp.EnableControls;
    except
      ShowMessage('Problem enablecontrols QGroupedSmp');
    end;
    try
      dmCor.SmpLoc.EnableControls;
    except
      ShowMessage('Problem enablecontrols SmpLoc');
    end;
    try
      dmCor.GroupedSmpLoc.EnableControls;
    except
      ShowMessage('Problem enablecontrols GroupedSmpLoc');
    end;
    try
      dmCor.QGroupedSmpLoc.EnableControls;
    except
      ShowMessage('Problem enablecontrols QGroupedSmpLoc');
    end;
//    dmCor.cdsCoranChem.First;
//    AssignChartDataSources('Close');
    scSimilarityChoice := scNone;
    PrepareComboBoxes(0);
    PrepareGraphs;
    //AssignChartDataSources('Open');
    try
      dmCor.GroupedSmpLoc.EnableControls;
    except
      ShowMessage('Problem enablecontrols GroupedSmpLoc again');
    end;
    sbMain.Panels[1].Text := 'New data imported';
    sbMain.Refresh;
  end else
  begin
    sbMain.Panels[1].Text := 'Import cancelled';
    sbMain.Refresh;
  end;
end;

procedure TfmCoranMain.bbEmptyCoranChemClick(Sender: TObject);
begin
  //dmCor.cdsCoranChem.DisableControls;
  //dmCor.cdsCoranRaw.DisableControls;
  dmCor.DeleteCoranChem.ExecSQL;
  dmCor.cdsCoranChem.Close;
  dmCor.cdsCoranChem.Open;
  dmCor.DeleteCoranRaw.ExecSQL;
  dmCor.cdsCoranRaw.Close;
  dmCor.cdsCoranRaw.Open;
  dmCor.cdsCoranChem.EnableControls;
  dmCor.cdsCoranRaw.EnableControls;
  dmCor.SmpLoc.Open;
  if not(dmCor.SmpLoc.Bof and dmCor.SmpLoc.Eof) then
  begin
    dmCor.SmpLoc.Last;
    repeat
      dmCor.SmpLoc.Delete;
      dmCor.SmpLoc.Next;
    until dmCor.SmpLoc.BOF;
  end;
  dmCor.FacLoadingsVar.Open;
  if not(dmCor.FacLoadingsVar.Bof and dmCor.FacLoadingsVar.Eof) then
  begin
    dmCor.FacLoadingsVar.Last;
    repeat
      dmCor.FacLoadingsVar.Delete;
      dmCor.FacLoadingsVar.Next;
    until dmCor.FacLoadingsVar.BOF;
  end;
  dmCor.cdsCoranChem.First;
  dmCor.cdsCoranRaw.First;
  dmCor.SmpLoc.First;
  dmCor.FacLoadingsVar.First;
end;

procedure TfmCoranMain.About1Click(Sender: TObject);
begin
  AboutForm := TAboutBox.Create(self);
  try
    AboutForm.ShowModal;
  finally
    AboutForm.Free;
  end;
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
  //zpath: array [0..MAX_PATH] of char;
  PublicPath : string;
begin
  PublicPath := TPath.GetHomePath;
  dmCor.CommonFilePath := IncludeTrailingPathDelimiter(PublicPath) + 'EggSoft\';
  dmCor.IniFilename := dmCor.CommonFilePath + 'Coran.ini';
  //IniFilePath := CommonFilePath;
  //ProgramFilePath := IniFilePath + 'Coran\';
  AppIni := TIniFile.Create(dmCor.IniFilename);
  //uses ShlObj
  // this gives access to all the systemed defined folders, no direct dependency on env. variables.
  //SHGetFolderPath(0, CSIDL_COMMON_APPDATA or CSIDL_FLAG_CREATE, 0, SHGFP_TYPE_CURRENT, @zpath);
  //dmCor.CommonFilePath := IncludeTrailingPathDelimiter(string(zpath)) + 'EggSoft\Coran\';
  //dmCor.IniFilename := IncludeTrailingPathDelimiter(string(zpath)) + 'EggSoft\Coran\'+'Coran.INI';
  //dmCor.ConnectionString1 := 'Provider=Microsoft.ACE.OLEDB.15.0;Data Source=';      //for Office 64 bit
  //dmCor.ConnectionString2 := ';Mode=ReadWrite;Persist Security Info=False;';
  dmCor.ConnectionString1 := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=';
  dmCor.ConnectionString2 := ';Mode=ReadWrite;Persist Security Info=False;';
  AppIni := TIniFile.Create(dmCor.IniFilename);
  try
    dmCor.ConnectionString := AppIni.ReadString('ADO','Connection string',dmCor.ConnectionString1+dmCor.CommonFilePath+'Coran\Database\Coran.mdb'+dmCor.ConnectionString2);
    //dmCor.ConnectionString := AppIni.ReadString('ADO','Connection string','Provider=Microsoft.ACE.OLEDB.15.0;Data Source=C:\ProgramData\EggSoft\Coran\Database\Coran.mdb;Mode=ReadWrite;Persist Security Info=False;');
    DataPath := AppIni.ReadString('File Paths','Data spreadsheets',dmCor.CommonFilePath+'Coran\Data\');
    //cdsPath := AppIni.ReadString('File Paths','Internal files',dmCor.CommonFilePath+'Data\');
    //ImportPath := AppIni.ReadString('File Paths','Data import path',dmCor.CommonFilePath);
    ExportPath := AppIni.ReadString('File Paths','Data export path',dmCor.CommonFilePath+'Coran\');
    FlexTemplatePath := AppIni.ReadString('File Paths','Template path',dmCor.CommonFilePath+'Coran\Templates\');
    //DataPath := AppIni.ReadString('Paths','Data path','C:\');
    //ExportPath := AppIni.ReadString('Paths','Spreadsheet exports path','C:\');
    //FlexTemplatePath := AppIni.ReadString('Paths','Spreadsheet template path','C:\ProgramFiles\EggSoft\Coran\Templates\');
    JPEGPath := AppIni.ReadString('Paths','JPEG exports path','C:\');
    ImportSpecNameColStr := AppIni.ReadString('ColumnDefinitions','ImportSpecNameColStr','A');
    PositionColStr := AppIni.ReadString('ColumnDefinitions','PositionColStr','B');
    CalledColStr := AppIni.ReadString('ColumnDefinitions','CalledColStr','C');
    ColumnColStr := AppIni.ReadString('ColumnDefinitions','ColumnColStr','D');
    TakeLogColStr := AppIni.ReadString('ColumnDefinitions','TakeLogColStr','E');
    WSumFacColStr := AppIni.ReadString('ColumnDefinitions','WSumFacColStr','F');
    tmpStr := AppIni.ReadString('Defaults','DefaultMinimum','1.0e-6');
    MakeQQPlot := AppIni.ReadBool('Defaults','CalculateQuantiles',false);
    Include4DVarData := AppIni.ReadBool('Defaults','Include4DVarData',false);
    Val(tmpStr,DefaultMinimum,iCode);
    if (iCode > 0) then DefaultMinimum := 1.0e-6;
  finally
    AppIni.Free;
  end;
  //used to force back to 32 bit, assuming this is how things are set up if using the JET engine
  //dmCor.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\ProgramData\EggSoft\Coran\Database\coran.mdb;Mode=ReadWrite;Persist Security Info=False;';
  dmCor.Coran.ConnectionString := dmCor.ConnectionString;
  dmCor.Coran.Connected := true;
end;

procedure TfmCoranMain.SetIniFile;
var
  AppIni   : TIniFile;
  //zpath: array [0..MAX_PATH] of char;
  //PublicPath : string;
begin
  //uses ShlObj
  // this gives access to all the systemed defined folders, no direct dependency on env. variables.
  //SHGetFolderPath(0, CSIDL_COMMON_APPDATA or CSIDL_FLAG_CREATE, 0, SHGFP_TYPE_CURRENT, @zpath);
  //dmCor.CommonFilePath := IncludeTrailingPathDelimiter(string(zpath)) + 'EggSoft\Coran\';
  //dmCor.IniFilename := IncludeTrailingPathDelimiter(string(zpath)) + 'EggSoft\Coran\'+'Coran.INI';
  //PublicPath := TPath.GetHomePath;
  //dmCor.CommonFilePath := IncludeTrailingPathDelimiter(PublicPath) + 'EggSoft\';
  //dmCor.IniFilename := dmCor.CommonFilePath + 'Coran.ini';
  AppIni := TIniFile.Create(dmCor.IniFilename);
  try
    AppIni.WriteString('ADO','Connection string',dmCor.ConnectionString);
    AppIni.WriteString('Paths','Data path',DataPath);
    AppIni.WriteString('Paths','Spreadsheet exports path',ExportPath);
    AppIni.WriteString('Paths','Spreadsheet template path',FlexTemplatePath);
    AppIni.WriteString('Paths','JPEG exports path',JPEGPath);
    AppIni.WriteString('ColumnDefinitions','ImportSpecNameColStr',ImportSpecNameColStr);
    AppIni.WriteString('ColumnDefinitions','PositionColStr',PositionColStr);
    AppIni.WriteString('ColumnDefinitions','CalledColStr',CalledColStr);
    AppIni.WriteString('ColumnDefinitions','ColumnColStr',ColumnColStr);
    AppIni.WriteString('ColumnDefinitions','TakeLogColStr',TakeLogColStr);
    AppIni.WriteString('ColumnDefinitions','WSumFacColStr',WSumFacColStr);
    AppIni.WriteString('Defaults','DefaultMinimum',FormatFloat('##0.0000e-00',DefaultMinimum));
    AppIni.WriteBool('Defaults','CalculateQuantiles',MakeQQPlot);
    AppIni.WriteBool('Defaults','Include4DVarData',Include4DVarData);
  finally
    AppIni.Free;
  end;
end;

procedure TfmCoranMain.FormShow(Sender: TObject);
var
  ii, j : integer;
begin
  MakeQQPlot := false;
  Include4DVarData := false;
  PrintGraph1.Enabled := false;
  PrintResults1.Checked := false;
  PrintData1.Checked := false;
  Project1.Enabled := false;
  tsResults.TabVisible := false;
  tsSimilarity.TabVisible := false;
  tsEigenValues.TabVisible := false;
  tsLoadings.TabVisible := false;
  tsScores.TabVisible := false;
  tsOutlierMap.TabVisible := false;
  tsGraph1.TabVisible := false;
  tsGraph2.TabVisible := false;
  tsGraph3.TabVisible := false;
  ts3D.TabVisible := false;
  tsCheck.TabVisible := false;
  pc1.ActivePage := tsControl;
  DefaultMinimum := 1.0e-6;
  NumEigenVectorsSelected := 10;
  DefaultRotation := 315;
  DefaultElevation := 350;
  DefaultPerspective := 0;
  DefaultZoom := 80;
  GetIniFile;
  tsLoc4D.TabVisible := Include4DVarData;
  Include4DVariableData1.Checked := Include4DVarData;
  CalculateQuantiles1.Checked := MakeQQPlot;
  //dmCor.ConnectionString1 := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=';
  //dmCor.ConnectionString2 := ';Mode=ReadWrite;Persist Security Info=False';
  {
  with dmCor do
  begin
    try
      Coran.Connected := false;
    except
    end;
    Coran.ConnectionString := dmCor.ConnectionString;
    //Coran.ConnectionString := 'FILE NAME='+ADODataLinkFile;
    //Coran.Provider := ADODataLinkFile;
    Coran.Connected := true;
    //Coran.Open('admin','');
  end;

  AssignChartDataSources('Close');
  with dmCor do
  begin
    cdsCoranChem.Open;
    cdsCoranRaw.Open;
    ElemNames.Open;
    QGroups.Open;
    QPlotGroups.Open;
    GroupedSmp.Open;
    QGroupedSmp.Open;
    SmpLoc.Open;
    GroupedSmpLoc.Open;
    QGroupedSmpLoc.Open;
    CoranStats.Open;
    cdsFacLoadingsSmp.Open;
    ElemNames.First;
    ii := ElemNames.RecordCount;
    if (ii > 0) then
    begin
      ii := 0;
      try
        repeat
          ii :=ii + 1;
          Oxidename[ElemNamesPos.AsInteger] := ElemNamesCalled.AsString;
          TakeLogs[ElemNamesPos.AsInteger] := ElemNamesTakeLog.AsString;
          DBGrid1.Columns[ii+2].Title.Caption := dmCor.ElemNamesCalled.AsString;
          DBGridStats.Columns[ii+1].Title.Caption := dmCor.ElemNamesCalled.AsString;
          dbgSimilarity.Columns[ii+1].Title.Caption := dmCor.ElemNamesCalled.AsString;
          ElemNames.Next;
        until ((ElemNames.Eof) or (ii >= MMax));
      except
      end;
      Nox := ii;
    end;
  end;
  if (NumEigenVectorsSelected > Nox) then NumEigenVectorsSelected := Nox;
  dmCor.ElemNames.First;
  for ii := Nox+1 to MMaxFields do
  begin
    DBGrid1.Columns[ii+2].Visible := false;
    DBGrid2.Columns[ii+2].Visible := false;
  end;
  for ii := Nox+1 to MMaxFields do
  begin
    DBGridStats.Columns[ii+1].Visible := false;
  end;
  AssignChartDataSources('Open');
  AddToComboBoxVar(Nox);
  //cbRawGraphVarChange(self);
  scSimilarityChoice := scNone;
  //cbVariableID.Clear;
  if Include4DVarData then
  begin
    tsLoc4D.Visible := true;
  end else
  begin
    tsLoc4D.Visible := false;
  end;
  //cbVariableID.Text := cbVariableID.Items[0];
  PrepareComboBoxes(0);
  try
    PrepareGraphs;
  except
  end;
  FromRowValueString := '2';
  ToRowValueString := '2';
  }
end;

procedure TfmCoranMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  SetIniFile;
end;

procedure TfmCoranMain.FormCreate(Sender: TObject);
begin
  //
end;

procedure TfmCoranMain.DataLinkFile1Click(Sender: TObject);
var
  tmpstring : string;
begin
  dmCor.ConnectionString := InputBox('ADO','Data connection string ',tmpString);
  dmCor.ConnectionString := UpperCase(dmCor.ConnectionString);
end;

procedure TfmCoranMain.PrepareComboBoxes(NumAttributes : integer);
var
  i, j : integer;
  A : SingleArrayM;
begin
  if (NumAttributes > 0) then
  begin
    dmCor.EigenVal.Open;
    dmCor.EigenVal.First;
    i := 0;
    repeat
      i := i + 1;
      A[i] := dmCor.EigenValEigenValuePct.AsFloat;
      dmCor.EigenVal.Next;
    until dmCor.EigenVal.Eof;
    dmCor.EigenVal.First;
  end;
   cbX.Items.Clear;
   cbY.Items.Clear;
   cbZ.Items.Clear;
   cbXLoc.Items.Clear;
   cbYLoc.Items.Clear;
   cbZLoc.Items.Clear;
   cbX4D.Items.Clear;
   cbY4D.Items.Clear;
   cbZ4D.Items.Clear;
   cbVariableID.Clear;
   if (scSimilarityChoice in [scCorrespondence,scRQVariance,scRQPearsonR,
                   scPCAVariance,scPCAPearsonR,scPCASpearmanR,
                   scPCAKendallR,scRQSpearmanR,scRQKendallR,
                   scRobPCA]) then cbVariableID.Items.Add('ScoreDistance');
   if (scSimilarityChoice in [scCorrespondence,scRQVariance,scRQPearsonR,
                   scPCAVariance,scPCAPearsonR,scPCASpearmanR,
                   scPCAKendallR,scRQSpearmanR,scRQKendallR,
                   scRobPCA]) then
   begin
     for i := 1 to NumAttributes do
     begin
       if ((i > 1) or (scSimilarityChoice > scCorrespondence)) then
       begin
         AddToComboBoxXYZ(i,'X');
         AddToComboBoxXYZ(i,'Y');
         AddToComboBoxXYZ(i,'Z');
         AddToComboBoxVariableID(A,i);
       end;
     end;
   end;
   if (scSimilarityChoice in [scDiscrim2Grp]) then
   begin
     AddToComboBoxXYZ(1,'X');
     AddToComboBoxXYZ(1,'Y');
     AddToComboBoxXYZ(1,'Z');
     AddToComboBoxXYZ(2,'Y');
     AddToComboBoxXYZ(2,'Z');
     AddToComboBoxVariableID(A,1);
   end;
   if Include4DVarData then
   begin
     dmCor.ElemNames.Open;
     dmCor.ElemNames.First;
     for j := 1 to Nox do
     begin
       cbVariableID.Items.Add(dmCor.ElemNamesCalled.AsString);
       dmCor.ElemNames.Next;
     end;
   end;
   cbVariableID.Text := cbVariableID.Items[0];
   cbXLoc.Items.Add('Longitude');
   cbXLoc.Items.Add('Latitude');
   cbXLoc.Items.Add('Elevation');
   cbZLoc.Items.Add('Latitude');
   cbZLoc.Items.Add('Longitude');
   cbZLoc.Items.Add('Elevation');
   cbYLoc.Items.Add('Elevation');
   cbYLoc.Items.Add('Longitude');
   cbYLoc.Items.Add('Latitude');
   cbX4D.Items.Add('Longitude');
   cbX4D.Items.Add('Latitude');
   cbX4D.Items.Add('Elevation');
   cbZ4D.Items.Add('Latitude');
   cbZ4D.Items.Add('Longitude');
   cbZ4D.Items.Add('Elevation');
   cbY4D.Items.Add('Elevation');
   cbY4D.Items.Add('Longitude');
   cbY4D.Items.Add('Latitude');
   cbX.Text := cbX.Items[0];
   cbY.Text := cbY.Items[1];
   cbZ.Text := cbZ.Items[2];
   if (scSimilarityChoice in [scDiscrim2Grp]) then cbZ.Text := cbZ.Items[1];
   cbXLoc.Text := cbXLoc.Items[0];
   cbYLoc.Text := cbYLoc.Items[0];
   cbZLoc.Text := cbZLoc.Items[0];
   cbX4D.Text := cbX4D.Items[0];
   cbY4D.Text := cbY4D.Items[0];
   cbZ4D.Text := cbZ4D.Items[0];
   cbVariableID.Text := cbVariableID.Items[0];
   AddToComboBoxScoreVar(Nox);
end;

procedure TfmCoranMain.PrepareGraphs;
var
  tmpStr : string;
  i, j, k : integer;
  iStart : integer;
begin
  SetSymbolDefaults;
  Series1601.WhiskerLength := 1.5;
  iStart := 1;
  if (scSimilarityChoice = scCorrespondence) then iStart := 2;
  j := 1;
  if (scSimilarityChoice = scCorrespondence) then j := 2;
  AssignChartDataSources('Close');
  sbMain.Panels[1].Text := 'Assign parameter values';
  sbMain.Refresh;
  dmCor.qDim4Smp1.Parameters.ParamByName('VariableID').Value := cbVariableID.Text;
  dmCor.qDim4Smp2.Parameters.ParamByName('VariableID').Value := cbVariableID.Text;
  dmCor.qDim4Smp3.Parameters.ParamByName('VariableID').Value := cbVariableID.Text;
  dmCor.qDim4Smp4.Parameters.ParamByName('VariableID').Value := cbVariableID.Text;
  dmCor.qDim4Smp5.Parameters.ParamByName('VariableID').Value := cbVariableID.Text;
  dmCor.qDim4Smp6.Parameters.ParamByName('VariableID').Value := cbVariableID.Text;
  sbMain.Panels[1].Text := 'Parameter values assigned';
  sbMain.Refresh;
  cbComponent.Items.Clear;
  for i := j to Nox do
  begin
    cbComponent.Items.Add('Component' + IntToStr(i));
  end;
  tmpStr := '4';
  if (scSimilarityChoice = scCorrespondence) then
  begin
    tmpStr := '5';
    if (Nox < 5) then tmpStr := IntToStr(Nox);
  end;
  if (Nox < 4) then tmpStr := IntToStr(Nox);
  cbComponent.Text := 'Component'+tmpStr;
  sbMain.Panels[1].Text := 'Graph the loadings';
  sbMain.Refresh;
  {loadings}
  try
    DBChart5.Series[0].DataSource := dmCor.FacLoadingsVarNoZero;
    DBChart6.Series[0].DataSource := dmCor.FacLoadingsVarNoZero;
    DBChart7.Series[0].DataSource := dmCor.FacLoadingsVarNoZero;
    DBChart10.Series[0].DataSource := dmCor.FacLoadingsVarNoZero;
    DBChart5.BottomAxis.Title.Caption := 'Component ' + IntToStr(iStart);
    DBChart6.BottomAxis.Title.Caption := 'Component ' + IntToStr(iStart+1);
    DBChart7.BottomAxis.Title.Caption := 'Component ' + IntToStr(iStart+2);
    DBChart10.BottomAxis.Title.Caption := 'Component '+tmpStr;
    DBChart5.Series[0].XValues.ValueSource := 'Vector' + IntToStr(iStart);
    DBChart5.Series[0].YValues.ValueSource := 'Pos';
    DBChart6.Series[0].XValues.ValueSource := 'Vector' + IntToStr(iStart+1);
    DBChart6.Series[0].YValues.ValueSource := 'Pos';
    DBChart7.Series[0].XValues.ValueSource := 'Vector' + IntToStr(iStart+2);
    DBChart7.Series[0].YValues.ValueSource := 'Pos';
    DBChart10.Series[0].XValues.ValueSource := 'Vector'+tmpStr;
    DBChart10.Series[0].YValues.ValueSource := 'Pos';
    case scSimilarityChoice of
      scCorrespondence : begin
      end;
      scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
      end;
      scPCAVariance,scPCAPearsonR,scPCASpearmanR,scPCAKendallR : begin
        DBChart5.Series[0].Active := false;
        DBChart6.Series[0].Active := false;
        DBChart7.Series[0].Active := false;
        DBChart10.Series[0].Active := false;
      end;
      scDiscrim2Grp : begin
        DBChart5.Series[0].Active := false;
        DBChart6.Series[0].Active := false;
        DBChart7.Series[0].Active := false;
        DBChart10.Series[0].Active := false;
      end;
    end;
    DBChart5.Series[0].XValues.Order := loNone;
    DBChart5.Series[0].YValues.Order := loNone;
    DBChart6.Series[0].XValues.Order := loNone;
    DBChart6.Series[0].YValues.Order := loNone;
    DBChart7.Series[0].XValues.Order := loNone;
    DBChart7.Series[0].YValues.Order := loNone;
    DBChart10.Series[0].XValues.Order := loNone;
    DBChart10.Series[0].YValues.Order := loNone;
    DBChart5.Series[0].Active := true;
    DBChart6.Series[0].Active := true;
    DBChart7.Series[0].Active := true;
    DBChart10.Series[0].Active := true;
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
    DBChart1.BottomAxis.Title.Caption := 'Component ' + IntToStr(iStart);
    DBChart1.LeftAxis.Title.Caption := 'Component ' + IntToStr(iStart+1);
    DBChart2.BottomAxis.Title.Caption := 'Component ' + IntToStr(iStart);
    DBChart2.LeftAxis.Title.Caption := 'Component ' + IntToStr(iStart+2);
    DBChart3.BottomAxis.Title.Caption := 'Component ' + IntToStr(iStart+1);
    DBChart3.LeftAxis.Title.Caption := 'Component ' + IntToStr(iStart+2);
    DBChart1.Series[0].XValues.ValueSource := 'Vector' + IntToStr(iStart);
    DBChart1.Series[0].YValues.ValueSource := 'Vector' + IntToStr(iStart+1);
    DBChart2.Series[0].XValues.ValueSource := 'Vector' + IntToStr(iStart);
    DBChart2.Series[0].YValues.ValueSource := 'Vector' + IntToStr(iStart+2);
    DBChart3.Series[0].XValues.ValueSource := 'Vector' + IntToStr(iStart+1);
    DBChart3.Series[0].YValues.ValueSource := 'Vector' + IntToStr(iStart+2);
    case scSimilarityChoice of
      scCorrespondence : begin
      end;
      scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
        DBChart1.Series[0].Active := true;
        DBChart2.Series[0].Active := true;
        DBChart3.Series[0].Active := true;
      end;
      scPCAVariance,scPCAPearsonR,scPCASpearmanR,scPCAKendallR : begin
        DBChart1.Series[0].Active := false;
        DBChart2.Series[0].Active := false;
        DBChart3.Series[0].Active := false;
      end;
      scDiscrim2Grp : begin
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
  sbMain.Panels[1].Text := 'Graph the samples';
  sbMain.Refresh;
  {all samples}
  try
    DBChart1.Series[1].DataSource := dmCor.cdsFacLoadingsSmp;
    DBChart2.Series[1].DataSource := dmCor.cdsFacLoadingsSmp;
    DBChart3.Series[1].DataSource := dmCor.cdsFacLoadingsSmp;
    DBChart1.Series[1].Active := true;
    DBChart2.Series[1].Active := true;
    DBChart3.Series[1].Active := true;
    DBChart1.Series[1].XValues.ValueSource := 'Vector' + IntToStr(iStart);
    DBChart1.Series[1].YValues.ValueSource := 'Vector' + IntToStr(iStart+1);
    DBChart2.Series[1].XValues.ValueSource := 'Vector' + IntToStr(iStart);
    DBChart2.Series[1].YValues.ValueSource := 'Vector' + IntToStr(iStart+2);
    DBChart3.Series[1].XValues.ValueSource := 'Vector' + IntToStr(iStart+1);
    DBChart3.Series[1].YValues.ValueSource := 'Vector' + IntToStr(iStart+2);
    case scSimilarityChoice of
      scCorrespondence : begin
      end;
      scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
      end;
      scPCAVariance,scPCAPearsonR,scPCASpearmanR, scPCAKendallR : begin
      end;
      scDiscrim2grp : begin
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
    DBChart1.Series[2].XValues.ValueSource := 'Vector' + IntToStr(iStart);
    DBChart1.Series[2].YValues.ValueSource := 'Vector' + IntToStr(iStart+1);
    DBChart2.Series[2].XValues.ValueSource := 'Vector' + IntToStr(iStart);
    DBChart2.Series[2].YValues.ValueSource := 'Vector' + IntToStr(iStart+2);
    DBChart3.Series[2].XValues.ValueSource := 'Vector' + IntToStr(iStart+1);
    DBChart3.Series[2].YValues.ValueSource := 'Vector' + IntToStr(iStart+2);
    case scSimilarityChoice of
      scCorrespondence : begin
      end;
      scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
      end;
      scPCAVariance,scPCAPearsonR,scPCASpearmanR, scPCAKendallR : begin
      end;
      scDiscrim2Grp : begin
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
    DBChart1.Series[3].XValues.ValueSource := 'Vector' + IntToStr(iStart);
    DBChart1.Series[3].YValues.ValueSource := 'Vector' + IntToStr(iStart+1);
    DBChart2.Series[3].XValues.ValueSource := 'Vector' + IntToStr(iStart);
    DBChart2.Series[3].YValues.ValueSource := 'Vector' + IntToStr(iStart+2);
    DBChart3.Series[3].XValues.ValueSource := 'Vector' + IntToStr(iStart+1);
    DBChart3.Series[3].YValues.ValueSource := 'Vector' + IntToStr(iStart+2);
    case scSimilarityChoice of
      scCorrespondence : begin
      end;
      scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
      end;
      scPCAVariance,scPCAPearsonR,scPCASpearmanR, scPCAKendallR : begin
      end;
      scDiscrim2Grp : begin
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
    DBChart1.Series[4].XValues.ValueSource := 'Vector' + IntToStr(iStart);
    DBChart1.Series[4].YValues.ValueSource := 'Vector' + IntToStr(iStart+1);
    DBChart2.Series[4].XValues.ValueSource := 'Vector' + IntToStr(iStart);
    DBChart2.Series[4].YValues.ValueSource := 'Vector' + IntToStr(iStart+2);
    DBChart3.Series[4].XValues.ValueSource := 'Vector' + IntToStr(iStart+1);
    DBChart3.Series[4].YValues.ValueSource := 'Vector' + IntToStr(iStart+2);
    case scSimilarityChoice of
      scCorrespondence : begin
      end;
      scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
      end;
      scPCAVariance,scPCAPearsonR,scPCASpearmanR, scPCAKendallR : begin
        DBChart1.Series[4].Active := false;
        DBChart2.Series[4].Active := false;
        DBChart3.Series[4].Active := false;
      end;
      scDiscrim2Grp : begin
        DBChart1.Series[4].Active := false;
        DBChart2.Series[4].Active := false;
        DBChart3.Series[4].Active := false;
        DBChart8.Series[0].Active := false;
        DBChart8.Series[4].Active := false;
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
  sbMain.Panels[1].Text := 'Graph the 2D localities';
  sbMain.Refresh;
  {all samples - localities}
  try
    DBChart4.Series[0].DataSource := dmCor.SmpLoc;
    DBChart4.Series[0].XValues.Order := loNone;
    DBChart4.Series[0].YValues.Order := loNone;
  except
  end;
  {all samples in group - localities}
  try
    DBChart4.Series[1].DataSource := dmCor.GroupedSmpLoc;
    DBChart4.Series[1].XValues.Order := loNone;
    DBChart4.Series[1].YValues.Order := loNone;
  except
  end;
  {current sample - locality}
  try
    DBChart4.Series[2].DataSource := dmCor.QGroupedSmpLoc;
    DBChart4.Series[2].XValues.Order := loNone;
    DBChart4.Series[2].YValues.Order := loNone;
  except
  end;
  DBChart8.BottomAxis.Title.Caption := cbX.Text;
  DBChart8.LeftAxis.Title.Caption := cbY.Text;
  DBChart8.DepthAxis.Title.Caption := cbZ.Text;
  sbMain.Panels[1].Text := 'Graph the 3D localities';
  sbMain.Refresh;
  {3D - variable loadings}
  try
    DBChart8.Series[0].DataSource := dmCor.FacLoadingsVar;
    DBChart8.Series[0].XValues.Order := loNone;
    DBChart8.Series[0].YValues.Order := loNone;
    (DBChart8.Series[0] as TPoint3DSeries).ZValues.Order := loNone;
    (DBChart8.Series[0] as TPoint3DSeries).XValues.ValueSource := 'Vector' + IntToStr(iStart);
    (DBChart8.Series[0] as TPoint3DSeries).YValues.ValueSource := 'Vector' + IntToStr(iStart+1);
    (DBChart8.Series[0] as TPoint3DSeries).ZValues.ValueSource := 'Vector' + IntToStr(iStart+2);
  except
  end;
  {3D - sample loadings}
  try
    DBChart8.Series[1].DataSource := dmCor.cdsFacLoadingsSmp;
    DBChart8.Series[1].XValues.Order := loNone;
    DBChart8.Series[1].YValues.Order := loNone;
    (DBChart8.Series[1] as TPoint3DSeries).ZValues.Order := loNone;
    (DBChart8.Series[1] as TPoint3DSeries).XValues.ValueSource := 'Vector' + IntToStr(iStart);
    (DBChart8.Series[1] as TPoint3DSeries).YValues.ValueSource := 'Vector' + IntToStr(iStart+1);
    (DBChart8.Series[1] as TPoint3DSeries).ZValues.ValueSource := 'Vector' + IntToStr(iStart+2);
  except
  end;
  {3D - grouped sample loadings}
  try
    DBChart8.Series[2].DataSource := dmCor.GroupedSmp;
    DBChart8.Series[2].XValues.Order := loNone;
    DBChart8.Series[2].YValues.Order := loNone;
    (DBChart8.Series[2] as TPoint3DSeries).ZValues.Order := loNone;
    (DBChart8.Series[2] as TPoint3DSeries).XValues.ValueSource := 'Vector' + IntToStr(iStart);
    (DBChart8.Series[2] as TPoint3DSeries).YValues.ValueSource := 'Vector' + IntToStr(iStart+1);
    (DBChart8.Series[2] as TPoint3DSeries).ZValues.ValueSource := 'Vector' + IntToStr(iStart+2);
  except
  end;
  {3D - current sample}
  try
    DBChart8.Series[3].DataSource := dmCor.QGroupedSmp;
    DBChart8.Series[3].XValues.Order := loNone;
    DBChart8.Series[3].YValues.Order := loNone;
    (DBChart8.Series[3] as TPoint3DSeries).ZValues.Order := loNone;
    (DBChart8.Series[3] as TPoint3DSeries).XValues.ValueSource := 'Vector' + IntToStr(iStart);
    (DBChart8.Series[3] as TPoint3DSeries).YValues.ValueSource := 'Vector' + IntToStr(iStart+1);
    (DBChart8.Series[3] as TPoint3DSeries).ZValues.ValueSource := 'Vector' + IntToStr(iStart+2);
  except
  end;
  {3D - current variable}
  try
    DBChart8.Series[4].DataSource := dmCor.FacLoadingsVarLinked;
    DBChart8.Series[4].XValues.Order := loNone;
    DBChart8.Series[4].YValues.Order := loNone;
    (DBChart8.Series[4] as TPoint3DSeries).ZValues.Order := loNone;
    (DBChart8.Series[4] as TPoint3DSeries).XValues.ValueSource := 'Vector' + IntToStr(iStart);
    (DBChart8.Series[4] as TPoint3DSeries).YValues.ValueSource := 'Vector' + IntToStr(iStart+1);
    (DBChart8.Series[4] as TPoint3DSeries).ZValues.ValueSource := 'Vector' + IntToStr(iStart+2);
  except
  end;
  {3D - all samples - localities}
  try
    DBChart9.Series[0].DataSource := dmCor.SmpLoc;
    DBChart9.Series[0].XValues.Order := loNone;
    DBChart9.Series[0].YValues.Order := loNone;
    (DBChart9.Series[0] as TPoint3DSeries).ZValues.Order := loNone;
    (DBChart9.Series[0] as TPoint3DSeries).XValues.ValueSource := cbXLoc.Text;
    (DBChart9.Series[0] as TPoint3DSeries).YValues.ValueSource := cbYLoc.Text;
    (DBChart9.Series[0] as TPoint3DSeries).ZValues.ValueSource := cbZLoc.Text;
  except
  end;
  {3D -all samples in group - localities}
  try
    DBChart9.Series[1].DataSource := dmCor.GroupedSmpLoc;
    DBChart9.Series[1].XValues.Order := loNone;
    DBChart9.Series[1].YValues.Order := loNone;
    (DBChart9.Series[1] as TPoint3DSeries).ZValues.Order := loNone;
    (DBChart9.Series[1] as TPoint3DSeries).XValues.ValueSource := cbXLoc.Text;
    (DBChart9.Series[1] as TPoint3DSeries).YValues.ValueSource := cbYLoc.Text;
    (DBChart9.Series[1] as TPoint3DSeries).ZValues.ValueSource := cbZLoc.Text;
  except
  end;
  {3D - current sample - locality}
  try
    DBChart9.Series[2].DataSource := dmCor.QGroupedSmpLoc;
    DBChart9.Series[2].XValues.Order := loNone;
    DBChart9.Series[2].YValues.Order := loNone;
    (DBChart9.Series[2] as TPoint3DSeries).ZValues.Order := loNone;
    (DBChart9.Series[2] as TPoint3DSeries).XValues.ValueSource := cbXLoc.Text;
    (DBChart9.Series[2] as TPoint3DSeries).YValues.ValueSource := cbYLoc.Text;
    (DBChart9.Series[2] as TPoint3DSeries).ZValues.ValueSource := cbZLoc.Text;
  except
  end;
  sbMain.Panels[1].Text := 'Graph the 4D localities';
  sbMain.Refresh;
  {4D - current sample - locality}
  try
    DBChart11.Series[6].DataSource := dmCor.QGroupedSmpLoc;
    (DBChart11.Series[6] as TPoint3DSeries).XValues.Order := loNone;
    (DBChart11.Series[6] as TPoint3DSeries).YValues.Order := loNone;
    (DBChart11.Series[6] as TPoint3DSeries).ZValues.Order := loNone;
    (DBChart11.Series[6] as TPoint3DSeries).XValues.ValueSource := cbX4D.Text;
    (DBChart11.Series[6] as TPoint3DSeries).YValues.ValueSource := cbY4D.Text;
    (DBChart11.Series[6] as TPoint3DSeries).ZValues.ValueSource := cbZ4D.Text;
  except
  end;
  {4D - hinge values}
  for k := 0 to 5 do
  begin
    try
      (DBChart11.Series[k] as TPoint3DSeries).XValues.Order := loNone;
      (DBChart11.Series[k] as TPoint3DSeries).YValues.Order := loNone;
      (DBChart11.Series[k] as TPoint3DSeries).ZValues.Order := loNone;
      (DBChart11.Series[k] as TPoint3DSeries).XValues.ValueSource := cbX4D.Text;
      (DBChart11.Series[k] as TPoint3DSeries).YValues.ValueSource := cbY4D.Text;
      (DBChart11.Series[k] as TPoint3DSeries).ZValues.ValueSource := cbZ4D.Text;
    except
    end;
  end;
  sbShow4DClick(Self);
  case scSimilarityChoice of
    scCorrespondence,scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
      DBChart8.Series[0].Visible := true;
    end;
    scPCAVariance,scPCAPearsonR,scPCASpearmanR, scPCAKendallR : begin
      DBGrid29.Enabled := false;
      DBNavigator21.Enabled := false;
      DBChart8.Series[0].Visible := false;
    end;
    scDiscrim2Grp,scDiscrimnGrp,scCluster : begin
      DBGrid29.Enabled := false;
      DBNavigator21.Enabled := false;
      DBChart8.Series[0].Visible := false;
      tsGraph2.TabVisible := false;
      tsGraph3.TabVisible := false;
      ts3D.TabVisible := false;
      tsLoc4D.TabVisible := Include4DVarData;
    end;
  end;
  {DBChart12}
  DBChart12.Series[0].Active := false;
  DBChart12.Series[1].Active := false;
  DBChart12.Series[2].Active := false;
  SeriesG1205.Active := false;
  SeriesG1206.Active := false;
  DBChart12.BottomAxis.Title.Caption := 'Score Distance from '+IntToStr(NumEigenVectorsSelected)+' of '+IntToStr(Nox)+' eigenvectors';
  DBChart12.LeftAxis.Title.Caption := 'Orthogonal Distance';

  DBChart12.Series[0].DataSource := dmCor.qSDOD;
  DBChart12.Series[0].XValues.Order := loNone;
  DBChart12.Series[0].YValues.Order := loNone;
  DBChart12.Series[0].XValues.ValueSource := 'ScoreDistance';
  DBChart12.Series[0].YValues.ValueSource := 'OrthogonalDistance';

  DBChart12.Series[1].DataSource := dmCor.qGroupedSDOD;
  DBChart12.Series[1].XValues.Order := loNone;
  DBChart12.Series[1].YValues.Order := loNone;
  DBChart12.Series[1].XValues.ValueSource := 'ScoreDistance';
  DBChart12.Series[1].YValues.ValueSource := 'OrthogonalDistance';

  DBChart12.Series[2].DataSource := dmCor.qGroupedSmpSDOD;
  DBChart12.Series[2].XValues.Order := loNone;
  DBChart12.Series[2].YValues.Order := loNone;
  DBChart12.Series[2].XValues.ValueSource := 'ScoreDistance';
  DBChart12.Series[2].YValues.ValueSource := 'OrthogonalDistance';

  DBChart12.Series[0].Active := true;
  DBChart12.Series[1].Active := true;
  DBChart12.Series[2].Active := true;
  SeriesG1205.Active := true;
  SeriesG1206.Active := true;
  if (NumEigenVectorsSelected = Nox) then
  begin
    DBChart12.LeftAxis.Title.Caption := 'Sample Sequence';
    SeriesG1206.Active := false;
  end;
  DBChart15.Series[0].Active := false;
  DBChart15.Series[0].XValues.Order := loNone;
  DBChart15.Series[0].YValues.Order := loNone;
  tmpStr := 'Param'+IntToStr(cbRawGraphVar.ItemIndex+1);
  DBChart15.Series[0].XValues.ValueSource := tmpStr;
  DBChart15.Series[0].Active := true;
  case scSimilarityChoice of
    scCorrespondence : begin
    end;
    scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
    end;
    scPCAVariance,scPCAPearsonR,scPCASpearmanR, scPCAKendallR : begin
    end;
    scDiscrim2Grp,scDiscrimnGrp,scCluster : begin
      DBChart12.Series[0].Active := false;
      DBChart12.Series[1].Active := false;
      DBChart12.Series[2].Active := false;
    end;
  end;
  if MakeQQPlot then DBChartQuantile.Visible := true
                else DBChartQuantile.Visible := false;
  DBChart8.View3DOptions.Rotation := DefaultRotation;
  DBChart8.View3DOptions.Elevation := DefaultElevation;
  DBChart8.View3DOptions.Perspective := DefaultPerspective;
  DBChart8.View3DOptions.Zoom := DefaultZoom;
  DBChart9.View3DOptions.Rotation := DefaultRotation;
  DBChart9.View3DOptions.Elevation := DefaultElevation;
  DBChart9.View3DOptions.Perspective := DefaultPerspective;
  DBChart9.View3DOptions.Zoom := DefaultZoom;
  DBChart11.View3DOptions.Rotation := DefaultRotation;
  DBChart11.View3DOptions.Elevation := DefaultElevation;
  DBChart11.View3DOptions.Perspective := DefaultPerspective;
  DBChart11.View3DOptions.Zoom := DefaultZoom;
  sbMain.Panels[1].Text := 'Open queries for graphs';
  sbMain.Refresh;
  AssignChartDataSources('Close');
  AssignChartDataSources('Open');
  cbScoreGraphVarChange(self);
end;

procedure TfmCoranMain.Button1Click(Sender: TObject);
begin
  PrepareGraphs;
end;

procedure TfmCoranMain.dbnGroupedSmpClick(Sender: TObject;
  Button: TNavigateBtn);
begin
  try
    dmCor.QGroupedSmp.Close;
    dmCor.QGroupedSmpLoc.Close;
    dmCor.QGroupedSmpSDOD.Close;
  except
  end;
  try
    dmCor.QGroupedSmp.Open;
    dmCor.QGroupedSmpLoc.Open;
    dmCor.QGroupedSmpSDOD.Open;
  except
  end;
end;

procedure TfmCoranMain.dbnGroupsClick(Sender: TObject;
  Button: TNavigateBtn);
begin
  try
    dmCor.GroupedSmp.Close;
    dmCor.GroupedSmpLoc.Close;
    dmCor.QGroupedSDOD.Close;
  except
  end;
  try
    dmCor.GroupedSmp.Open;
    dmCor.GroupedSmpLoc.Open;
    dmCor.QGroupedSDOD.Open;
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
  if (pc1.ActivePage = tsEigenValues) then ChartEigenValue.Print;
  if (pc1.ActivePage = tsScores) then DBChart10.Print;
  if (pc1.ActivePage = tsLocalities) then DBChart4.Print;
  if (pc1.ActivePage = ts3D) then DBChart8.Print;
  if (pc1.ActivePage = tsLoc3D) then DBChart9.Print;
  if (pc1.ActivePage = tsLoc4D) then DBChart11.Print;
  if (pc1.ActivePage = tsOutlierMap) then DBChart12.Print;
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
  ii, I, J, K, L : integer;
  n : integer;
  tmpStr : string;
  tmpstr10 : string[10];
  blankstr : string;
  {A is MxM, X is NxMx2, A2 is MxM, C is 2xM, NS is 2 matrix}
begin
  blankstr := '  ';
  DBChart7.Visible := false;
  Panel40.Visible := false;
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
  if (PrintResults1.Checked=true) then IPrn := 'Y'
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
  dmCor.ElemNames.Open;
  Nox := dmCor.ElemNames.RecordCount;
  TotalRecs := dmCor.cdsCoranChem.RecordCount;
  //M := Nox;
  for ii := Nox+1 to MMaxFields do
  begin
    DBGrid1.Columns[ii+2].Visible := false;
  end;
  {
  for ii := 2 to MMaxFields do
  begin
    DBGrid2.Columns[ii+2].Visible := false;
  end;
  }
  for ii := Nox+1 to MMaxFields do
  begin
    dbgFacLoadingsVar.Columns[ii+1].Visible := false;
    DBGridStats.Columns[ii+1].Visible := false;
  end;
  for ii := 2 to MMaxFields do
  begin
    dbgFacLoadingsSmp.Columns[ii+2].Visible := false;
  end;
  for ii := Nox+1 to MMaxFields do
  begin
    dbgEigenVec.Columns[ii].Visible := false;
  end;
  for ii := Nox+1 to MMaxFields do
  begin
    dbgSimilarity.Columns[ii+1].Visible := false;
  end;
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
    ComponentStr[I]:=dmCor.GroupedChemSAMPLENUM.AsString;
    for J:=1 to Nox do
    begin
      X[i,J] := dmCor.GroupedChem.Fields[J+2].AsVariant;
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
  dmCor.GroupedChem.Close;
  dmCor.GroupedChem.Open;
  {read group 2 data}
  N:=1;
  sbmain.Panels[1].Text :='Reading data for '+IntToStr(Nox)+' elements for second group';
  sbMain.Refresh;
  dmCor.GroupedChem.First;
  repeat
    i:=i+1;
    ComponentStr[I]:=dmCor.GroupedChemSAMPLENUM.AsString;
    for J:=1 to Nox do
    begin
      X[I,J] := dmCor.GroupedChem.Fields[J+2].AsVariant;
    end;
    sbmain.Panels[1].Text :='Processing record '+IntToStr(N)+' of '+IntToStr(TotalRecs)+' Group 2';
    sbMain.Refresh;
    N:=N+1;
    dmCor.GroupedChem.Next;
  until (dmCor.GroupedChem.Eof);
  Ndig[2] := i - Ndig[1];
  sbMain.Refresh;
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
  for i := 1 to Nox do
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
    tmpStr := FormatFloat('#####0.00000',A3[1,k]/Ndig[1]);
    ResultsArray[k,3] := CharStream(10-Length(tmpstr),32)+tmpstr + blankstr;
    tmpStr := FormatFloat('#####0.00000',A3[2,k]/Ndig[2]);
    ResultsArray[k,4] := CharStream(10-Length(tmpstr),32)+tmpstr + blankstr;
    tmpStr := FormatFloat('#####0.00000',A2[k,k]);
    ResultsArray[k,5] := CharStream(10-Length(tmpstr),32)+tmpstr + blankstr;
  end;
  SLE(A1,A2,Nox,MM,DefaultZeroLimit);
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
  if (dmCor.CoranSimilarity.RecordCount > 0) then
  begin
    dmCor.CoranSimilarity.Last;
    repeat
      dmCor.CoranSimilarity.Delete;
    until dmCor.CoranSimilarity.Eof;
  end;
  for i := 1 to Nox do
  begin
    dmCor.CoranSimilarity.Append;
    dmCor.CoranSimilarityParamNum.AsInteger := i;
    dmCor.CoranSimilarityParamCalled.AsString := OxideName[i];
    for j := 1 to Nox do
    begin
      dmCor.CoranSimilarity.Fields[j+1].AsVariant := A1[i,j];
    end;
    dmCor.CoranSimilarity.Post;
    dbgSimilarity.Columns[i+1].Title.Caption := Oxidename[i];
  end;
  FillSimilaritySpreadSheet;
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
  AM := 1.0*Nox;
  F := (((AN1 + AN2 - AM - 1.0)*AN1*AN2)/(AN3*AM*(AN1 + AN2)))*D2;
  ND1 := Nox;
  ND2 := Ndig[1] + Ndig[2] - Nox - 1;
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
  for i := 1 to Nox do
  begin
    if (D2 <> 0.0) then
    begin
      E := (A2[i,i]*(A3[1,i]/AN1 - A3[2,i]/AN2)/D2)*100.0;
    end else
    begin
      E := -100.0;
    end;
    {
    MemoResults.Lines.Add(' '+FormatFloat('000',i)+' '+OxideName[i]+'          '
       +FormatFloat('#######0.0000',A2[i,i])+'          '+FormatFloat('#######0.0###',E));
    }
    if IPrn = 'Y' then Writeln(Lst,OxideName[i]:15,A2[i,i]:14:4,E:12:2);
  end;
  {add data for another two columns}
  for i := 1 to Nox do
  begin
    if (D2 <> 0.0) then
    begin
      E := (A2[i,i]*(A3[1,i]/AN1 - A3[2,i]/AN2)/D2)*100.0;
    end else
    begin
      E := -100.0;
    end;
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
  CreateEigSprdShts;
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
  pVar.Visible := false;
  pGraph1Var.Visible := false;
  pGraph2Var.Visible := false;
  pGraph2Var.Visible := false;
  pOutlierMapVar.Visible := false;
  p3DVar.Visible := false;
  pLocalitiesVar.Visible := false;
  PrepareGraphs;
end;

procedure TfmCoranMain.CalcDiscrmScores;
var
  ii, I, J, K, L : integer;
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
  j := 1;
  j := 2;
  ii := 0;
  for i := 1 to Nox do
  begin
    ii := ii + 1;
    dmCor.FacLoadingsVar.Append;
    dmCor.FacLoadingsVarPos.AsInteger := i;
    dmCor.FacLoadingsVarCalled.AsString := OxideName[i];
    if (D2 <> 0.0) then
    begin
      E := (A2[i,i]*(A3[1,i]/AN1 - A3[2,i]/AN2)/D2)*100.0;
    end else
    begin
      E := -100.0;
    end;
    for j := 1 to 2 do
    begin
      case j of
        1 : begin
          dmCor.FacLoadingsVarVector1.AsFloat := A2[i,i];
        end;
        2 : begin
          dmCor.FacLoadingsVarVector2.AsFloat := E;
        end;
      end;
    end;
    dmCor.FacLoadingsVar.Post;
  end;
  dmCor.FacLoadingsVar.First;
  {clear factor loadings table for samples}
  dmCor.cdsFacLoadingsSmp.Last;
  if not (dmCor.cdsFacLoadingsSmp.Bof and dmCor.cdsFacLoadingsSmp.Eof) then
  begin
    dmCor.cdsFacLoadingsSmp.Last;
    repeat
      dmCor.cdsFacLoadingsSmp.Delete;
    until dmCor.cdsFacLoadingsSmp.Bof;
  end;
  sbmain.Panels[1].Text :='Calculating scores for group 1';
  sbMain.Refresh;
  dmCor.QGroups.Close;
  dmCor.QGroups.Open;
  dmCor.QGroups.First;
  dmCor.GroupedChem.Close;
  dmCor.GroupedChem.Open;
  dmCor.GroupedChem.First;
  dmCor.QPlotGroups.Close;
  dmCor.QPlotGroups.Open;
  dmCor.QGroupedSmp.Close;
  dmCor.QGroupedSmp.Open;
  {fill factor loadings table for samples for first group}
  j := 1;
  j := 2;
  i := 1;
  k := 1;
  repeat
    sbmain.Panels[1].Text :='Calculating scores for group 1'+dmCor.GroupedChemSAMPLENUM.AsString;
    sbMain.Refresh;
    D := 0.0;
    for l := 1 to Nox do
    begin
      D := D + A2[l,l] * X[k,l];
    end;
    dmCor.QGroupedSmp.Append;
    dmCor.QGroupedSmpGroupName.AsString := dmCor.GroupedChemGroupName.AsString;
    dmCor.QGroupedSmpPlotGroupName.AsString := dmCor.GroupedChemPlotGroupName.AsString;
    dmCor.QGroupedSmpSampleNum.AsString := dmCor.GroupedChemSAMPLENUM.AsString;
    j := 1;
    dmCor.QGroupedSmpVector1.AsFloat := D;
    j := 2;
    dmCor.QGroupedSmpVector2.AsFloat := 1.0;
    dmCor.QGroupedSmp.Post;
    dmCor.GroupedChem.Next;
    i := i + 1;
    k := k + 1;
  until dmCor.GroupedChem.Eof;
  dmCor.GroupedChem.First;
  dmCor.cdsFacLoadingsSmp.First;
  {fill factor loadings table for samples for second group}
  sbmain.Panels[1].Text :='Calculating scores for group 2';
  sbMain.Refresh;
  dmCor.QGroups.Next;
  dmCor.GroupedChem.Close;
  dmCor.GroupedChem.Open;
  dmCor.GroupedChem.First;
  dmCor.QPlotGroups.Close;
  dmCor.QPlotGroups.Open;
  dmCor.QGroupedSmp.Close;
  dmCor.QGroupedSmp.Open;
  dmCor.cdsFacLoadingsSmp.Close;
  dmCor.cdsFacLoadingsSmp.Open;
  repeat
    sbmain.Panels[1].Text :='Calculating scores for group 2'+dmCor.GroupedChemSAMPLENUM.AsString;
    sbMain.Refresh;
    D := 0.0;
    for l := 1 to Nox do
    begin
      D := D + A2[l,l] * X[k,l];
    end;
    dmCor.QGroupedSmp.Append;
    dmCor.QGroupedSmpGroupName.AsString := dmCor.GroupedChemGroupName.AsString;
    dmCor.QGroupedSmpPlotGroupName.AsString := dmCor.GroupedChemPlotGroupName.AsString;
    dmCor.QGroupedSmpSampleNum.AsString := dmCor.GroupedChemSAMPLENUM.AsString;
    j := 1;
    dmCor.QGroupedSmpVector1.AsFloat := D;
    j := 2;
    dmCor.QGroupedSmpVector2.AsFloat := 2.0;
    dmCor.QGroupedSmp.Post;
    dmCor.GroupedChem.Next;
    i := i + 1;
    k := k + 1;
  until dmCor.GroupedChem.Eof;
  dmCor.GroupedChem.First;
  dmCor.cdsFacLoadingsSmp.First;
  dmCor.cdsFacLoadingsSmp.Close;
  dmCor.cdsFacLoadingsSmp.Open;
  dmCor.GroupedChem.Close;
end;

procedure TfmCoranMain.DiscrimProject;
var
  ii, I, J, K, L : integer;
  n : integer;
  tmpStr : string;
  tmpstr10 : string[10];
  blankstr : string;
  {A is MxM, X is NxMx2, A2 is MxM, C is 2xM, NS is 2 matrix}
begin
  blankstr := '  ';
  DBChart7.Visible := false;
  Panel40.Visible := false;
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
  if (PrintResults1.Checked=true) then IPrn := 'Y'
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
  dmCor.ElemNames.Open;
  Nox := dmCor.ElemNames.RecordCount;
  TotalRecs := dmCor.cdsCoranChem.RecordCount;
  //M := Nox;
  for ii := Nox+1 to MMaxFields do
  begin
    DBGrid1.Columns[ii+2].Visible := false;
  end;
  for ii := Nox+1 to MMaxFields do
  begin
    dbgFacLoadingsVar.Columns[ii+1].Visible := false;
    DBGridStats.Columns[ii+1].Visible := false;
  end;
  for ii := 2 to MMaxFields do
  begin
    dbgFacLoadingsSmp.Columns[ii+2].Visible := false;
  end;
  for ii := Nox+1 to MMaxFields do
  begin
    dbgEigenVec.Columns[ii].Visible := false;
  end;
  for ii := Nox+1 to MMaxFields do
  begin
    dbgSimilarity.Columns[ii+1].Visible := false;
  end;
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
    dmCor.cdsCoranChem.Open;
  except
  end;
  dmCor.cdsCoranChem.First;
  {read all data}
  i:=0;
  n:=1;
  sbmain.Panels[1].Text :='Reading data for '+IntToStr(Nox)+' elements';
  sbMain.Refresh;
  dmCor.cdsCoranChem.First;
  repeat
    i:=i+1;
    ComponentStr[I]:=dmCor.cdsCoranChemSAMPLENUM.AsString;
    for J:=1 to Nox do
    begin
      X[i,J] := dmCor.cdsCoranChem.Fields[J+2].AsVariant;
    end;
    sbmain.Panels[1].Text :='Processing record '+IntToStr(N)+' of '+IntToStr(TotalRecs);
    sbMain.Refresh;
    n:=n+1;
    dmCor.cdsCoranChem.Next;
  until (dmCor.cdsCoranChem.Eof);

  dmCor.EigenVec.Open;
  sbmain.Panels[1].Text :='Reading discriminant factors for '+IntToStr(Nox)+' elements';
  sbMain.Refresh;
  dmCor.EigenVec.First;
  for J:=1 to Nox do
  begin
    A2[j,j] := dmCor.EigenVec.Fields[1].AsFloat;
    //ShowMessage(IntToStr(j)+ ' '+FormatFloat('###0.0000',A2[j,j]));
    dmCor.EigenVec.Next;
  end;
  CalcProjectedDiscrmScores;
  tsGraph1.TabVisible := true;
  tsGraph2.TabVisible := false;
  tsGraph3.TabVisible := false;
  tsScores.TabVisible := true;
  tsLocalities.TabVisible := true;
  pVar.Visible := false;
  pGraph1Var.Visible := false;
  pGraph2Var.Visible := false;
  pGraph2Var.Visible := false;
  pOutlierMapVar.Visible := false;
  p3DVar.Visible := false;
  pLocalitiesVar.Visible := false;
  PrepareGraphs;
end;

procedure TfmCoranMain.CalcProjectedDiscrmScores;
var
  ii, I, J, K, L : integer;
  E : double;
begin
  {clear factor loadings table for variables}
  dmCor.FacLoadingsVar.Open;
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
  j := 1;
  ii := 0;
  for i := 1 to Nox do
  begin
    ii := ii + 1;
    dmCor.FacLoadingsVar.Append;
    dmCor.FacLoadingsVarPos.AsInteger := i;
    dmCor.FacLoadingsVarCalled.AsString := OxideName[i];
    for j := 1 to 2 do
    begin
      case j of
        1 : begin
          dmCor.FacLoadingsVarVector1.AsFloat := A2[i,i];
        end;
        2 : begin
        end;
      end;
    end;
    dmCor.FacLoadingsVar.Post;
  end;
  dmCor.FacLoadingsVar.First;
  {clear factor loadings table for samples}
  dmCor.cdsFacLoadingsSmp.Open;
  dmCor.cdsFacLoadingsSmp.Last;
  if not (dmCor.cdsFacLoadingsSmp.Bof and dmCor.cdsFacLoadingsSmp.Eof) then
  begin
    dmCor.cdsFacLoadingsSmp.Last;
    repeat
      dmCor.cdsFacLoadingsSmp.Delete;
    until dmCor.cdsFacLoadingsSmp.Bof;
  end;
  sbmain.Panels[1].Text :='Calculating scores for all data';
  sbMain.Refresh;
  dmCor.cdsCoranChem.Close;
  dmCor.cdsCoranChem.Open;
  dmCor.cdsCoranChem.First;
  dmCor.QGroupedSmp.Open;
  dmCor.QGroupedSmp.First;
  {fill factor loadings table for samples for all data}
  j := 1;
  j := 2;
  i := 1;
  k := 1;
  repeat
    sbmain.Panels[1].Text :='Calculating scores for '+dmCor.cdsCoranChemSAMPLENUM.AsString;
    sbMain.Refresh;
    D := 0.0;
    for l := 1 to Nox do
    begin
      D := D + A2[l,l] * X[k,l];
    end;
    dmCor.QGroupedSmp.Append;
    dmCor.QGroupedSmpGroupName.AsString := dmCor.cdsCoranChemGROUPNAME.AsString;
    dmCor.QGroupedSmpPlotGroupName.AsString := dmCor.cdsCoranChemPlotGroupName.AsString;
    dmCor.QGroupedSmpSampleNum.AsString := dmCor.cdsCoranChemSAMPLENUM.AsString;
    j := 1;
    dmCor.QGroupedSmpVector1.AsFloat := D;
    j := 2;
    dmCor.QGroupedSmpVector2.AsFloat := 1.0;
    dmCor.QGroupedSmp.Post;
    dmCor.cdsCoranChem.Next;
    i := i + 1;
    k := k + 1;
  until dmCor.cdsCoranChem.Eof;
  dmCor.cdsCoranChem.First;
  dmCor.QGroupedSmp.Close;
  dmCor.QGroups.Open;
  dmCor.QGroups.First;
  dmCor.QGroups.Filtered := false;
  i := 0;
  repeat
    i := i + 1;
    dmCor.cdsFacLoadingsSmp.Close;
    dmCor.cdsFacLoadingsSmp.Filter := 'GROUPNAME = '+''''+dmCor.QGroupsGroupName.AsString+'''';
    dmCor.cdsFacLoadingsSmp.Open;
    dmCor.cdsFacLoadingsSmp.Filtered := true;
    dmCor.cdsFacLoadingsSmp.First;
    repeat
      dmCor.cdsFacLoadingsSmp.Edit;
      dmCor.cdsFacLoadingsSmpVector2.AsFloat := 1.0*i;
      dmCor.cdsFacLoadingsSmp.Post;
      dmCor.cdsFacLoadingsSmp.Next;
    until dmCor.cdsFacLoadingsSmp.Eof;
    dmCor.QGroups.Next;
  until dmCor.QGroups.Eof;
  dmCor.cdsFacLoadingsSmp.Filtered := false;
  dmCor.cdsFacLoadingsSmp.Close;
  dmCor.QGroups.Filtered := false;
  dmCor.QGroups.Close;
end;

procedure TfmCoranMain.ImportDataDefinitions1Click(Sender: TObject);
begin
  try
    ImportForm2 := TfmSheetImport2.Create(Self);
    ImportForm2.OpenDialogSprdSheet.FileName := 'CoranDefinitions';
    if (ImportForm2.ShowModal = mrOK) then
    begin
      sbMain.Panels[1].Text := 'New definitions imported';
      sbMain.Refresh;
    end else
    begin
      sbMain.Panels[1].Text := 'Cancelled import of new definitions';
      sbMain.Refresh;
    end;
  finally
    ImportForm2.Free;
  end;
  //ShowMessage('1');
end;

procedure TfmCoranMain.bbExitClick(Sender: TObject);
begin
  dmCor.Coran.Connected := false;
  Close;
end;

procedure TfmCoranMain.DiscrimMulti;
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
begin
{}
end;

procedure TfmCoranMain.udRotationClick(Sender: TObject;
  Button: TUDBtnType);
begin
  DBChart8.View3DOptions.Rotation := udRotation.Position;
end;

procedure TfmCoranMain.udElevationClick(Sender: TObject;
  Button: TUDBtnType);
begin
  DBChart8.View3DOptions.Elevation := udElevation.Position;
end;

procedure TfmCoranMain.udPerspectiveClick(Sender: TObject;
  Button: TUDBtnType);
begin
  DBChart8.View3DOptions.Perspective := udPerspective.Position;
end;

procedure TfmCoranMain.pc1Change(Sender: TObject);
begin
  if (pc1.ActivePageIndex in [1,4,5,6,7,8,9,10,11,12,13,14])
    then sbMain.Panels[1].Text := 'Graphs : left click and drag to zoom; double click to export'
    else sbMain.Panels[1].Text := '';
  if (pc1.ActivePageIndex in [4,6,7,8,9,10,11,12,13,14])
    then PrintGraph1.Enabled := true
    else PrintGraph1.Enabled := false;
  if ((pc1.ActivePageIndex in [8]) and ScaleAxesEqually1.Checked)
    then ScaleAxesEquallyGraph1;
  if ((pc1.ActivePageIndex in [9]) and ScaleAxesEqually1.Checked)
    then ScaleAxesEquallyGraph2;
  if ((pc1.ActivePageIndex in [10]) and ScaleAxesEqually1.Checked)
    then ScaleAxesEquallyGraph3;
end;

procedure TfmCoranMain.udZoomClick(Sender: TObject; Button: TUDBtnType);
begin
  DBChart8.View3DOptions.Zoom := udZoom.Position;
end;

procedure TfmCoranMain.sbShow3DClick(Sender: TObject);
begin
  DBChart8.BottomAxis.Title.Caption := cbX.Text;
  DBChart8.LeftAxis.Title.Caption := cbY.Text;
  DBChart8.DepthAxis.Title.Caption := cbZ.Text;

  DBChart8.Series[0].Active := false;
  DBChart8.Series[1].Active := false;
  DBChart8.Series[2].Active := false;
  DBChart8.Series[3].Active := false;
  DBChart8.Series[4].Active := false;

  DBChart8.Series[0].DataSource := dmCor.FacLoadingsVar;
  DBChart8.Series[0].XValues.ValueSource := cbX.Text;
  DBChart8.Series[0].YValues.ValueSource := cbY.Text;
  (DBChart8.Series[0] as TPoint3DSeries).ZValues.ValueSource := cbZ.Text;
  case scSimilarityChoice of
    scCorrespondence,scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
      DBChart8.Series[0].Visible := true;
    end;
    scPCAVariance,scPCAPearsonR,scPCASpearmanR, scPCAKendallR : begin
      DBGrid29.Enabled := false;
      DBNavigator21.Enabled := false;
      DBChart8.Series[0].Visible := false;
    end;
    scDiscrim2Grp,scDiscrimnGrp,scCluster : begin
      DBGrid29.Enabled := false;
      DBNavigator21.Enabled := false;
      DBChart8.Series[0].Visible := false;
    end;
  end;

  DBChart8.Series[1].DataSource := dmCor.cdsFacLoadingsSmp;
  DBChart8.Series[1].XValues.ValueSource := cbX.Text;
  DBChart8.Series[1].YValues.ValueSource := cbY.Text;
  (DBChart8.Series[1] as TPoint3DSeries).ZValues.ValueSource := cbZ.Text;
  DBChart8.Series[1].Active := true;

  DBChart8.Series[2].DataSource := dmCor.GroupedSmp;
  DBChart8.Series[2].XValues.ValueSource := cbX.Text;
  DBChart8.Series[2].YValues.ValueSource := cbY.Text;
  (DBChart8.Series[2] as TPoint3DSeries).ZValues.ValueSource := cbZ.Text;
  DBChart8.Series[2].Active := true;

  DBChart8.Series[3].DataSource := dmCor.QGroupedSmp;
  DBChart8.Series[3].XValues.ValueSource := cbX.Text;
  DBChart8.Series[3].YValues.ValueSource := cbY.Text;
  (DBChart8.Series[3] as TPoint3DSeries).ZValues.ValueSource := cbZ.Text;
  DBChart8.Series[3].Active := true;

  DBChart8.Series[4].DataSource := dmCor.FacLoadingsVarLinked;
  DBChart8.Series[4].XValues.ValueSource := cbX.Text;
  DBChart8.Series[4].YValues.ValueSource := cbY.Text;
  (DBChart8.Series[4] as TPoint3DSeries).ZValues.ValueSource := cbZ.Text;
  DBChart8.Series[4].Active := true;
end;

procedure TfmCoranMain.AddToComboBoxXYZ(i : integer; Choice : string);
var
  tmpStr : string;
begin
  tmpStr := 'Vector' + IntToStr(i);
  if (Choice = 'X') then cbX.Items.Add(tmpStr);
  if (Choice = 'Y') then cbY.Items.Add(tmpStr);
  if (Choice = 'Z') then cbZ.Items.Add(tmpStr);
end;

procedure TfmCoranMain.AddToComboBoxVar(Nox : integer);
var
  i : integer;
begin
  dmCor.ElemNames.Open;
  dmCor.ElemNames.First;
  cbRawGraphVar.Clear;
  for i := 1 to Nox do
  begin
    //if (i = 1) then cbRawGraphVar.Text := dmCor.ElemNamesCalled.AsString;
    cbRawGraphVar.Items.Add(dmCor.ElemNamesCalled.AsString);
    dmCor.ElemNames.Next;
  end;
  dmCor.ElemNames.First;
  cbRawGraphVar.ItemIndex := 0;
end;

procedure TfmCoranMain.AddToComboBoxScoreVar(Nox : integer);
var
  i : integer;
begin
  cbScoreGraphVar.Clear;
  if (scSimilarityChoice in [scCorrespondence,scRQVariance,scRQPearsonR,
                   scPCAVariance,scPCAPearsonR,scPCASpearmanR,
                   scPCAKendallR,scRQSpearmanR,scRQKendallR,
                   scRobPCA]) then
  begin
    for i := 1 to Nox do
    begin
      cbScoreGraphVar.Items.Add('Vector'+IntToStr(i));
    end;
  end;
  if (scSimilarityChoice in [scDiscrim2Grp]) then
  begin
    cbScoreGraphVar.Items.Add('Vector'+IntToStr(1));
  end;
  cbScoreGraphVar.ItemIndex := 0;
end;

procedure TfmCoranMain.AddToComboBoxVariableID(A : SingleArrayM; i : integer);
var
  tmpStr : string;
begin
  tmpStr := 'Vector' + IntToStr(i);
  if (A[i] > 0.001) then cbVariableID.Items.Add(tmpStr);
end;

procedure TfmCoranMain.ImportEigenValues1Click(Sender: TObject);
begin
  EigenValuesImported := false;
  try
    ImportFormEigVal := TfmSheetImportEigVal.Create(Self);
    ImportFormEigVal.OpenDialogSprdSheet.FileName := '';
    ImportFormEigVal.ShowModal;
    EigenValuesImported := true;
  finally
    ImportFormEigVal.Free;
  end;
  Project1.Enabled := (EigenValuesImported and EigenVectorsImported and MeanValuesImported);
  ProjectCorrespondenceanalysis2.Enabled := (EigenValuesImported and EigenVectorsImported and MeanValuesImported);
  PrincipalComponentAnalysis1.Enabled := (EigenValuesImported and EigenVectorsImported and MeanValuesImported);
  SimultaneousRandQmode1.Enabled :=(EigenValuesImported and EigenVectorsImported and MeanValuesImported);
  DiscriminantFunctionAnalysis2.Enabled := false;
  sbMain.Panels[1].Text := 'Eigen values imported';
  sbMain.Refresh;
end;

procedure TfmCoranMain.ImportEigenvectors1Click(Sender: TObject);
begin
  EigenVectorsImported := true;
  try
    ImportFormEigVec := TfmSheetImportEigVec.Create(Self);
    ImportFormEigVec.OpenDialogSprdSheet.FileName := '';
    ImportFormEigVec.ShowModal;
    EigenValuesImported := true;
  finally
    ImportFormEigVec.Free;
  end;
  Project1.Enabled := (EigenValuesImported and EigenVectorsImported and MeanValuesImported);
  ProjectCorrespondenceanalysis2.Enabled := (EigenValuesImported and EigenVectorsImported and MeanValuesImported);
  PrincipalComponentAnalysis1.Enabled := (EigenValuesImported and EigenVectorsImported and MeanValuesImported);
  SimultaneousRandQmode1.Enabled :=(EigenValuesImported and EigenVectorsImported and MeanValuesImported);
  DiscriminantFunctionAnalysis2.Enabled := false;
  sbMain.Panels[1].Text := 'Eigen vectors imported';
  sbMain.Refresh;
end;

procedure TfmCoranMain.udRotationLocClick(Sender: TObject; Button: TUDBtnType);
begin
  DBChart9.View3DOptions.Rotation := udRotationLoc.Position;
end;

procedure TfmCoranMain.udElevationLocClick(Sender: TObject; Button: TUDBtnType);
begin
  DBChart9.View3DOptions.Elevation := udElevationLoc.Position;
end;

procedure TfmCoranMain.udPerspectiveLocClick(Sender: TObject; Button: TUDBtnType);
begin
  DBChart9.View3DOptions.Perspective := udPerspectiveLoc.Position;
end;

procedure TfmCoranMain.udZoomLocClick(Sender: TObject; Button: TUDBtnType);
begin
  DBChart9.View3DOptions.Zoom := udZoomLoc.Position;
end;

procedure TfmCoranMain.sbShow3DLocClick(Sender: TObject);
begin
  DBChart9.BottomAxis.Title.Caption := cbXLoc.Text;
  DBChart9.LeftAxis.Title.Caption := cbYLoc.Text;
  DBChart9.DepthAxis.Title.Caption := cbZLoc.Text;

  DBChart9.Series[0].Active := false;
  DBChart9.Series[1].Active := false;
  DBChart9.Series[2].Active := false;

  DBChart9.Series[0].DataSource := dmCor.SmpLoc;
  DBChart9.Series[0].XValues.ValueSource := cbXLoc.Text;
  DBChart9.Series[0].YValues.ValueSource := cbYLoc.Text;
  (DBChart9.Series[0] as TPoint3DSeries).ZValues.ValueSource := cbZLoc.Text;
  DBChart9.Series[0].Active := true;

  DBChart9.Series[1].DataSource := dmCor.GroupedSmpLoc;
  DBChart9.Series[1].XValues.ValueSource := cbXLoc.Text;
  DBChart9.Series[1].YValues.ValueSource := cbYLoc.Text;
  (DBChart9.Series[1] as TPoint3DSeries).ZValues.ValueSource := cbZLoc.Text;
  DBChart9.Series[1].Active := true;

  DBChart9.Series[2].DataSource := dmCor.QGroupedSmpLoc;
  DBChart9.Series[2].XValues.ValueSource := cbXLoc.Text;
  DBChart9.Series[2].YValues.ValueSource := cbYLoc.Text;
  (DBChart9.Series[2] as TPoint3DSeries).ZValues.ValueSource := cbZLoc.Text;
  DBChart9.Series[2].Active := true;
end;

procedure TfmCoranMain.CalculateProjections;
var
  ii, jj, Noxtemp : integer;
  n :integer;
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
  Nox := dmCor.ElemNames.RecordCount;
  for ii := Nox to MM-1 do
  begin
    DBGrid1.Columns[ii+3].Visible := false;
    DBGrid2.Columns[ii+3].Visible := false;
  end;
  for ii := Nox to MM-1 do
  begin
    dbgFacLoadingsVar.Columns[ii+2].Visible := false;
    DBGridStats.Columns[ii+2].Visible := false;
  end;
  for ii := Nox to MM-1 do
  begin
    dbgFacLoadingsSmp.Columns[ii+3].Visible := false;
  end;
  for ii := Nox to MM-1 do
  begin
    dbgEigenVec.Columns[ii+1].Visible := false;
  end;
  for ii := Nox to MM-1 do
  begin
    dbgSimilarity.Columns[ii+2].Visible := false;
  end;
  I:=0;
  n:=1;
  TooMuch:=false;
  TotalRecs := dmCor.cdsCoranChem.RecordCount;
  sbmain.Panels[1].Text :='Reading data for '+IntToStr(Nox)+' elements';
  sbMain.Refresh;
  dmCor.cdsCoranChem.First;
  repeat
    if ((I<(NN-10))) then
    begin
      I:=I+1;
      ComponentStr[I]:=dmCor.cdsCoranChemSampleNum.AsString;
      for J:=1 to Nox do
      begin
        X[I,J] := dmCor.cdsCoranChem.Fields[J+1].AsVariant;
      end;
    end;
    if ((I>=(NN-10))) then TooMuch:=true;
    sbmain.Panels[1].Text :='Processing record '+IntToStr(N)+' of '+IntToStr(TotalRecs)+' Total included '+IntToStr(I);
    sbMain.Refresh;
    n:=n+1;
    dmCor.cdsCoranChem.Next;
  until ((n>TotalRecs) or (I>=NN-10) or dmCor.cdsCoranChem.eof);
  if TooMuch then
  begin
    MessageDlg('Data overflow. Truncating!!',mtWarning,[mbOK],0);
  end;
  NumSamples:=I;
  //M:=Nox;
  for I:=1 to Nox-1 do begin
    VarbSymbol[I]:=48+I;
  end;
  VarbSymbol[Nox]:=48;
  if (NumSamples<Nox) then begin
    Nox:=NumSamples;
    MessageDlg('Insufficient data in file. Decreasing variables',mtWarning,[mbOK],0);
  end;
  i := 1;
  dmCor.ElemNames.First;
  repeat
    if (dmCor.ElemNamesPos.AsInteger > 0) then
    begin
      OxideName[dmCor.ElemNamesPos.AsInteger] := dmCor.ElemNamesCalled.AsString;
    end;
    i := i + 1;
    dmCor.ElemNames.Next;
  until ((dmCor.ElemNames.Eof) or (i > Nox));
  dmCor.ElemNames.First;
  sbmain.Panels[1].Text :='Calculating similarity matrix';
  sbMain.Refresh;
  case scSimilarityChoice of
    scCorrespondence : begin
      ScaleCorAnal(NumSamples,Nox);
    end;
    scRQVariance : begin
      ScaleRQAnalVar(NumSamples,Nox);
    end;
    scRQPearsonR : begin
      ScaleRQAnalStd(NumSamples,Nox);
    end;
    scPCAVariance : begin
      Cov2(X,A3,NumSamples,Nox);
    end;
    scPCAPearsonR : begin
      Stand(X,NumSamples,Nox);
      RCoef2(X,A3,NumSamples,Nox);
    end;
    scDiscrim2Grp : begin
    end;
    scDiscrimnGrp : begin
    end;
    scCluster : begin
    end;
    scPCASpearmanR : begin
      SpearmanRho(X,A3,NumSamples,Nox);
    end;
    scPCAKendallR : begin
      KendallTau(X,A3,NumSamples,Nox);
    end;
    scRQSpearmanR : begin
    end;
    scRQKendallR : begin
    end;
  end;
  for I:=1 to NumSamples do
  begin
    for J:=1 to Nox do
    begin
      if (X[I,J]=0.0) then X[I,J]:=0.001;
      X1[i,j] := X[i,j];
    end;
  end;
  for I:=1 to Nox do
  begin
    Str(I:10,tempstr);
    EigName[I]:=tempstr;
  end;
  case scSimilarityChoice of
    scCorrespondence,scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
      for i:= 1 to NumSamples do
      begin
        for j:= 1 to Nox do
        begin
          X1[i,j] := W[i,j];
        end;
      end;
      Transp(W,WP,NumSamples,Nox);
    end;
    scPCAVariance,scPCAPearsonR,scPCASpearmanR,scPCAKendallR : begin
      {nothing required here}
    end;
  end;
  Memoresults.Lines.Add(' ');
  {try writing all in separate columns}
  for i := 1 to Nox do
  begin
    tmpstr := '';
    tmpstr := tmpstr + FormatFloat('00',i);
    ResultsArray[i,1] := tmpstr + '   ';
    tmpstr := OxideName[i];
    ResultsArray[i,2] := tmpstr + CharStream(10-Length(tmpstr),32) + blankstr;
    tmpStr := TakeLogs[i];
    ResultsArray[i,3] := ' ' + tmpstr + '  ';
    for J:=1 to Nox do
    begin
      if (A3[i,J] <= 9000000.0) then tmpStr := FormatFloat('###0.000',A3[i,J]);
      if (A3[i,J] > 9000000.0) then tmpStr := FormatFloat('###0.00',A3[i,J]);
      ResultsArray[i,j+3] := CharStream(12-Length(tmpstr),32)+tmpstr + blankstr;
    end;
  end;
  tmpstr := '                     ';
  for i := 1 to Nox do
  begin
    tmpstr := tmpstr + FormatFloat('    00      ',i);
  end;
  MemoResults.Lines.Add(tmpstr);
  Memoresults.Lines.Add(' ');
  for i := 1 to Nox do
  begin
    tmpstr := '';
    for j := 1 to Nox+3 do
    begin
      tmpstr := tmpstr + ResultsArray[i,j];
    end;
    MemoResults.Lines.Add(tmpstr);
  end;
  MemoResults.Lines.Add('  ');
  MemoResults.Lines.Add('  ');
  sbmain.Panels[1].Text :='Calculating eigenvectors and eigenvalues';
  sbMain.Refresh;
  SumE:=0.0;
  for I:=1 to Nox do
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
  for I:=1 to Nox do
  begin
    SumEE:=SumEE+Abs(A1[I,1]);
    A1[I,2]:=Abs(A1[I,1])*100.0/SumE;
    A1[I,3]:=SumEE*100.0/SumE;
    if not ((scSimilarityChoice=scCorrespondence) and (I=1)) then
    begin
      ChartEigenvalue.Series[0].AddXY(1.0*I,A1[I,2]);
    end;
  end;
  if (scSimilarityChoice=scCorrespondence) then
  begin
    A3[1,1]:=1.0;
    A1[1,1]:=1.0;
  end;
  if (scSimilarityChoice=scCorrespondence) then
  begin
    Memoresults.Lines.Add('Eigenvalue 1 and Eigenvector 1 are artifices of the method. Ignore them!');
    Memoresults.Lines.Add(' ');
  end;
  MemoResults.Lines.Add(' ');
  MemoResults.Lines.Add('Component   Eigen       % of      Cumulative');
  MemoResults.Lines.Add('            value       trace     % of trace');
  MemoResults.Lines.Add(' ');
  {try writing all in separate columns}
  for i := 1 to Nox do
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
  for i := 1 to Nox do
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
  sbmain.Panels[1].Text :='Calculating factor loadings';
  sbMain.Refresh;
  case scSimilarityChoice of
    scCorrespondence : begin
      ScaleCorAnal(NumSamples,Nox);
      {
      for J:=1 to M do
      begin
        A3[J,J]:=Sqrt(Abs(A3[J,J]));
      end;
      }
      MmultC(A2,A3,TempC,Nox,Nox,Nox);  {factor loadings for variables}
      for J:=1 to Nox do
      begin
        A3[J,J]:=1.0/A3[J,J];
      end;
      MmultC(DC,TempC,A1,Nox,Nox,Nox);  {scaled factor loadings for variables}
      Mmult(W,A2,B,NumSamples,Nox,Nox);
      Mmult(B,A3,W,NumSamples,Nox,Nox);
      MmultR(DR,W,TempM,NumSamples,Nox);     {factor loadings for samples}
      for J:=1 to Nox do
      begin
        A3[J,J]:=1.0/A3[J,J];
      end;
      Mmult(TempM,A3,B,NumSamples,Nox,Nox);    {scaled factor loadings for samples}
    end;
    scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
      if (scSimilarityChoice = scRQVariance) then ScaleRQAnalVar(NumSamples,Nox);
      if (scSimilarityChoice = scRQPearsonR) then ScaleRQAnalStd(NumSamples,Nox);
      for J:=1 to Nox do
      begin
        A3[J,J]:=Sqrt(Abs(A3[J,J]));
      end;
      MmultC(A2,A3,A1,Nox,Nox,Nox);  {factor loadings for variables}
      Mmult(W,A2,B,NumSamples,Nox,Nox);     {factor loadings for samples}
    end;
    scPCAVariance,scPCAPearsonR,scPCASpearmanR,scPCAKendallR : begin
      Mmult(X,A2,B,NumSamples,Nox,Nox);     {factor scores for samples}
    end;
  end;
  CalcComponentScores;
  dmCor.QGroups.Open;
  dmCor.QPlotGroups.Open;
  if ((scSimilarityChoice=scCorrespondence) and (Nox>4)) then
  begin
    for I:=1 to NumSamples do
    begin
     DR[I]:=1.0/DR[I];
     DR[I]:=DR[I]*DR[I];
    end;
    for J:=1 to Nox do
    begin
     DC[J,J]:=1.0/DC[J,J];
     DC[J,J]:=DC[J,J]*DC[J,J];
     A3[J,J]:=A3[J,J]*A3[J,J];
    end;
  end;
  pVar.Visible := true;
  pGraph1Var.Visible := true;
  pGraph2Var.Visible := true;
  pGraph3Var.Visible := true;
  pOutlierMapVar.Visible := false;
  p3DVar.Visible := false;
  pLocalitiesVar.Visible := false;
  case scSimilarityChoice of
    scCorrespondence : begin
       lIgnoreLoadings.Visible := true;
       lIgnoreLoadingsVar.Visible := true;
       lIgnoreLoadingsSmp.Visible := true;
       NX:=2;
    end;
    scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
       lIgnoreLoadings.Visible := false;
       lIgnoreLoadingsVar.Visible := false;
       lIgnoreLoadingsSmp.Visible := false;
       NX:=1;
    end;
    scPCAVariance,scPCAPearsonR,scPCASpearmanR, scPCAKendallR : begin
       lIgnoreLoadings.Visible := false;
       lIgnoreLoadingsVar.Visible := false;
       lIgnoreLoadingsSmp.Visible := false;
       pVar.Visible := true;
       pGraph1Var.Visible := false;
       pGraph2Var.Visible := false;
       pGraph3Var.Visible := false;
       pOutlierMapVar.Visible := false;
       p3DVar.Visible := false;
       pLocalitiesVar.Visible := false;
       NX:=1;
    end;
    scDiscrim2Grp,scDiscrimnGrp,scCluster : begin
       lIgnoreLoadings.Visible := false;
       lIgnoreLoadingsVar.Visible := false;
       lIgnoreLoadingsSmp.Visible := false;
       NX:=1;
    end;
  end;
  case scSimilarityChoice of
    scCorrespondence,scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
       pVar.Visible := true;
       Xmin:=A1[1,NX];
       Xmax:=Xmin;
       Ymin:=A1[1,NX+1];
       Ymax:=Ymin;
       Zmin:=A1[1,NX+2];
       Zmax:=Zmin;
       for J:=1 to Nox do
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
    scPCAVariance,scPCAPearsonR,scPCASpearmanR,scPCAKendallR : begin
       pVar.Visible := true;
       Xmin:=B[1,NX];
       Xmax:=Xmin;
       Ymin:=B[1,NX+1];
       Ymax:=Ymin;
       Zmin:=B[1,NX+2];
       Zmax:=Zmin;
    end;
  end;
  for I:=1 to NumSamples do
  begin
    if (Xmin > B[I,NX]) then Xmin:=B[I,NX];
    if (Xmax < B[I,NX]) then Xmax:=B[I,NX];
    if (Ymin > B[I,NX+1]) then Ymin:=B[I,NX+1];
    if (Ymax < B[I,NX+1]) then Ymax:=B[I,NX+1];
    if (Zmin > B[I,NX+2]) then Zmin:=B[I,NX+2];
    if (Zmax < B[I,NX+2]) then Zmax:=B[I,NX+2];
    for K:=1 to 5 do
    begin
      //X[I+Nox,K]:=B[I,K];  //as it used to be Oct 2006. Not sure why I+Nox
      X[I,K]:=B[I,K];
    end;
  end;
  tsGraph1.TabVisible := false;
  tsGraph2.TabVisible := false;
  tsGraph3.TabVisible := false;
  tsLocalities.TabVisible := false;
  tsScores.TabVisible := false;
  pVar.Visible := true;
  pGraph1Var.Visible := true;
  pOutlierMapVar.Visible := true;
  p3DVar.Visible := true;
  pLocalitiesVar.Visible := true;
  case scSimilarityChoice of
    scCorrespondence : begin
      tsGraph1.TabVisible := true;
      tsGraph2.TabVisible := true;
      tsGraph3.TabVisible := true;
      tsLocalities.TabVisible := true;
      tsScores.TabVisible := true;
      PrepareGraphs;
    end;
    scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
      tsGraph1.TabVisible := true;
      tsGraph2.TabVisible := true;
      tsGraph3.TabVisible := true;
      tsScores.TabVisible := true;
      tsLocalities.TabVisible := true;
      PrepareGraphs;
    end;
    scPCAVariance,scPCAPearsonR,scPCASpearmanR, scPCAKendallR : begin
      tsGraph1.TabVisible := true;
      tsGraph2.TabVisible := true;
      tsGraph3.TabVisible := true;
      tsScores.TabVisible := false;
      tsLocalities.TabVisible := true;
      pVar.Visible := true;
      pGraph1Var.Visible := false;
      pOutlierMapVar.Visible := false;
      p3DVar.Visible := false;
      pLocalitiesVar.Visible := false;
      PrepareGraphs;
    end;
  end;
end;

procedure TfmCoranMain.SelectProcess(Sender: TObject);
begin
  ChartEigenValue.Visible := true;
  Panel45.Visible := true;
  dmCor.cdsCoranChem.Open;
  dmCor.ElemNames.Open;
  NumSamples := dmCor.cdsCoranChem.RecordCount;
  Nox := dmCor.ElemNames.RecordCount;
  NumEigenVectorsSelected := 10;
  if (NumEigenVectorsSelected > Nox) then NumEigenVectorsSelected := Nox;
  if (Sender = Correspondenceanalysis1) then scSimilarityChoice := scCorrespondence;
  if (Sender = RQmodevariance1) then scSimilarityChoice := scRQVariance;
  if (Sender = RQmodePearsonscorrelation1) then scSimilarityChoice := scRQPearsonR;
  if (Sender = PCAvariance1) then scSimilarityChoice := scPCAVariance;
  if (Sender = PCAPearsonscorrelation1) then scSimilarityChoice := scPCAPearsonR;
  if (Sender = Discriminantanalysis2group1) then scSimilarityChoice := scDiscrim2Grp;
  if (Sender = Discriminantanalysisngroup1) then scSimilarityChoice := scDiscrimnGrp;
  if (Sender = Clusteranalysis1) then scSimilarityChoice := scCluster;
  if (Sender = PCASpearmanscorrelation1) then scSimilarityChoice := scPCASpearmanR;
  if (Sender = PCAKendallscorrelation1) then scSimilarityChoice := scPCAKendallR;
  if (Sender = RQmodeSpearmanscorrelation1) then scSimilarityChoice := scRQSpearmanR;
  if (Sender = RQmodeKendallscorrelation1) then scSimilarityChoice := scRQKendallR;
  if ((Sender = RQmodeSpearmanscorrelation1) or
      (Sender = RQmodeKendallscorrelation1)) then
  begin
    MessageDlg('Relative scaling for samples and variables has not yet been correctly implemented',mtWarning,[mbOK],0);
  end;
  if ((eTitle.Text = '') or (Pos('_',eTitle.Text) = 1)) then
  begin
    if (Sender = Correspondenceanalysis1) then eTitle.Text := '_Correspondence Analysis';
    if (Sender = RQmodevariance1) then eTitle.Text := '_Simultaneous R- and Q-mode Component Analysis (Variance)';
    if (Sender = RQmodePearsonscorrelation1) then eTitle.Text := '_Simultaneous R- and Q-mode Component Analysis (Pearson R)';
    if (Sender = PCAvariance1) then eTitle.Text := '_Principal Component Analysis (Variance)';
    if (Sender = PCAPearsonscorrelation1) then eTitle.Text := '_Principal Component Analysis (Pearson R)';
    if (Sender = Discriminantanalysis2group1) then eTitle.Text := '_Discriminant Analysis (2 group)';
    if (Sender = Discriminantanalysisngroup1) then eTitle.Text := '_Discriminant Analysis (n group)';
    if (Sender = Clusteranalysis1) then eTitle.Text := '_Cluster Analysis';
    if (Sender = PCASpearmanscorrelation1) then eTitle.Text := '_Principal Component Analysis (Spearman rho)';
    if (Sender = PCAKendallscorrelation1) then eTitle.Text := '_Principal Component Analysis (Kendall tau)';
    if (Sender = RQmodeSpearmanscorrelation1) then eTitle.Text := '_Simultaneous R- and Q-mode Component Analysis (Spearman rho)';
    if (Sender = RQmodeKendallscorrelation1) then eTitle.Text := '_Simultaneous R- and Q-mode Component Analysis (Kendall tau)';
  end;
  //DisableEnableControls('Disable');
  CheckColumnTotalsStats(NumSamples,Nox,TotalsOK);
  //ShowMessage('Finished CheckColumnTotalStats');
  if (TotalsOK) then
  begin
    NumEigenvectorsSelected := Nox;
    AssignChartDataSources('Close');
    pc1.Enabled := false;
    dmCor.cdsFacLoadingsSmp.Open;
    dmCor.FacLoadingsVar.Open;
    dmCor.cdsCoranChem.Open;
    dmCor.CoranSimilarity.Open;
    TooMuch := false;
    EigenValuesImported := false;
    EigenVectorsImported := false;
    Project1.Enabled := (EigenValuesImported and EigenVectorsImported);
    try
      dmCor.QGroups.Open;
      dmCor.QPlotGroups.Open;
    except
    end;
    dmCor.EigenVal.Open;
    dmCor.EigenVec.Open;
    dmCor.CoranSimilarity.Open;
    DBChart7.Visible := true;
    Panel40.Visible := true;
    if (scSimilarityChoice in [scCorrespondence,scRQVariance,scRQPearsonR,
                     scPCAVariance,scPCAPearsonR]) then DoProcess;
    if (scSimilarityChoice = scDiscrim2Grp) then Discrm;
    if (scSimilarityChoice = scDiscrimnGrp) then DiscrimMulti;
    if (scSimilarityChoice = scCluster) then Cluster;
    if (scSimilarityChoice in [scPCASpearmanR, scPCAKendallR,scRQSpearmanR,
                    scRQKendallR]) then DoProcess;
    //ShowMessage('Finished calculating similarity matrices');
    dmCor.cdsFacLoadingsSmp.Open;
    dmCor.cdsFacLoadingsSmp.First;
    dmCor.cdsCoranChem.First;
    tsResults.TabVisible := true;
    tsSimilarity.TabVisible := true;
    tsEigenValues.TabVisible := true;
    tsLoadings.TabVisible := true;
    tsScores.TabVisible := true;
    tsOutlierMap.TabVisible := true;
    tsGraph1.TabVisible := true;
    tsGraph2.TabVisible := true;
    tsGraph3.TabVisible := true;
    ts3D.TabVisible := true;
    tsLoc3D.TabVisible := true;
    tsLoc4D.TabVisible := Include4DVarData;
    DBGrid29.Enabled := true;
    DBNavigator21.Enabled := true;
    sbMain.Panels[1].Text := 'Completed';
    case scSimilarityChoice of
      scCorrespondence,scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
      end;
      scPCAVariance,scPCAPearsonR,scPCASpearmanR, scPCAKendallR : begin
      end;
      scDiscrim2Grp : begin
        tsSimilarity.TabVisible := false;
        tsEigenValues.TabVisible := false;
        tsScores.TabVisible := false;
        tsOutlierMap.TabVisible := false;
        tsGraph2.TabVisible := false;
        tsGraph3.TabVisible := false;
      end;
      scDiscrimnGrp,scCluster : begin
        tsSimilarity.TabVisible := false;
        tsEigenValues.TabVisible := false;
        tsScores.TabVisible := false;
        tsOutlierMap.TabVisible := false;
        tsGraph2.TabVisible := false;
        tsGraph3.TabVisible := false;
      end;
    end;
    //ShowMessage('Before PrepareComboBoxes');
    PrepareComboBoxes(NumEigenVectorsSelected);
    //ShowMessage('After PrepareComboBoxes and Before CalculateHingeStatistics');
    CalculateHingeStatistics(NumEigenVectorsSelected);
    //ShowMessage('After CalculateHingeStatistics and Before PrepareGraphes');
    PrepareGraphs;
    pc1.Enabled := true;
    DisableEnableControls('Enable');
    AssignChartDataSources('Close');
    AssignChartDataSources('Open');
    sbMain.Panels[1].Text := 'Completed';
    sbMain.Refresh;
    pc1.ActivePage := tsControl;
  end;
end;

procedure TfmCoranMain.ProjectSelectProcess(Sender: TObject);
var
  i, j : integer;
begin
  dmCor.ElemNames.Open;
  NumSamples := dmCor.cdsCoranChem.RecordCount;
  NumEigenVectorsSelected := 10;
  if (NumEigenVectorsSelected > Nox) then NumEigenVectorsSelected := Nox;
  AssignChartDataSources('Close');
  DisableEnableControls('Disable');
  FillChar(A3,SizeOf(A3),0);
  FillChar(A1,SizeOf(A1),0);
  FillChar(A2,SizeOf(A2),0);
  dmCor.EigenVal.Open;
  dmCor.EigenVec.Open;
  dmCor.EigenVal.First;
  dmCor.EigenVec.First;
  for i := 1 to Nox do
  begin
    A3[i,i] := dmCor.EigenValEigenValue.AsFloat;
    for j := 1 to Nox do
    begin
      A2[i,j] := dmCor.EigenVec.Fields[j].AsFloat;
    end;
    dmCor.EigenVal.Next;
    dmCor.EigenVec.Next;
  end;
  dmCor.EigenVal.First;
  dmCor.EigenVec.First;
  ChartEigenValue.Visible := true;
  Panel45.Visible := true;
  if (Sender = ProjectCorrespondenceanalysis2) then scSimilarityChoice := scCorrespondence;
  if (Sender = ProjectRQmodevariance2) then scSimilarityChoice := scRQVariance;
  if (Sender = ProjectRQmodePearsonscorrelation2) then scSimilarityChoice := scRQPearsonR;
  if (Sender = ProjectPCAvariance2) then scSimilarityChoice := scPCAVariance;
  if (Sender = ProjectPCAPearsonscorrelation2) then scSimilarityChoice := scPCAPearsonR;
  if (Sender = ProjectPCASpearmanscorrelation2) then scSimilarityChoice := scPCASpearmanR;
  if (Sender = ProjectPCAKendallscorrelation2) then scSimilarityChoice := scPCAKendallR;
  if (Sender = ProjectRQmodeSpearmanscorrelation2) then scSimilarityChoice := scRQSpearmanR;
  if (Sender = ProjectRQmodeKendallscorrelation2) then scSimilarityChoice := scRQKendallR;
  if ((Sender = ProjectRQmodeSpearmanscorrelation2) or
      (Sender = ProjectRQmodeKendallscorrelation2)) then
  begin
    MessageDlg('Relative scaling for samples and variables has not yet been correctly implemented',mtWarning,[mbOK],0);
  end;
  try
    dmCor.QGroups.Open;
    dmCor.QPlotGroups.Open;
  except
  end;
  dmCor.EigenVal.Open;
  dmCor.EigenVec.Open;
  dmCor.CoranSimilarity.Open;
  DBChart7.Visible := true;
  if (scSimilarityChoice < scDiscrim2Grp) then CalculateProjections;
  dmCor.cdsFacLoadingsSmp.First;
  dmCor.cdsCoranChem.First;
  PrepareComboBoxes(NumEigenVectorsSelected);
  CalculateHingeStatistics(NumEigenVectorsSelected);
  PrepareGraphs;
  tsResults.TabVisible := true;
  tsSimilarity.TabVisible := true;
  tsEigenValues.TabVisible := true;
  tsLoadings.TabVisible := true;
  tsScores.TabVisible := true;
  tsOutlierMap.TabVisible := true;
  tsGraph1.TabVisible := true;
  tsGraph2.TabVisible := true;
  tsGraph3.TabVisible := true;
  ts3D.TabVisible := true;
  tsLoc3D.TabVisible := true;
  tsLoc4D.TabVisible := Include4DVarData;
  AssignChartDataSources('Open');
  DisableEnableControls('Enable');
  sbMain.Panels[1].Text := 'Completed';
  sbMain.Refresh;
end;

procedure TfmCoranMain.cbComponentChange(Sender: TObject);
var
  tmpStr : string;
  i : integer;
begin
  i := Pos('t',cbComponent.Text)+1;
  tmpStr := 'Vector'+Copy(cbComponent.Text,i,length(cbComponent.Text)-i+1);
  DBChart10.Series[0].Active := false;
  DBChart10.Series[0].DataSource := dmCor.FacLoadingsVarNoZero;
  DBChart10.Series[0].XValues.ValueSource := tmpStr;
  {
  DBChart10.Series[0].YValues.ValueSource := dmCor.FacLoadingsVarPos.AsVariant;
  }
  DBChart10.Series[0].Active := true;
end;

procedure TfmCoranMain.CreateEigSprdShts;
var
  ii, I, J : integer;
  {
  A1t : RealArrayC;
  }
  SumE, SumEE : double;
begin
  {clear factor loadings table for variables}
  if (scSimilarityChoice < scRobPCA) then
  begin
    dmCor.EigenVal.First;
    i := 0;
    repeat
      i := i+1;
      TempC[i,1] := dmCor.EigenValEigenValue.AsFloat;     // was A1t
      TempC[i,2] := dmCor.EigenValEigenValuePct.AsFloat;
      TempC[i,3] := dmCor.EigenValEigenValueCumPct.AsFloat;
      dmCor.EigenVal.Next;
    until dmCor.EigenVal.Eof;
    dmCor.EigenVal.First;
    if (scSimilarityChoice=scCorrespondence) then TempC[1,1]:=0.0;
    SumE:=0.0;
    for I:=1 to Nox do
    begin
      SumE:=SumE+Abs(TempC[I,1]);
    end;
    TempC[1,2]:=0.0;
    TempC[1,3]:=0.0;
    SumEE:=0.0;
    for I:=1 to Nox do
    begin
      SumEE:=SumEE+Abs(TempC[I,1]);
      TempC[I,2]:=Abs(TempC[I,1])*100.0/SumE;
      TempC[I,3]:=SumEE*100.0/SumE;
    end;
    if (scSimilarityChoice=scCorrespondence) then
    begin
      TempC[1,1]:=1.0;
    end;
    {fill eigen values in spreadsheet}
  end;
  if (scSimilarityChoice < scRobPCA) then
  begin
    dmCor.EigenVec.First;
    for i := 1 to Nox do
    begin
      for j := 1 to Nox do
      begin
        TempC[i,j] := dmCor.EigenVec.Fields[j+1].AsFloat;
      end;
      dmCor.EigenVec.Next;
    end;
    dmCor.EigenVec.First;
    if (scSimilarityChoice=scCorrespondence) then
    begin
      for i := 1 to Nox do TempC[i,1]:=0.0;
    end;
    {fill eigen vectors in spreadsheet}
    dmCor.EigenVec.First;
  end;
end;

procedure TfmCoranMain.PrintResults1Click(Sender: TObject);
begin
  Printresults1.Checked := not Printresults1.Checked;
end;

procedure TfmCoranMain.PrintData1Click(Sender: TObject);
begin
  Printdata1.Checked := not Printdata1.Checked;
end;

procedure TfmCoranMain.ExportGraph1Click(Sender: TObject);
begin
  SaveDialogJPEG.InitialDir := JPEGPath;
  if (pc1.ActivePage = tsGraph1) then
  begin
    SaveDialogJPEG.FileName := 'Component_1_vs_2.jpg';
  end;
  if (pc1.ActivePage = tsGraph2) then
  begin
    SaveDialogJPEG.FileName := 'Component_1_vs_3.jpg';
  end;
  if (pc1.ActivePage = tsGraph3) then
  begin
    SaveDialogJPEG.FileName := 'Component_2_vs_3.jpg';
  end;
  if (pc1.ActivePage = tsEigenValues) then
  begin
    SaveDialogJPEG.FileName := 'Eigenvalues.jpg';
  end;
  if (pc1.ActivePage = tsScores) then
  begin
    SaveDialogJPEG.FileName := 'VariableLoadings_4.jpg';
  end;
  if (pc1.ActivePage = tsLocalities) then
  begin
    SaveDialogJPEG.FileName := 'Localities.jpg';
  end;
  if (pc1.ActivePage = ts3D) then
  begin
    SaveDialogJPEG.FileName := 'Scores3D_1_vs_2_vs_3.jpg';
  end;
  if (pc1.ActivePage = tsLoc3D) then
  begin
    SaveDialogJPEG.FileName := 'Localities3D.jpg';
  end;
  if (pc1.ActivePage = tsLoc4D) then
  begin
    SaveDialogJPEG.FileName := 'Localities4D.jpg';
  end;

  if SaveDialogJPEG.Execute then
  begin
    JPEGPath := ExtractFilePath(SaveDialogJPEG.FileName);
    if (pc1.ActivePage = tsGraph1) then
    begin
      //TeeSaveToJPEGFile(DBChart1,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart1.Width,DBChart1.Height);
    end;
    if (pc1.ActivePage = tsGraph2) then
    begin
      //TeeSaveToJPEGFile(DBChart2,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart2.Width,DBChart2.Height);
    end;
    if (pc1.ActivePage = tsGraph3) then
    begin
      //TeeSaveToJPEGFile(DBChart3,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart3.Width,DBChart3.Height);
    end;
    if (pc1.ActivePage = tsEigenValues) then
    begin
      //TeeSaveToJPEGFile(ChartEigenValue,SaveDialogJPEG.FileName,False,jpBestQuality,100,ChartEigenValue.Width,ChartEigenValue.Height);
    end;
    if (pc1.ActivePage = tsScores) then
    begin
      //TeeSaveToJPEGFile(DBChart10,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart10.Width,DBChart10.Height);
    end;
    if (pc1.ActivePage = tsLocalities) then
    begin
      //TeeSaveToJPEGFile(DBChart4,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart4.Width,DBChart4.Height);
    end;
    if (pc1.ActivePage = ts3D) then
    begin
      //TeeSaveToJPEGFile(DBChart8,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart8.Width,DBChart8.Height);
    end;
    if (pc1.ActivePage = tsLoc3D) then
    begin
      //TeeSaveToJPEGFile(DBChart9,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart9.Width,DBChart9.Height);
    end;
    if (pc1.ActivePage = tsLoc4D) then
    begin
      //TeeSaveToJPEGFile(DBChart11,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart11.Width,DBChart11.Height);
    end;
  end;
end;

procedure TfmCoranMain.FillSimilaritySpreadSheet;
var
  i, j : integer;
  {
  A1t : RealArrayC;
  }
begin
  {clear contents of spreadsheet}
  if (scSimilarityChoice in [scCorrespondence,scRQVariance,scRQPearsonR,
                   scPCAVariance,scPCAPearsonR,scPCASpearmanR,
                   scPCAKendallR,scRQSpearmanR,scRQKendallR]) then
  begin
    dmCor.CoranSimilarity.First;
    for i := 1 to Nox do
    begin
      for j := 1 to Nox do
      begin
        TempC[i,j] := dmCor.CoranSimilarity.Fields[j+1].AsFloat;  // was A1t
      end;
      dmCor.CoranSimilarity.Next;
    end;
    dmCor.CoranSimilarity.First;
    if (scSimilarityChoice=scCorrespondence) then
    begin
      for i := 1 to Nox do TempC[i,1]:=0.0;
    end;
    {fill similarity matrix values in spreadsheet}
    dmCor.CoranSimilarity.First;
  end;
end;

procedure TfmCoranMain.sbShow4DClick(Sender: TObject);
var
  j : integer;
  tmpStr : string;
  i : integer;
begin
  DBChart11.BottomAxis.Title.Caption := cbX4D.Text;
  DBChart11.LeftAxis.Title.Caption := cbY4D.Text;
  DBChart11.DepthAxis.Title.Caption := cbZ4D.Text;

  DBChart11.BottomAxis.Automatic := true;
  DBChart11.LeftAxis.Automatic := true;
  DBChart11.DepthAxis.Automatic := true;

  {
  try
    BottomAxisMin := 0.0;
    BottomAxisMax := 0.0;
    try
      GetMinMax('X',BottomAxisMin,BottomAxisMax);
      DBChart11.BottomAxis.SetMinMax(BottomAxisMin,BottomAxisMax);
    except
    end;
    //ShowMessage('X'+'  '+FormatFloat('####0.00',BottomAxisMin)+'  '+FormatFloat('####0.00',BottomAxisMax));
    DBChart11.BottomAxis.Automatic := false;
  except
    DBChart11.BottomAxis.Automatic := true;
  end;
  try
    LeftAxisMin := 0.0;
    LeftAxisMax := 0.0;
    try
      GetMinMax('Y',LeftAxisMin,LeftAxisMax);
      DBChart11.LeftAxis.SetMinMax(LeftAxisMin,LeftAxisMax);
    except
    end;
    //ShowMessage('Y'+'  '+FormatFloat('####0.00',LeftAxisMin)+'  '+FormatFloat('####0.00',LeftAxisMax));
    DBChart11.LeftAxis.Automatic := false;
  except
    DBChart11.LeftAxis.Automatic := true;
  end;
  try
    DepthTopAxisMin := 0.0;
    DepthTopAxisMax := 0.0;
    try
      GetMinMax('Z',DepthTopAxisMin,DepthTopAxisMax);
      DBChart11.DepthAxis.SetMinMax(DepthTopAxisMin,DepthTopAxisMax);
    except
    end;
    //ShowMessage('Z'+'  '+FormatFloat('####0.00',DepthTopAxisMin)+'  '+FormatFloat('####0.00',DepthTopAxisMax));
    DBChart11.DepthAxis.Automatic := false;
  except
    DBChart11.DepthAxis.Automatic := true;
  end;
  }

  DBChart11.Series[6].Active := false;
  DBChart11.Series[6].DataSource := dmCor.QGroupedSmpLoc;
  (DBChart11.Series[6] as TPoint3DSeries).XValues.ValueSource := cbX4D.Text;
  (DBChart11.Series[6] as TPoint3DSeries).YValues.ValueSource := cbY4D.Text;
  (DBChart11.Series[6] as TPoint3DSeries).ZValues.ValueSource := cbZ4D.Text;
  DBChart11.Series[6].Active := true;

  for j := 0 to 5 do
  begin
    DBChart11.Series[j].Active := false;
  end;
  for j := 0 to 5 do
  begin
    case j of
      0 : DBChart11.Series[j].DataSource := dmCor.qDim4Smp1;
      1 : DBChart11.Series[j].DataSource := dmCor.qDim4Smp2;
      2 : DBChart11.Series[j].DataSource := dmCor.qDim4Smp3;
      3 : DBChart11.Series[j].DataSource := dmCor.qDim4Smp4;
      4 : DBChart11.Series[j].DataSource := dmCor.qDim4Smp5;
      5 : DBChart11.Series[j].DataSource := dmCor.qDim4Smp6;
    end;
    (DBChart11.Series[j] as TPoint3DSeries).XValues.Order := loNone;
    (DBChart11.Series[j] as TPoint3DSeries).YValues.Order := loNone;
    (DBChart11.Series[j] as TPoint3DSeries).ZValues.Order := loNone;
    (DBChart11.Series[j] as TPoint3DSeries).XValues.ValueSource := cbX4D.Text;
    (DBChart11.Series[j] as TPoint3DSeries).YValues.ValueSource := cbY4D.Text;
    (DBChart11.Series[j] as TPoint3DSeries).ZValues.ValueSource := cbZ4D.Text;
  end;

  AssignChartDataSources('Close');
  dmCor.qDim4Smp1.Parameters.ParamByName('VariableID').Value := cbVariableID.Text;
  dmCor.qDim4Smp2.Parameters.ParamByName('VariableID').Value := cbVariableID.Text;
  dmCor.qDim4Smp3.Parameters.ParamByName('VariableID').Value := cbVariableID.Text;
  dmCor.qDim4Smp4.Parameters.ParamByName('VariableID').Value := cbVariableID.Text;
  dmCor.qDim4Smp5.Parameters.ParamByName('VariableID').Value := cbVariableID.Text;
  dmCor.qDim4Smp6.Parameters.ParamByName('VariableID').Value := cbVariableID.Text;
  sbMain.Panels[1].Text := 'Parameter values assigned';
  sbMain.Refresh;
  AssignChartDataSources('Open');
  for j := 0 to 5 do
  begin
    DBChart11.Series[j].Active := true;
  end;
  if (Pos('Vector',cbVariableID.Text) > 0) then
  begin
    DBChart14.Visible := true;
    dmCor.FacLoadingsVarNoZero.Close;
    i := Pos('r',cbVariableID.Text)+1;
    tmpStr := 'Vector'+Copy(cbVariableID.Text,i,length(cbVariableID.Text)-i+1);
    DBChart14.Series[0].Active := false;
    DBChart14.Series[0].XValues.Order := loNone;
    DBChart14.Series[0].YValues.Order := loNone;
    DBChart14.Series[0].DataSource := dmCor.FacLoadingsVarNoZero;
    DBChart14.Series[0].YValues.ValueSource := tmpStr;
    DBChart14.Series[0].XValues.ValueSource := 'Pos';
    dmCor.FacLoadingsVarNoZero.Open;
    DBChart14.Series[0].Active := true;
  end else
  begin
    DBChart14.Series[0].Active := false;
    DBChart14.Visible := false;
  end;
end;

procedure TfmCoranMain.CalculateHingeStatistics(NumAttributes : integer);
var
  i, j : integer;
  n : integer;
  k : integer;
  t : double;
  XMed, HingeL, HingeU,
  WhiskerL, WhiskerU : double;
  TriMean : double;
  ml : integer;
  {
  Loc : RealArrayM;
  Mah : RealArrayR;
  Atmp : RealArrayC;
  }
begin
  sbMain.Panels[1].Text := 'Calculating 4D attributes';
  sbMain.Refresh;
  with dmCor do
  begin
    DeleteDim4Smp.SQL.Clear;
    DeleteDim4Smp.SQL.Add('delete * from Dimension4Smp');
    DeleteDim4Smp.SQL.Add('where Dimension4Smp.VariableID like '+''''+'Vector%'+'''');
    DeleteDim4Smp.ExecSQL;
    DeleteDim4Smp.SQL.Clear;
    DeleteDim4Smp.SQL.Add('delete * from Dimension4Smp');
    DeleteDim4Smp.SQL.Add('where Dimension4Smp.VariableID like '+''''+'Score%'+'''');
    DeleteDim4Smp.ExecSQL;
    cdsqDim4Smp.Close;
    cdsqDim4Smp.Open;
    sbMain.Panels[1].Text := 'Calculating 4D attributes - score distances and vectors';
    sbMain.Refresh;
    FillChar(DD,sizeof(DD),0);
    //square roots of eigenvalues in A3, eigenvectors in A2, scores in B
    //first, calculate score distances
    for j := 1 to Nox do
    begin
      TempC[j,j] := A3[j,j]*A3[j,j];
    end;
    Mahalanobis(B,DD,TempC,DR,NumSamples,Nox);
    for i := 1 to NumSamples do
    begin
      if (DR[i] > 0.0) then
        DR[i] := sqrt(DR[i])
      else
        DR[i] := 0.0;
    end;
    dmCor.cdsFacLoadingsSmp.Open;
    dmCor.cdsqDim4Smp.Open;
    dmCor.cdsFacLoadingsSmp.First;
    k := 0;
    repeat
      k := k + 1;
      try
        cdsqDim4Smp.Append;
        dmCor.cdsqDim4SmpGROUPNAME.AsString := dmCor.cdsFacLoadingsSmpGROUPNAME.AsString;
        dmCor.cdsqDim4SmpPlotGroupName.AsString := dmCor.cdsFacLoadingsSmpPlotGroupName.AsString;
        dmCor.cdsqDim4SmpSAMPLENUM.AsString := dmCor.cdsFacLoadingsSmpSAMPLENUM.AsString;
        dmCor.cdsqDim4SmpVariableID.AsString := 'ScoreDistance';
        //create dummy record for this sample and the score distance
        dmCor.cdsqDim4SmpClassificationID.AsString := '0';
        dmCor.cdsqDim4SmpSmpValue.AsFloat := DR[k];
        dmCor.cdsqDim4Smp.Post;
      except
      end;
      for j := 1 to NumAttributes do
      begin
        try
          dmCor.cdsqDim4Smp.Append;
          dmCor.cdsqDim4SmpGROUPNAME.AsString := dmCor.cdsFacLoadingsSmpGROUPNAME.AsString;
          dmCor.cdsqDim4SmpPlotGroupName.AsString := dmCor.cdsFacLoadingsSmpPlotGroupName.AsString;
          dmCor.cdsqDim4SmpSAMPLENUM.AsString := dmCor.cdsFacLoadingsSmpSAMPLENUM.AsString;
          dmCor.cdsqDim4SmpVariableID.AsString := 'Vector' + IntToStr(j);
          //create dummy record for sample and this vector
          dmCor.cdsqDim4SmpClassificationID.AsString := '0';
          dmCor.cdsqDim4SmpSmpValue.AsFloat := dmCor.cdsFacLoadingsSmp.Fields[j+2].AsFloat;
          dmCor.cdsqDim4Smp.Post;
        except
        end;
      end;
      dmCor.cdsFacLoadingsSmp.Next;
    until dmCor.cdsFacLoadingsSmp.Eof;
    //dmCor.qDim4Smp.Close;
    dmCor.cdsCoranRaw.Open;
    //dmCor.qDim4Smp.Open;
    dmCor.cdsCoranRaw.First;
    dmCor.ElemNames.Open;
    dmCor.ElemNames.First;
    //dmCor.qDim4Smp.Close;
    dmCor.cdsqDim4Smp.Filtered := false;
    dmCor.cdsqDim4Smp.Filter := 'VariableID = '+''''+'ScoreDistance'+'''';
    dmCor.cdsqDim4Smp.Filtered := true;
    //dmCor.qDim4Smp.Open;
    sbMain.Panels[1].Text := 'Calculating 4D attributes - score distances';
    sbMain.Refresh;
    dmCor.cdsqDim4Smp.First;
    //calculate median
    i := 1;
    repeat
      DR[i] :=  dmCor.cdsqDim4SmpSmpValue.AsFloat;
      i := i + 1;
      dmCor.cdsqDim4Smp.Next;
    until dmCor.cdsqDim4Smp.Eof;
    //ShowMessage(IntToStr(NumSamples)+'**'+IntToStr(i-1));
    //n := NumSamples;
    n := i-1;
    Sort(DR,n,1);
    ml := (n+1) div 2;
    //ShowMessage(IntToStr(ml)+'**'+IntToStr(n-ml+1));
    XMed := 0.5*(DR[ml] + DR[n-ml+1]);
    //calculate lower and upper hinge
    Hinges(DR,HingeL,HingeU,NumSamples);
    TriMean := 0.5*XMed + 0.25*HingeL + 0.25*HingeU;
    Whiskers(DR,XMed,WhiskerL,WhiskerU,n);
    dmCor.cdsqDim4Smp.First;
    repeat
      t := dmCor.cdsqDim4SmpSmpValue.AsFloat;
      dmCor.cdsqDim4Smp.Edit;
      if (t >= XMed) then
      begin
        if (t >= HingeU) then
        begin
          if (t > WhiskerU) then
          begin
            dmCor.cdsqDim4SmpClassificationID.AsString := '1UE';
          end else
          begin
            dmCor.cdsqDim4SmpClassificationID.AsString := '2UW';
          end;
        end else
        begin
          dmCor.cdsqDim4SmpClassificationID.AsString := '3UB';
        end;
      end else
      begin
        if (t >= WhiskerL) then
        begin
          if (t >= HingeL) then
          begin
            dmCor.cdsqDim4SmpClassificationID.AsString := '4LB';
          end else
          begin
            dmCor.cdsqDim4SmpClassificationID.AsString := '5LW';
          end;
        end else
        begin
          dmCor.cdsqDim4SmpClassificationID.AsString := '6LE';
        end;
      end;
      dmCor.cdsqDim4Smp.Next;
    until dmCor.cdsqDim4Smp.Eof;
    sbMain.Panels[1].Text := 'Calculating 4D attributes - vectors';
    sbMain.Refresh;
    for j := 1 to NumAttributes do
    begin
      sbMain.Panels[0].Text := IntToStr(j);
      sbMain.Refresh;
      dmCor.cdsqDim4Smp.Filtered := false;
      dmCor.cdsqDim4Smp.First;
      dmCor.cdsqDim4Smp.Filter := 'VariableID = '+''''+'Vector'+IntToStr(j)+'''';
      dmCor.cdsqDim4Smp.Filtered := true;
      //calculate median
      i := 1;
      repeat
        DR[i] :=  dmCor.cdsqDim4SmpSmpValue.AsFloat;
        i := i + 1;
        dmCor.cdsqDim4Smp.Next;
      until dmCor.cdsqDim4Smp.Eof;
      //n:= NumSamples;
      n := i-1;
      Sort(DR,n,1);
      ml := (n+1) div 2;
      XMed := 0.5*(DR[ml] + DR[n-ml+1]);
      //calculate lower and upper hinge
      Hinges(DR,HingeL,HingeU,NumSamples);
      TriMean := 0.5*XMed + 0.25*HingeL + 0.25*HingeU;
      Hinges(DR,WhiskerL,WhiskerU,n);
      Whiskers(DR,XMed,WhiskerL,WhiskerU,n);
      dmCor.cdsqDim4Smp.First;
      repeat
        t := dmCor.cdsqDim4SmpSmpValue.AsFloat;
        dmCor.cdsqDim4Smp.Edit;
        if (t >= XMed) then
        begin
          if (t >= HingeU) then
          begin
            if (t >= WhiskerU) then
            begin
              dmCor.cdsqDim4SmpClassificationID.AsString := '1UE';
            end else
            begin
              dmCor.cdsqDim4SmpClassificationID.AsString := '2UW';
            end;
          end else
          begin
            dmCor.cdsqDim4SmpClassificationID.AsString := '3UB';
          end;
        end else
        begin
          if (t >= WhiskerL) then
          begin
            if (t >= HingeL) then
            begin
              dmCor.cdsqDim4SmpClassificationID.AsString := '4LB';
            end else
            begin
              dmCor.cdsqDim4SmpClassificationID.AsString := '5LW';
            end;
          end else
          begin
            dmCor.cdsqDim4SmpClassificationID.AsString := '6LE';
          end;
        end;
        dmCor.cdsqDim4Smp.Next;
      until dmCor.cdsqDim4Smp.Eof;
    end;
    sbMain.Panels[0].Text := '';
    sbMain.Panels[1].Text := 'Saving 4D component data to database';
    sbMain.Refresh;
    dmCor.cdsqDim4Smp.Filtered := false;
    dmCor.cdsqDim4Smp.First;
    dmCor.cdsqDim4Smp.ApplyUpdates(-1);
    sbMain.Panels[0].Text := '';
    sbMain.Panels[1].Text := '';
    sbMain.Refresh;
  end;
end;

procedure TfmCoranMain.CalculateHingeStatisticsForOriginalvariables(NumAttributes : integer);
var
  i, j : integer;
  k : integer;
  t : double;
  XMed, HingeL, HingeU,
  WhiskerL, WhiskerU : double;
  TriMean : double;
  ml : integer;
begin
  sbMain.Panels[1].Text := 'Calculating 4D attributes';
  sbMain.Refresh;
  with dmCor do
  begin
    //dmCor.cdsqDim4Smp.DisableControls;
    //dmCor.cdsCoranRaw.DisableControls;
    //dmCor.ElemNames.DisableControls;
    dmCor.cdsqDim4Smp.Open;
    dmCor.cdsCoranRaw.Open;
    //dmCor.qDim4Smp.Open;
    dmCor.cdsCoranRaw.First;
    dmCor.ElemNames.Open;
    dmCor.ElemNames.First;
    if Include4DVarData then
    begin
      sbMain.Panels[1].Text := 'Preparing 4D attributes - original variables';
      sbMain.Refresh;
      k := 0;
      repeat
        k := k + 1;
        dmCor.ElemNames.First;
        for j := 1 to NumAttributes do
        begin
          //sbMain.Panels[0].Text := '  '+dmCor.ElemNamesCalled.AsString;
          //sbMain.Refresh;
          try
            dmCor.cdsqDim4Smp.Append;
            dmCor.cdsqDim4SmpGROUPNAME.AsString := dmCor.cdsCoranRawGROUPNAME.AsString;
            dmCor.cdsqDim4SmpPlotGroupName.AsString := dmCor.cdsCoranRawPlotGroupName.AsString;
            dmCor.cdsqDim4SmpSAMPLENUM.AsString := dmCor.cdsCoranRawSAMPLENUM.AsString;
            dmCor.cdsqDim4SmpVariableID.AsString := dmCor.ElemNamesCalled.AsString;
            //create dummy record for sample and this vector
            dmCor.cdsqDim4SmpClassificationID.AsString := 'O';
            dmCor.cdsqDim4SmpSmpValue.AsFloat := dmCor.cdsCoranRaw.Fields[j+2].AsFloat;
            dmCor.cdsqDim4Smp.Post;
          except
          end;
          dmCor.ElemNames.Next;
        end;
        dmCor.cdsCoranRaw.Next;
      until dmCor.cdsCoranRaw.Eof;
    end;
    //dmCor.cdsqDim4Smp.ApplyUpdates(-1);
    //dmCor.qDim4Smp.Close;
    dmCor.cdsqDim4Smp.Filtered := false;
    if Include4DVarData then
    begin
      sbMain.Panels[1].Text := 'Calculating 4D attributes - original variables';
      sbMain.Refresh;
      FillChar(DD,sizeof(DD),0);
      dmCor.cdsCoranRaw.Open;
      //dmCor.qDim4Smp.Open;
      dmCor.cdsCoranRaw.First;
      dmCor.ElemNames.Open;
      dmCor.ElemNames.First;
      for j := 1 to NumAttributes do
      begin
        sbMain.Panels[0].Text := '  '+dmCor.ElemNamesCalled.AsString;
        sbMain.Refresh;
        dmCor.cdsqDim4Smp.Filtered := false;
        dmCor.cdsqDim4Smp.First;
        dmCor.cdsqDim4Smp.Filter := 'VariableID = '+''''+dmCor.ElemNamesCalled.AsString+'''';
        dmCor.cdsqDim4Smp.Filtered := true;
        //calculate median
        i := 1;
        repeat
          DR[i] :=  dmCor.cdsqDim4SmpSmpValue.AsFloat;
          //if (i > 47) then ShowMessage('median '+IntToStr(i)+'  '+FormatFloat('###0.0000',dmCor.cdsqDim4SmpSmpValue.AsFloat));
          i := i + 1;
          dmCor.cdsqDim4Smp.Next;
        until dmCor.cdsqDim4Smp.Eof;
        //ShowMessage('before Sort '+IntToStr(NumSamples));
        Sort(DR,NumSamples,1);
        ml := (NumSamples+1) div 2;
        XMed := 0.5*(DR[ml] + DR[NumSamples-ml+1]);
        //calculate lower and upper hinge
        Hinges(DR,HingeL,HingeU,NumSamples);
        TriMean := 0.5*XMed + 0.25*HingeL + 0.25*HingeU;
        Hinges(DR,WhiskerL,WhiskerU,NumSamples);
        Whiskers(DR,XMed,WhiskerL,WhiskerU,NumSamples);
        dmCor.cdsqDim4Smp.First;
        repeat
          t := dmCor.cdsqDim4SmpSmpValue.AsFloat;
          dmCor.cdsqDim4Smp.Edit;
          if (t >= XMed) then
          begin
            if (t >= HingeU) then
            begin
              if (t >= WhiskerU) then
              begin
                dmCor.cdsqDim4SmpClassificationID.AsString := '1UE';
              end else
              begin
                dmCor.cdsqDim4SmpClassificationID.AsString := '2UW';
              end;
            end else
            begin
              dmCor.cdsqDim4SmpClassificationID.AsString := '3UB';
            end;
          end else
          begin
            if (t >= WhiskerL) then
            begin
              if (t >= HingeL) then
              begin
                dmCor.cdsqDim4SmpClassificationID.AsString := '4LB';
              end else
              begin
                dmCor.cdsqDim4SmpClassificationID.AsString := '5LW';
              end;
            end else
            begin
              dmCor.cdsqDim4SmpClassificationID.AsString := '6LE';
            end;
          end;
          dmCor.cdsqDim4Smp.Next;
        until dmCor.cdsqDim4Smp.Eof;
        dmCor.ElemNames.Next;
      end;
    end;
    sbMain.Panels[0].Text := '';
    sbMain.Panels[1].Text := 'Saving 4D variable data to database';
    sbMain.Refresh;
    dmCor.cdsqDim4Smp.Filtered := false;
    dmCor.cdsqDim4Smp.First;
    dmCor.cdsqDim4Smp.ApplyUpdates(-1);
    dmCor.cdsqDim4Smp.EnableControls;
    dmCor.cdsCoranRaw.EnableControls;
    dmCor.ElemNames.EnableControls;
    sbMain.Panels[0].Text := '';
    sbMain.Panels[1].Text := '';
    sbMain.Refresh;
  end;
end;


procedure TfmCoranMain.Discretise(NumAttributes : integer);
var
  i, j : integer;
  k : integer;
  t, t1 : double;
  XMed, HingeL, HingeU,
  WhiskerL, WhiskerU : double;
  NumSamples : integer;
begin
  sbMain.Panels[1].Text := 'Discretising transformed variables using boxplot boundaries';
  sbMain.Refresh;
  with dmCor do
  begin
    DeleteDim4.ExecSQL;
    qDim4.Open;
    //dmCor.qCoranChem.DisableControls;
    //dmCor.ElemNames.DisableControls;
    //dmCor.CoranStats.DisableControls;
    dmCor.qCoranChem.Open;
    NumSamples := dmCor.qCoranChem.RecordCount;
    dmCor.qCoranChem.First;
    dmCor.ElemNames.Open;
    dmCor.ElemNames.First;
    dmCor.CoranStats.Open;
    dmCor.CoranStats.First;
    for i := 1 to NumSamples do
    begin
      dmCor.qDim4.Append;
      dmCor.qDim4GROUPNAME.AsString := dmCor.qCoranChemGROUPNAME.AsString;
      dmCor.qDim4PlotGroupName.AsString := dmCor.qCoranChemPlotGroupName.AsString;
      dmCor.qDim4SAMPLENUM.AsString := dmCor.qCoranChemSAMPLENUM.AsString;
      dmCor.qDim4Sequence.AsInteger := dmCor.qCoranChemSequence.AsInteger;
      for j := 1 to MMaxFields do
      begin
        dmCor.qDim4.Fields[j+2].AsFloat := 0.0;
      end;
      dmCor.qDim4.Post;
      dmCor.qCoranChem.Next;
    end;
    dmCor.qCoranChem.First;
    for j := 1 to NumAttributes do
    begin
      dmCor.qDim4.First;
      sbMain.Panels[0].Text := '  '+dmCor.ElemNamesCalled.AsString;
      sbMain.Refresh;
      dmCor.CoranStats.Locate('Summary','Upper fence',[]);
      WhiskerU := dmCor.CoranStats.Fields[j+1].AsVariant;
      dmCor.CoranStats.Locate('Summary','Upper hinge',[]);
      HingeU := dmCor.CoranStats.Fields[j+1].AsVariant;
      dmCor.CoranStats.Locate('Summary','Median',[]);
      XMed := dmCor.CoranStats.Fields[j+1].AsVariant;
      dmCor.CoranStats.Locate('Summary','Lower hinge',[]);
      HingeL := dmCor.CoranStats.Fields[j+1].AsVariant;
      dmCor.CoranStats.Locate('Summary','Lower fence',[]);
      WhiskerL := dmCor.CoranStats.Fields[j+1].AsVariant;
      dmCor.qCoranChem.First;
      for i := 1 to NumSamples do
      begin
        {
        sbMain.Panels[2].Text := IntToStr(i);
        sbMain.Refresh;
        }
        t := dmCor.qCoranChem.Fields[j+2].AsFloat;
        {
        if (j in [10,11,12]) then
        begin
          if (i < 5) then ShowMessage(IntToStr(j)+'  '+IntToStr(i)+'  '+FormatFloat('###0.0000',t));
        end;
        }
        if (t >= XMed) then
        begin
          if (t >= HingeU) then
          begin
            if (t >= WhiskerU) then
            begin
              t1 := 6.0;
            end else
            begin
              t1 := 5.0;
            end;
          end else
          begin
            t1 := 4.0;
          end;
        end else
        begin
          if (t >= WhiskerL) then
          begin
            if (t >= HingeL) then
            begin
              t1 := 3.0;
            end else
            begin
              t1 := 2.0;
            end;
          end else
          begin
            t1 := 1.0;
          end;
        end;
        try
          {
          if (j in [10,11,12]) then
          begin
            t := dmCor.qDim4.Fields[j+2].AsFloat;
            if (i < 5) then ShowMessage('Dim4 '+IntToStr(j)+'  '+IntToStr(i)+'  '+dmCor.qDim4SAMPLENUM.AsString+'  '+ FormatFloat('###0.0000',t));
          end;
          }
          dmCor.qDim4.Edit;
          {
          if (j in [10,11,12]) then
          begin
            if (i < 5) then ShowMessage('after edit');
          end;
          }
          dmCor.qDim4.Fields[j+2].AsFloat := t1;
          {
          if (j in [10,11,12]) then
          begin
            if (i < 5) then ShowMessage('before post');
          end;
          }
          dmCor.qDim4.Next;
        except
        end;
        {
          if (j in [10,11,12]) then
          begin
            if (i < 5) then ShowMessage('after post');
          end;
        }
        dmCor.qCoranChem.Next;
        {
          if (j in [10,11,12]) then
          begin
            if (i < 5) then ShowMessage('after next 1');
          end;
        //dmCor.qDim4.Next;
          if (j in [10,11,12]) then
          begin
            if (i < 5) then ShowMessage('after next 2');
          end;
        }
      end;
      dmCor.ElemNames.Next;
    end;
    dmCor.qCoranChem.First;
    dmCor.qCoranChem.Close;
    dmCor.qDim4.Close;
    dmCor.ElemNames.First;
    dmCor.CoranStats.First;
    dmCor.qCoranChem.EnableControls;
    dmCor.ElemNames.EnableControls;
    dmCor.CoranStats.EnableControls;
    sbMain.Panels[0].Text := '';
    sbMain.Panels[1].Text := '';
    sbMain.Panels[2].Text := '';
    sbMain.Refresh;
  end;
end;

procedure TfmCoranMain.udRotation4DClick(Sender: TObject;
  Button: TUDBtnType);
begin
  DBChart11.View3DOptions.Rotation := udRotation4D.Position;
end;

procedure TfmCoranMain.udElevation4DClick(Sender: TObject;
  Button: TUDBtnType);
begin
  DBChart11.View3DOptions.Elevation := udElevation4D.Position;
end;

procedure TfmCoranMain.udPerspective4DClick(Sender: TObject;
  Button: TUDBtnType);
begin
  DBChart11.View3DOptions.Perspective := udPerspective4D.Position;
end;

procedure TfmCoranMain.udZoom4DClick(Sender: TObject; Button: TUDBtnType);
begin
  DBChart11.View3DOptions.Zoom := udZoom4D.Position;
end;

procedure TfmCoranMain.cbVariableIDChange(Sender: TObject);
begin
  sbShow4DClick(Sender);
end;

procedure TfmCoranMain.Showvariablelabelsingraphs1Click(Sender: TObject);
begin
  ShowVariableLabelsInGraphs1.Checked := not ShowVariableLabelsInGraphs1.Checked;
  if ShowVariableLabelsInGraphs1.Checked then
  begin
    DBChart1.Series[0].Marks.Visible := true;
    DBChart2.Series[0].Marks.Visible := true;
    DBChart3.Series[0].Marks.Visible := true;
    (DBChart8.Series[0] as TPoint3DSeries).Marks.Visible := true;
  end else
  begin
    DBChart1.Series[0].Marks.Visible := false;
    DBChart2.Series[0].Marks.Visible := false;
    DBChart3.Series[0].Marks.Visible := false;
    (DBChart8.Series[0] as TPoint3DSeries).Marks.Visible := false;
  end;
end;

procedure TfmCoranMain.Defaultvalues1Click(Sender: TObject);
begin
  try
    DefaultForm := TfmDefaults.Create(Self);
    DefaultForm.ShowModal;
    SetSymbolDefaults;
  finally
    DefaultForm.Free;
  end;
end;

procedure TfmCoranMain.SetSymbolDefaults;
var
  tSizeVariables, tSizeSamples,
  tSizeGroup, tSizeSample,
  tSizeVariable, tSizeUpperExtreme,
  tSizeUpperWhisker, tSizeUpperBox,
  tSizeLowerBox, tSizeLowerWhisker,
  tSizeLowerExtreme : integer;
begin
  dmCor.qSymbolDefaults.Open;
  dmCor.qSymbolDefaults.Locate('SymbolName','Variables',[]);
  tSizeVariables := dmCor.qSymbolDefaultsSymbolSize.AsInteger;
  dmCor.qSymbolDefaults.Locate('SymbolName','samples',[]);
  tSizeSamples := dmCor.qSymbolDefaultsSymbolSize.AsInteger;
  dmCor.qSymbolDefaults.Locate('SymbolName','Group',[]);
  tSizeGroup := dmCor.qSymbolDefaultsSymbolSize.AsInteger;
  dmCor.qSymbolDefaults.Locate('SymbolName','Sample',[]);
  tSizeSample := dmCor.qSymbolDefaultsSymbolSize.AsInteger;
  dmCor.qSymbolDefaults.Locate('SymbolName','Variable',[]);
  tSizeVariable := dmCor.qSymbolDefaultsSymbolSize.AsInteger;
  dmCor.qSymbolDefaults.Locate('SymbolName','Upper extreme',[]);
  tSizeUpperExtreme := dmCor.qSymbolDefaultsSymbolSize.AsInteger;
  dmCor.qSymbolDefaults.Locate('SymbolName','Upper whisker',[]);
  tSizeUpperWhisker := dmCor.qSymbolDefaultsSymbolSize.AsInteger;
  dmCor.qSymbolDefaults.Locate('SymbolName','Upper box',[]);
  tSizeUpperBox := dmCor.qSymbolDefaultsSymbolSize.AsInteger;
  dmCor.qSymbolDefaults.Locate('SymbolName','Lower box',[]);
  tSizeLowerBox := dmCor.qSymbolDefaultsSymbolSize.AsInteger;
  dmCor.qSymbolDefaults.Locate('SymbolName','Lower whisker',[]);
  tSizeLowerWhisker := dmCor.qSymbolDefaultsSymbolSize.AsInteger;
  dmCor.qSymbolDefaults.Locate('SymbolName','Lower extreme',[]);
  tSizeLowerExtreme := dmCor.qSymbolDefaultsSymbolSize.AsInteger;
  SeriesG10.Pointer.Size:= tSizeVariables;
  SeriesG11.Pointer.Size:= tSizeSamples;
  SeriesG12.Pointer.Size:= tSizeGroup;
  SeriesG13.Pointer.Size:= tSizeSample;
  SeriesG14.Pointer.Size:= tSizeVariable;
  SeriesG20.Pointer.Size:= tSizeVariables;
  SeriesG21.Pointer.Size:= tSizeSamples;
  SeriesG22.Pointer.Size:= tSizeGroup;
  SeriesG23.Pointer.Size:= tSizeSample;
  SeriesG24.Pointer.Size:= tSizeVariable;
  SeriesG30.Pointer.Size:= tSizeVariables;
  SeriesG31.Pointer.Size:= tSizeSamples;
  SeriesG32.Pointer.Size:= tSizeGroup;
  SeriesG33.Pointer.Size:= tSizeSample;
  SeriesG34.Pointer.Size:= tSizeVariable;
  SeriesG41.Pointer.Size:= tSizeSamples;
  SeriesG42.Pointer.Size:= tSizeGroup;
  SeriesG43.Pointer.Size:= tSizeSample;
  SeriesG80.Pointer.Size:= tSizeVariables;
  SeriesG81.Pointer.Size:= tSizeSamples;
  SeriesG82.Pointer.Size:= tSizeGroup;
  SeriesG83.Pointer.Size:= tSizeSample;
  SeriesG84.Pointer.Size:= tSizeVariable;
  SeriesG91.Pointer.Size:= tSizeSamples;
  SeriesG92.Pointer.Size:= tSizeGroup;
  SeriesG93.Pointer.Size:= tSizeSample;
  SeriesG1103.Pointer.Size:= tSizeSample;
  SeriesG1105.Pointer.Size:= tSizeUpperExtreme;
  SeriesG1106.Pointer.Size:= tSizeUpperWhisker;
  SeriesG1107.Pointer.Size:= tSizeUpperBox;
  SeriesG1108.Pointer.Size:= tSizeLowerBox;
  SeriesG1109.Pointer.Size:= tSizeLowerWhisker;
  SeriesG1110.Pointer.Size:= tSizeLowerExtreme;
  SeriesG1201.Pointer.Size:= tSizeSamples;
  SeriesG1202.Pointer.Size:= tSizeGroup;
  SeriesG1203.Pointer.Size:= tSizeSample;
end;

procedure TfmCoranMain.AssignChartDataSources(OpenClose : string);
begin
  if (UpperCase(OpenClose) = 'OPEN') then
  begin
    try
      dmCor.cdsCoranChem.Open;
    except
      MessageDlg('Could not open CoranChem',mtWarning,[mbOK],0);
    end;
    try
      dmCor.ElemNames.Open;
    except
      MessageDlg('Could not open ElemNames',mtWarning,[mbOK],0);
    end;
    try
      dmCor.cdsCoranRaw.Open;
    except
      MessageDlg('Could not open CoranRaw',mtWarning,[mbOK],0);
    end;
    try
      dmCor.QHist.Open;
    except
      MessageDlg('Could not open QHist',mtWarning,[mbOK],0);
    end;
    try
      dmCor.CoranFac.Open;
    except
      MessageDlg('Could not open CoranFac',mtWarning,[mbOK],0);
    end;
    try
      dmCor.SmpLoc.Open;
    except
      MessageDlg('Could not open SmpLoc',mtWarning,[mbOK],0);
    end;
    try
      dmCor.cdsFacLoadingsSmp.Open;
    except
      //MessageDlg('Could not open FacLoadingsSmp',mtWarning,[mbOK],0);
    end;
    try
      dmCor.QGroups.Open;
    except
      MessageDlg('Could not open QGroups',mtWarning,[mbOK],0);
    end;
    try
      dmCor.QPlotGroups.Open;
    except
      MessageDlg('Could not open QPlotGroups',mtWarning,[mbOK],0);
    end;
    try
      dmCor.CoranVecLinked.Open;
    except
      MessageDlg('Could not open CoranVecLinked',mtWarning,[mbOK],0);
    end;
    try
      dmCor.FacLoadingsVar.Open;
    except
      MessageDlg('Could not open FacLoadingsVarLinked',mtWarning,[mbOK],0);
    end;
    try
      dmCor.FacLoadingsVarLinked.Open;
    except
      MessageDlg('Could not open FacLoadingsVarLinked',mtWarning,[mbOK],0);
    end;
    try
      //dmCor.FacLoadingsVarLinked2.Open;
    except
      MessageDlg('Could not open FacLoadingsVarLinked2',mtWarning,[mbOK],0);
    end;
    try
      dmCor.FacLoadingsVarNoZero.Open;
    except
      MessageDlg('Could not open FacLoadingsVarNoZero',mtWarning,[mbOK],0);
    end;
    try
      dmCor.GroupedSmp.Open;
    except
      MessageDlg('Could not open GroupedSmp',mtWarning,[mbOK],0);
    end;
    try
      dmCor.GroupedSmpLoc.Open;
    except
      MessageDlg('Could not open GroupedSmpLoc',mtWarning,[mbOK],0);
    end;
    try
      dmCor.QGroupedSmpLoc.Open;
    except
      MessageDlg('Could not open QGroupedSmpLoc',mtWarning,[mbOK],0);
    end;
    try
      dmCor.QGroupedSmp.Open;
    except
      MessageDlg('Could not open QGroupedSmp',mtWarning,[mbOK],0);
    end;
    try
      dmCor.cdsCoranQuantile.Open;
    except
      MessageDlg('Could not open CoranQuantile',mtWarning,[mbOK],0);
    end;
    try
      dmCor.qSDOD.Open;
    except
      MessageDlg('Could not open qSDOD',mtWarning,[mbOK],0);
    end;
    try
      dmCor.qGroupedSDOD.Open;
    except
      MessageDlg('Could not open qGroupedSDOD',mtWarning,[mbOK],0);
    end;
    try
      dmCor.qGroupedSmpSDOD.Open;
    except
      MessageDlg('Could not open qGroupedSmpSDOD',mtWarning,[mbOK],0);
    end;
    try
      dmCor.GroupedChem.Open;
    except
      MessageDlg('Could not open GroupedChem',mtWarning,[mbOK],0);
    end;
    try
      dmCor.qDim4Smp1.Open;
    except
      MessageDlg('Could not open qDim4Smp1',mtWarning,[mbOK],0);
    end;
    try
      dmCor.qDim4Smp2.Open;
    except
      MessageDlg('Could not open qDim4Smp2',mtWarning,[mbOK],0);
    end;
    try
      dmCor.qDim4Smp3.Open;
    except
      MessageDlg('Could not open qDim4Smp3',mtWarning,[mbOK],0);
    end;
    try
      dmCor.qDim4Smp4.Open;
    except
      MessageDlg('Could not open qDim4Smp4',mtWarning,[mbOK],0);
    end;
    try
      dmCor.qDim4Smp5.Open;
    except
      MessageDlg('Could not open qDim4Smp5',mtWarning,[mbOK],0);
    end;
    try
      dmCor.qDim4Smp6.Open;
    except
      MessageDlg('Could not open qDim4Smp6',mtWarning,[mbOK],0);
    end;
  end else
  begin
    try
      dmCor.ElemNames.Close;
      dmCor.CoranFac.Close;
      dmCor.SmpLoc.Close;
      dmCor.cdsFacLoadingsSmp.Close;
      dmCor.QGroups.Close;
      dmCor.QPlotGroups.Close;
      dmCor.CoranVecLinked.Close;
      dmCor.FacLoadingsVar.Close;
      dmCor.FacLoadingsVarLinked.Close;
      //dmCor.FacLoadingsVarLinked2.Close;
      dmCor.FacLoadingsVarNoZero.Close;
      dmCor.GroupedSmp.Close;
      dmCor.GroupedSmpLoc.Close;
      dmCor.QGroupedSmpLoc.Close;
      dmCor.QGroupedSmp.Close;
      dmCor.qSDOD.Close;
      dmCor.qGroupedSDOD.Close;
      dmCor.qGroupedSmpSDOD.Close;
      dmCor.GroupedChem.Close;
      dmCor.cdsCoranQuantile.Close;
      dmCor.qDim4Smp1.Close;
      dmCor.qDim4Smp2.Close;
      dmCor.qDim4Smp3.Close;
      dmCor.qDim4Smp4.Close;
      dmCor.qDim4Smp5.Close;
      dmCor.qDim4Smp6.Close;
      //dmCor.Coran.Connected := false;
    except
    end;
  end;
end;

procedure TfmCoranMain.DisableEnableControls(DisableEnable : string);
begin
  if (DisableEnable = 'Disable') then
  begin
     {
    dmCor.cdsCoranChem.DisableControls;
    dmCor.CoranVecLinked.DisableControls;
    dmCor.GroupedSmp.DisableControls;
    dmCor.GroupedSmpLoc.DisableControls;
    dmCor.GroupedChem.DisableControls;
    dmCor.QGroups.DisableControls;
    dmCor.QPlotGroups.DisableControls;
    dmCor.QGroupedSmpLoc.DisableControls;
    dmCor.QGroupedSmp.DisableControls;
    dmCor.cdsFacLoadingsSmp.DisableControls;
    dmCor.FacLoadingsVar.DisableControls;
    dmCor.qSDOD.DisableControls;
    }
  end else
  begin
    dmCor.cdsCoranChem.EnableControls;
    dmCor.CoranVecLinked.EnableControls;
    dmCor.GroupedSmp.EnableControls;
    dmCor.GroupedSmpLoc.EnableControls;
    dmCor.GroupedChem.EnableControls;
    dmCor.QGroups.EnableControls;
    dmCor.QPlotGroups.EnableControls;
    dmCor.QGroupedSmpLoc.EnableControls;
    dmCor.QGroupedSmp.EnableControls;
    dmCor.cdsFacLoadingsSmp.EnableControls;
    dmCor.FacLoadingsVar.EnableControls;
    dmCor.qSDOD.EnableControls;
  end;
end;

procedure TfmCoranMain.GetMinMax(AxisChosen : string;
                             var AxisMin : extended;
                             var AxisMax : extended);
var
  tmp : double;
  tmpStr : string;
begin
  if (AxisChosen = 'X') then
  begin
    tmpStr := cbX4D.Text;
  end;
  if (AxisChosen = 'Y') then
  begin
    tmpStr := cbY4D.Text;
  end;
  if (AxisChosen = 'Z') then
  begin
    tmpStr := cbZ4D.Text;
  end;
  dmCor.qDim4SmpTst.Parameters.ParamByName('VariableID').Value := cbVariableID.Text;
  dmCor.qDim4SmpTst.Open;
  dmCor.qDim4SmpTst.First;
  tmp := dmCor.qDim4SmpTst.FieldValues[tmpStr];
  AxisMin := tmp;
  AxisMax := tmp;
  repeat
    tmp := dmCor.qDim4SmpTst.FieldValues[tmpStr];
    if (tmp < AxisMin) then AxisMin := tmp;
    if (tmp > AxisMax) then AxisMax := tmp;
    dmCor.qDim4SmpTst.Next;
  until dmCor.qDim4SmpTst.Eof;
  dmCor.qDim4SmpTst.Close;
  if (AxisMin = AxisMax) then AxisMax := AxisMin + 0.1 * AxisMin;
end;


procedure TfmCoranMain.cbVariableIDOutlierMapChange(Sender: TObject);
begin
  sbShow4DClick(Sender);
end;

procedure TfmCoranMain.sbShowOutlierMapClick(Sender: TObject);
begin
  DBChart12.Series[0].Active := false;
  DBChart12.Series[1].Active := false;
  DBChart12.Series[2].Active := false;
  AssignChartDataSources('Close');
  AssignChartDataSources('Open');
  DBChart12.Series[0].Active := true;
  DBChart12.Series[1].Active := true;
  DBChart12.Series[2].Active := true;
end;

procedure TfmCoranMain.CalculateOrthogonalDistances(NumEigenVectorsSelected : integer);
var
  i, j : integer;
  tmp, t, tSum : double;
  ml : integer;
  XMed, TriMean, HingeL, HingeU : double;
begin
  FillChar(DD,sizeof(DD),0);
  {square roots of eigenvalues in A3, eigenvectors in A2, scores in B}
  sbMain.Panels[1].Text := 'Deleting orthogonal distances';
  sbMain.Refresh;
  dmCor.DeleteSDOD.ExecSQL;
  {first, calculate score distances}
    // Mahalanobis distance in classical PCA subspace
    //Tclas=classic.Xc*classic.P(:,1:out.k);
    //out.classic.sd=sqrt(mahalanobis(Tclas,zeros(size(Tclas,2),1),'invcov',1./classic.L(1:out.k)))';
    //classic.P        : loadings (eigenvectors)
    //classic.L        : eigenvalues
    //classic.M        : center of the data
    //classic.T        : scores
    //classic.Xc       : mean-subtracted data
  sbMain.Panels[1].Text := 'Reading data';
  sbMain.Refresh;
  dmCor.cdsCoranChem.Open;
  dmCor.cdsCoranChem.First;
  for i := 1 to NumSamples do
  begin
    for j := 1 to Nox do
    begin
      X1[i,j] := dmCor.cdsCoranChem.Fields[j+2].AsFloat;
      //if (i < 11) then ShowMessage('read - X1 '+IntToStr(i)+IntToStr(j)+' '+FormatFloat('####0.0000',X1[i,j]));
    end;
    dmCor.cdsCoranChem.Next;
  end;
  dmCor.cdsCoranChem.First;
  case scSimilarityChoice of
    scCorrespondence : begin
      ScaleCorAnal(NumSamples,Nox);
      for i := 1 to NumSamples do
      begin
        for j := 1 to Nox do
        begin
          X1[i,j] := W[i,j];
        end;
      end;
    end;
    scPCAVariance : begin
    end;
    scPCAPearsonR : begin
      Stand(X,NumSamples,Nox);
    end;
    scPCASpearmanR : begin
      SpearmanRho(X,A3,NumSamples,Nox);
    end;
    scPCAKendallR : begin
      KendallTau(X,A3,NumSamples,Nox);
    end;
    scRQVariance : begin
      {
      ScaleRQAnalVar(NumSamples,Nox);
      for i := 1 to NumSamples do
      begin
        for j := 1 to Nox do
        begin
          X1[i,j] := W[i,j];
          //if (i < 11) then ShowMessage('scaled - X1 '+IntToStr(i)+IntToStr(j)+' '+FormatFloat('####0.0000',X1[i,j]));
        end;
      end;
      }
    end;
    scRQPearsonR : begin
      Stand(X,NumSamples,Nox);
      {
      ScaleRQAnalStd(NumSamples,Nox);
      for i := 1 to NumSamples do
      begin
        for j := 1 to Nox do
        begin
          X1[i,j] := W[i,j];
        end;
      end;
      }
    end;
    scRQSpearmanR : begin
      ScaleRQAnalSpearman(NumSamples,Nox);
      for i := 1 to NumSamples do
      begin
        for j := 1 to Nox do
        begin
          X1[i,j] := W[i,j];
        end;
      end;
    end;
    scRQKendallR : begin
    end;
    scDiscrim2Grp : begin
    end;
    scDiscrimnGrp : begin
    end;
    scCluster : begin
    end;
  end;
  sbMain.Panels[1].Text := 'Reading eigenvectors';
  sbMain.Refresh;
  dmCor.EigenVec.Open;
  dmCor.EigenVec.First;
  for i := 1 to Nox do
  begin
    for j := 1 to Nox do
    begin
      A2[i,j] := dmCor.EigenVec.Fields[j].AsFloat;
      //if (i < 3) then ShowMessage('read - A2 '+IntToStr(i)+IntToStr(j)+' '+FormatFloat('####0.0000',A2[i,j]));
    end;
    dmCor.EigenVec.Next;
  end;
  dmCor.EigenVec.First;
  sbMain.Panels[1].Text := 'Reading eigenvalues';
  sbMain.Refresh;
  dmCor.EigenVal.Open;
  dmCor.EigenVal.First;
  for i := 1 to Nox do
  begin
    for j := 1 to Nox do
    begin
      A3[i,j] := 0.0;
    end;
    A3[i,i] := dmCor.EigenVal.Fields[1].AsFloat;
    //if (i < 3) then ShowMessage('read - A3 '+IntToStr(i)+' '+FormatFloat('####0.0000',A3[i,i]));
    dmCor.EigenVal.Next;
  end;
  dmCor.EigenVal.First;
  sbMain.Panels[1].Text := 'Calculating score distances';
  sbMain.Refresh;
  for j := 1 to Nox do
  begin
    for i := 1 to Nox do
    begin
      TempC[i,j] := 0.0;
    end;
    //put eigenvalues in TempC
    TempC[j,j] := A3[j,j];
    //if (j < 4) then ShowMessage('TempC '+IntToStr(j)+' '+FormatFloat('####0.0000',TempC[j,j]));
    DD[j] := 0.0;
    for i := 1 to NumSamples do
    begin
      DD[j] := DD[j] + X1[i,j];
      //if (i < 4) then ShowMessage('X1 '+IntToStr(i)+IntToStr(j)+' '+FormatFloat('####0.0000',X1[j,j]));
    end;
    DD[j] := DD[j]/(1.0*NumSamples);
    //if (j < 5) then ShowMessage('DD '+IntToStr(j)+' '+FormatFloat('####0.0000',DD[j]));
    if (scSimilarityChoice in [scPCAVariance,scPCAPearsonR,scRQVariance,scRQPearsonR]) then   // added RQ to get same results as PCA
    begin
      for i := 1 to NumSamples do
      begin
        X[i,j] := X1[i,j] - DD[j];
        //if (i < 4) then ShowMessage('calc - X '+IntToStr(i)+IntToStr(j)+' '+FormatFloat('####0.0000',X[i,j]));
      end;
    end;
    DD[j] := 0.0;
  end;
    //Tclas=classic.Xc*classic.P(:,1:out.k);
  MMult(X,A2,TempM,NumSamples,Nox,Nox);  //must be Nox,Nox here
  for i := 1 to NumSamples do
  begin
    for j := 1 to 2 do
    begin
      //if (i < 5) then ShowMessage('MMult - TempM '+IntToStr(i)+IntToStr(j)+' '+FormatFloat('####0.0000',TempM[i,j]));
    end;
  end;

  for j := 1 to Nox do
  begin
    for i := 1 to NumSamples do
    begin
      //if (TempC[i,j] > 0.0) then TempC[i,j] := 1.0/TempC[i,j];
      //if (j < 4) then ShowMessage('TempC '+IntToStr(i)+IntToStr(j)+' '+FormatFloat('####0.0000',TempC[i,j]));
    end;
    //if (j < 4) then ShowMessage('DD '+IntToStr(j)+' '+FormatFloat('####0.0000',DD[j]));
  end;

    //out.classic.sd=sqrt(mahalanobis(Tclas,zeros(size(Tclas,2),1),'invcov',1./classic.L(1:out.k)))';
    //classic.P        : loadings (eigenvectors)
    //classic.L        : eigenvalues
    //classic.M        : center of the data
    //classic.T        : scores
    //classic.Xc       : mean-subtracted data
    //ShowMessage('check this Mahalanobis');
  Mahalanobis(TempM,DD,TempC,DR,NumSamples,NumEigenVectorsSelected);
  //Mahalanobis(B,DD,TempC,DR,NumSamples,NumEigenVectorsSelected);
  for i := 1 to NumSamples do
  begin
    if (DR[i] > 0.0) then
      DR[i] := sqrt(DR[i])
    else
      DR[i] := 0.0;
  end;
  sbMain.Panels[1].Text := 'Calculating orthogonal distances';
  sbMain.Refresh;
  {now calculate orthogonal distances}
  FillChar(DD,sizeof(DD),0);
  dmCor.cdsCoranChem.Open;
  dmCor.cdsCoranChem.First;
  for i := 1 to NumSamples do
  begin
    for j := 1 to Nox do
    begin
      X1[i,j] := dmCor.cdsCoranChem.Fields[j+2].AsFloat;
    end;
    dmCor.cdsCoranChem.Next;
  end;
  dmCor.cdsCoranChem.First;
  case scSimilarityChoice of
    scCorrespondence : begin
      ScaleCorAnal(NumSamples,Nox);
      for i := 1 to NumSamples do
      begin
        for j := 1 to Nox do
        begin
          X1[i,j] := W[i,j];
        end;
      end;
    end;
    scPCAVariance : begin
    end;
    scPCAPearsonR : begin
      Stand(X,NumSamples,Nox);
    end;
    scPCASpearmanR : begin
      SpearmanRho(X,A3,NumSamples,Nox);
    end;
    scPCAKendallR : begin
      KendallTau(X,A3,NumSamples,Nox);
    end;
    scRQVariance : begin
      {
      ScaleRQAnalVar(NumSamples,Nox);
      for i := 1 to NumSamples do
      begin
        for j := 1 to Nox do
        begin
          X1[i,j] := W[i,j];
        end;
      end;
      }
    end;
    scRQPearsonR : begin
      Stand(X,NumSamples,Nox);
      {
      ScaleRQAnalStd(NumSamples,Nox);
      for i := 1 to NumSamples do
      begin
        for j := 1 to Nox do
        begin
          X1[i,j] := W[i,j];
        end;
      end;
      }
    end;
    scRQSpearmanR : begin
      ScaleRQAnalSpearman(NumSamples,Nox);
      for i := 1 to NumSamples do
      begin
        for j := 1 to Nox do
        begin
          X1[i,j] := W[i,j];
        end;
      end;
    end;
    scRQKendallR : begin
    end;
    scDiscrim2Grp : begin
    end;
    scDiscrimnGrp : begin
    end;
    scCluster : begin
    end;
  end;
  dmCor.EigenVec.Open;
  dmCor.EigenVec.First;
  for i := 1 to Nox do
  begin
    for j := 1 to Nox do
    begin
      A2[i,j] := dmCor.EigenVec.Fields[j].AsFloat;
    end;
    dmCor.EigenVec.Next;
  end;
  dmCor.EigenVec.First;
  dmCor.EigenVal.Open;
  dmCor.EigenVal.First;
  for i := 1 to Nox do
  begin
    for j := 1 to Nox do
    begin
      A3[i,j] := 0.0;
    end;
    A3[i,i] := dmCor.EigenVal.Fields[1].AsFloat;
    //if (i < 3) then ShowMessage('read - A3 '+IntToStr(i)+' '+FormatFloat('####0.0000',A3[i,i]));
    dmCor.EigenVal.Next;
  end;
  dmCor.EigenVal.First;
  for j := 1 to Nox do
  begin
    for i := 1 to Nox do
    begin
      TempC[i,j] := 0.0;
    end;
    //put eigenvalues in TempC
    TempC[j,j] := A3[j,j];
  end;
  case scSimilarityChoice of
    scCorrespondence, scPCAVariance, scPCAPearsonR,
    scRQVariance, scRQPearsonR : begin
      for j := 1 to Nox do
      begin
        DD[j] := 0.0;
        for i := 1 to NumSamples do
        begin
          DD[j] := DD[j] + X1[i,j];
        end;
        DD[j] := DD[j]/(1.0*NumSamples);
        for i := 1 to NumSamples do
        begin
          X[i,j] := X1[i,j] - DD[j];
        end;
        DD[j] := 0.0;
      end;
        //Tclas=classic.Xc*classic.P(:,1:out.k);
      MMult(X,A2,TempM,NumSamples,Nox,Nox);  //must be Nox,Nox here
    end;
    scPCASpearmanR, scPCAKendallR,
    scRQSpearmanR, scRQKendallR : begin
      for j := 1 to NumEigenVectorsSelected do
      begin
        //dmCor.qDim4Smp.Close;
        dmCor.cdsqDim4Smp.Filtered := false;
        dmCor.cdsqDim4Smp.Filter := 'VariableID = '+''''+'Vector'+IntToStr(j)+'''';
        dmCor.cdsqDim4Smp.Filtered := true;
        dmCor.cdsqDim4Smp.Open;
        for i := 1 to NumSamples do
        begin
          DR[i] :=  dmCor.cdsqDim4SmpSmpValue.AsFloat;
          dmCor.cdsqDim4Smp.Next;
        end;
        Sort(DR,NumSamples,1);
        ml := (NumSamples+1) div 2;
        XMed := 0.5*(DR[ml] + DR[NumSamples-ml+1]);
        Hinges(DR,HingeL,HingeU,NumSamples);
        TriMean := 0.5*XMed + 0.25*HingeL + 0.25*HingeU;
        DD[j] := TriMean;
      end;
      for i := 1 to NumSamples do
      begin
        for j := 1 to NumEigenVectorsSelected do
        begin
          TempM[i,j] := DD[j];
        end;
      end;
      //dmCor.qDim4Smp.Close;
      dmCor.cdsqDim4Smp.Filtered := false;
      //dmCor.qDim4Smp.Open;
      for j := 1 to Nox do
      begin
        for i := 1 to NumSamples do
        begin
          X[i,j] := X1[i,j] - DD[j];
        end;
        DD[j] := 0.0;
      end;
        //Tclas=classic.Xc*classic.P(:,1:out.k);
      MMult(X,A2,TempM,NumSamples,Nox,Nox);  //must be Nox,Nox here
    end;
    scRobPCA : begin
    end;
  end;
    //Xtilde=Tclas*classic.P(:,1:out.k)';
  sbMain.Panels[1].Text := 'Calculate Xtilde = Tclas*classic.P(:,1:out.k)';
  sbMain.Refresh;
  {transpose the eigenvectors into a temporary matrix}
  for i := 1 to Nox do
  begin
    for j := 1 to Nox do
    begin
      TempC[j,i] := A2[i,j];
      //if (i < 6) then ShowMessage('A2 '+IntToStr(i)+IntToStr(j)+' '+FormatFloat('####0.0000',A2[i,j]));
    end;
  end;
  MMult(TempM,TempC,W,NumSamples,NumEigenVectorsSelected,NumEigenVectorsSelected);
  //MMult(B,TempC,W,NumSamples,NumEigenVectorsSelected,NumEigenVectorsSelected);
  for i := 1 to NumSamples do
  begin
    for j := 1 to NumEigenVectorsSelected do
    begin
      //if (i < 6) then ShowMessage('W '+IntToStr(i)+IntToStr(j)+' '+FormatFloat('####0.0000',W[i,j]));
    end;
  end;
    //Cdiff=classic.Xc-Xtilde;
    //for i=1:n
    //    out.classic.od(i,1)=norm(Cdiff(i,:));
    //end
  sbMain.Panels[1].Text := 'Calculate Cdiff = Xc - Xtilde';
  sbMain.Refresh;
  Fillchar(DR2,sizeof(DR2),0);
  for i := 1 to NumSamples do
  begin
    for j := 1 to NumEigenVectorsSelected do
    begin
      TempM[i,j] := X[i,j] - W[i,j];
      //if (i < 6) then ShowMessage('TempM '+IntToStr(i)+IntToStr(j)+' '+FormatFloat('####0.0000',TempM[i,j]));
    end;
    DR2[i] := DR2[i] + TempM[i,NumEigenVectorsSelected]*TempM[i,NumEigenVectorsSelected];
  end;
  for i := 1 to NumSamples do
  begin
    if (DR2[i] > 0.0) then DR2[i] := sqrt(DR2[i]);
  end;
  sbMain.Panels[1].Text := 'Calculate orthogonal distances';
  sbMain.Refresh;
  dmCor.cdsFacLoadingsSmp.First;
  dmCor.tSDOD.Open;
  dmCor.tSDOD.First;
  i := 0;
  repeat
    i := i + 1;
    dmCor.tSDOD.Append;
    dmCor.tSDODGROUPNAME.AsString := dmCor.cdsFacLoadingsSmpGROUPNAME.AsString;
    dmCor.tSDODPlotGroupName.AsString := dmCor.cdsFacLoadingsSmpPlotGroupName.AsString;
    dmCor.tSDODSAMPLENUM.AsString := dmCor.cdsFacLoadingsSmpSAMPLENUM.AsString;
    dmCor.tSDODSeq.AsInteger := i;
    {calculate score distance}
    dmCor.tSDODScoreDistance.AsFloat := DR[i];
    tmp := DR2[i];
    if (NumEigenVectorsSelected = Nox) then tmp := i;
    dmCor.tSDODOrthogonalDistance.AsFloat := tmp;
    //if (i < 6) then ShowMessage('OD '+IntToStr(i)+' '+FormatFloat('####0.0000',tmp));
    dmCor.tSDOD.Post;
    dmCor.cdsFacLoadingsSmp.Next
  until dmCor.cdsFacLoadingsSmp.Eof;
  CalculateScoreOrthogonalCutoffs(NumEigenVectorsSelected);
  dmCor.tSDOD.Close;
end;

procedure TfmCoranMain.CalculateScoreOrthogonalCutoffs(NumEigenVectorsSelected : integer);
var
  i : integer;
  tx, ty, tmpx, tmpy : double;
  Chix, Chiy : double;
  t, tOD, tOD2, MeanOD, SDevOD : double;
  ta : RealArrayR;
begin
  {
  get Ymin and Ymax for orthogonal distances
  calculate chi-squared value for 97.5 % probability and num deg freedom
  append chi-squared value in DBChart12 from Ymin to YMax

  get Xmin and Xmax for score distances
  calculate chi-squared value for 97.5 % probability and num deg freedom
  append chi-squared value in DBChart12 from Xmin to XMax
  }

  ChiX := ChiInv(0.975,NumEigenVectorsSelected);
  if (NumEigenVectorsSelected < Nox) then ChiY := ChiInv(0.975,NumEigenVectorsSelected)
                                     else ChiY := 0.0;
  ChiX := Sqrt(InvChiSqValues[NumEigenVectorsSelected]);
  dmCor.tSDOD.Open;
  dmCor.tSDOD.First;
  XMin := dmCor.tSDODScoreDistance.AsFloat;
  YMin := dmCor.tSDODOrthogonalDistance.AsFloat;
  XMax := XMin;
  YMax := YMin;
  tOD := 0.0;
  tOD2 := 0.0;
  //ShowMessage(IntToStr(NumSamples));
  i := 0;
  repeat
    i := i + 1;
    if (XMin > dmCor.tSDODScoreDistance.AsFloat) then XMin := dmCor.tSDODScoreDistance.AsFloat;
    if (XMax < dmCor.tSDODScoreDistance.AsFloat) then XMax := dmCor.tSDODScoreDistance.AsFloat;
    if (YMin > dmCor.tSDODOrthogonalDistance.AsFloat) then YMin := dmCor.tSDODOrthogonalDistance.AsFloat;
    if (YMax < dmCor.tSDODOrthogonalDistance.AsFloat) then YMax := dmCor.tSDODOrthogonalDistance.AsFloat;
    //ShowMessage(IntToStr(i)+' OD '+FormatFloat('###0.00000',dmCor.tSDODOrthogonalDistance.AsFloat));
    t := Power(0.666666,dmCor.tSDODOrthogonalDistance.AsFloat);
    //ShowMessage(IntToStr(i)+' t '+FormatFloat('###0.00000',t));
    tOD := tOD + t;
    ta[i] := t;
    dmCor.tSDOD.Next;
  until dmCor.tSDOD.Eof;
  // m=mean(out.classic.od.^(2/3));
  // s=sqrt(var(out.classic.od.^(2/3)));
  // out.classic.cutoff.od = sqrt(norminv(cutoff,m,s)^3);
  MeanOD := tOD/(1.0*NumSamples);
  tOD2 := 0.0;
  for i := 1 to NumSamples do
  begin
    t := ta[i] - MeanOD;
    t := t*t;
    tOD2 := tOD2 + t;
  end;
  SDevOD := tOD2/(1.0*(NumSamples-1));
  SDevOD := Sqrt(SDevOD);
  //ShowMessage('Mean '+FormatFloat('###0.00000',MeanOD)+'   SDev '+FormatFloat('###0.00000',SDevOD));
  t := invCumNorm(Cutoff);
  //ShowMessage(FormatFloat('###0.00000',t));
  t := t*SDevOD + MeanOD;
  ChiY := Power(1.5,t);
  //ShowMessage('ChiY '+FormatFloat('###0.00000',ChiY));
  if (XMax < ChiX) then XMax := ChiX;
  if (YMax < ChiY) then YMax := ChiY;
  tx := (Xmax - XMin)/10.0;
  ty := (Ymax - YMin)/10.0;
  tmpx := XMin;
  tmpy := YMin;
  SeriesG1205.Clear;  //Score distance cutoff
  SeriesG1206.Clear;  //Orthogonal distance cutoff
  repeat
    SeriesG1205.AddXY(ChiX,tmpy);
    SeriesG1206.AddXY(tmpx,ChiY);
    tmpx := tmpx + tx;
    tmpy := tmpy + ty;
  until (tmpx >= XMax);
  SeriesG1205.AddXY(ChiX,YMax);
  SeriesG1206.AddXY(XMax,ChiY);
  dmCor.tSDOD.Close;
end;



procedure TfmCoranMain.Button2Click(Sender: TObject);
begin
  dmCor.qSDOD.Open;
end;

procedure TfmCoranMain.bCloseAllClick(Sender: TObject);
begin
  dmCor.Coran.Close;
end;

procedure TfmCoranMain.boOpenAllClick(Sender: TObject);
begin
  dmCor.cdsCoranChem.Open;
  dmCor.cdsCoranRaw.Open;
  dmCor.CoranStats.Open;
  dmCor.CoranSimilarity.Open;
  dmCor.EigenVec.Open;
  dmCor.EigenVal.Open;
  AssignChartDataSources('Open');
end;

procedure TfmCoranMain.Button3Click(Sender: TObject);
begin
  dmCor.FacLoadingsVar.Close;
  dmCor.FacLoadingsVarNoZero.Close;
  dmCor.FacLoadingsVar.Open;
  dmCor.FacLoadingsVarNoZero.Open;
end;

procedure TfmCoranMain.ScaleAxesEqually1Click(Sender: TObject);
begin
  ScaleAxesEqually1.Checked := not ScaleAxesEqually1.Checked;
  ScaleAxesEqually;
end;

procedure TfmCoranMain.ScaleAxesEqually;
begin
  if ((pc1.ActivePageIndex in [8]) and ScaleAxesEqually1.Checked)
    then ScaleAxesEquallyGraph1;
  if ((pc1.ActivePageIndex in [9]) and ScaleAxesEqually1.Checked)
    then ScaleAxesEquallyGraph2;
  if ((pc1.ActivePageIndex in [10]) and ScaleAxesEqually1.Checked)
    then ScaleAxesEquallyGraph3;
  if ((pc1.ActivePageIndex in [13]) and ScaleAxesEqually1.Checked)
    then ScaleAxesEquallyGraph3D;

  if ScaleAxesEqually1.Checked then
  begin
  end else
  begin
    DBChart1.BottomAxis.Automatic := true;
    DBChart1.LeftAxis.Automatic := true;
    DBChart2.BottomAxis.Automatic := true;
    DBChart2.LeftAxis.Automatic := true;
    DBChart3.BottomAxis.Automatic := true;
    DBChart3.LeftAxis.Automatic := true;
    DBChart8.BottomAxis.Automatic := true;
    DBChart8.LeftAxis.Automatic := true;
    DBChart8.DepthAxis.Automatic := true;
  end;
end;

procedure TfmCoranMain.ScaleAxesEquallyGraph1;
var
  XMax1, XMin1, YMax1, YMin1,
  tMax, tMin : double;
begin
  XMax1 := DBChart1.BottomAxis.Maximum;
  XMin1 := DBChart1.BottomAxis.Minimum;
  YMax1 := DBChart1.LeftAxis.Maximum;
  YMin1 := DBChart1.LeftAxis.Minimum;
  if ScaleAxesEqually1.Checked then
  begin
    tMin := Min(XMin1,YMin1);
    tMax := Max(XMax1,YMax1);
    tMin := tMin - (tMax-tMin)/100.0;
    tMax := tMax + (tMax-tMin)/100.0;
    DBChart1.BottomAxis.SetMinMax(tMin,tMax);
    DBChart1.LeftAxis.SetMinMax(tMin,tMax);
    DBChart1.BottomAxis.Automatic := false;
    DBChart1.LeftAxis.Automatic := false;
  end else
  begin
    DBChart1.BottomAxis.Automatic := true;
    DBChart1.LeftAxis.Automatic := true;
  end;
end;

procedure TfmCoranMain.ScaleAxesEquallyGraph2;
var
  XMax2, XMin2, YMax2, YMin2,
  tMax, tMin : double;
begin
  XMax2 := DBChart2.BottomAxis.Maximum;
  XMin2 := DBChart2.BottomAxis.Minimum;
  YMax2 := DBChart2.LeftAxis.Maximum;
  YMin2 := DBChart2.LeftAxis.Minimum;

  if ScaleAxesEqually1.Checked then
  begin
    tMin := Min(XMin2,YMin2);
    tMax := Max(XMax2,YMax2);
    tMin := tMin - (tMax-tMin)/100.0;
    tMax := tMax + (tMax-tMin)/100.0;
    DBChart2.BottomAxis.SetMinMax(tMin,tMax);
    DBChart2.LeftAxis.SetMinMax(tMin,tMax);
    DBChart2.BottomAxis.Automatic := false;
    DBChart2.LeftAxis.Automatic := false;
  end else
  begin
    DBChart2.BottomAxis.Automatic := true;
    DBChart2.LeftAxis.Automatic := true;
  end;
end;

procedure TfmCoranMain.ScaleAxesEquallyGraph3;
var
  XMax3, XMin3, YMax3, YMin3,
  tMax, tMin : double;
begin
  XMax3 := DBChart3.BottomAxis.Maximum;
  XMin3 := DBChart3.BottomAxis.Minimum;
  YMax3 := DBChart3.LeftAxis.Maximum;
  YMin3 := DBChart3.LeftAxis.Minimum;
  if ScaleAxesEqually1.Checked then
  begin
    tMin := Min(XMin3,YMin3);
    tMax := Max(XMax3,YMax3);
    tMin := tMin - (tMax-tMin)/100.0;
    tMax := tMax + (tMax-tMin)/100.0;
    DBChart3.BottomAxis.SetMinMax(tMin,tMax);
    DBChart3.LeftAxis.SetMinMax(tMin,tMax);
    DBChart3.BottomAxis.Automatic := false;
    DBChart3.LeftAxis.Automatic := false;
  end else
  begin
    DBChart3.BottomAxis.Automatic := true;
    DBChart3.LeftAxis.Automatic := true;
  end;
end;

procedure TfmCoranMain.ScaleAxesEquallyGraph3D;
var
  XMax, XMin, YMax, YMin,
  ZMax, ZMin : double;
  tMax, tMin : double;
begin
  XMax := DBChart8.BottomAxis.Maximum;
  XMin := DBChart8.BottomAxis.Minimum;
  ZMax := DBChart8.LeftAxis.Maximum;
  ZMin := DBChart8.LeftAxis.Minimum;
  YMax := DBChart8.DepthAxis.Maximum;
  YMin := DBChart8.DepthAxis.Minimum;
  if ScaleAxesEqually1.Checked then
  begin
    tMin := Min(XMin,YMin);
    tMin := Min(tMin,ZMin);
    tMax := Max(XMax,YMax);
    tMax := Max(tMax,ZMax);
    tMin := tMin - (tMax-tMin)/100.0;
    tMax := tMax + (tMax-tMin)/100.0;
    DBChart8.BottomAxis.SetMinMax(tMin,tMax);
    DBChart8.LeftAxis.SetMinMax(tMin,tMax);
    DBChart8.DepthAxis.SetMinMax(tMin,tMax);
    DBChart8.BottomAxis.Automatic := false;
    DBChart8.LeftAxis.Automatic := false;
    DBChart8.DepthAxis.Automatic := false;
  end else
  begin
    DBChart8.BottomAxis.Automatic := true;
    DBChart8.LeftAxis.Automatic := true;
    DBChart8.DepthAxis.Automatic := true;
  end;
end;

procedure TfmCoranMain.CheckColumnTotals(N, M : integer; var AllOK : boolean);
var
  i, j : integer;
  Total : double;
begin
  AllOK := true;
  //dmCor.cdsCoranRaw.DisableControls;
  dmCor.cdsCoranRaw.Open;
  for j := 1 to M do
  begin
    Total := 0.0;
    dmCor.cdsCoranRaw.First;
    for i := 1 to N do
    begin
      Total := Total + dmCor.cdsCoranRaw.Fields[j+2].AsFloat;
      dmCor.cdsCoranRaw.Next;
    end;
    if (Total < DefaultZeroLimit) then
    begin
      MessageDlg('Column '+IntToStr(j)+' ('+DBGrid1.Columns[j+2].Title.Caption+') does not contain any data! Can not continue. Total = '+FormatFloat('#0.0000',Total),mtWarning,[mbOK],0);
      AllOK := false;
      Exit;
    end;
  end;
  dmCor.cdsCoranRaw.First;
  dmCor.cdsCoranRaw.EnableControls;
end;

procedure TfmCoranMain.CheckColumnTotalsStats(N, M : integer; var AllOK : boolean);
var
  j : integer;
  tMax, tMin : double;
begin
  AllOK := true;
  //dmCor.CoranStats.DisableControls;
  dmCor.CoranStats.Open;
  for j := 1 to M do
  begin
    dmCor.CoranStats.Locate('Summary','Minimum',[]);
    tMin := dmCor.CoranStats.Fields[j+1].AsVariant;
    dmCor.CoranStats.Locate('Summary','Maximum',[]);
    tMax := dmCor.CoranStats.Fields[j+1].AsVariant;
    if ((tMax - tMin) <= DefaultZeroLimit) then
    begin
      MessageDlg('Column '+IntToStr(j)+' ('+DBGrid1.Columns[j+2].Title.Caption+') does not contain any data! Can not continue.',mtWarning,[mbOK],0);
      AllOK := false;
      Exit;
    end;
  end;
  dmCor.CoranStats.First;
  dmCor.CoranStats.EnableControls;
end;

procedure TfmCoranMain.cbRawGraphVarChange(Sender: TObject);
var
  tmpStr : string;
begin
  if (cbRawGraphVar.ItemIndex < 0) then cbRawGraphVar.Items.Add('Param1');
  if (cbRawGraphVar.ItemIndex < 0) then cbRawGraphVar.ItemIndex := 0;
  dmCor.cdsCoranChem.Open;
  NumSamples := dmCor.cdsCoranChem.RecordCount;
  //Boxplot
  DBChartBox.Series[0].Active := false;
  tmpStr := 'Param'+IntToStr(cbRawGraphVar.ItemIndex+1);
  DBChartBox.Series[0].XValues.ValueSource := tmpStr;
  DBChartBox.Series[0].Active := true;
  //Histogram
  DBChartHist.Series[0].Active := false;
  tmpStr := 'Param'+IntToStr(cbRawGraphVar.ItemIndex+1);
  DBChartHist.Series[0].YValues.ValueSource := tmpStr;
  DBChartHist.Series[0].Active := true;
  //sequence plot
  DBChart15.Series[0].Active := false;
  DBChart15.Series[0].XValues.Order := loNone;
  DBChart15.Series[0].YValues.Order := loNone;
  tmpStr := 'Param'+IntToStr(cbRawGraphVar.ItemIndex+1);
  DBChart15.Series[0].XValues.ValueSource := tmpStr;
  DBChart15.Series[0].Active := true;
  //Quantile plot
  if MakeQQPlot then
  begin
    DBChartQuantile.Series[0].Active := false;
    DBChartQuantile.Series[0].XValues.Order := loNone;
    DBChartQuantile.Series[0].YValues.Order := loNone;
    tmpStr := 'Param'+IntToStr(cbRawGraphVar.ItemIndex+1);
    DBChartQuantile.Series[0].XValues.ValueSource := tmpStr;
    tmpStr := 'Norm'+IntToStr(cbRawGraphVar.ItemIndex+1);
    DBChartQuantile.Series[0].YValues.ValueSource := tmpStr;
    DBChartQuantile.Series[0].Active := true;

    DBChartQuantile.Series[1].Active := false;
    DBChartQuantile.Series[1].XValues.Order := loNone;
    DBChartQuantile.Series[1].YValues.Order := loNone;
    tmpStr := 'Norm'+IntToStr(cbRawGraphVar.ItemIndex+1);
    DBChartQuantile.Series[1].XValues.ValueSource := tmpStr;
    DBChartQuantile.Series[1].YValues.ValueSource := tmpStr;
    DBChartQuantile.Series[1].Active := true;
  end;
end;

procedure TfmCoranMain.Connecttodatabase1Click(Sender: TObject);
var
  ii : integer;
begin
  //ShowMessage('Connection String is');
  //ShowMessage(dmCor.ConnectionString);
  with dmCor do
  begin
    try
      Coran.Connected := false;
    except
    end;
    Coran.ConnectionString := dmCor.ConnectionString;
    //Coran.ConnectionString := 'FILE NAME='+ADODataLinkFile;
    //Coran.Provider := ADODataLinkFile;
    Coran.Connected := true;
    //Coran.Open('admin','');
  end;

  AssignChartDataSources('Close');
  with dmCor do
  begin
    cdsCoranChem.Open;
    cdsCoranRaw.Open;
    ElemNames.Open;
    QGroups.Open;
    QPlotGroups.Open;
    GroupedSmp.Open;
    QGroupedSmp.Open;
    SmpLoc.Open;
    GroupedSmpLoc.Open;
    QGroupedSmpLoc.Open;
    CoranStats.Open;
    cdsFacLoadingsSmp.Open;
    ElemNames.First;
    ii := ElemNames.RecordCount;
    if (ii > 0) then
    begin
      ii := 0;
      try
        repeat
          ii :=ii + 1;
          Oxidename[ElemNamesPos.AsInteger] := ElemNamesCalled.AsString;
          TakeLogs[ElemNamesPos.AsInteger] := ElemNamesTakeLog.AsString;
          DBGrid1.Columns[ii+2].Title.Caption := dmCor.ElemNamesCalled.AsString;
          DBGridStats.Columns[ii+1].Title.Caption := dmCor.ElemNamesCalled.AsString;
          dbgSimilarity.Columns[ii+1].Title.Caption := dmCor.ElemNamesCalled.AsString;
          ElemNames.Next;
        until ((ElemNames.Eof) or (ii >= MMax));
      except
      end;
      Nox := ii;
    end;
  end;
  if (NumEigenVectorsSelected > Nox) then NumEigenVectorsSelected := Nox;
  dmCor.ElemNames.First;
  for ii := Nox+1 to MMaxFields do
  begin
    DBGrid1.Columns[ii+2].Visible := false;
    DBGrid2.Columns[ii+2].Visible := false;
  end;
  for ii := Nox+1 to MMaxFields do
  begin
    DBGridStats.Columns[ii+1].Visible := false;
  end;
  AssignChartDataSources('Open');
  AddToComboBoxVar(Nox);
  //cbRawGraphVarChange(self);
  scSimilarityChoice := scNone;
  //cbVariableID.Clear;
  if Include4DVarData then
  begin
    tsLoc4D.Visible := true;
  end else
  begin
    tsLoc4D.Visible := false;
  end;
  //cbVariableID.Text := cbVariableID.Items[0];
  PrepareComboBoxes(0);
  try
    PrepareGraphs;
  except
  end;
  FromRowValueString := '2';
  ToRowValueString := '2';
end;

procedure TfmCoranMain.ConstructHistogram(iVariable : integer);
const
  numbins = 9;
var
  i,j : integer;
  tMin, tMax, tRange,
  tmp : double;
  tBinUpperBoundary : array[1..numbins-1] of single;
  tBinCount : array[1..numbins] of integer;
begin
  //determine minimum
  //determine maximum
  //subdivide into fixed number of bins
  //allocate data to bins
  //store in datbase table
  for i := 1 to numbins do
  begin
    tBinCount[i] := 0;
  end;
  dmCor.CoranStats.Open;
  dmCor.CoranStats.First;
  dmCor.CoranStats.Locate('Summary','Minimum',[]);
  tMin := dmCor.CoranStats.Fields[iVariable+1].AsVariant;
  dmCor.CoranStats.Locate('Summary','Maximum',[]);
  tMax := dmCor.CoranStats.Fields[iVariable+1].AsVariant;
  tRange := (tMax - tMin)/(1.0*numbins);
  for i := 1 to numbins-1 do
  begin
    tBinUpperBoundary[i] := tMin + tRange*(1.0*i);
    //if (iVariable < 2) then ShowMessage('uprbound '+FormatFloat('####0.000',tBinUpperBoundary[i]));
  end;
  dmCor.cdsCoranChem.Close;
  dmCor.cdsCoranChem.Open;
  dmCor.cdsCoranChem.First;
  j := 0;
  repeat
    j := j + 1;
    //ShowMessage(intToStr(j));
    tmp := dmCor.cdsCoranChem.Fields[iVariable+2].AsVariant;
    if (tmp <= tBinUpperBoundary[1]) then tbinCount[1] := tBinCount[1] + 1;
    if (tmp > tBinUpperBoundary[numbins-1]) then tbinCount[numbins] := tBinCount[numbins] + 1;
    for i := 2 to numbins-1 do
    begin
      if ((tmp <= tBinUpperBoundary[i]) and (tmp > tBinUpperBoundary[i-1])) then
        tbinCount[i] := tBinCount[i] + 1;
    end;
    dmCor.cdsCoranChem.Next;
  until dmCor.cdsCoranChem.Eof;
  dmCor.QHist.Open;
  dmCor.QHist.First;
  for i := 1 to numbins do
  begin
    dmCor.QHist.Edit;
    dmCor.QHist.Fields[iVariable].AsVariant := 0;
    dmCor.QHist.Fields[iVariable].AsVariant := tBinCount[i];
    dmCor.QHist.Next;
  end;
end;

procedure TfmCoranMain.cbScoreGraphVarChange(Sender: TObject);
var
  tmpStr : string;
begin
  if (cbScoreGraphVar.ItemIndex < 0) then cbScoreGraphVar.Items.Add('Vector1');
  if (cbScoreGraphVar.ItemIndex < 0) then cbScoreGraphVar.ItemIndex := 0;
  //sequence plot
  DBChart13.Series[0].Active := false;
  DBChart13.Series[0].XValues.Order := loNone;
  DBChart13.Series[0].YValues.Order := loNone;
  tmpStr := 'Vector'+IntToStr(cbScoreGraphVar.ItemIndex+1);
  DBChart13.Series[0].XValues.ValueSource := tmpStr;
  DBChart13.Series[0].Active := true;
  //Boxplot
  DBChartBoxScore.Series[0].Active := false;
  tmpStr := 'Vector'+IntToStr(cbScoreGraphVar.ItemIndex+1);
  DBChartBoxScore.Series[0].XValues.ValueSource := tmpStr;
  DBChartBoxScore.Series[0].Active := true;
end;

procedure TfmCoranMain.CalculateQuantiles(Nox : integer);
var
  i, j, n : integer;
  tMean, tSDev,
  tmp : double;
begin
  //get mean and standard deviation
  //sort
  //calculate z-scores
  //allocate data to bins
  //store in datbase table
  //pc1.ActivePage := tsCheck;
  //dmCor.cdsCoranChem.DisableControls;
  //dmCor.CoranStats.DisableControls;
  //dmCor.cdsCoranQuantile.DisableControls;
  dmCor.CoranStats.Open;
  dmCor.cdsCoranChem.Open;
  n := dmCor.cdsCoranChem.RecordCount;
  for i := 1 to n do
  begin
    tmp := (1.0*i-0.375)/(1.0*n+0.25);
    // another example uses tmp := (1.0*i-3.0/8.0)/(1.0*n+1.0/4.0)
    //tmp := (1.0*i)/(1.0*n+1.0);
    DR[i] := invCumNorm(tmp);
  end;
  dmCor.DeleteQuantiles.ExecSQL;
  dmCor.cdsCoranQuantile.Close;
  dmCor.cdsCoranQuantile.Open;
  dmCor.cdsCoranQuantile.First;
  dmCor.cdsCoranChem.First;
  sbMain.Panels[0].Text := IntToStr(0);
  sbMain.Panels[1].Text := 'Calculating quantiles - default values';
  sbMain.Refresh;
  i := 0;
  repeat
    i := i + 1;
    if (i mod 100 = 0) then
    begin
      sbMain.Panels[0].Text := '  smp '+IntToStr(i);
      sbMain.Refresh;
    end;
    dmCor.cdsCoranQuantile.Append;
    dmCor.cdsCoranQuantileGROUPNAME.AsString := dmCor.cdsCoranChemGROUPNAME.AsString;
    dmCor.cdsCoranQuantilePlotGroupName.AsString := dmCor.cdsCoranChemPlotGroupName.AsString;
    dmCor.cdsCoranQuantileSAMPLENUM.AsString := dmCor.cdsCoranChemSAMPLENUM.AsString;
    for j := 1 to Nox do
    begin
      dmCor.cdsCoranQuantile.Fields[j+2].AsVariant := dmCor.cdsCoranChem.Fields[j+2].AsVariant;
    end;
    dmCor.cdsCoranQuantile.Post;
    dmCor.cdsCoranChem.Next;
  until dmCor.cdsCoranChem.Eof;
  sbMain.Panels[1].Text := 'Calculating quantiles - distribution values';
  sbMain.Panels[0].Text := '  j '+IntToStr(0);
  sbMain.Refresh;
  for j := 1 to Nox do
  begin
    if (j mod 5 = 0) then
    begin
      sbMain.Panels[0].Text := '  var '+IntToStr(j);
      sbMain.Refresh;
    end;
    dmCor.CoranStats.First;
    dmCor.CoranStats.Locate('Summary','Mean',[]);
    tMean := dmCor.CoranStats.Fields[j+1].AsVariant;
    dmCor.CoranStats.Locate('Summary','Std dev',[]);
    tSDev := dmCor.CoranStats.Fields[j+1].AsVariant;
    dmCor.cdsCoranQuantile.First;
    repeat
      dmCor.cdsCoranQuantile.Edit;
      tmp := dmCor.cdsCoranQuantile.Fields[j+2].AsVariant;
      if (tSDev > 0.0) then dmCor.cdsCoranQuantile.Fields[j+2].AsVariant := (tmp-tMean)/tSDev
                       else dmCor.cdsCoranQuantile.Fields[j+2].AsVariant := -99.0;
      dmCor.cdsCoranQuantile.Post;
      dmCor.cdsCoranQuantile.Next;
    until dmCor.cdsCoranQuantile.Eof;
  end;
  dmCor.cdsCoranChem.First;
  dmCor.cdsCoranQuantile.First;
  dmCor.CoranStats.First;
  sbMain.Panels[1].Text := 'Calculating quantiles - normal values';
  sbMain.Panels[0].Text := '  j '+IntToStr(0);
  sbMain.Refresh;
  for j := 1 to Nox do
  begin
    //if (j mod 5 = 0) then
    //begin
      sbMain.Panels[0].Text := '  var '+IntToStr(j);
      sbMain.Refresh;
    //end;
    dmCor.cdsCoranQuantile.IndexFieldNames := 'Param'+IntToStr(j);
    dmCor.cdsCoranQuantile.First;
    for i := 1 to n do
    begin
      dmCor.cdsCoranQuantile.Edit;
      dmCor.cdsCoranQuantile.Fields[MMaxFields+j+2].AsVariant := DR[i];
      dmCor.cdsCoranQuantile.Post;
      dmCor.cdsCoranQuantile.Next;
    end;
  end;
  sbMain.Panels[1].Text := 'Saving quantiles to database';
  sbMain.Refresh;
  dmCor.cdsCoranQuantile.ApplyUpdates(-1);
  dmCor.cdsCoranQuantile.IndexFieldNames := '';
  dmCor.cdsCoranChem.EnableControls;
  dmCor.CoranStats.EnableControls;
  dmCor.cdsCoranQuantile.EnableControls;
  sbMain.Panels[1].Text := '';
  sbMain.Panels[0].Text := '  ';
  sbMain.Refresh;
end;


procedure TfmCoranMain.CalculateQuantilesA(Nox : integer);
var
  i, j, n : integer;
  tMean, tSDev,
  tmp : double;
begin
  //no longer needed

  //get mean and standard deviation
  //sort
  //calculate z-scores
  //allocate data to bins
  //store in datbase table
  //pc1.ActivePage := tsCheck;
  //dmCor.cdsCoranChem.DisableControls;
  //dmCor.CoranStats.DisableControls;
  //dmCor.CoranQuantile.DisableControls;
  dmCor.CoranStats.Open;
  dmCor.cdsCoranChem.Open;
  n := dmCor.cdsCoranChem.RecordCount;
  for i := 1 to n do
  begin
    tmp := (1.0*i)/(1.0*n);
    if (i = n) then tmp := 0.9999;
    DR[i] := invCumNorm(tmp);
  end;
  dmCor.CoranQuantile.Open;
  dmCor.CoranQuantile.First;
  if not(dmCor.CoranQuantile.Bof and dmCor.CoranQuantile.Eof) then
  begin
    dmCor.CoranQuantile.Last;
    repeat
      dmCor.CoranQuantile.Delete;
      dmCor.CoranQuantile.Next;
    until dmCor.CoranQuantile.BOF;
  end;
  dmCor.cdsCoranChem.First;
  sbMain.Panels[0].Text := IntToStr(0);
  sbMain.Panels[1].Text := 'Calculating quantiles';
  sbMain.Refresh;
  i := 0;
  repeat
    i := i + 1;
    if (i mod 100 = 0) then
    begin
      sbMain.Panels[0].Text := '  i '+IntToStr(i);
      sbMain.Refresh;
    end;
    for j := 1 to Nox do
    begin
      X[i,j] := dmCor.cdsCoranChem.Fields[j+2].AsFloat;
    end;
    dmCor.cdsCoranChem.Next;
  until dmCor.cdsCoranChem.Eof;
  for j := 1 to Nox do
  begin
    if (j mod 3 = 0) then
    begin
      sbMain.Panels[0].Text := '  j '+IntToStr(j);
      sbMain.Refresh;
    end;
    dmCor.CoranStats.First;
    dmCor.CoranStats.Locate('Summary','Mean',[]);
    tMean := dmCor.CoranStats.Fields[j+1].AsVariant;
    dmCor.CoranStats.Locate('Summary','Std dev',[]);
    tSDev := dmCor.CoranStats.Fields[j+1].AsVariant;
    for i := 1 to n do
    begin
      tmp := X[i,j];
      if (tSDev > 0.0) then X[i,j] := (tmp-tMean)/tSDev
                       else X[i,j] := -99.0;
    end;
  end;
  dmCor.cdsCoranChem.First;
  i := 0;
  repeat
    i := i + 1;
    if (i mod 100 = 0) then
    begin
      sbMain.Panels[0].Text := '  i '+IntToStr(i);
      sbMain.Refresh;
    end;
    dmCor.CoranQuantile.Append;
    dmCor.CoranQuantileGROUPNAME.AsString := dmCor.cdsCoranChemGROUPNAME.AsString;
    dmCor.CoranQuantilePlotGroupName.AsString := dmCor.cdsCoranChemPlotGroupName.AsString;
    dmCor.CoranQuantileSAMPLENUM.AsString := dmCor.cdsCoranChemSAMPLENUM.AsString;
    for j := 1 to Nox do
    begin
      dmCor.CoranQuantile.Fields[j+2].AsFloat := X[i,j];
    end;
    dmCor.CoranQuantile.Post;
    dmCor.cdsCoranChem.Next;
  until dmCor.cdsCoranChem.Eof;
  dmCor.CoranQuantile.Close;
  dmCor.CoranQuantile.Open;
  dmCor.CoranQuantile.First;
  dmCor.cdsCoranChem.First;
  dmCor.CoranStats.First;
  for j := 1 to Nox do
  begin
    if (j mod 2 = 0) then
    begin
      sbMain.Panels[0].Text := '  order j '+IntToStr(j);
      sbMain.Refresh;
    end;
    dmCor.CoranQuantile.Close;
    dmCor.CoranQuantile.SQL.Clear;
    dmCor.CoranQuantile.SQL.Add('select * from coranquantile');
    dmCor.CoranQuantile.SQL.Add('order by '+'Param'+IntToStr(j));
    dmCor.CoranQuantile.Open;
    for i := 1 to n do
    begin
      dmCor.CoranQuantile.Edit;
      dmCor.CoranQuantile.Fields[MMaxFields+j+2].AsVariant := DR[i];
      dmCor.CoranQuantile.Post;
      dmCor.CoranQuantile.Next;
    end;
  end;
  dmCor.cdsCoranChem.EnableControls;
  dmCor.CoranStats.EnableControls;
  dmCor.CoranQuantile.EnableControls;
  sbMain.Panels[0].Text := '  ';
  sbMain.Refresh;
end;


procedure TfmCoranMain.CalculateQuantile(iField,N : integer);
var
  i : integer;
  tMean, tSDev,
  tmp : double;
begin
  //get mean and standard deviation
  //sort
  //calculate z-scores
  //allocate data to bins
  //store in datbase table
  //pc1.ActivePage := tsCheck;
  //dmCor.cdsCoranChem.DisableControls;
  //dmCor.CoranStats.DisableControls;
  //dmCor.Quantile1.DisableControls;
  dmCor.CoranStats.Open;
  dmCor.cdsCoranChem.Open;
  dmCor.Quantile1.Open;
  dmCor.Quantile1.First;
  if not(dmCor.Quantile1.Bof and dmCor.Quantile1.Eof) then
  begin
    dmCor.Quantile1.Last;
    repeat
      dmCor.Quantile1.Delete;
      dmCor.Quantile1.Next;
    until dmCor.Quantile1.BOF;
  end;
  dmCor.cdsCoranChem.First;
  i := 0;
  repeat
    dmCor.Quantile1.Append;
    i := i + 1;
    dmCor.Quantile1Seq.AsInteger := i;
    dmCor.Quantile1Param1.AsFloat := dmCor.cdsCoranChem.Fields[iField+2].AsFloat;
    dmCor.Quantile1.Post;
    dmCor.cdsCoranChem.Next;
  until dmCor.cdsCoranChem.Eof;
  dmCor.Quantile1.Close;
  dmCor.Quantile1.Open;
  dmCor.Quantile1.First;
  i := 0;
  repeat
    dmCor.Quantile1.Edit;
    i := i + 1;
    tmp := (1.0*i)/(1.0*N);
    if (i < N) then
    begin
      dmCor.Quantile1Prob1.AsFloat := tmp;
      dmCor.Quantile1Norm1.AsFloat := invCumNorm(tmp);
    end else
    begin
      tmp := 0.9999;
      dmCor.Quantile1Prob1.AsFloat := tmp;
      dmCor.Quantile1Norm1.AsFloat := invCumNorm(tmp);
    end;
    dmCor.Quantile1.Post;
    dmCor.Quantile1.Next;
  until dmCor.Quantile1.Eof;
  dmCor.CoranStats.First;
  dmCor.CoranStats.Locate('Summary','Mean',[]);
  tMean := dmCor.CoranStats.Fields[iField+1].AsVariant;
  dmCor.CoranStats.Locate('Summary','Std dev',[]);
  tSDev := dmCor.CoranStats.Fields[iField+1].AsVariant;
  dmCor.Quantile1.First;
  repeat
    dmCor.Quantile1.Edit;
    tmp := dmCor.Quantile1Param1.AsFloat;
    if (tSDev > 0.0) then dmCor.Quantile1ZScore.AsFloat := (tmp-tMean)/tSDev
                     else dmCor.Quantile1ZScore.AsFloat := -99.0;
    dmCor.Quantile1.Post;
    dmCor.Quantile1.Next;
  until dmCor.Quantile1.Eof;
  dmCor.cdsCoranChem.First;
  dmCor.Quantile1.First;
  dmCor.CoranStats.First;
  dmCor.cdsCoranChem.EnableControls;
  dmCor.CoranStats.EnableControls;
  dmCor.Quantile1.EnableControls;
end;

procedure TfmCoranMain.Button4Click(Sender: TObject);
begin
  dmCor.cdsCoranChem.Open;
  NumSamples := dmCor.cdsCoranChem.RecordCount;
  CalculateQuantile(1, NumSamples);
end;

procedure TfmCoranMain.Button5Click(Sender: TObject);
begin
  //CalculateQuantilesA(Nox);
  GetIniFile;
  ShowMessage(dmCor.CommonFilePath);
  ShowMessage(dmCor.ConnectionString);
  with dmCor do
  begin
    try
      Coran.Connected := false;
    except
    end;
    Coran.ConnectionString := dmCor.ConnectionString;
    ShowMessage(Coran.ConnectionString);
  end;
end;

procedure TfmCoranMain.Calculatequantiles1Click(Sender: TObject);
begin
  CalculateQuantiles1.Checked := not CalculateQuantiles1.Checked;
  MakeQQPlot := CalculateQuantiles1.Checked;
end;

procedure TfmCoranMain.ExportGraph2Click(Sender: TObject);
var
  clTemp : TColor;
begin
  try
    clTemp := (Sender as TDBChart).Color;
    (Sender as TDBChart).Color := clWhite;
  except
    try
      clTemp := (Sender as TChart).Color;
      (Sender as TChart).Color := clWhite;
    except
    end;
  end;
  SaveDialogJPEG.InitialDir := JPEGPath;
  if (Sender = DBChartHist) then
  begin
    //DBChartHist.Color := clWhite;
    SaveDialogJPEG.FileName := 'Data_Hist.jpg';
  end;
  if (Sender = DBChart15) then
  begin
    SaveDialogJPEG.FileName := 'Data_Sequence.jpg';
  end;
  if (Sender = DBChartQuantile) then
  begin
    SaveDialogJPEG.FileName := 'Data_Quantile.jpg';
  end;
  if (Sender = DBChartBox) then
  begin
    SaveDialogJPEG.FileName := 'Data_Box.jpg';
  end;
  if (Sender = DBChartBoxScore) then
  begin
    SaveDialogJPEG.FileName := 'Score_Box.jpg';
  end;
  if (Sender = DBChartHistScore) then
  begin
    SaveDialogJPEG.FileName := 'Score_Hist.jpg';
  end;
  if (Sender = DBChartQuantileScore) then
  begin
    SaveDialogJPEG.FileName := 'Score_Quantile.jpg';
  end;
  if (Sender = DBChart1) then
  begin
    SaveDialogJPEG.FileName := 'Component_1_vs_2.jpg';
  end;
  if (Sender = DBChart2) then
  begin
    SaveDialogJPEG.FileName := 'Component_1_vs_3.jpg';
  end;
  if (Sender = DBChart3) then
  begin
    SaveDialogJPEG.FileName := 'Component_2_vs_3.jpg';
  end;
  if (Sender = ChartEigenValue) then
  begin
    SaveDialogJPEG.FileName := 'Eigenvalues.jpg';
  end;
  if (Sender = DBChart5) then
  begin
    SaveDialogJPEG.FileName := 'VariableLoadings_1.jpg';
  end;
  if (Sender = DBChart6) then
  begin
    SaveDialogJPEG.FileName := 'VariableLoadings_2.jpg';
  end;
  if (Sender = DBChart7) then
  begin
    SaveDialogJPEG.FileName := 'VariableLoadings_3.jpg';
  end;
  if (Sender = DBChart10) then
  begin
    SaveDialogJPEG.FileName := 'VariableLoadings_4.jpg';
  end;
  if (Sender = DBChart4) then
  begin
    SaveDialogJPEG.FileName := 'Localities.jpg';
  end;
  if (Sender = DBChart8) then
  begin
    SaveDialogJPEG.FileName := 'Scores3D_1_vs_2_vs_3.jpg';
  end;
  if (Sender = DBChart9) then
  begin
    SaveDialogJPEG.FileName := 'Localities3D.jpg';
  end;
  if (Sender = DBChart11) then
  begin
    SaveDialogJPEG.FileName := 'Localities4D.jpg';
  end;
  if (Sender = DBChart12) then
  begin
    SaveDialogJPEG.FileName := 'OutlierMap.jpg';
  end;

  if SaveDialogJPEG.Execute then
  begin
    JPEGPath := ExtractFilePath(SaveDialogJPEG.FileName);
    if (Sender = DBChartHist) then
    begin
      //TeeSaveToJPEGFile(DBChartHist,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartHist.Width,DBChartHist.Height);
    end;
    if (Sender = DBChartQuantile) then
    begin
      //TeeSaveToJPEGFile(DBChartQuantile,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartQuantile.Width,DBChartQuantile.Height);
    end;
    if (Sender = DBChartBox) then
    begin
      //TeeSaveToJPEGFile(DBChartBox,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartBox.Width,DBChartBox.Height);
    end;
    if (Sender = DBChartHistScore) then
    begin
      //TeeSaveToJPEGFile(DBChartHistScore,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartHistScore.Width,DBChartHistScore.Height);
    end;
    if (Sender = DBChartQuantileScore) then
    begin
      //TeeSaveToJPEGFile(DBChartQuantileScore,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartQuantileScore.Width,DBChartQuantileScore.Height);
    end;
    if (Sender = DBChartBoxScore) then
    begin
      //TeeSaveToJPEGFile(DBChartBoxScore,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartBoxScore.Width,DBChartBoxScore.Height);
    end;
    if (Sender = DBChart1) then
    begin
      //TeeSaveToJPEGFile(DBChart1,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart1.Width,DBChart1.Height);
    end;
    if (Sender = DBChart2) then
    begin
      //TeeSaveToJPEGFile(DBChart2,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart2.Width,DBChart2.Height);
    end;
    if (Sender = DBChart3) then
    begin
      //TeeSaveToJPEGFile(DBChart3,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart3.Width,DBChart3.Height);
    end;
    if (Sender = ChartEigenValue) then
    begin
      //TeeSaveToJPEGFile(ChartEigenValue,SaveDialogJPEG.FileName,False,jpBestQuality,100,ChartEigenValue.Width,ChartEigenValue.Height);
    end;
    if (Sender = DBChart5) then
    begin
      //TeeSaveToJPEGFile(DBChart5,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart5.Width,DBChart5.Height);
    end;
    if (Sender = DBChart6) then
    begin
      //TeeSaveToJPEGFile(DBChart6,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart6.Width,DBChart6.Height);
    end;
    if (Sender = DBChart7) then
    begin
      //TeeSaveToJPEGFile(DBChart7,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart7.Width,DBChart7.Height);
    end;
    if (Sender = DBChart10) then
    begin
      //TeeSaveToJPEGFile(DBChart10,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart10.Width,DBChart10.Height);
    end;
    if (Sender = DBChart4) then
    begin
      //TeeSaveToJPEGFile(DBChart4,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart4.Width,DBChart4.Height);
    end;
    if (Sender = DBChart8) then
    begin
      //TeeSaveToJPEGFile(DBChart8,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart8.Width,DBChart8.Height);
    end;
    if (Sender = DBChart9) then
    begin
      //TeeSaveToJPEGFile(DBChart9,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart9.Width,DBChart9.Height);
    end;
    if (Sender = DBChart11) then
    begin
      //TeeSaveToJPEGFile(DBChart11,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart11.Width,DBChart11.Height);
    end;
    if (Sender = DBChart12) then
    begin
      //TeeSaveToJPEGFile(DBChart12,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart12.Width,DBChart12.Height);
    end;
  end;
  try
    (Sender as TDBChart).Color := clTemp;
  except
    try
      (Sender as TChart).Color := clTemp;
    except
    end;
  end;
end;

procedure TfmCoranMain.PrintGraph2Click(Sender: TObject);
begin
  if (Sender = DBChart1) then DBChart1.Print;
  if (Sender = DBChart2) then DBChart2.Print;
  if (Sender = DBChart3) then DBChart3.Print;
  if (Sender = ChartEigenValue) then ChartEigenValue.Print;
  if (Sender = DBChart5) then DBChart5.Print;
  if (Sender = DBChart6) then DBChart6.Print;
  if (Sender = DBChart7) then DBChart7.Print;
  if (Sender = DBChart10) then DBChart10.Print;
  if (Sender = DBChart4) then DBChart4.Print;
  if (Sender = DBChart8) then DBChart8.Print;
  if (Sender = DBChart9) then DBChart9.Print;
  if (Sender = DBChart11) then DBChart11.Print;
  if (Sender = DBChart12) then DBChart12.Print;
  if (Sender = DBChartHist) then DBChartHist.Print;
  if (Sender = DBChartQuantile) then DBChartQuantile.Print;
  if (Sender = DBChartBox) then DBChartBox.Print;
  if (Sender = DBChartHistScore) then DBChartHistScore.Print;
  if (Sender = DBChartQuantileScore) then DBChartQuantileScore.Print;
  if (Sender = DBChartBoxScore) then DBChartBoxScore.Print;
end;

procedure TfmCoranMain.TablesDblClick(Sender: TObject);
begin
  if (Sender = DBGridStats) then MessageDlg('Not yet implemented',mtInformation,[mbOK],0);
  if (Sender = DBGrid1) then MessageDlg('Not yet implemented',mtInformation,[mbOK],0);
end;

procedure TfmCoranMain.DBChartHistDblClick(Sender: TObject);
begin
  SaveDialogJPEG.InitialDir := JPEGPath;
  if (Sender = DBChartHist) then
  begin
    SaveDialogJPEG.FileName := 'Data_Hist.jpg';
  end;

  if SaveDialogJPEG.Execute then
  begin
    JPEGPath := ExtractFilePath(SaveDialogJPEG.FileName);
    if (Sender = DBChartHist) then
    begin
      //TeeSaveToJPEGFile(DBChartHist,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartHist.Width,DBChartHist.Height);
    end;
  end;
end;

procedure TfmCoranMain.DBChartQuantileDblClick(Sender: TObject);
begin
  SaveDialogJPEG.InitialDir := JPEGPath;
  if (Sender = DBChartQuantile) then
  begin
    SaveDialogJPEG.FileName := 'Data_Quantile.jpg';
  end;

  if SaveDialogJPEG.Execute then
  begin
    JPEGPath := ExtractFilePath(SaveDialogJPEG.FileName);
    if (Sender = DBChartQuantile) then
    begin
      //TeeSaveToJPEGFile(DBChartQuantile,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartQuantile.Width,DBChartQuantile.Height);
    end;
  end;
end;

procedure TfmCoranMain.DBChartBoxDblClick(Sender: TObject);
begin
  SaveDialogJPEG.InitialDir := JPEGPath;
  if (Sender = DBChartBox) then
  begin
    SaveDialogJPEG.FileName := 'Data_Box.jpg';
  end;

  if SaveDialogJPEG.Execute then
  begin
    JPEGPath := ExtractFilePath(SaveDialogJPEG.FileName);
    if (Sender = DBChartBox) then
    begin
      //TeeSaveToJPEGFile(DBChartBox,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartBox.Width,DBChartBox.Height);
    end;
  end;
end;

procedure TfmCoranMain.ChartEigenvalueDblClick(Sender: TObject);
begin
  SaveDialogJPEG.InitialDir := JPEGPath;
  if (Sender = ChartEigenValue) then
  begin
    SaveDialogJPEG.FileName := 'Eigenvalues.jpg';
  end;

  if SaveDialogJPEG.Execute then
  begin
    JPEGPath := ExtractFilePath(SaveDialogJPEG.FileName);
    if (Sender = ChartEigenValue) then
    begin
      //TeeSaveToJPEGFile(ChartEigenValue,SaveDialogJPEG.FileName,False,jpBestQuality,100,ChartEigenValue.Width,ChartEigenValue.Height);
    end;
  end;
end;

procedure TfmCoranMain.DBChartHistScoreDblClick(Sender: TObject);
begin
  SaveDialogJPEG.InitialDir := JPEGPath;
  if (Sender = DBChartHistScore) then
  begin
    SaveDialogJPEG.FileName := 'Score_Hist.jpg';
  end;

  if SaveDialogJPEG.Execute then
  begin
    JPEGPath := ExtractFilePath(SaveDialogJPEG.FileName);
    if (Sender = DBChartHistScore) then
    begin
      //TeeSaveToJPEGFile(DBChartHistScore,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartHistScore.Width,DBChartHistScore.Height);
    end;
  end;
end;

procedure TfmCoranMain.DBChartQuantileScoreDblClick(Sender: TObject);
begin
  SaveDialogJPEG.InitialDir := JPEGPath;
  if (Sender = DBChartQuantileScore) then
  begin
    SaveDialogJPEG.FileName := 'Score_Quantile.jpg';
  end;

  if SaveDialogJPEG.Execute then
  begin
    JPEGPath := ExtractFilePath(SaveDialogJPEG.FileName);
    if (Sender = DBChartQuantileScore) then
    begin
      //TeeSaveToJPEGFile(DBChartQuantileScore,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartQuantileScore.Width,DBChartQuantileScore.Height);
    end;
  end;
end;

procedure TfmCoranMain.DBChartBoxScoreDblClick(Sender: TObject);
begin
  SaveDialogJPEG.InitialDir := JPEGPath;
  if (Sender = DBChartBoxScore) then
  begin
    SaveDialogJPEG.FileName := 'Score_Box.jpg';
  end;

  if SaveDialogJPEG.Execute then
  begin
    JPEGPath := ExtractFilePath(SaveDialogJPEG.FileName);
    if (Sender = DBChartBoxScore) then
    begin
      //TeeSaveToJPEGFile(DBChartBoxScore,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartBoxScore.Width,DBChartBoxScore.Height);
    end;
  end;
end;

procedure TfmCoranMain.DBChart5DblClick(Sender: TObject);
begin
  SaveDialogJPEG.InitialDir := JPEGPath;
  if (Sender = DBChart5) then
  begin
    SaveDialogJPEG.FileName := 'VariableLoadings_1.jpg';
  end;
  if (Sender = DBChart6) then
  begin
    SaveDialogJPEG.FileName := 'VariableLoadings_2.jpg';
  end;
  if (Sender = DBChart7) then
  begin
    SaveDialogJPEG.FileName := 'VariableLoadings_3.jpg';
  end;
  if (Sender = DBChart10) then
  begin
    SaveDialogJPEG.FileName := 'VariableLoadings_4.jpg';
  end;

  if SaveDialogJPEG.Execute then
  begin
    JPEGPath := ExtractFilePath(SaveDialogJPEG.FileName);
    if (Sender = DBChart5) then
    begin
      //TeeSaveToJPEGFile(DBChart5,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart5.Width,DBChart5.Height);
    end;
    if (Sender = DBChart6) then
    begin
      //TeeSaveToJPEGFile(DBChart6,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart6.Width,DBChart6.Height);
    end;
    if (Sender = DBChart7) then
    begin
      //TeeSaveToJPEGFile(DBChart7,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart7.Width,DBChart7.Height);
    end;
    if (Sender = DBChart10) then
    begin
      //TeeSaveToJPEGFile(DBChart10,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart10.Width,DBChart10.Height);
    end;
  end;
end;

procedure TfmCoranMain.DBChart12DblClick(Sender: TObject);
begin
  SaveDialogJPEG.InitialDir := JPEGPath;
  if (Sender = DBChartHist) then
  begin
    SaveDialogJPEG.FileName := 'Data_Hist.jpg';
  end;
  if (Sender = DBChartQuantile) then
  begin
    SaveDialogJPEG.FileName := 'Data_Quantile.jpg';
  end;
  if (Sender = DBChartBox) then
  begin
    SaveDialogJPEG.FileName := 'Data_Box.jpg';
  end;
  if (Sender = DBChartBoxScore) then
  begin
    SaveDialogJPEG.FileName := 'Score_Box.jpg';
  end;
  if (Sender = DBChartHistScore) then
  begin
    SaveDialogJPEG.FileName := 'Score_Hist.jpg';
  end;
  if (Sender = DBChartQuantileScore) then
  begin
    SaveDialogJPEG.FileName := 'Score_Quantile.jpg';
  end;
  if (Sender = DBChart1) then
  begin
    SaveDialogJPEG.FileName := 'Component_1_vs_2.jpg';
  end;
  if (Sender = DBChart2) then
  begin
    SaveDialogJPEG.FileName := 'Component_1_vs_3.jpg';
  end;
  if (Sender = DBChart3) then
  begin
    SaveDialogJPEG.FileName := 'Component_2_vs_3.jpg';
  end;
  if (Sender = ChartEigenValue) then
  begin
    SaveDialogJPEG.FileName := 'Eigenvalues.jpg';
  end;
  if (Sender = DBChart5) then
  begin
    SaveDialogJPEG.FileName := 'VariableLoadings_1.jpg';
  end;
  if (Sender = DBChart6) then
  begin
    SaveDialogJPEG.FileName := 'VariableLoadings_2.jpg';
  end;
  if (Sender = DBChart7) then
  begin
    SaveDialogJPEG.FileName := 'VariableLoadings_3.jpg';
  end;
  if (Sender = DBChart10) then
  begin
    SaveDialogJPEG.FileName := 'VariableLoadings_4.jpg';
  end;
  if (Sender = DBChart4) then
  begin
    SaveDialogJPEG.FileName := 'Localities.jpg';
  end;
  if (Sender = DBChart8) then
  begin
    SaveDialogJPEG.FileName := 'Scores3D_1_vs_2_vs_3.jpg';
  end;
  if (Sender = DBChart9) then
  begin
    SaveDialogJPEG.FileName := 'Localities3D.jpg';
  end;
  if (Sender = DBChart11) then
  begin
    SaveDialogJPEG.FileName := 'Localities4D.jpg';
  end;
  if (Sender = DBChart12) then
  begin
    SaveDialogJPEG.FileName := 'OutlierMap.jpg';
  end;

  if SaveDialogJPEG.Execute then
  begin
    JPEGPath := ExtractFilePath(SaveDialogJPEG.FileName);
    if (Sender = DBChartHist) then
    begin
      //TeeSaveToJPEGFile(DBChartHist,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartHist.Width,DBChartHist.Height);
    end;
    if (Sender = DBChartQuantile) then
    begin
      //TeeSaveToJPEGFile(DBChartQuantile,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartQuantile.Width,DBChartQuantile.Height);
    end;
    if (Sender = DBChartBox) then
    begin
      //TeeSaveToJPEGFile(DBChartBox,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartBox.Width,DBChartBox.Height);
    end;
    if (Sender = DBChartHistScore) then
    begin
      //TeeSaveToJPEGFile(DBChartHistScore,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartHistScore.Width,DBChartHistScore.Height);
    end;
    if (Sender = DBChartQuantileScore) then
    begin
      //TeeSaveToJPEGFile(DBChartQuantileScore,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartQuantileScore.Width,DBChartQuantileScore.Height);
    end;
    if (Sender = DBChartBox) then
    begin
      //TeeSaveToJPEGFile(DBChartBoxScore,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChartBoxScore.Width,DBChartBoxScore.Height);
    end;
    if (Sender = DBChart1) then
    begin
      //TeeSaveToJPEGFile(DBChart1,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart1.Width,DBChart1.Height);
    end;
    if (Sender = DBChart2) then
    begin
      //TeeSaveToJPEGFile(DBChart2,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart2.Width,DBChart2.Height);
    end;
    if (Sender = DBChart3) then
    begin
      //TeeSaveToJPEGFile(DBChart3,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart3.Width,DBChart3.Height);
    end;
    if (Sender = ChartEigenValue) then
    begin
      //TeeSaveToJPEGFile(ChartEigenValue,SaveDialogJPEG.FileName,False,jpBestQuality,100,ChartEigenValue.Width,ChartEigenValue.Height);
    end;
    if (Sender = DBChart5) then
    begin
      //TeeSaveToJPEGFile(DBChart5,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart5.Width,DBChart5.Height);
    end;
    if (Sender = DBChart6) then
    begin
      //TeeSaveToJPEGFile(DBChart6,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart6.Width,DBChart6.Height);
    end;
    if (Sender = DBChart7) then
    begin
      //TeeSaveToJPEGFile(DBChart7,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart7.Width,DBChart7.Height);
    end;
    if (Sender = DBChart10) then
    begin
      //TeeSaveToJPEGFile(DBChart10,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart10.Width,DBChart10.Height);
    end;
    if (Sender = DBChart4) then
    begin
      //TeeSaveToJPEGFile(DBChart4,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart4.Width,DBChart4.Height);
    end;
    if (Sender = DBChart8) then
    begin
      //TeeSaveToJPEGFile(DBChart8,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart8.Width,DBChart8.Height);
    end;
    if (Sender = DBChart9) then
    begin
      //TeeSaveToJPEGFile(DBChart9,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart9.Width,DBChart9.Height);
    end;
    if (Sender = DBChart11) then
    begin
      //TeeSaveToJPEGFile(DBChart11,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart11.Width,DBChart11.Height);
    end;
    if (Sender = DBChart12) then
    begin
      //TeeSaveToJPEGFile(DBChart12,SaveDialogJPEG.FileName,False,jpBestQuality,100,DBChart12.Width,DBChart12.Height);
    end;
  end;
end;

procedure TfmCoranMain.Include4DVariableData1Click(Sender: TObject);
begin
  Include4DVariableData1.Checked := not Include4DVariableData1.Checked;
  Include4DVarData := Include4DVariableData1.Checked;
end;

procedure TfmCoranMain.ImportDiscriminantFactors1Click(Sender: TObject);
begin
  DiscriminantfactorsImported := false;
  try
    ImportFormDiscrimFac := TfmSheetImportDiscrimFac.Create(Self);
    ImportFormDiscrimFac.OpenDialogSprdSheet.FileName := '';
    ImportFormDiscrimFac.ShowModal;
    DiscriminantFactorsImported := true;
  finally
    ImportFormDiscrimFac.Free;
  end;
  Project1.Enabled := (DiscriminantfactorsImported);
  ProjectCorrespondenceanalysis2.Enabled := false;
  PrincipalComponentAnalysis1.Enabled := false;
  SimultaneousRandQmode1.Enabled := false;
  DiscriminantFunctionAnalysis2.Enabled :=(DiscriminantfactorsImported);
end;

procedure TfmCoranMain.ProjectDiscriminantanalysis2group2Click(Sender: TObject);
var
  i, ii : integer;
begin
  tsEigenValues.TabVisible := true;
  dmCor.EigenVec.Open;
  for ii := 2 to MMaxFields do
  begin
    dbgEigenVec.Columns[ii].Visible := false;
  end;
  ChartEigenValue.Visible := false;
  Panel45.Visible := false;

  dmCor.cdsCoranChem.Open;
  NumSamples := dmCor.cdsCoranChem.RecordCount;
  NumEigenVectorsSelected := 10;
  if (NumEigenVectorsSelected > Nox) then NumEigenVectorsSelected := Nox;
  AssignChartDataSources('Close');
  DisableEnableControls('Disable');
  FillChar(A3,SizeOf(A3),0);
  FillChar(A1,SizeOf(A1),0);
  FillChar(A2,SizeOf(A2),0);
  dmCor.EigenVec.Open;
  dmCor.EigenVec.First;
  for i := 1 to Nox do
  begin
    A2[i,i] := dmCor.EigenVec.Fields[1].AsFloat;
    dmCor.EigenVec.Next;
  end;
  dmCor.EigenVec.First;
  if (Sender = ProjectDiscriminantanalysis2group2) then scSimilarityChoice := scDiscrim2Grp;
  try
    dmCor.QGroups.Open;
    dmCor.QPlotGroups.Open;
  except
  end;
  dmCor.EigenVec.Open;
  if (scSimilarityChoice in [scDiscrim2Grp]) then DiscrimProject;
  dmCor.cdsFacLoadingsSmp.First;
  dmCor.cdsCoranChem.First;
  PrepareComboBoxes(NumEigenVectorsSelected);
  CalculateHingeStatistics(NumEigenVectorsSelected);
  PrepareGraphs;
  tsResults.TabVisible := true;
  tsSimilarity.TabVisible := false;
  tsEigenValues.TabVisible := true;
  tsLoadings.TabVisible := true;
  tsScores.TabVisible := true;
  tsOutlierMap.TabVisible := false;
  tsGraph1.TabVisible := true;
  tsGraph2.TabVisible := false;
  tsGraph3.TabVisible := false;
  ts3D.TabVisible := true;
  tsLoc3D.TabVisible := true;
  tsLoc4D.TabVisible := Include4DVarData;
  AssignChartDataSources('Open');
  DisableEnableControls('Enable');
  sbMain.Panels[1].Text := 'Completed';
  sbMain.Refresh;


end;

procedure TfmCoranMain.bDeleteClick(Sender: TObject);
var
  i : integer;
begin
  with dmCor do
  begin
    cdsqDim4Smp.Open;
    try
      DeleteDim4Smp.SQL.Clear;
      DeleteDim4Smp.SQL.Add('delete * from Dimension4Smp');
      DeleteDim4Smp.SQL.Add('where Dimension4Smp.VariableID like '+''''+'Vector%'+'''');
      DeleteDim4Smp.ExecSQL;
    except
      ShowMessage('delete error 1');
    end;
    try
      DeleteDim4Smp.SQL.Clear;
      DeleteDim4Smp.SQL.Add('delete * from Dimension4Smp');
      DeleteDim4Smp.SQL.Add('where Dimension4Smp.VariableID like '+''''+'Score%'+'''');
      DeleteDim4Smp.ExecSQL;
    except
      ShowMessage('delete error 2');
    end;
    {
    for i := 1 to Nox do
    begin
      try
        DeleteDim4Smp.SQL.Clear;
        DeleteDim4Smp.SQL.Add('delete * from Dimension4Smp');
        DeleteDim4Smp.SQL.Add('where Dimension4Smp.VariableID = '+''''+'Vector'+IntToStr(i)+'''');
        DeleteDim4Smp.ExecSQL;
      except
        ShowMessage('delete error 3');
      end;
    end;
    try
      DeleteDim4Smp.SQL.Clear;
      DeleteDim4Smp.SQL.Add('delete * from Dimension4Smp');
      DeleteDim4Smp.SQL.Add('where Dimension4Smp.VariableID = '+''''+'ScoreDistance'+'''');
      DeleteDim4Smp.ExecSQL;
    except
      ShowMessage('delete error 4');
    end;
    }
    cdsqDim4Smp.Close;
    cdsqDim4Smp.Open;
  end;
end;

procedure TfmCoranMain.bDim4SmpClick(Sender: TObject);
begin
  dmCor.cdsqDim4Smp.Open;
end;

procedure TfmCoranMain.Transformeddata1Click(Sender: TObject);
var
  fr : TFlexcelReport;
  TemplateNameStr : string;
begin
  //dmFlex.FlexCoranChem;
  try
    fr := TFlexcelReport.Create(true);
    fr.AddTable('FlexCoranChem', dmFlex.CoranChem);
    TemplateNamestr := FlexTemplatePath+'Data_Transformed.xlsx';
    //dmFlex.FlexCoranChem.Template := FlexTemplatePath+'Data_Transformed.xls';
    dmFlex.SaveDialogFlex.InitialDir := ExportPath;
    dmFlex.SaveDialogFlex.FileName := 'Data_Transformed_1.xlsx';
    if dmFlex.SaveDialogFlex.Execute then
    begin
      ExportPath := ExtractFilePath(dmFlex.SaveDialogFlex.FileName);
      //dmFlex.FlexCoranChem.FileName := dmFlex.SaveDialogFlex.FileName;
      try
        // delete the file if it already exists
        DeleteFile(dmFlex.SaveDialogFlex.FileName);
      except
      end;
      //dmFlex.FlexCoranChem.Run;
    end;
    finally
      fr.Free;
    end;
end;

procedure TfmCoranMain.Summarydata1Click(Sender: TObject);
var
  fr : TFlexcelReport;
  TemplateNameStr : string;
begin
  try
    fr := TFlexcelReport.Create(true);
    fr.AddTable('FlexCoranStats', dmFlex.CoranStats);
    TemplateNamestr := FlexTemplatePath+'Data_Summary.xlsx';
  //dmFlex.FlexCoranStats.Template := FlexTemplatePath+'Data_Summary.xls';
    dmFlex.SaveDialogFlex.InitialDir := ExportPath;
    dmFlex.SaveDialogFlex.FileName := 'Data_Summary_1.xlsx';
    if dmFlex.SaveDialogFlex.Execute then
    begin
      ExportPath := ExtractFilePath(dmFlex.SaveDialogFlex.FileName);
      //dmFlex.FlexCoranStats.FileName := dmFlex.SaveDialogFlex.FileName;
      try
        // delete the file if it already exists
        DeleteFile(dmFlex.SaveDialogFlex.FileName);
      except
      end;
      //dmFlex.FlexCoranStats.Run;
    end;
  finally
    fr.Free;
  end;
end;

procedure TfmCoranMain.N4Ddata1Click(Sender: TObject);
var
  fr : TFlexcelReport;
  TemplateNameStr : string;
begin
  try
    fr := TFlexcelReport.Create(true);
    fr.AddTable('FlexDim4SmpNum', dmFlex.qDim4SmpNum);
    TemplateNamestr := FlexTemplatePath+'4D.xlsx';
    //dmFlex.FlexDim4SmpNum.Template := FlexTemplatePath+'4D.xls';
    dmFlex.SaveDialogFlex.InitialDir := ExportPath;
    dmFlex.SaveDialogFlex.FileName := 'Data_4D_1.xlsx';
    if dmFlex.SaveDialogFlex.Execute then
    begin
      ExportPath := ExtractFilePath(dmFlex.SaveDialogFlex.FileName);
      //dmFlex.FlexDim4SmpNum.FileName := dmFlex.SaveDialogFlex.FileName;
      try
        // delete the file if it already exists
        DeleteFile(dmFlex.SaveDialogFlex.FileName);
      except
      end;
      //dmFlex.FlexDim4SmpNum.Run;
    end;
  finally
    fr.Free;
  end;
end;

procedure TfmCoranMain.Samplescores1Click(Sender: TObject);
var
  fr : TFlexcelReport;
  TemplateNameStr : string;
begin
  try
    fr := TFlexcelReport.Create(true);
    fr.AddTable('FlexfacLoadingsSmp', dmFlex.FacLoadingsSmp);
    TemplateNamestr := FlexTemplatePath+'Scores.xlsx';
    //dmFlex.FlexfacLoadingsSmp.Template := FlexTemplatePath+'Scores.xls';
    dmFlex.SaveDialogFlex.InitialDir := ExportPath;
    dmFlex.SaveDialogFlex.FileName := 'Scores_1.xlsx';
    if dmFlex.SaveDialogFlex.Execute then
    begin
      ExportPath := ExtractFilePath(dmFlex.SaveDialogFlex.FileName);
      //dmFlex.FlexfacLoadingsSmp.FileName := dmFlex.SaveDialogFlex.FileName;
      try
        // delete the file if it already exists
        DeleteFile(dmFlex.SaveDialogFlex.FileName);
      except
      end;
      //dmFlex.FlexfacLoadingsSmp.Run;
    end;
  finally
    fr.Free;
  end;
end;

procedure TfmCoranMain.Eigenvalues1Click(Sender: TObject);
var
  fr : TFlexcelReport;
  TemplateNameStr : string;
begin
  try
    fr := TFlexcelReport.Create(true);
    fr.AddTable('FlexEigenVal', dmFlex.EigenVal);
    TemplateNamestr := FlexTemplatePath+'Eigenvalues.xlsx';
    //dmFlex.FlexEigenVal.Template := FlexTemplatePath+'Eigenvalues.xls';
    dmFlex.SaveDialogFlex.InitialDir := ExportPath;
    dmFlex.SaveDialogFlex.FileName := 'Eigenvalues_1.xlsx';
    if dmFlex.SaveDialogFlex.Execute then
    begin
      ExportPath := ExtractFilePath(dmFlex.SaveDialogFlex.FileName);
      //dmFlex.FlexEigenVal.FileName := dmFlex.SaveDialogFlex.FileName;
      try
        // delete the file if it already exists
        DeleteFile(dmFlex.SaveDialogFlex.FileName);
      except
      end;
      //dmFlex.FlexEigenVal.Run;
      fr.Run(TemplateNameStr,TPath.Combine(ExportPath,dmFlex.SaveDialogFlex.FileName));
    end;
  finally
    fr.Free;
  end;
end;

procedure TfmCoranMain.Eigenvectors1Click(Sender: TObject);
var
  fr : TFlexcelReport;
  TemplateNameStr : string;
begin
  try
    fr := TFlexcelReport.Create(true);
    fr.AddTable('FlexEigenVec', dmFlex.EigenVec);
    TemplateNamestr := FlexTemplatePath+'Eigenvectors.xlsx';
    //dmFlex.FlexEigenVec.Template := FlexTemplatePath+'Eigenvectors.xls';
    dmFlex.SaveDialogFlex.InitialDir := ExportPath;
    dmFlex.SaveDialogFlex.FileName := 'Eigenvectors_1.xlsx';
    if dmFlex.SaveDialogFlex.Execute then
    begin
      ExportPath := ExtractFilePath(dmFlex.SaveDialogFlex.FileName);
      //dmFlex.FlexEigenVec.FileName := dmFlex.SaveDialogFlex.FileName;
      try
        // delete the file if it already exists
        DeleteFile(dmFlex.SaveDialogFlex.FileName);
      except
      end;
      //dmFlex.FlexEigenVec.Run;
    end;
  finally
    fr.Free;
  end;
end;

procedure TfmCoranMain.Similaritymatrix1Click(Sender: TObject);
var
  fr : TFlexcelReport;
  TemplateNameStr : string;
begin
  try
    fr := TFlexcelReport.Create(true);
    fr.AddTable('FlexCoranSimilarity', dmFlex.CoranSimilarity);
    TemplateNamestr := FlexTemplatePath+'Similarity.xlsx';
    //dmFlex.FlexCoranSimilarity.Template := FlexTemplatePath+'Similarity.xls';
    dmFlex.SaveDialogFlex.InitialDir := ExportPath;
    dmFlex.SaveDialogFlex.FileName := 'Similarity_1.xlsx';
    if dmFlex.SaveDialogFlex.Execute then
    begin
      ExportPath := ExtractFilePath(dmFlex.SaveDialogFlex.FileName);
      //dmFlex.FlexCoranSimilarity.FileName := dmFlex.SaveDialogFlex.FileName;
      try
        // delete the file if it already exists
        DeleteFile(dmFlex.SaveDialogFlex.FileName);
      except
      end;
      //dmFlex.FlexCoranSimilarity.Run;
    end;
  finally
    fr.Free;
  end;
end;

procedure TfmCoranMain.VariableLoadings1Click(Sender: TObject);
var
  fr : TFlexcelReport;
  TemplateNameStr : string;
begin
  try
    fr := TFlexcelReport.Create(true);
    fr.AddTable('FlexFacLoadingsVarNoZero', dmFlex.FacLoadingsVarNoZero);
    TemplateNamestr := FlexTemplatePath+'Eigenvalues.xlsx';
    //dmFlex.FlexFacLoadingsVarNoZero.Template := FlexTemplatePath+'Loadings.xls';
    dmFlex.SaveDialogFlex.InitialDir := ExportPath;
    dmFlex.SaveDialogFlex.FileName := 'Loadings_1.xlsx';
    if dmFlex.SaveDialogFlex.Execute then
    begin
      ExportPath := ExtractFilePath(dmFlex.SaveDialogFlex.FileName);
      //dmFlex.FlexFacLoadingsVarNoZero.FileName := dmFlex.SaveDialogFlex.FileName;
      try
        // delete the file if it already exists
        DeleteFile(dmFlex.SaveDialogFlex.FileName);
      except
      end;
      //dmFlex.FlexFacLoadingsVarNoZero.Run;
    end;
  finally
    fr.Free;
  end;
end;

procedure TfmCoranMain.Discretisedata1Click(Sender: TObject);
var
  NumAttributes : integer;
  fr : TFlexcelReport;
  TemplateNameStr : string;
begin
  try
    fr := TFlexcelReport.Create(true);
    fr.AddTable('FlexDim4', dmFlex.qDim4);
    NumAttributes := dmCor.ElemNames.RecordCount;
    Discretise(NumAttributes);
    TemplateNamestr := FlexTemplatePath+'Data_transformed_discretised.xlsx';
    //dmFlex.FlexDim4.Template := FlexTemplatePath+'Data_transformed_discretised.xls';
    dmFlex.SaveDialogFlex.InitialDir := ExportPath;
    dmFlex.SaveDialogFlex.FileName := 'Data_transformed_discretised_1.xlsx';
    if dmFlex.SaveDialogFlex.Execute then
    begin
      ExportPath := ExtractFilePath(dmFlex.SaveDialogFlex.FileName);
      //dmFlex.FlexDim4.FileName := dmFlex.SaveDialogFlex.FileName;
      try
        // delete the file if it already exists
        DeleteFile(dmFlex.SaveDialogFlex.FileName);
      except
      end;
      //dmFlex.FlexDim4.Run;
    end;
  finally
    fr.Free;
  end;
end;

procedure TfmCoranMain.ImportMeans1Click(Sender: TObject);
begin
  MeanValuesImported := false;
  try
    ImportFormMean := TfmSheetImportMean.Create(Self);
    ImportFormMean.OpenDialogSprdSheet.FileName := '';
    ImportFormMean.ShowModal;
    MeanValuesImported := true;
  finally
    ImportFormMean.Free;
  end;
  Project1.Enabled := (EigenValuesImported and EigenVectorsImported and MeanvaluesImported);
  ProjectCorrespondenceanalysis2.Enabled := (EigenValuesImported and EigenVectorsImported and MeanValuesImported);
  PrincipalComponentAnalysis1.Enabled := (EigenValuesImported and EigenVectorsImported and MeanValuesImported);
  SimultaneousRandQmode1.Enabled :=(EigenValuesImported and EigenVectorsImported and MeanValuesImported);
  DiscriminantFunctionAnalysis2.Enabled := false;
  sbMain.Panels[1].Text := 'Location values imported';
  sbMain.Refresh;
end;


end.


