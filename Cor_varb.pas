unit Cor_varb;

interface

uses
  SysUtils, AllSorts, NumRecipes_varb;

const
  NN          = NMax;
  MM          = MMax;
  MMaxFields  = 30; // maximum number of fields in database, not in arrays
  CoranVersion = '3.0.0';
  zero         = zerovalue;
  DefaultZeroLimit = 1.0e-14;
  Cutoff = 0.975;
  InvChiSqValues : array[1..70] of double = (
    5.0239,7.3778,9.3484,11.1433,12.8325,14.4494,16.0128,17.5345,
   19.0228,20.4832,21.9200,23.3367,24.7356,26.1189,27.4884,28.8454,
   30.1910,31.5264,32.8523,34.1696,35.4789,36.7807,38.0756,39.3641,
   40.6465,41.9232,43.1945,44.4608,45.7223,46.9792,48.2319,49.4804,
   50.7251,51.9660,53.2033,54.4373,55.6680,56.8955,58.1201,59.3417,
   60.5606,61.7768,62.9904,64.2015,65.4102,66.6165,67.8206,
   69.0226,70.2224,71.4202,72.6160,73.8099,75.0019,76.1920,
   77.3805,78.5672,79.7522,80.9356,82.1174,83.2977,84.4764,
   85.6537,86.8296,88.0041,89.1771,90.3489,91.5194,92.6885,
   93.8565,95.0232);

type
  TSimilarityChoices = (scCorrespondence=1,scRQVariance=2,scRQPearsonR=3,
                   scPCAVariance=4,scPCAPearsonR=5,scDiscrim2Grp=6,
                   scDiscrimnGrp=7,scCluster=8,scPCASpearmanR=9,
                   scPCAKendallR=10,scRQSpearmanR=11,scRQKendallR=12,
                   scRobPCA=13,scNone=100);
  TTabSheetPageIndices = (piControl,piSummary,piSimilarity,piResults,
                   piEigenValues,piLoadings,piScores,piOutlierMap,
                   piGraph1,piGraph2,piGraph3,piSpreadSheets,
                   piLocalities,pi3D,piLoc3D,piLoc4D,piCheck);
  RealArrayX  = array[1..NN,1..MM] of double;
  RealArrayXP = array[1..MM,1..NN] of double;
  RealArrayC  = array[1..MM,1..MM] of double;
  RealArrayR  = narray;
  SingleArrayM  = array[1..MM] of single;
  RealArrayM  = array[1..MM] of double;
  IntArray    = array[1..MM] of integer;
  StringArray8 = array[1..MM,1..MM+3] of string[12];
  String1Array10  = array[1..MM] of string[1];
  String10Array10  = array[1..MM] of string[10];
  String15Array100  = array[1..NN] of string[15];

var
  scSimilarityChoice : TSimilarityChoices;
  VarbSymbol                  : array[0..MM] of shortint;
  ComponentStr   : String15Array100;
  EigName     : String10Array10;
  Filename    : string[8];
  Title       : string[40];
  tempstr     : string[10];
  //N, M        : integer;
  Total       : double;
  X1, X, B, W,
  TempM       : RealArrayX;
  WP          : RealArrayXP;
  DR, DR2     : RealArrayR;
  DC, A1, A2,
  A3, TempC   : RealArrayC;
  DD          : RealArrayM;
  CA, CRR     : array[1..5] of double;
  Ndig        : IntArray;
  C, D,
  SumE, SumEE : double;
  I, J, K, KP : integer;
  Xmin, Xmax,
  Ymin, Ymax,
  Zmin, Zmax  : double;
  NDG, NX,
  NY, NZ      : integer;
  Nox : integer;
  TooMuch : boolean;
  IPrn : string;
  OxideName : String10Array10;
  TakeLogs  : String1Array10;
  AN1,AN2,AN3,AM : double;
  ND1,ND2 : integer;
  R0,R1,R2,D2,F,E : double;
  ResultsArray : StringArray8;
  DefaultMinimum : double;
  DefaultRotation, DefaultElevation,
  DefaultPerspective,
  DefaultZoom : integer;
  //NumAttributes : integer;
  BottomAxisMin, BottomAxisMax,
  LeftAxisMin, LeftAxisMax,
  DepthTopAxisMin, DepthTopAxisMax : extended;
  NumEigenVectorsSelected : integer;
  NumSamples : integer;
  TotalsOK : boolean;
  MakeQQPlot : boolean;
  Include4DVarData : boolean;

var
   done                  : boolean;
   Lst                   : TextFile;
   AnyKey                : char;
   Toggle100             : byte;
   FilePrepared          : boolean;
   ElementPos            : array [1..MM+6] of integer;
   FullFileName         : string;
   FlexTemplatePath,
   JPEGPath,
   ExportPath,
   ADODataLinkFile, DataPath   : string;
   TotalRecs                   : Integer;
   RowCount             : array[1..10] of integer;
   ImportSheetNumber,
   WSumFacCol,
   ImportSpecNameCol,
   PositionCol, CalledCol,
   ColumnCol,
   TakeLogCol           : integer;
   WSumFacColStr,
   ImportSpecNameColStr,
   PositionColStr, CalledColStr,
   ColumnColStr,
   TakeLogColStr        : string;
   FromRowValueString, ToRowValueString : string;
   FromRow, ToRow : integer;
   EigenValuesImported, EigenVectorsImported,
   DiscriminantFactorsImported, MeanValuesImported : boolean;

  function ConvertCol2Int(AnyString : string) : integer;

implementation


function ConvertCol2Int(AnyString : string) : integer;
var
  itmp    : integer;
  tmpStr  : string;
  tmpChar : char;
begin
    AnyString := UpperCase(AnyString);
    tmpStr := AnyString;
    ClearNull(tmpStr);
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


end.