unit Cor_varb;
interface
uses
  kem_varb, getfield;

const
  NN          = 100;
  MM          =  10;

type
  str1            = string[1];
  Str80       = string[80];

type
  RealArrayX  = array[1..NN,1..MM] of real;
  RealArrayXP = array[1..MM,1..NN] of real;
  RealArrayC  = array[1..MM,1..MM] of real;
  RealArrayR  = array[1..NN] of real;
  IntArray    = array[1..9] of integer;
  String10Array100  = array[1..100] of string[10];
  PlotValRecType    = record
       XCap,YCap   : string[20];
       XMini,XMaxi,YMini,YMaxi,
       XTics,YTics : real;
       XLog,YLog,
       XNum,YNum   : byte;
       XLab,YLab   : byte;
                      end;

var
  SmpSymbol                   : array[0..NN] of integer;
  VarbSymbol                  : array[0..MM] of integer;
  Component   : String10Array100;
  EigName     : String10Array10;
  Filename    : string[8];
  Title       : string[40];
  tempstr     : string[10];
  N, M        : integer;
  Total       : real;
  X, B, W,
  TempM       : RealArrayX;
  WP          : RealArrayXP;
  DR          : RealArrayR;
  DC, A1, A2,
  A3, TempC   : RealArrayC;
  CA, CRR     : array[1..5] of real;
  Ndig        : IntArray;
  C, D,
  SumE, SumEE : real;
  I, J, K, KP : integer;
  Xmin, Xmax,
  Ymin, Ymax,
  Zmin, Zmax  : real;
  NDG, NX,
  NY, NZ      : integer;
  SimilarityChoice : integer;
  PlotValRec  : array[1..4] of PlotValRecType;

implementation

end.