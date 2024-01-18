unit Cor_Grfr;
interface

uses
printer,
    sys_def,
    allsorts, mathproc,
    getfield, kem_varb,
    cor_varb;
var
   SessionFile : text;


procedure OpenSessionFile(Snum : byte; var Code:integer);
procedure Make_GRF(CodeForSession,PltNum : byte;
                   GrafType    :  char;
                   X_AxisName  :  string;
                   XAxisType,XFormat : byte;
                   LowX,HighX,TicEveryX : real;
                   LabelEveryX : byte;
                   Y_AxisName  : string;
                   YAxisType,YFormat : byte;
                   LowY,HighY,TicEveryY : real;
                   LabelEveryY : byte;
                   var Check : integer);
procedure DumpToGrapher(PltType : char; PltNum : byte;
                         var Check : integer);


implementation

procedure OpenSessionFile(Snum : byte; var Code:integer);
begin
  Assign(SessionFile,Drive3+'\SESSION'+chr(64+Snum)+'.LST');
{$I-}
  Rewrite(SessionFile); IOCheck;
  if IOErr then Code:=IOVal
           else
           begin
             Code:=0;
             Writeln(SessionFile,'Plotting session : ',Chr(64+Snum));
             Writeln(SessionFile,'===========================');
           end;
{$I+}
end;

procedure ParseStr(S1 : string; var S2 : string);
type
    OrdSet = set of byte;
const
LinkSet : array[1..15] of byte=
    (97,98,67,112,82,114,108,115,85,72,88,100,0,0,101);
var
   GreekSet : OrdSet;
   i,j,L : byte;
   anychr : char;
begin
GreekSet:=[224..235,238];
 L:=Length(S1);
 S1:=S1+'    ';
 S2:=CharStream(80,32);
 { check for superscripts }
 i:=0;j:=0;
 while (i<=L) do
 begin
  i:=i+1;
  if not (ord(S1[i]) in [48..57]) then
     begin
      if (ord(S1[i]) in GreekSet) then
         begin
          j:=j+1; S2[j]:=chr(159);
          j:=j+1; S2[j]:=chr(LinkSet[ord(S1[i])-223]);
          j:=j+1; S2[j]:=chr(159);
         end else
          begin
           j:=j+1; S2[j]:=S1[i];
         end;
    end else
     begin
      if (ord(S1[i+1]) in [48..57]) then
         begin
          j:=j+1;S2[j]:=chr($18);
          j:=j+1; S2[j]:=S1[i];
          while (ord(S1[i+1]) in [48..57]) do
          begin
           j:=j+1; S2[j]:=S1[i+1];
           i:=i+1;
          end;
          j:=j+1; S2[j]:=chr($F7);
         end else
         begin
          j:=j+1; S2[j]:=chr($19);
          j:=j+1; S2[j]:=S1[i];
          j:=j+1; S2[j]:=chr($F7);
         end;
     end;
 end;
 S2[0]:=chr(j);
end;

procedure Make_GRF(CodeForSession,PltNum : byte;
                   GrafType    :  char;
                   X_AxisName  :  string;
                   XAxisType,XFormat : byte;
                   LowX,HighX,TicEveryX : real;
                   LabelEveryX : byte;
                   Y_AxisName  : string;
                   YAxisType,YFormat : byte;
                   LowY,HighY,TicEveryY : real;
                   LabelEveryY : byte;
                   var Check : integer);
const
 {Data-related constants}
     PltType : array[1..2] of string[4] = ('arith','log ');
     DatFileName : string[5] = 'CORG_';
     NamFileName : string[5] = 'GRFN_';
     Row_1 = '66 67 48 3 "NO"  48';      {data columns...}
     Row_2 = '"NO" "SOLID"  1.500e-001  1'; {LINE CONNECTING PTS?}
     Row_2a = '"YES" "SOLID"  1.500e-001  2'; {LINE CONNECTING PTS?}
     Row_3 = '"NO"  33   1.500e-001  1  1'; {centred symbols?}
     Row_3a = '"YES"  39   1.500e-001  2  1'; {centred symbols?}
     Row_4 = '65 0.000e+000 0.000e+000  0.000e+000 "GEOPLOT"  2.000e-001 1';
     Row_4a = '65 0.200e+000 0.000e+000  0.000e+000 "SET1"  2.000e-001 2';
     Row_5 = '"SOLID" 0 1.500e-001 9.900e+029 9.900e+029 200 2.000e+000 1';
     Row_6 = '9.900e+028 9.900e+028 9.900e+028 9.900e+028 1.500e-001';
     Row_6a= '9.900e+029 9.900e+029 9.900e+029 9.900e+029 1.500e-001';

   Top : array[1..8] of string =
('TOP',
'1.500e+000 6.500e+000 6.000e+000 88',
'9.900e+028 9.900e+028 9.900e+028 1 1',
'0.000e+000 0.000e+000 1.500e-001 1 1',
'10 2 1',
'1 9.900e+028 0.000e+000 9.900e+028 2.500e-001',
'"DEFAULT" "DEFAULT" ""',
'4.000e-002');

   Right : array[1..8] of string =
('RIGHT',
'7.500e+000 1.500e+000 5.000e+000 89',
'9.900e+028 9.900e+028 9.900e+028 1 1',
'0.000e+000 0.000e+000 1.500e-001 1 1',
'10 2 1',
'1 9.900e+028 0.000e+000 9.900e+028 2.500e-001',
'"DEFAULT" "DEFAULT" ""',
'4.000e-002');

 {Axis-related constants}
     Axis_FontsStr = '"SET1" "SET4"';
     TicDistanceStr = '5.000e-002';
     TicHeightStr = '1.000e-001';

  {1."X"}
     StartX = '1.500e+000 1.500e+000 6.000e+000 88';
     X_AngleStr = '0.000e+000';

  {2."Y"}
     StartY = '1.500e+000 1.500e+000 5.000e+000 89';
     Y_AngleStr = '2.700e+002';
var
   GRF_File : text;
   AnyStr,AxisName_X,AxisName_Y,GDatFileName,GrafFileName : string;
   W,D,SpidNum : byte;
   NumTics : integer;
   GAxisType : string;
   Fx : word;
   YDat,YVal : real;

   procedure FileCheck;
   begin
     IOCheck;  Check:=IOVal;
   end;

   function Okay : integer;
   begin
     if Check<1 then
     begin
       FileCheck; Okay:=IOVal;
     end else OKay:=0;
   end;

begin
  Check:=0;
  if GrafType='C' then
    GrafFileName:=Drive3+'\'+DatFileName + chr(CodeForSession+64)+'_C.GRF'
  else
    GrafFileName:=Drive3+'\'+DatFileName + chr(CodeForSession+64)+Int2Str(PltNum)+'.GRF';
  ClearNull(GrafFileName);
  {$I-}
  Assign(GRF_File,GrafFileName); Rewrite(GRF_File); FileCheck;
  if Check>0 then Exit;
  {
  if GrafType='C' then
    Writeln(SessionFile,'Caption file written')
  else
    Writeln(SessionFile,'Plot #',PltNum,':  ',
         PltType[XAxisType],' [X]: ',X_AxisName,
         ' - vs -  ',PltType[YAxisType],' [Y]: ',Y_AxisName);
  }
  IOCheck;
  {-------------------writing GRF file----------------}
  {0. Header info}
  if Okay<1 then Writeln(GRF_File,'1243');
  if Okay<1 then
  begin
     case GrafType of
        '1','2': begin
           case SimilarityChoice of
              1..3 : Writeln(GRF_File,'2 4 1 0 0');
              4..5 : Writeln(GRF_File,'1 4 1 0 0');
           end;
        end;
        'C' : Writeln(GRF_File,'2 2 1 0 0');
     end;
  end;
  {1.Data file info}
  if GrafType='C' then
  begin
    GDatFileName:= DatFileName + chr(CodeForSession+64)+'_S';
  end else
  begin
    case SimilarityChoice of
      1..3 : begin
         GDatFileName:= DatFileName + chr(SimilarityChoice+64)+'V'+Int2Str(PltNum);
      end;
      4..5 : begin
         GDatFileName:= DatFileName + chr(CodeForSession+64)+Int2Str(PltNum);
      end;
    end;
  end;
  ClearNull(GDatFileName);
  if Okay<1 then Writeln(GRF_File,GDatFileName);
  if Okay<1 then Writeln(GRF_File,Row_1);
  case SimilarityChoice of
    1..3 : begin
       if Okay<1 then Writeln(GRF_File,Row_2a);
    end;
    4..5 : begin
       if Okay<1 then Writeln(GRF_File,Row_2);
    end;
  end;
  if Okay<1 then Writeln(GRF_File,Row_3);
  if Okay<1 then Writeln(GRF_File,Row_4);
  if Okay<1 then Writeln(GRF_File,Row_5);
  if Okay<1 then Writeln(GRF_File,Row_6);
  if GrafType='C' then
  begin
     {2. Group names "data"}
     GDatFileName:= DatFileName
                 +chr(CodeForSession+64)+'_N';
     ClearNull(GDatFileName);
     if Okay<1 then Writeln(GRF_File,GDatFileName);
     if Okay<1 then Writeln(GRF_File,Row_1);
     if Okay<1 then Writeln(GRF_File,Row_2);
     if Okay<1 then Writeln(GRF_File,Row_3);
     if Okay<1 then Writeln(GRF_File,
        '65   0.000e+000  -1.000e-001  0.000e+000  "SET4"  2.000e-001  1');
     if Okay<1 then Writeln(GRF_File,Row_5);
     if Okay<1 then Writeln(GRF_File,Row_6a);
  end;
  if GrafType<>'C' then
  begin
    if SimilarityChoice in [1..3] then
    begin
      GDatFileName:= DatFileName + chr(CodeForSession+64)+Int2Str(PltNum);
      ClearNull(GDatFileName);
      if Okay<1 then Writeln(GRF_File,GDatFileName);
      if Okay<1 then Writeln(GRF_File,Row_1);
      if Okay<1 then Writeln(GRF_File,Row_2);
      if Okay<1 then Writeln(GRF_File,Row_3);
      if Okay<1 then Writeln(GRF_File,Row_4);
      if Okay<1 then Writeln(GRF_File,Row_5);
      if Okay<1 then Writeln(GRF_File,Row_6);
    end;
  end;
  {Axis info : X}
  if XAxisType=2  then GAxisType:=' 2 1'
                  else GAxisType:=' 1 1';
  AxisName_X:='X_'+chr(CodeForSession+64)+Int2Str(PltNum);
  ClearNull(AxisName_X);
  if Okay<1 then Writeln(GRF_File,AxisName_X);
  if Okay<1 then
  begin
    if GrafType='C' then
       Writeln(GRF_File,'-1.000e+000 -1.000e+000 9.500e+000 88')
    else  Writeln(GRF_File,StartX);
  end;
  W:=(XFormat div 10);
  D:=(XFormat mod 10);
  if Okay<1 then Writeln(GRF_File,
           Real2Str(LowX,W+3,D+2),'  ',
           Real2Str(HighX,W+3,D+2),' 9.900e+028 ',GAxisType);
  if Okay<1 then
  begin
    if XAxisType=1 then
         Writeln(GRF_File, '0.000e+000  ',Real2Str(TicEveryX,W+3,D+2),
                       ' 1.000e-001  1  1')
    else
         Writeln(GRF_File, '0.000e+000  ',' 9.900e+028  ',
                       ' 1.500e-001  1  1');
  end;
  if Okay<1 then
  begin
    if XAxisType=2 then
        Writeln(GRF_File,' 5  ',Int2Str(D),'  2')
    else
        Writeln(GRF_File,Int2Str(LabelEveryX),'  ',Int2Str(D),'  1');
  end;
  if Okay<1 then
  begin
    if (GrafType in ['1','2']) then
       Writeln(GRF_File,'1 9.900e+028 0.000e+000 9.900e+028 2.500e-001')
    else
       Writeln(GRF_File,'1 3.000e+000 4.800e+000 9.900e+028 2.500e-001');
  end;
  if Okay<1 then
  begin
    if GrafType in ['1','2'] then
          Writeln(GRF_File,'"SET1"  "SET4"  "',PlotValRec[KP-1].XCap,'"')
    else
          Writeln(GRF_File,'"SET1"  "SET4"  "',X_AxisName,'"');
  end;
  if Okay<1 then Writeln(GRF_File,'5.000e-002');
  {Axis info : Y}
  AxisName_Y:='Y_'+chr(CodeForSession+64)+Int2Str(PltNum);
  ClearNull(AxisName_Y);
  if Okay<1 then Writeln(GRF_File,AxisName_Y);
  if Okay<1 then
  begin
    if GrafType='C' then
       Writeln(GRF_File,'-1.000e+000 -1.000e+000 1.300e+001 89')
    else  Writeln(GRF_File,StartY);
  end;
  W:=(YFormat div 10);
  D:=(YFormat mod 10);
  if YAxisType=2  then GAxisType:=' 2 1'
                  else GAxisType:=' 1 1';
  if Okay<1 then Writeln(GRF_File,
           Real2Str(LowY,W+3,D+2),'  ',
           Real2Str(HighY,W+3,D+2),' 9.900e+028 ',GAxisType);
  if Okay<1 then
  begin
    if YAxisType=1 then
        Writeln(GRF_File, '2.700e+002  ',Real2Str(TicEveryY,W+3,D+2),
                      ' 1.000e-001  1  1')
    else
        Writeln(GRF_File, '2.700e+002  ',' 9.900e+028  ',
                      ' 1.500e-001  1  1')
  end;
  if Okay<1 then
  begin
    if YAxisType=2 then
         Writeln(GRF_File,' 5  ',Int2Str(D),'  2')
    else
         Writeln(GRF_File,Int2Str(LabelEveryY),'  ',Int2Str(D),'  1');
  end;
  if Okay<1 then
       Writeln(GRF_File,'1 9.900e+028 0.000e+000 9.900e+028 2.500e-001');
  if Okay<1 then
  begin
    if GrafType in ['1','2'] then
         Writeln(GRF_File,'"SET1"  "SET4"  "',PlotValRec[KP-1].YCap,'"')
    else
         Writeln(GRF_File,'"SET1"  "SET4"  "',Y_AxisName,'"');
  end;
  if Okay<1 then Writeln(GRF_File,'5.000e-002');
  if GrafType<>'C' then
  begin
    for Fx:=1 to 8 do
    begin
      if Okay<1 then Writeln(GRF_File,Top[Fx]);
    end;
    for Fx:=1 to 8 do
    begin
      if Okay<1 then Writeln(GRF_File,Right[Fx]);
    end;
  end;
  {Write out Title TXT files}
  if GrafType in ['C','1','2'] then
  begin
    if Okay<1 then Writeln(GRF_File,'TITLE');
    if Okay<1 then Writeln(GRF_File,'"SET4" "SET5" 84 1');
    if Okay<1 then
       Writeln(GRF_File,'1.650e+000 6.000e+000 0.000e+000 4.000e-001');
    if Okay<1 then Writeln(GRF_File,'1');
    ParseStr(Title,AnyStr);
    if Okay<1 then Writeln(GRF_File,'0 ',Int2Str(Length(AnyStr)),
                             '"',AnyStr,'"');
  end;
  Close(GRF_File);
  if Okay>0 then
  begin
    Beep; Exit;
  end;
  {$I-}
end;


procedure DumpToGrapher(PltType : char; PltNum : byte;
                         var Check : integer);
const
 GRFSymbols : array[0..20] of string[4] =
 (#34#76#34#32,
  #34#65#34#32,#34#66#34#32,#34#67#34#32,#34#68#34#32,
  #34#69#34#32,#34#96#34#32,#34#97#34#32,#34#78#34#32,
  #34#79#34#32,#34#80#34#32,#34#81#34#32,#34#98#34#32,
  #34#86#34#32,
  #34#87#34#32,#34#88#34#32,#34#89#34#32,#34#90#34#32,
  #34#91#34#32,#34#92#34#32,#34#75#34#32);
 {Data-related constants}
     DatFileName : string[5] = 'CORG_';
var
   GRF_File,DAT_File : text;
   GrafFileName,GDatFileName : string;
   W,D,XType,YType : byte;
   GAxisType : string;
   Fx : word;
   SymbolIs : string[4];
   indx1,indx2 : integer;
   XVal,YVal,YDat : real;

          procedure FileCheck;
          begin
            IOCheck;
            Check:=IOVal;
          end;

          function Okay : integer;
          begin
           if Check<1 then
            begin
             FileCheck; Okay:=IOVal;
            end
            else OKay:=0;
          end;
begin
  {----------------Write GRF file----------------}
  case PltType of
    'C','1' : begin
           XType:=1; YType:=1;
         end;
    '2' : begin
           XType:=2; YType:=2;
         end;
    '4' : begin
           XType:=1; YType:=2;
         end;
  end;
  case PltType of
    '1','2' : begin
       Make_GRF(SimilarityChoice,PltNum,PltType,
            PlotValRec[KP-1].XCap,PlotValRec[KP-1].XLog+1,PlotValRec[KP-1].XNum,
            PlotValRec[KP-1].Xmini,PlotValRec[KP-1].XMaxi,PlotValRec[KP-1].XTics,1,
            PlotValRec[KP-1].YCap,PlotValRec[KP-1].YLog+1,PlotValRec[KP-1].YNum,
            PlotValRec[KP-1].Ymini,PlotValRec[KP-1].YMaxi,PlotValRec[KP-1].YTics,1,
              Check);
    end;
    'C' :
       Make_GRF(SimilarityChoice,PltNum,PltType,
            '  ',1,0,-1.0,8.5,0.0,10,
            '  ',1,0,-1.0,15.0,0.0,10,
              Check);
  end;
  if Check>0 then  Exit;
  case PltType of
    '1','2' : begin
        GrafFileName:=Drive3+'\'+DatFileName +  chr(SimilarityChoice+64)+
                       'V'+Int2Str(PltNum)+'.DAT';
        ClearNull(GrafFileName);
        Assign(DAT_File,GrafFileName); Rewrite(DAT_File); FileCheck;
        if Check>0 then Exit;
        {writing DAT file for variables}
        SymbolIs:=GRFSymbols[0];
        for Fx:=1 to M do
        begin
             Writeln(DAT_File,SymbolIs,'  0.0    0.0');
             Writeln(DAT_File,SymbolIs,'  ',A1[Fx,NX]:15:6,A1[Fx,NY]:15:6);
        end;
        Close(DAT_File);
        if Okay>0 then
        begin
          Beep; Exit;
        end;
        GrafFileName:=Drive3+'\'+DatFileName
                       +chr(SimilarityChoice+64)+Int2Str(PltNum)+'.DAT';
        ClearNull(GrafFileName);
        Assign(DAT_File,GrafFileName); Rewrite(DAT_File); FileCheck;
        if Check>0 then Exit;
        {writing DAT file for data}
        for Fx:=1 to N do
        begin
           if OKay<1 then
           begin
             if (SmpSymbol[Fx] in [0..20]) then
               SymbolIs:=GRFSymbols[SmpSymbol[Fx]]
             else
               SymbolIs:=GRFSymbols[0];
             Writeln(DAT_File,SymbolIs,'  ',B[Fx,NX]:15:6,B[Fx,NY]:15:6);
           end
           else Fx:=N;
        end; {of FX do loop}
        Close(DAT_File);
        if Okay>0 then
        begin
          Beep; Exit;
        end;
    end;  {of 1,2,4}
    'C' : begin
        GrafFileName:=Drive3+'\'+DatFileName
                       +chr(SimilarityChoice+64)+'_S'+'.DAT';
        ClearNull(GrafFileName);
        Assign(DAT_File,GrafFileName); Rewrite(DAT_File); FileCheck;
        if Check>0 then Exit;
        {writing Symbols file}
        for Fx:=1 to NoGroups do
        begin
          YVal:=11.0-(0.4*Fx);
            if Okay<1 then Writeln(DAT_File,GRFSymbols[GroupSymbol[Fx]],
                           '  2.0 ',Real2Str(YVal,10,2));
        end;
              Close(DAT_File);
              if Okay>0 then
                 begin
                   Beep; Exit;
                 end;
        GrafFileName:=Drive3+'\'+DatFileName
                       +chr(SimilarityChoice+64)+'_N'+'.DAT';
        ClearNull(GrafFileName);
        Assign(DAT_File,GrafFileName); Rewrite(DAT_File); FileCheck;
        if Check>0 then Exit;
        {writing SymbolCaptions file}
        for Fx:=1 to NoGroups do
        begin
          YVal:=11.0-(0.4*Fx);
            if Okay<1 then Writeln(DAT_File,'"',GroupTitle[Fx],'"',
                           '  2.5 ',Real2Str(YVal,10,2));
        end;
              Close(DAT_File);
              if Okay>0 then
                 begin
                   Beep; Exit;
                 end;

    end;  {of type 'C'}
  end; {of case plottype}
end; {of procedure}

end.
