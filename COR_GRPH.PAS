unit Cor_grph;

interface
uses
 crt, dos,
 allsorts, sys_def,
 xvideo, inputset,
 kem_varb,
 graph, cor_varb,
 getfield,
 easyplot;

procedure Get_Axis_Values( var XMin, XMax, YMin, YMax : real);
procedure Coran_Graph;
procedure Group_Banner;

implementation

const
  StatusLineMsg = 'F1=select  F2=scaling  F10=quit  CtrlF4=Grapher  AltF4=hz dump  AltF8=vrt dump';

{----------CORAN graphics routines----------}


procedure Group_Banner;
var
  ComplexKey  : boolean;
  KeyCode     : byte;
var
   PlotSessionName : string;
   InptChr,InptChr2,InptChr3
             :   char;
   Cntrl     :   byte;
   TitleFont,TitleSize,NameFont,NameSize : word;
   YStart : integer;
begin
  SetGraphMode(GraphMode);
  GetAspectRatio(XAsp,YAsp);
  HorizPixels:=Round(0.005*MaxX);
  VertPixels:=AspY(HorizPixels);
  SaveViewPort(0);
  TitleFont:=AxisFont;
  TitleSize:=AxisLabelScale;
  NameFont:=TicFont;
  NameSize:=TicLabelScale+1;
  PlotSessionName:= 'Names and Symbols for file : '+FileName;
  SetTextStyle(TitleFont,HorizDir,TitleSize);
  SetTextJustify(CenterText,CenterText);
  OutTextXY(Round(0.5*MaxX),Round(0.75*TextHeight('M')),
            PlotSessionName);
  OutTextXY(Round(0.5*MaxX),Round(1.5*TextHeight('M')),
            StringOf(Length(PlotSessionName),ord('-')));
  YStart:=Round(2.0*TextHeight('M'));
  SetTextStyle(NameFont,HorizDir,NameSize);
  SetTextJustify(LeftText,CenterText);
  for Indx:=1 to NoGroups do
  begin
     SetColor(SymbolColor[Indx]);
     PlotSymbol(Round(0.2*MaxX),YStart+Round(1.5*Indx*TextHeight('M')),
                                                        GroupSymbol[Indx]);
     OutTextXY(Round(0.3*MaxX),YStart+Round(1.5*Indx*TextHeight('M')),
                                                        GroupTitle[Indx]);
  end;
  GetKey(ComplexKey,KeyCode);
  case ComplexKey of
    true : begin
      case KeyCode of
            altf8key : begin
               ScreenDumpVert(10);
               Tone(400,250);
            end;
            altf4key : begin
               ScreenDumpHoz(10);
               Tone(400,250);
            end;
            cntrlf4key : begin
              LoadViewPort(0);
              SetTextStyle(DefaultFont,HorizDir,1);
              SetTextJustify(LeftText,TopText);
              SetFillStyle(EmptyFill,0);
              Bar(Round(MaxX*0.05),MaxY-(TextHeight('M')+4),
                  Round(0.95*MaxX),MaxY);
              SetColor(SymbolColor[2]);
              Rectangle(Round(MaxX*0.05),MaxY-(TextHeight('M')+4),
                       Round(0.95*MaxX),MaxY);
              OutTextXY(Round(MaxX*0.05)+TextWidth('M'),
                   MaxY-(TextHeight('M')+2),'Generating GRAPHER file.  Please wait ...');
              DumpToGrapher('C',1,IOVal);
              if IOVal=0 then
              begin
                     Bar(Round(MaxX*0.05),MaxY-(TextHeight('M')+4),
                                 Round(0.95*MaxX),MaxY);
                     Rectangle(Round(MaxX*0.05),MaxY-(TextHeight('M')+4),
                                 Round(0.95*MaxX),MaxY);
                     ClearNull(Drive3);
                     OutTextXY(Round(MaxX*0.05)+TextWidth('M'),
                        MaxY-(TextHeight('M')+2),'GRAPHER file generated successfully = '
                        +Drive3+'\CORG_'+chr(SimilarityChoice+64)+chr(KP+48)+'.GRF');
              end
              else begin
                     Bar(Round(MaxX*0.05),MaxY-(TextHeight('M')+4),
                                 Round(0.95*MaxX),MaxY);
                     Rectangle(Round(MaxX*0.05),MaxY-(TextHeight('M')+4),
                                 Round(0.95*MaxX),MaxY);
                     ClearNull(Drive3);
                     OutTextXY(Round(MaxX*0.05)+TextWidth('M'),
                        MaxY-(TextHeight('M')+2),'Error generating GRAPHER file '
                        +Drive3+'\CORG_'+chr(SimilarityChoice+64)+chr(KP+48)+'.GRF');
              end;
              BeepBarp;
              Delay(1000);
            end;
      end;
    end;
    false : begin
    end;
  end;
  RestoreCrtMode;
  Delay(600);
end;



procedure Get_Axis_Values ( var XMin, XMax, YMin, YMax : real);
const
  LMax  =  8;
  LMin  =  1;
var
  L  : integer;
  ChStr : string[1];
  XWide, YWide : integer;
begin
    StartAtX:=XMin;
    StartAtY:=YMin;
    EndAtX:=XMax;
    EndAtY:=YMax;
    EndAtX:=EndAtX;
    EndAtY:=EndAtY;
    XWide:=PlotValRec[KP-1].XNum;
    YWide:=PlotValRec[KP-1].YNum;
    if (EndAtX<=StartAtX) then EndAtX:=StartAtX+0.005*StartAtX;
    if (EndAtY<=StartAtY) then EndAtY:=StartAtY+0.005*StartAtY;
    TicsAtX:=(EndAtX-StartAtX)/10.0;
    TicsAtY:=(EndAtY-StartAtY)/10.0;
    TicsAtX := PlotValRec[KP-1].XTics;
    TicsAtY := PlotValRec[KP-1].YTics;
    Draw_Block( 8, 5,11,75,1,NormBld);
    Centre( 8,' Graph Options ',RevNorm);
    ClearBlock( 9, 6,10,74,true);
    PrintS(10, 7,'Title',NormBld);
    Draw_Block(12, 5,18,75,1,NormBld);
    ClearBlock(13, 6,17,74,true);
    PrintS(13,25,'X Axis',UndlnBld);
    PrintS(13,50,'Y Axis',UndlnBld);
    At(14, 7); Write('Minimum    ',StartAtX:13:6);
    At(14,45); Write(StartAtY:13:6);
    At(15, 7); Write('Maximum    ',EndAtX:13:6);
    At(15,45); Write(EndAtY:13:6);
    At(16, 7); Write('Tick every ',TicsAtX:13:6);
    At(16,45); Write(TicsAtY:13:6);
    At(17, 7); Write('Field specif.   ',(Xwide div 10):1,':',(Xwide mod 10):1);
    At(17,45); Write((Ywide div 10):1,':',(Ywide mod 10):1);
    TCI:=1;
    CPos:=1;
    ReadStr(10,18,StringOf(40,60),Norm,RevNorm,
                 CPos,TCI,Title);
    L:=1;
    repeat
      case L of
        1 : begin
          CPos:=0;
          ReadReal(14,18,'@@@@@@.@@@@@@',Norm,RevNorm,-99999.99,99999.999,
                 CPos,TCI,StartAtX);
        end;
        2 : begin
          CPos:=0;
          ReadReal(14,45,'@@@@@@.@@@@@@',Norm,RevNorm,-99999.99,99999.999,
                 CPos,TCI,StartAtY);
        end;
        3 : begin
          CPos:=0;
          ReadReal(15,18,'@@@@@@.@@@@@@',Norm,RevNorm,-99999.99,99999.999,
                 CPos,TCI,EndAtX);
        end;
        4 : begin
          CPos:=0;
          ReadReal(15,45,'@@@@@@.@@@@@@',Norm,RevNorm,-99999.99,99999.999,
                 CPos,TCI,EndAtY);
        end;
        5 : begin
          CPos:=0;
          ReadReal(16,18,'@@@@@@.@@@@@@',Norm,RevNorm,-99999.99,99999.999,
                 CPos,TCI,TicsAtX);
        end;
        6 : begin
          CPos:=0;
          ReadReal(16,45,'@@@@@@.@@@@@@',Norm,RevNorm,-99999.99,99999.999,
                 CPos,TCI,TicsAtY);
        end;
        7 : begin
          CPos:=0;
          ReadInt(17,23,'#:#',Norm,RevNorm,10,97,
                 CPos,TCI,Xwide);
        end;
        8 : begin
          CPos:=0;
          ReadInt(17,45,'#:#',Norm,RevNorm,10,97,
                 CPos,TCI,Ywide);
        end;
      end;
      case TCI of
        tabkey, enterkey : begin
          L:=L+1;
          if L>LMax then L:=LMin;
        end;
        shiftabkey : begin
          L:=L-1;
          if L<LMin then L:=LMax;
        end;
        arrowdnkey : begin
          L:=L+2;
          if L>LMax then L:=LMin;
        end;
        arrowupkey : begin
          L:=L-2;
          if L<LMin then L:=LMax;
        end;
      end;
    until ((TCI=f10key) or (TCI=esckey));
    XMin:=StartAtX;
    XMax:=EndAtX;
    YMin:=StartAtY;
    YMax:=EndAtY;
    PlotValRec[KP-1].XMini:=XMin;
    PlotValRec[KP-1].XMaxi:=XMax;
    PlotValRec[KP-1].YMini:=YMin;
    PlotValRec[KP-1].YMaxi:=YMax;
    PlotValRec[KP-1].XTics:=TicsAtX;
    PlotValRec[KP-1].YTics:=TicsAtY;
    PlotValRec[KP-1].XNum:=XWide;
    PlotValRec[KP-1].YNum:=YWide;
end;


procedure Coran_Graph;
var
  Indx,SmpIndx        : integer;
  ComplexKey  : boolean;
  KeyCode     : byte;
  EndHelp     : boolean;
  BkColor,FgColor : integer;

procedure SetUpNormNorm(XLabel,YLabel : string;
                    MinmX,MaxmX,XTicEvery,XSmallTicEvery     :real;
                    XWide,XDec : byte;
                    MinmY,MaxmY,YTicEvery,YSmallTicEvery     : real;
                    YWide,YDec : byte);
var
 Xo,Yo,Dummy : integer;
begin
 SaveViewPort(0);
 SetAxis_X(MinmX,MaxmX,XTicEvery,XSmallTicEvery,XWide,XDec,XLabel,Xo,Dummy);
 SetAxis_Y(MinmY,MaxmY,YTicEvery,YSmallTicEvery,YWide,YDec,YLabel,Dummy,Yo);
 SetViewPort(Xo,Yo-LengthY_Axis,Xo+LengthX_Axis+HorizPixels,Yo,true);
 SaveViewPort(1);
end;


procedure Show_Graph;
var
  Indx            : integer;
  NfieldX, NdecX,
  NfieldY, NdecY  : integer;
  t1, t2, temp    : real;
  AxisTit1, AxisTit2  : string[10];
begin
    SetColor(White);
    Nfieldx:=PlotValRec[KP-1].XNum div 10;
    Ndecx:=PlotValRec[KP-1].XNum mod 10;
    Nfieldy:=PlotValRec[KP-1].YNum div 10;
    Ndecy:=PlotValRec[KP-1].YNum mod 10;
    SetGraphMode(GraphMode);
    GetAspectRatio(XAsp,YAsp);
    HorizPixels:=Round(0.005*MaxX);
    VertPixels:=AspY(HorizPixels);
    case NX of
      1 : AxisTit1:='Factor 1';
      2 : AxisTit1:='Factor 2';
      3 : AxisTit1:='Factor 3';
      4 : AxisTit1:='Factor 4';
    end;
    case NY of
      1 : AxisTit2:='Factor 1';
      2 : AxisTit2:='Factor 2';
      3 : AxisTit2:='Factor 3';
      4 : AxisTit2:='Factor 4';
    end;
    PlotValRec[KP-1].XCap:=AxisTit1;
    PlotValRec[KP-1].YCap:=AxisTit2;
    SetUpNormNorm(AxisTit1,AxisTit2,
                   StartAtX,EndAtX,TicsAtX,0.5*TicsAtX,NfieldX,NdecX,
                   StartAtY,EndAtY,TicsAtY,0.5*TicsAtY,NfieldY,NdecY);
    DrawBorder;
    LoadViewPort(0);
    SetTextStyle(DefaultFont,HorizDir,1);
    SetTextJustify(CenterText,CenterText);
    SetFillStyle(EmptyFill,0);
    Bar(Round(MaxX*0.05),1,
        Round(0.95*MaxX),(TextHeight('M')+4));
    OutTextXY(Round(MaxX/2.0),Round((TextHeight('M')+2)/2.0),Title);
    LoadViewPort(1);
    {Draw line for origin}
    SetLineStyle(DashedLn,0,NormWidth);
    MoveTo(XPos(0.0),YPos(StartAtY));
    LineTo(XPos(0.0),YPos(EndAtY));
    MoveTo(XPos(StartAtX),YPos(0.0));
    LineTo(XPos(EndAtX),YPos(0.0));
    SetLineStyle(SolidLn,0,NormWidth);
    {PlotPoints}
    if (SimilarityChoice<4) then begin
      SetColor(White);
      for Indx:=1 to M do
      begin
          PlotSymbol(XPos(A1[Indx,NX]),YPos(A1[Indx,NY]),VarbSymbol[Indx]);
      end;
    end;
    for Indx:=1 to N do
    begin
        SetColor(SymbolColor[Indx]);
        PlotSymbol(XPos(B[Indx,NX]),YPos(B[Indx,NY]),
                   SmpSymbol[Indx]);
    end;
    SetColor(White);
    StatusLine(StatusLineMsg);
end;


begin
 if GraphError=0 then begin
    Show_Graph;
    StatusLine(StatusLineMsg);
    GraphAlready:=true;
    Graphdone:=false;
    Indx:=0;
    repeat
      StatusLine(StatusLineMsg);
      GetKey(ComplexKey,KeyCode);
      case ComplexKey of
        true : begin
          case KeyCode of
            esckey : Graphdone:=true;
            f10key : Graphdone:=true;
            f1key  : begin
                     EndHelp:=false;
                     KeyCode:=1;
                     SmpIndx:=1;
                     SetTextStyle(DefaultFont,HorizDir,1);
                     SetTextJustify(LeftText,TopText);
                     SetFillStyle(EmptyFill,0);
                  while not EndHelp do
                  begin
                   LoadViewPort(0);
                   if ((KeyCode<>arrowlfkey) or (KeyCode<>arrowrtkey)) then begin
                     SetColor(SymbolColor[SmpIndx]);
                     Bar(Round(MaxX*0.05),MaxY-(TextHeight('M')+4),
                                 Round(0.95*MaxX),MaxY);
                     Rectangle(Round(MaxX*0.05),MaxY-(TextHeight('M')+4),
                                 Round(0.95*MaxX),MaxY);
                     OutTextXY(Round(MaxX*0.05)+TextWidth('M'),
                               MaxY-(TextHeight('M')+2),'Sample: '+Component[SmpIndx]);
                     OutTextXY(Round(MaxX*0.35)+TextWidth('M'),
                               MaxY-(TextHeight('M')+2),
                               'X: '+Real2Str(B[SmpIndx,NX],9,5));
                     OutTextXY(Round(MaxX*0.65)+TextWidth('M'),
                               MaxY-(TextHeight('M')+2),
                               'Y: '+Real2Str(B[SmpIndx,NY],9,5));
                   end;
                   LoadViewPort(1);
                   FlikSymbol(XPos(B[SmpIndx,NX]),YPos(B[SmpIndx,NY]),
                              SmpSymbol[SmpIndx]);
                   GetKey(ComplexKey,KeyCode);
                   if ComplexKey then
                   begin
                    case KeyCode of
                     f1key : EndHelp:=true;
                     f10key : begin
                              EndHelp:=true;
                              GraphDone:=true;
                              end;
                     arrowupkey,arrowlfkey : SmpIndx:=SmpIndx-1;
                     endkey : SmpIndx:=N;
                     homekey : SmpIndx:=1;
                     arrowdnkey,arrowrtkey : SmpIndx:=SmpIndx+1;
                    end;
                   end;
                   if not ComplexKey then
                   begin
                    case KeyCode of
                     esckey : begin
                              EndHelp:=true;
                              GraphDone:=true;
                              end;
                     end;
                   end;
                   if SmpIndx<1 then SmpIndx:=N;
                   if SmpIndx>N then SmpIndx:=1;
                  end;  {of while not EndHelp}
                   LoadViewPort(0);
                   SetColor(White);
                   Bar(Round(MaxX*0.05),MaxY-(TextHeight('M')+4),
                               Round(0.95*MaxX),MaxY);
                   StatusLine(StatusLineMsg);
                     end;
            f2key  : begin
               RestoreCrtMode;
               Delay(1000);
               Get_Axis_Values(StartAtX,EndAtX,StartAtY,EndAtY);
               Show_Graph;
               StatusLine(StatusLineMsg);
            end;
            altf8key : begin
               LoadViewPort(0);
               ScreenDumpVert(10);
               Tone(400,250);
            end;
            altf4key : begin
               LoadViewPort(0);
               ScreenDumpHoz(10);
               Tone(400,250);
            end;
            cntrlf4key : begin
              LoadViewPort(0);
              SetTextStyle(DefaultFont,HorizDir,1);
              SetTextJustify(LeftText,TopText);
              SetFillStyle(EmptyFill,0);
              Bar(Round(MaxX*0.05),MaxY-(TextHeight('M')+4),
                  Round(0.95*MaxX),MaxY);
              SetColor(SymbolColor[2]);
              Rectangle(Round(MaxX*0.05),MaxY-(TextHeight('M')+4),
                       Round(0.95*MaxX),MaxY);
              OutTextXY(Round(MaxX*0.05)+TextWidth('M'),
                   MaxY-(TextHeight('M')+2),'Generating GRAPHER file.  Please wait ...');
              DumpToGrapher('1',1,IOVal);
              if IOVal=0 then
              begin
                     Bar(Round(MaxX*0.05),MaxY-(TextHeight('M')+4),
                                 Round(0.95*MaxX),MaxY);
                     Rectangle(Round(MaxX*0.05),MaxY-(TextHeight('M')+4),
                                 Round(0.95*MaxX),MaxY);
                     ClearNull(Drive3);
                     OutTextXY(Round(MaxX*0.05)+TextWidth('M'),
                        MaxY-(TextHeight('M')+2),'GRAPHER file generated successfully = '
                        +Drive3+'\CORG_'+chr(SimilarityChoice+64)+chr(KP+48)+'.GRF');
              end
              else begin
                     Bar(Round(MaxX*0.05),MaxY-(TextHeight('M')+4),
                                 Round(0.95*MaxX),MaxY);
                     Rectangle(Round(MaxX*0.05),MaxY-(TextHeight('M')+4),
                                 Round(0.95*MaxX),MaxY);
                     ClearNull(Drive3);
                     OutTextXY(Round(MaxX*0.05)+TextWidth('M'),
                        MaxY-(TextHeight('M')+2),'Error generating GRAPHER file '
                        +Drive3+'\CORG_'+chr(SimilarityChoice+64)+chr(KP+48)+'.GRF');
              end;
              BeepBarp;
              Delay(1000);
            end;
          end;
        end;
      false : begin
          case KeyCode of
            esckey : Graphdone:=true;
          end;
        end;
      end;
    until Graphdone;
    GraphAlready:=true;
    RestoreCrtMode;
    Delay(1000);
 end
 else begin
   ClrScr;
   Centre(13,'. . Graphics mode not available on this computer . .',NormBld);
   Delay(2000);
 end;
 ClrScr;
end;

end.




