program Coran;
{$B-} {boolean evaluation}
{$D-} {debug information}
{$E+} {8087 emulation}
{$L-} {local symbol information}
{$N+} {numeric processing}
{$R-} {range checking}
{$S-} {stack checking}
uses
  crt, dos, sys_def,
  sys_def2,
  xvideo, allsorts,
  readsys2, kem_varb,
  protect, findfile,
  graph, easyplot,
  getfield, lptcheck,
  cor_varb, cor_hard,
  inputset, filselec,
  matrix;

const
   LocalProgName     = 'CORAN';
   LocalUserCode     = '2.1:000:030792';
   LocalYearStr      = '1992';
   LocalMonthStr     = '12';
   LocalDayStr       = '31';
   OvrName           = 'CORAN.OVR';
   BombChar          = 'N';
const
  Version        =  '  ';
  Finished       : boolean = false;
  GetNewElements : boolean = true;
  GetNewGroups   : boolean = true;
var
  TooMuch          : boolean;
  X1               : RealArrayX;


procedure FileBanner;
const
  ListLines = 21;
  StartLine = 2;
  List : array[1..ListLines] of string =
(
'     ÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛ   ',
'     ÛÛ ÉÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ» ÛÛ±± ',
'     ÛÛ º     **   C O R A N   *   C O R A N   *   C O R A N   **    º ÛÛ±± ',
'     ÛÛ º    *   C O R A N   ***   C O R A N   ***   C O R A N   *   º ÛÛ±± ',
'     ÛÛ ÇÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¶ ÛÛ±± ',
'     ÛÛ º   A package for multivariate components/factor analysis    º ÛÛ±± ',
'     ÛÛ º   with the following features:                             º ÛÛ±± ',
'     ÛÛ º                                                            º ÛÛ±± ',
'     ÛÛ º       û  Accesses data from files created by KEMFILE       º ÛÛ±± ',
'     ÛÛ º       û  10 user selectable elements for modelling         º ÛÛ±± ',
'     ÛÛ º       û  Various measures of similarity are available for  º ÛÛ±± ',
'     ÛÛ º                Correspondence Analysis                     º ÛÛ±± ',
'     ÛÛ º                Principal Components Analysis               º ÛÛ±± ',
'     ÛÛ º                Combined R-/Q-mode Analysis                 º ÛÛ±± ',
'     ÛÛ ÇÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¶ ÛÛ±± ',
'     ÛÛ º   Author`s private copy                                    º ÛÛ±± ',
'     ÛÛ ÇÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¶ ÛÛ±± ',
'     ÛÛ º    Version 2.1/180292            c) KEMSYS Software 1992   º ÛÛ±± ',
'     ÛÛ ÈÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼ ÛÛ±± ',
'     ÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛ±± ',
'        ±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±± '
);

var
 i : byte;
 ch : integer;

begin
    ClrScr;    OffScr;
    for i:=1 to ListLines do
      PrintS(i+StartLine,2,List[i],11);
    Centre(ListLines+StartLine+1,'Press [Space Bar] to Continue, or [Esc] to exit',10);
    Draw_Block(1,1,25,80,1,14);
(*       At({ListLines+StartLine+3}25,37);Write('(',MagicN,':',LocalMaxUse,')');*)
    ChangeAtt(ListLines+StartLine-3,15,53,9);
    ChangeAtt(StartLine+3,12,57,14);
    ChangeAtt(StartLine+4,12,57,14);
    OnScr;
    At(25,40);
    repeat
     ch:=Ord(ReadKey);
    until (ch in [32,esckey]);
    ClrScr;
     if ch=esckey then
     begin
       Terminate;
       Halt;
     end;
end;


procedure OpenBanner;
begin
  FileBanner;
end;{proc OpenBanner}


procedure SelectFile1;
var
   Test,Okay         : Char;
begin
  Okay   :=   'N';
  ClrScr;
  {$I-}
  FileName:='        ';
  SExtension:='KEM';
  SelectFile(1,1,'Data file',10,'N',SDirec,FileName,SExtension,DataFile);
  if (FileName<>'ABORT') then begin
    TotalRecs:=FileSize(DataFile)-3;
    RecCount:=FileSize(DataFile);
  end;
  {$I+}
end;            {of SelectFile1 }



procedure PlotCorAnal;
begin
  KP:=2;
  repeat
    ClrScr;
    Draw_Block( 8,10,17,70,1,NormBld);
    ClearBlock( 9,11,16,69,true);
    PrintS( 9,12,'[ 0 ] - Return to previous menu ',NormBld);
    PrintS(11,12,'[ 1 ] - Plot group symbols      ',NormBld);
    PrintS(12,12,'[ 2 ] - Factor 2 vs Factor 3    ',NormBld);
    PrintS(13,12,'[ 3 ] - Factor 2 vs Factor 4    ',NormBld);
    PrintS(14,12,'[ 4 ] - Factor 3 vs Factor 4    ',NormBld);
    PrintS(16,12,'Your choice?   ',NormBld);
    ReadInt(16,25,'#',Norm,RevNorm,0,4,CPos,TCI,KP);
    case KP of
      1 : begin
         Group_Banner;
         KP:=1;
      end;
      2 : begin
         NX:=2;
         NY:=3;
         Get_Axis_Values(XMin,XMax,YMin,YMax);
         Coran_Graph;
         KP:=2;
      end;
      3 : begin
         NX:=2;
         NY:=4;
         Get_Axis_Values(XMin,XMax,ZMin,ZMax);
         Coran_Graph;
         KP:=3;
      end;
      4 : begin
         NX:=3;
         NY:=4;
         Get_Axis_Values(YMin,YMax,ZMin,ZMax);
         Coran_Graph;
         KP:=4;
      end;
    end;
  until (KP=0);
end;

procedure PlotRQAnal;
begin
  KP:=2;
  repeat
    ClrScr;
    Draw_Block( 8,10,17,70,1,NormBld);
    ClearBlock( 9,11,16,69,true);
    PrintS( 9,12,'[ 0 ] - Return to previous menu ',NormBld);
    PrintS(11,12,'[ 1 ] - Plot group symbols      ',NormBld);
    PrintS(12,12,'[ 2 ] - Factor 1 vs Factor 2    ',NormBld);
    PrintS(13,12,'[ 3 ] - Factor 1 vs Factor 3    ',NormBld);
    PrintS(14,12,'[ 4 ] - Factor 2 vs Factor 3    ',NormBld);
    PrintS(16,12,'Your choice?   ',NormBld);
    ReadInt(16,25,'#',Norm,RevNorm,0,4,CPos,TCI,KP);
    case KP of
      1 : begin
         Group_Banner;
         KP:=1;
      end;
      2 : begin
         NX:=1;
         NY:=2;
         Get_Axis_Values(XMin,XMax,YMin,YMax);
         Coran_Graph;
         KP:=2;
      end;
      3 : begin
         NX:=1;
         NY:=3;
         Get_Axis_Values(XMin,XMax,ZMin,ZMax);
         Coran_Graph;
         KP:=3;
      end;
      4 : begin
         NX:=2;
         NY:=3;
         Get_Axis_Values(YMin,YMax,ZMin,ZMax);
         Coran_Graph;
         KP:=4;
      end;
    end;
  until (KP=0);
end;

procedure DoProcess;
begin
          ClrScr;
          Draw_Block(1,10,3,70,1,NormBld);
          ClearBlock(2,11,2,69,true);
          PrintS( 2,12,'Title',NormBld);
          CPos:=1;
          TCI:=0;
          repeat
            ReadStr( 2,20,StringOf(40,60),Norm,RevNorm,CPos,TCI,Title);
          until ((TCI=enterkey) or (TCI=f10key));
          Draw_Block(4,10,7,70,1,NormBld);
          ClearBlock(5,11,6,69,true);
          Centre(5,'Set print device to top of page and press Y when ready.',NormBld);
          Centre(6,'Press N if no printout is required. ',NormBld);
          At(25,40);
          repeat
            ReadKbd(IPrn);
            IPrn:=UpCase(IPrn);
          until (IPrn in ['Y','N','S']);
          if (IPrn in ['Y','S']) then
          begin
            ClrScr;
            Draw_Block(9,15,11,65,1,NormBld);
            Centre(10,'... Check printer ready ...',NormBld);
            CheckPrinter;
            if Prndone then IPrn:='N';
          end;
          ClrScr;
          Draw_Block( 8,10,18,70,1,NormBld);
          ClearBlock( 9,11,17,69,true);
          PrintS(10,12,'[ 1 ] = Correspondence analysis ',NormBld);
          PrintS(11,12,'[ 2 ] = R- Q- Mode variance     ',NormBld);
          PrintS(12,12,'[ 3 ] = R- Q- Mode correlation  ',NormBld);
          PrintS(13,12,'[ 4 ] = PCA - covariance        ',NormBld);
          PrintS(14,12,'[ 5 ] = PCA - correlation       ',NormBld);
          PrintS(16,12,'Your choice?                ',NormBld);
          ReadInt(16,25,'#',Norm,RevNorm,1,5,CPos,TCI,SimilarityChoice);
          ClrScr;
          FillChar(X,SizeOf(X),0);
          At(13,15); Write('Reading data for ',Nox:2,' elements');
          I:=0;
          N:=1;
          TooMuch:=false;
          repeat
            Seek(DataFile,(N+2));
            Read(DataFile,DataField);
            if ((I<(NN-10)) and (GroupToPlot[DataField.GroupNumber]>0)) then
            begin
              I:=I+1;
              SmpSymbol[I]:=PlottingSymbol[DataField.GroupNumber];
              Component[I]:=DataField.SampleNo;
              for J:=1 to Nox do begin
                X[I,J]:=DataField.FileOxData[XtraElemPos[J,1]]*XtraFactor[J,1]
                        +DataField.FileOxData[XtraElemPos[J,2]]*XtraFactor[J,2];
              end;
            end;
            if ((I>=(NN-10)) and (GroupToPlot[DataField.GroupNumber]>0)) then
                    TooMuch:=true;
            At(17,15);
            Write('Processing record ',
                     N:4,'  of  ',TotalRecs:4,'   Total included: ',I);
            N:=N+1;
          until ((N>TotalRecs) or (I>=NN-10));
          if TooMuch then begin
            Draw_Block(10,20,12,60,1,NormBld);
            ClearBlock(11,21,11,59,true);
            Centre(11,'.. Data overflow. Truncating!! ..',NormBld);
            Beep; Beep;
            Delay(2000);
          end;
          ClrScr;
          N:=I;
          M:=Nox;
          for I:=1 to 9 do begin
            VarbSymbol[I]:=48+I;
          end;
          VarbSymbol[10]:=48;
          if (N<M) then begin
            M:=N;
            ClrScr;
            Draw_Block(11,10,13,70,1,NormBld);
            Centre(12,'... Insufficient data in file. Decreasing elements ...',Norm);
            At(25,40);
            Beep; Beep;
            Delay(2000);
          end;
          for I:=1 to N do
          begin
            for J:=1 to M do
            begin
              if (X[I,J]=0.0) then X[I,J]:=0.001;
              X1[i,j] := X[i,j];
            end;
          end;
          ClrScr;
          if (IPrn='Y') then
          begin
            Centre(13,'. . . Printing data matrix . . .',NormBldFlik);
            At(25,40);
            Writeln(ListFile,CondPrn);
            Write(ListFile,StringOf(70,32)+'CORAN version '+LocalUserCode);
            Writeln(ListFile,StringOf(10,32),'Printed on ',DateString);
            Writeln(ListFile);
            Writeln(ListFile,'    ',Title);
            Writeln(ListFile);
            Writeln(ListFile);
            Writeln(ListFile,'Data Matrix');
            PrintM(Component,X,N,M);
          end;
          for I:=1 to M do
          begin
            Str(I:10,tempstr);
            EigName[I]:=tempstr;
          end;
          ClrScr;
          Centre(13,'. . . Calculating similarity matrix . . .',NormBldFlik);
          At(25,40);
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
          ClrScr;
          if (IPrn='Y') then
          begin
            Centre(13,'. . . Printing similarity matrix . . .',NormBldFlik);
            At(25,40);
            Writeln(ListFile);
            Writeln(ListFile);
            case SimilarityChoice of
              1 : Writeln(ListFile,'Similarity Matrix - Correspondence Analysis');
              2 : Writeln(ListFile,'Similarity Matrix - R- Q-mode variance');
              3 : Writeln(ListFile,'Similarity Matrix - R- Q-mode correlation');
              4 : Writeln(ListFile,'Similarity Matrix - Principal components - covariance');
              5 : Writeln(ListFile,'Similarity Matrix - Principal components - correlation');
            end;
            PrintMC(OxideName,A3,M,M);
          end;
          ClrScr;
          Centre(13,'. . . Calculating eigenvectors and eigenvalues . . .',NormBldFlik);
          At(25,40);
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
          for I:=1 to M do
          begin
            SumEE:=SumEE+Abs(A1[I,1]);
            A1[I,2]:=Abs(A1[I,1])*100.0/SumE;
            A1[I,3]:=SumEE*100.0/SumE;
          end;
          if (SimilarityChoice=1) then
          begin
            A3[1,1]:=1.0;
            A1[1,1]:=1.0;
          end;
          ClrScr;
          if (IPrn='N') then
          begin
            ClrScr;
            Writeln(Title);
            Writeln;
            case SimilarityChoice of
              1 : Writeln('Similarity Matrix - Correspondence Analysis');
              2 : Writeln('Similarity Matrix - R- Q-mode variance');
              3 : Writeln('Similarity Matrix - R- Q-mode correlation');
              4 : Writeln('Similarity Matrix - Principal components - covariance');
              5 : Writeln('Similarity Matrix - Principal components - correlation');
            end;
            Writeln;
            if (SimilarityChoice=1) then
            begin
              Writeln('Eigenvalue 1 and Eigenvector 1 are artifices of the method. Ignore them!');
              Writeln;
            end;
            Writeln('Column 1 = Eigenvalues,  Column 2 = % of trace,  Column 3 = cum. % of trace');
            Writeln;
            for I:=1 to M do
            begin
              Write(EigName[I]:10);
              for J:=1 to 3 do
              begin
                Write(A1[I,J]:12:4);
              end;
              Writeln;
            end;
            At(25,40);
            WaitAnyKey;
          end;
          if (IPrn<>'N') then
          begin
            if (IPrn='Y') then Writeln(ListFile,FF);
            Writeln(ListFile,CondPrn);
            Write(ListFile,StringOf(70,32)+'CORAN version '+LocalUserCode);
            Writeln(ListFile,StringOf(10,32),'Printed on ',DateString);
            Writeln(ListFile);
            Writeln(ListFile,'    ',Title);
            Writeln(ListFile);
            Writeln(ListFile);
            case SimilarityChoice of
              1 : Writeln(ListFile,'Similarity Matrix - Correspondence Analysis');
              2 : Writeln(ListFile,'Similarity Matrix - R- Q-mode variance');
              3 : Writeln(ListFile,'Similarity Matrix - R- Q-mode correlation');
              4 : Writeln(ListFile,'Similarity Matrix - Principal components - covariance');
              5 : Writeln(ListFile,'Similarity Matrix - Principal components - correlation');
            end;
            Writeln(ListFile);
            if (SimilarityChoice=1) then
            begin
              Writeln(ListFile,'Eigenvalue 1 and Eigenvector 1 are artifices of the method. Ignore them!');
              Writeln(ListFile);
            end;
            Writeln(ListFile,'Column 1 = Eigenvalues,  Column 2 = % of trace,  Column 3 = cum. % of trace');
            PrintMC(EigName,A1,M,3);
            Writeln(ListFile);
            Writeln(ListFile);
            Writeln(ListFile,'Principal axis matrix  -  columns = eigenvectors, rows = variables');
            PrintMC(OxideName,A2,M,M);
          end;
          if (IPrn='S') then Writeln(ListFile,FF);
          ClrScr;
          Centre(13,'. . . Calculating factor loadings . . .',NormBldFlik);
          At(25,40);
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
              (*
              {Calculate factors used to calculate the sample factor loadings}
              Writeln(ListFile);
              Writeln(ListFile);
              Writeln(ListFile,'Variable factors used for loadings -  columns = factors, rows = variables');
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
              PrintMX(OxideName,W,M,M);
              MmultX(X,W,B,N,M,M);
              Writeln(ListFile);
              Writeln(ListFile);
              Writeln(ListFile,'Factor loadings  -  columns = factors, rows = samples');
              PrintM(Component,B,N,M);
              Writeln(ListFile);
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
          ClrScr;
          if (IPrn='Y') then
          begin
            Centre(13,'. . . Printing factor loadings . . .',NormBldFlik);
            At(25,40);
            Writeln(ListFile,FF);
            if (SimilarityChoice in [1..3]) then
            begin
              Writeln(ListFile,'Factor loadings  -  columns = factors, rows = variables');
              PrintMC(OxideName,A1,M,M);
              Writeln(ListFile);
              Writeln(ListFile);
            end;
            Writeln(ListFile,'Factor loadings  -  columns = factors, rows = samples');
            PrintM(Component,B,N,M);
            Writeln(ListFile);
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
             ClrScr;
             if (IPrn='Y') then
             begin
               Centre(13,'. . . Printing absolute and relative contributions . . .',NormBldFlik);
               At(25,40);
               Writeln(ListFile,FF);
               Writeln(ListFile,'Absolute and relative contributions for variables');
               Writeln(ListFile);
               Write(ListFile,'                                  2         ');
               Write(ListFile,'            3         ');
               Write(ListFile,'            4         ');
               Write(ListFile,'            5         ');
               Writeln(ListFile);
               Write(ListFile,'                Weight      Abs       Rel   ');
               Write(ListFile,'      Abs       Rel   ');
               Write(ListFile,'      Abs       Rel   ');
               Write(ListFile,'      Abs       Rel   ');
               Writeln(ListFile);
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
                 Write(ListFile,OxideName[J]:11,DC[J,J]:11:4);
                 for K:=1 to 4 do
                 begin
                   Write(ListFile,CA[K]:11:4,CRR[K]:11:4);
                 end;
                 Writeln(ListFile);
               end;
               Writeln(ListFile);
               Writeln(ListFile,'Absolute and relative contributions for samples');
               Writeln(ListFile);
               Write(ListFile,'                                  2         ');
               Write(ListFile,'            3         ');
               Write(ListFile,'            4         ');
               Write(ListFile,'            5         ');
               Writeln(ListFile);
               Write(ListFile,'                Weight      Abs       Rel   ');
               Write(ListFile,'      Abs       Rel   ');
               Write(ListFile,'      Abs       Rel   ');
               Write(ListFile,'      Abs       Rel   ');
               Writeln(ListFile);
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
                 Write(ListFile,Component[I]:11,DR[I]:11:4);
                 for K:=1 to 4 do
                 begin
                   Write(ListFile,CA[K]:11:4,CRR[K]:11:4);
                 end;
                 Writeln(ListFile);
               end;
             end;
          end;
          case SimilarityChoice of
            1 : begin
               NX:=2;
            end;
            2..5 : begin
               NX:=1;
            end;
          end;
          case SimilarityChoice of
            1..3 : begin
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
          if (IPrn='Y') then Writeln(ListFile,FF);
          case SimilarityChoice of
            1 : begin
              PlotCorAnal;
            end;
            2..5 : begin
              PlotRQAnal;
            end;
          end;
          ClrScr;
end;

begin
  CheckSystem(BombChar,LocalYearStr,LocalMonthStr,LocalDayStr,LocalProgName+'.EXE',LocalUserCode);
  ProgVersion:=LocalUserCode;
  SDirec:=Drive2;
  TFileName:='        ';
  TExtension:='DAT';
  Initialize;
  RestoreCrtMode;
  ClrScr;
  for KP:=2 to 4 do
  begin
    PlotValRec[KP-1].XCap:='     ';
    PlotValRec[KP-1].YCap:='     ';
    PlotValRec[KP-1].XMini:=10.0e6;
    PlotValRec[KP-1].YMini:=10.0e6;
    PlotValRec[KP-1].XMaxi:=-10.0e6;
    PlotValRec[KP-1].YMaxi:=-10.0e6;
    PlotValRec[KP-1].XLog:=0;
    PlotValRec[KP-1].YLog:=0;
    PlotValRec[KP-1].XNum:=62;
    PlotValRec[KP-1].YNum:=62;
    PlotValRec[KP-1].XLab:=3;
    PlotValRec[KP-1].YLab:=3;
  end;
  PickFileName:='CORAN';
  done:=false;
  for J:=1 to 10 do begin
   ID[J]:=0;
  end;
  Title:='              ';
  ElemName[0]:='..........';
  SimilarityChoice:=1;
  KP:=1;
  Openbanner;
  ClrScr;
  SetOxideNames;
  ElementOrder;
  SelectFile1;
  if (FileName<>'ABORT') then
  begin
    GetNames(DataFile);
    NDG:=NoGroups;
    for J:=1 to 10 do
    begin
      ElemPos[J]:=J;
      OxideName[J]:=ElemName[J];
      XtraElemPos[J,1]:=J;
      XtraFactor[J,1]:=1.0;
      XtraElemPos[J,2]:=1;
      XtraFactor[J,2]:=0.0;
    end;
    Finished:=false;
    TCI:=0;
    repeat
      GetPickFile(PickFileName);
      GetFieldNumbers;
    until (TCI=f10key);
    NoGroups:=NDG;
    GetNewElements:=false;
    ClrScr;
    InitPlotGroups;
    SelectPlotGroups;
    Group_Banner;
    GetNewGroups:=false;
    repeat {until Finished}
      GetNewElements:=false;
      repeat {until not GetNewElements}
        ClrScr;
        GetNewGroups:=false;
        repeat {until not GetNewGroups}
          DoProcess;
          K:=1;
          repeat
            ClrScr;
            Draw_Block( 5,10,12,70,1,NormBld);
            ClearBlock( 6,11,11,69,true);
            PrintS( 6,12,'[ 0 ] - Exit program                 ',NormBld);
            PrintS( 8,12,'[ 1 ] - Process different groups     ',NormBld);
            PrintS( 9,12,'[ 2 ] - Access new pick file         ',NormBld);
            PrintS(11,12,'Your choice?                ',NormBld);
            ReadInt(11,25,'#',Norm,RevNorm,0,2,CPos,TCI,K);
            case K of
              0 : begin
                Finished:=true;
                GetNewElements:=false;
                GetNewGroups := false;
              end;
              1 : begin
                GetNewElements:=false;
                GetNewGroups:=true;
                SelectPlotGroups;
                Group_Banner;
                GetNewGroups:=false;
                ClrScr;
              end;
              2 : begin
                GetNewGroups:=true;
                GetNewElements:=true;
              end;
            end;
          until (K in [0..2]);
          ClrScr;
          K:=2;
          if not Finished then
          begin
            if GetNewElements then
            begin
              repeat
                ClrScr;
                Draw_Block( 5,10,14,70,1,NormBld);
                ClearBlock( 6,11,13,69,true);
                PrintS( 6,12,'[ 0 ] - Exit program                 ',NormBld);
                PrintS( 8,12,'[ 1 ] - Process data from same file  ',NormBld);
                PrintS( 9,12,'[ 2 ] - Access new file              ',NormBld);
                PrintS(10,12,'[ 3 ] - Make ADD2KEM file            ',NormBld);
                PrintS(11,12,'[ 4 ] - ASCII file for factor calc.  ',NormBld);
                {
                PrintS(10,12,'[ 3 ] - Generate Roland plot file    ',NormBld);
                }
                PrintS(13,12,'Your choice?                ',NormBld);
                ReadInt(13,25,'#',Norm,RevNorm,0,4,CPos,TCI,K);
                case K of
                  0 : begin
                    Close(DataFile);
                    GetNewElements:=false;
                    GetNewGroups:=false;
                    Finished:=true;
                  end;
                  1 : begin
                    GetNames(DataFile);
                    GetNewElements:=true;
                    TCI:=0;
                    repeat
                      GetPickFile(PickFileName);
                      GetFieldNumbers;
                    until (TCI=f10key);
                    NoGroups:=NDG;
                    GetNewGroups:=true;
                    InitPlotGroups;
                    SelectPlotGroups;
                    Group_Banner;
                    GetNewGroups:=false;
                    ClrScr;
                  end;
                  2 : begin
                    Close(DataFile);
                    SelectFile1;
                    if (FileName<>'ABORT') then begin
                      GetNames(DataFile);
                      for J:=1 to 10 do begin
                        ElemPos[J]:=J;
                        OxideName[J]:=ElemName[J];
                        XtraElemPos[J,1]:=J;
                        XtraFactor[J,1]:=1.0;
                        XtraElemPos[J,2]:=1;
                        XtraFactor[J,2]:=0.0;
                      end;
                      GetNewElements:=true;
                      GetNewGroups:=true;
                      TCI:=0;
                      repeat
                        GetPickFile(PickFileName);
                        GetFieldNumbers;
                      until (TCI=f10key);
                      NoGroups:=NDG;
                      GetNewGroups:=true;
                      InitPlotGroups;
                      SelectPlotGroups;
                      Group_Banner;
                      GetNewGroups:=false;
                      ClrScr;
                    end
                    else Finished:=true;
                  end;
                  3 : begin
                     ClrScr;
                     {$I-}
                     TDirec:=SDirec;
                     TExtension:='DAT';
                     SelectFileT(1,1,'Text file',2,'Y',TDirec,FileName,TExtension,TextFile1);
                     if (FileName<>'ABORT') then
                     begin
                       Writeln(TextFile1,'1 4');
                       Writeln(TextFile1,'54 52 53 54');
                       Writeln(TextFile1,'   ');
                       for i:=1 to N do
                       begin
                         Write(TextFile1,Component[i]:10,'   ');
                         for j:=1 to 4 do
                         begin
                           Write(TextFile1,B[i,j]:15:6,' ');
                         end;
                         Writeln(TextFile1);
                       end;
                       Close(TextFile1);
                     end;
                     {$I+}
                  end;
                  {
                  3: begin
                    One;
                    Close(TextFile2);
                    K:=10;
                  end;
                  }
                  4 : begin
                     ClrScr;
                     {$I-}
                     TDirec:=SDirec;
                     TExtension:='DAT';
                     SelectFileT(1,1,'Text file',2,'Y',TDirec,FileName,TExtension,TextFile1);
                     if (FileName<>'ABORT') then
                     begin
                       k:=2;
                       Writeln(TextFile1,'Stretched, scaled data');
                       for i:=1 to N do
                       begin
                         Write(TextFile1,Component[i]:10,'   ');
                         for j:=1 to M do
                         begin
                           Write(TextFile1,X1[i,j]:10:5,' ');
                         end;
                         Writeln(TextFile1);
                       end;
                       Writeln(TextFile1,'Factor loadings');
                       for i:=1 to N do
                       begin
                         Write(TextFile1,B[i,k]:10:5,' ');
                         Writeln(TextFile1);
                       end;
                       Writeln(TextFile1);
                       Close(TextFile1);
                     end;
                     {$I+}
                  end;
                end;
              until (K in [0..2]);
            end;
          end;
        until not GetNewGroups;
      until not GetNewElements;
    until Finished;
  end;
  Terminate;
end.


