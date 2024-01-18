unit Unit1;

interface

implementation


procedure TfmCoranMain.DoProcess;
var
  ii, jj, Noxtemp : integer;
  I, J, K, KP : integer;
  tmpStr : string;
  blankstr : string;
  {
  DDD : narray;
  VVV : nbynarray;
  }
begin
  Noxtemp := M;
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
  Nox := dmCor.ElemNames.RecordCount;
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
      for J:=1 to Nox do
      begin
        X[I,J] := dmCor.CoranChem.Fields[J+2].AsVariant;
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
      PrintM(Component,X,N,M);
    end;
  end;
  for I:=1 to M do
  begin
    Str(I:10,tempstr);
    EigName[I]:=tempstr;
  end;
  sbmain.Panels[1].Text :='Calculating similarity matrix';
  sbMain.Refresh;
  case scSimilarityChoice of
    scCorrespondence : begin
      ScaleCorAnal;
    end;
    scPCAVariance : begin
      Cov2(X,A3,N,M);
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
      Stand(X,N,M);
      RCoef(X,A3,N,M);
    end;
    scPCASpearmanR : begin
      SpearmanRho(X,A3,N,M);
    end;
    scPCAKendallR : begin
      KendallTau(X,A3,N,M);
    end;
    scRQVariance : begin
      ScaleRQAnalVar;
    end;
    scRQPearsonR : begin
      ScaleRQAnalStd;
    end;
    scRQSpearmanR : begin
      ScaleRQAnalSpearman;
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
    scPCAVariance,scPCAPearsonR : begin
    end;
    scPCASpearmanR,scPCAKendallR : begin
    end;
    scRQSpearmanR,scRQKendallR : begin
      for i:= 1 to N do
      begin
        for j:= 1 to M do
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
    PrintMC(OxideName,A3,M,M);
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

  // procedure from Davis (1973)
  EigenJ(A3,A2,M);   //eigenvalues returned in A3, eigenvectors in A2

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
  for i := 1 to M do
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
    if (scSimilarityChoice=scCorrespondence) then
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
      dmCor.EigenVec.Fields[jj].AsVariant := A2[ii,jj];
    end;
    dmCor.EigenVec.Post;
  end;
  sbmain.Panels[1].Text :='Calculating factor loadings';
  sbMain.Refresh;
  case scSimilarityChoice of
    scCorrespondence : begin
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
    scPCAVariance,scPCAPearsonR : begin
      Mmult(X,A2,B,N,M,M);     {factor scores for samples}
      for J:=1 to M do
      begin
        A3[J,J]:=Sqrt(Abs(A3[J,J]));
      end;
      MmultC(A2,A3,A1,M,M,M);  {factor loadings for variables}
    end;
    scPCASpearmanR,scPCAKendallR : begin
      Mmult(X,A2,B,N,M,M);     {factor scores for samples}
      for J:=1 to M do
      begin
        A3[J,J]:=Sqrt(Abs(A3[J,J]));
      end;
      MmultC(A2,A3,A1,M,M,M);  {factor loadings for variables}
    end;
    scRQVariance,scRQPearsonR : begin
      for J:=1 to M do
      begin
        A3[J,J]:=Sqrt(Abs(A3[J,J]));
      end;
      MmultC(A2,A3,A1,M,M,M);  {factor loadings for variables}
      Mmult(W,A2,B,N,M,M);     {factor loadings for samples}
    end;
    scRQSpearmanR,scRQKendallR : begin
      for J:=1 to M do
      begin
        A3[J,J]:=Sqrt(Abs(A3[J,J]));
      end;
      MmultC(A2,A3,A1,M,M,M);  {factor loadings for variables}
      Mmult(W,A2,B,N,M,M);     {factor loadings for samples}
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
      PrintMC(OxideName,A1,M,M);
    end;
    if PrintData1.Checked then
    begin
      Writeln(Lst,'Factor loadings  -  columns = factors, rows = samples');
      PrintM(Component,B,N,M);
    end;
  end;
  if ((scSimilarityChoice=scCorrespondence) and (M>4)) then
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
  pc1.ActivePage := tsEigenValues;
  NumEigenVectorsSelected := 2;  // need to make this variable ***********
  try
    SelectNumvectorsForm := TfmSelectNumvectors.Create(Self);
    SelectNumvectorsForm.ShowModal;
  finally
    SelectNumvectorsForm.Free;
  end;
  pc1.ActivePage := tsSummary;
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
  tsGraph1.TabVisible := false;
  sbmain.Panels[1].Text :='Preparing graphs';
  sbMain.Refresh;
  tsGraph2.TabVisible := false;
  tsGraph3.TabVisible := false;
  tsLocalities.TabVisible := false;
  tsSpreadsheets.TabVisible := false;
  tsScores.TabVisible := false;
  pVar.Visible := true;
  pGraph1Var.Visible := true;
  pSaveVar.Visible := true;
  pSaveSmp.Visible := true;
  pOutlierMapVar.Visible := true;
  p3DVar.Visible := true;
  pLocalitiesVar.Visible := true;
  pSaveEigenVal.Visible := true;
  pSaveEigenVec.Visible := true;
  case scSimilarityChoice of
    scCorrespondence : begin
      tsGraph1.TabVisible := true;
      tsGraph2.TabVisible := true;
      tsGraph3.TabVisible := true;
      tsLocalities.TabVisible := true;
      tsSpreadsheets.TabVisible := true;
      tsScores.TabVisible := true;
      pSaveVar.Visible := true;
      pSaveSmp.Visible := true;
      pSaveEigenVal.Visible := true;
      pSaveEigenVec.Visible := true;
    end;
    scRQVariance,scRQPearsonR,scRQSpearmanR,scRQKendallR : begin
      tsGraph1.TabVisible := true;
      tsGraph2.TabVisible := true;
      tsGraph3.TabVisible := true;
      tsScores.TabVisible := true;
      tsLocalities.TabVisible := true;
      tsSpreadsheets.TabVisible := true;
      pSaveVar.Visible := true;
      pSaveSmp.Visible := true;
      pSaveEigenVal.Visible := true;
      pSaveEigenVec.Visible := true;
    end;
    scPCAVariance,scPCAPearsonR,scPCASpearmanR, scPCAKendallR : begin
      tsGraph1.TabVisible := true;
      tsGraph2.TabVisible := true;
      tsGraph3.TabVisible := true;
      tsScores.TabVisible := false;
      tsLocalities.TabVisible := true;
      tsSpreadsheets.TabVisible := true;
      pSaveVar.Visible := true;
      pSaveSmp.Visible := true;
      pSaveEigenVal.Visible := true;
      pSaveEigenVec.Visible := true;
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



end.
 