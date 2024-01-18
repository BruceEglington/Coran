program Coran_cds;

uses
  Forms,
  Cor_ShtIm in 'Cor_ShtIm.pas' {fmSheetImport},
  Cor_varb in 'Cor_varb.pas',
  Cor_About in 'Cor_About.pas' {AboutBox},
  MATRIX in '..\MATRIX.PAS',
  mathproc in '..\Mathproc.pas',
  AllSorts in '..\Allsorts.pas',
  Cor_ShtIm2 in 'Cor_ShtIm2.pas' {fmSheetImport2},
  Cor_ShtImdiscrimFac in 'Cor_ShtImdiscrimFac.pas' {fmSheetImportDiscrimFac},
  Cor_ShtImMean in 'Cor_ShtImMean.pas' {fmSheetImportMean},
  NumRecipes_varb in '..\NumRecipes_varb.pas',
  NumRecipes in '..\NumRecipes.pas',
  TeeJPEG in '..\..\Program Files\Steema Software\TeeChart Pro 7.07 Full Source Code\Sources\TeeJPEG.pas' {TeeJPEGOptions},
  Cor_RobustEstimator in 'Cor_RobustEstimator.pas',
  Cor_def in 'Cor_def.pas' {fmDefaults},
  Cor_SelLV in 'Cor_SelLV.pas' {fmSelectNumVectors},
  icnorm in '..\icnorm.pas',
  Cor_ShtImEigVec in 'Cor_ShtImEigVec.pas' {fmSheetImportEigVec},
  Cor_ShtImEigVal in 'Cor_ShtImEigVal.pas' {fmSheetImportEigVal},
  Cor_dm_acs in 'Cor_dm_acs.pas' {dmCor: TDataModule},
  Cor_mn in 'Cor_mn.pas' {fmCoranMain},
  Cor_dm_flex in '..\Coran\Cor_dm_flex.pas' {dmFlex: TDataModule};

{$R *.res}

{
Correspondence analysis procedures follow the mathematical formulae
of Joreskog et al (1976).

Principal component analysis procedures for both variance and Pearson R
correlation matrices follow the mathematical formulae
of Joreskog et al (1976), using the second solution for scores and loadings
provided on page 70. The second solution is preferred in order to maintain
compatibility with routines in MATLAB (basic routines plus those in LIBRA)

Simultaneous R- and Q-mode component analysis procedures for both variance
and Pearson R correlation matrices follow the mathematical formulae of
Davis (2002)

Two-group discriminant function analysis procedures utilise routines
provided by Davis(1973)

Score and orthogonal distances are computed following procedures described
by Hubert et al (2005) and have been checked using MATLAB code from the
LIBRA library for classical components analysis

The RobPCA procedure is being implemented using routines translated from
the TRobustEstimator procedures from the ROOT package


Start program
Open database
Determine number of records and number of variables
Create appropriate data arrays

}

begin
  Application.Initialize;
  Application.Title := 'Coran_cds';
  Application.CreateForm(TdmCor, dmCor);
  Application.CreateForm(TdmFlex, dmFlex);
  Application.CreateForm(TfmCoranMain, fmCoranMain);
  Application.CreateForm(TfmSheetImportEigVal, fmSheetImportEigVal);
  Application.Run;
end.
