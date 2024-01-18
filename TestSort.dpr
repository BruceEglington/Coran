program TestSort;

uses
  Forms,
  tst_sort in 'tst_sort.pas' {Form1},
  Matrix in '..\MATRIX.PAS',
  Cor_varb in 'Cor_varb.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
