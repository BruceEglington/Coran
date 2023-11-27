unit Cor_SelLV;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, Cor_varb;

type
  TfmSelectNumVectors = class(TForm)
    Label1: TLabel;
    eNumEigenVectorsSelected: TEdit;
    bbOK: TBitBtn;
    lMax: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure bbOKClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmSelectNumVectors: TfmSelectNumVectors;

implementation

{$R *.dfm}

procedure TfmSelectNumVectors.FormCreate(Sender: TObject);
begin
  eNumEigenVectorsSelected.Text := IntToStr(Nox - 1);
  lMax.Caption := '(Maximum is '+IntToStr(Nox)+')';
end;

procedure TfmSelectNumVectors.bbOKClick(Sender: TObject);
begin
  try
    NumEigenvectorsSelected := StrToInt(eNumEigenvectorsSelected.Text);
  except
    NumEigenvectorsSelected := Nox;
  end;
end;

end.
