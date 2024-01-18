unit tst_sort;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm1 = class(TForm)
    Memo1: TMemo;
    Button1: TButton;
    Memo2: TMemo;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses Cor_varb, Matrix;

{$R *.dfm}

procedure TForm1.FormShow(Sender: TObject);
var
  i : integer;
begin
  DR[1] := 9.8;
  DR[2] := 6.7;
  DR[3] := 13.8;
  DR[4] := 2.8;
  DR[5] := 5.6;
  DR[6] := 21.7;
  DR[7] := 3.3;
  N := 7;
  Memo1.Lines.Clear;
  Memo2.Lines.Clear;
  for i := 1 to 7 do
  begin
    Memo1.Lines.Add(FormatFloat('#0.0',DR[i]));
  end;
end;

procedure TForm1.Button1Click(Sender: TObject);
var
  i : integer;
begin
  Memo2.Lines.Clear;
  Sort (DR, N, 1);
  for i := 1 to 7 do
  begin
    Memo2.Lines.Add(FormatFloat('#0.0',DR[i]));
  end;
end;

end.
