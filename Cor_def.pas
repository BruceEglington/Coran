unit Cor_def;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, DBCtrls, Grids, DBGrids;

type
  TfmDefaults = class(TForm)
    Panel1: TPanel;
    bbClose: TBitBtn;
    Panel2: TPanel;
    Label1: TLabel;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    DBComboBox2: TDBComboBox;
    Label2: TLabel;
    procedure DBComboBox2Change(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure bbCloseClick(Sender: TObject);
    procedure DBNavigator1Click(Sender: TObject; Button: TNavigateBtn);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmDefaults: TfmDefaults;

implementation

uses Cor_dm_acs;

{$R *.dfm}

procedure TfmDefaults.DBComboBox2Change(Sender: TObject);
begin
  dmCor.qSymbolDefaults.Post;
end;

procedure TfmDefaults.FormCreate(Sender: TObject);
begin
  dmCor.qSymbolDefaults.Open;
  dmCor.qSymbolDefaults.First;
  dmCor.qSymbolDefaults.Edit;
end;

procedure TfmDefaults.bbCloseClick(Sender: TObject);
begin
  dmCor.qSymbolDefaults.Next;
  dmCor.qSymbolDefaults.Close;
end;

procedure TfmDefaults.DBNavigator1Click(Sender: TObject;
  Button: TNavigateBtn);
begin
  dmCor.qSymbolDefaults.Edit;
end;

end.
