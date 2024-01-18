object Form1: TForm1
  Left = -2
  Top = 110
  Width = 870
  Height = 500
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object DBChart1: TDBChart
    Left = 36
    Top = 28
    Width = 769
    Height = 421
    Title.Text.Strings = (
      'TDBChart')
    LeftAxis.Inverted = True
    View3D = False
    TabOrder = 0
    object Series3: TPointSeries
      Marks.Callout.Brush.Color = clBlack
      Marks.Visible = False
      DataSource = dmCor.FacLoadingsVar
      ClickableLine = False
      Pointer.InflateMargins = True
      Pointer.Style = psRectangle
      Pointer.Visible = True
      XValues.Name = 'X'
      XValues.Order = loAscending
      XValues.ValueSource = 'Pos'
      YValues.Name = 'Y'
      YValues.Order = loNone
      YValues.ValueSource = 'Vector3'
    end
    object Series1: TBarSeries
      Marks.Callout.Brush.Color = clBlack
      Marks.Visible = False
      DataSource = dmCor.FacLoadingsVar
      Gradient.Direction = gdTopBottom
      XValues.Name = 'X'
      XValues.Order = loAscending
      XValues.ValueSource = 'Pos'
      YValues.Name = 'Bar'
      YValues.Order = loNone
      YValues.ValueSource = 'Vector1'
    end
    object Series2: THorizBarSeries
      Marks.Callout.Brush.Color = clBlack
      Marks.Visible = True
      DataSource = dmCor.FacLoadingsVar
      Gradient.Direction = gdLeftRight
      XValues.Name = 'Bar'
      XValues.Order = loNone
      XValues.ValueSource = 'Vector3'
      YValues.Name = 'Y'
      YValues.Order = loAscending
      YValues.ValueSource = 'Pos'
    end
  end
end
