object FormDemo1: TFormDemo1
  Left = 0
  Top = 0
  Caption = 'AutomateApp4pas'
  ClientHeight = 376
  ClientWidth = 613
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  ShowHint = True
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBoxApp: TGroupBox
    AlignWithMargins = True
    Left = 8
    Top = 145
    Width = 597
    Height = 223
    Margins.Left = 8
    Margins.Top = 8
    Margins.Right = 8
    Margins.Bottom = 8
    Align = alClient
    Caption = 'Spreadsheet'
    TabOrder = 0
    ExplicitHeight = 206
    DesignSize = (
      597
      223)
    object LabelDescription: TLabel
      Left = 256
      Top = 32
      Width = 321
      Height = 13
      Anchors = [akLeft, akTop, akRight]
      AutoSize = False
    end
    object ButtonTest: TButton
      Left = 24
      Top = 50
      Width = 217
      Height = 25
      Caption = 'Check for app'
      TabOrder = 0
      OnClick = ButtonTestClick
    end
    object ButtonVersion: TButton
      Left = 24
      Top = 81
      Width = 217
      Height = 25
      Caption = 'Show app version'
      TabOrder = 1
      OnClick = ButtonVersionClick
    end
    object ButtonOpen: TButton
      Left = 24
      Top = 112
      Width = 217
      Height = 25
      Caption = 'Open document'
      TabOrder = 2
      OnClick = ButtonOpenClick
    end
    object ButtonEditCell: TButton
      Left = 24
      Top = 143
      Width = 217
      Height = 25
      Hint = 'Edit cell E4 and set memo text to this cell.'
      Caption = 'Edit cell content'
      TabOrder = 3
      OnClick = ButtonEditCellClick
    end
    object Memo1: TMemo
      Left = 256
      Top = 52
      Width = 321
      Height = 149
      Anchors = [akLeft, akTop, akRight, akBottom]
      Lines.Strings = (
        'Memo1')
      ScrollBars = ssVertical
      TabOrder = 4
    end
    object ButtonGetContent: TButton
      Left = 24
      Top = 174
      Width = 217
      Height = 25
      Hint = 'Read the doc cell by cell and fill memo.'
      Caption = 'Get document content'
      TabOrder = 5
      OnClick = ButtonGetContentClick
    end
  end
  object GroupBoxSettings: TGroupBox
    AlignWithMargins = True
    Left = 8
    Top = 8
    Width = 597
    Height = 121
    Margins.Left = 8
    Margins.Top = 8
    Margins.Right = 8
    Margins.Bottom = 8
    Align = alTop
    Caption = 'Settings'
    TabOrder = 1
    DesignSize = (
      597
      121)
    object RadioGroupDriver: TRadioGroup
      Left = 24
      Top = 27
      Width = 553
      Height = 46
      Anchors = [akLeft, akTop, akRight]
      Caption = 'Driver'
      ItemIndex = 0
      Items.Strings = (
        'HojaCalc')
      TabOrder = 0
    end
    object CheckBoxShowApp: TCheckBox
      Left = 24
      Top = 88
      Width = 542
      Height = 17
      Anchors = [akLeft, akTop, akRight]
      Caption = 'Show Application (while automating)'
      TabOrder = 1
      OnClick = CheckBoxShowAppClick
    end
  end
  object FileOpenDialog1: TFileOpenDialog
    FavoriteLinks = <>
    FileTypes = <>
    Options = []
    Left = 552
    Top = 8
  end
end
