object FormDemo1: TFormDemo1
  Left = 0
  Top = 0
  Caption = 'AutomateApp4pas'
  ClientHeight = 328
  ClientWidth = 613
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBoxOutlook: TGroupBox
    AlignWithMargins = True
    Left = 8
    Top = 145
    Width = 597
    Height = 175
    Margins.Left = 8
    Margins.Top = 8
    Margins.Right = 8
    Margins.Bottom = 8
    Align = alClient
    Caption = 'Outlook'
    TabOrder = 0
    DesignSize = (
      597
      175)
    object ButtonTest: TButton
      Left = 24
      Top = 32
      Width = 217
      Height = 25
      Caption = 'Check for Outlook'
      TabOrder = 0
      OnClick = ButtonTestClick
    end
    object ButtonVersion: TButton
      Left = 24
      Top = 63
      Width = 217
      Height = 25
      Caption = 'Show Outlook version'
      TabOrder = 1
      OnClick = ButtonVersionClick
    end
    object ButtonCount: TButton
      Left = 24
      Top = 94
      Width = 217
      Height = 25
      Caption = 'Count Outlook appointments'
      TabOrder = 2
      OnClick = ButtonCountClick
    end
    object ButtonExport: TButton
      Left = 24
      Top = 125
      Width = 217
      Height = 25
      Caption = 'Export all Outlook appointments'
      TabOrder = 3
      OnClick = ButtonExportClick
    end
    object Memo1: TMemo
      Left = 256
      Top = 34
      Width = 321
      Height = 116
      Anchors = [akLeft, akTop, akRight, akBottom]
      Lines.Strings = (
        'Memo1')
      TabOrder = 4
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
        'Outlook')
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
    end
  end
end
