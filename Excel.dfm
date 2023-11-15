object Form2: TForm2
  Left = 0
  Top = 0
  Caption = 'Form2'
  ClientHeight = 399
  ClientWidth = 696
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  TextHeight = 15
  object ButtonCreate: TButton
    Left = 8
    Top = 8
    Width = 150
    Height = 33
    Caption = #1057#1086#1079#1076#1072#1090#1100' Excel'
    TabOrder = 0
    OnClick = ButtonCreateClick
  end
  object ButtonClose: TButton
    Left = 248
    Top = 344
    Width = 150
    Height = 25
    Caption = #1047#1072#1082#1088#1099#1090#1100' Excel'
    TabOrder = 1
    OnClick = ButtonCloseClick
  end
  object ButtonOpenExcel: TButton
    Left = 8
    Top = 56
    Width = 150
    Height = 25
    Caption = #1054#1090#1082#1088#1099#1090#1100' Excel'
    TabOrder = 2
    OnClick = ButtonOpenExcelClick
  end
  object ButtonOpenSheet1: TButton
    Left = 8
    Top = 96
    Width = 150
    Height = 25
    Caption = #1054#1090#1082#1088#1099#1090#1100' '#1051#1080#1089#1090' 2'
    TabOrder = 3
    OnClick = ButtonOpenSheet1Click
  end
  object ButtonOpenSheet2: TButton
    Left = 8
    Top = 136
    Width = 150
    Height = 25
    Caption = #1054#1090#1082#1088#1099#1090#1100' '#1051#1080#1089#1090' 2.2'
    TabOrder = 4
    OnClick = ButtonOpenSheet2Click
  end
  object Edit1: TEdit
    Left = 208
    Top = 13
    Width = 121
    Height = 23
    TabOrder = 5
  end
  object ButtonAddC3: TButton
    Left = 368
    Top = 12
    Width = 150
    Height = 25
    Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1074' '#1103#1095#1077#1081#1082#1091' '#1057'3 '
    TabOrder = 6
    OnClick = ButtonAddC3Click
  end
  object Edit2: TEdit
    Left = 208
    Top = 57
    Width = 121
    Height = 23
    TabOrder = 7
  end
  object ButtonAddA2: TButton
    Left = 368
    Top = 56
    Width = 150
    Height = 25
    Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1074' '#1103#1095#1077#1081#1082#1091' '#1040'2'
    TabOrder = 8
    OnClick = ButtonAddA2Click
  end
  object ButtonAddFormula: TButton
    Left = 8
    Top = 176
    Width = 150
    Height = 25
    Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1092#1086#1088#1084#1091#1083#1091' '#1074'  '#1040'5'
    TabOrder = 9
    OnClick = ButtonAddFormulaClick
  end
  object ButtonFillCell: TButton
    Left = 8
    Top = 216
    Width = 150
    Height = 25
    Caption = #1054#1082#1088#1072#1089#1080#1090#1100' '#1103#1095#1077#1081#1082#1091' C5'
    TabOrder = 10
    OnClick = ButtonFillCellClick
  end
  object ButtonFillColStr: TButton
    Left = 208
    Top = 112
    Width = 310
    Height = 25
    Caption = #1054#1082#1088#1072#1089#1080#1090#1100' '#1089#1090#1086#1083#1073#1077#1094' '#1080' '#1089#1090#1088#1086#1082#1091
    TabOrder = 11
    OnClick = ButtonFillColStrClick
  end
  object ButtonСellParam: TButton
    Left = 8
    Top = 256
    Width = 150
    Height = 25
    Caption = #1055#1072#1088#1072#1084#1077#1090#1088#1099' '#1096#1088#1080#1092#1090#1072
    TabOrder = 12
    OnClick = ButtonСellParamClick
  end
  object ButtonCellColParam: TButton
    Left = 8
    Top = 296
    Width = 161
    Height = 25
    Caption = #1055#1072#1088#1072#1084#1077#1090#1088#1099' '#1096#1088#1080#1092#1090#1072' '#1089#1090#1086#1083#1073#1094#1072
    TabOrder = 13
    OnClick = ButtonCellColParamClick
  end
end
