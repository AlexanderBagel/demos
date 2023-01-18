object DropSourceMainForm: TDropSourceMainForm
  Left = 734
  Top = 367
  BorderIcons = [biSystemMenu]
  Caption = 'FWOleDragDrop Source demo'
  ClientHeight = 187
  ClientWidth = 519
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 8
    Width = 241
    Height = 57
    AutoSize = False
    Caption = 
      #1042#1099#1076#1077#1083#1080#1090#1077' '#1086#1076#1080#1085' '#1080#1083#1080' '#1085#1077#1089#1082#1086#1083#1100#1082#1086' '#1074#1080#1088#1090#1091#1072#1083#1100#1085#1099#1093' '#1092#1072#1081#1083#1086#1074' '#1080#1079' '#1085#1080#1079#1083#1077#1078#1072#1097#1077#1075#1086' '#1089#1087 +
      #1080#1089#1082#1072' '#1080' '#1087#1077#1088#1077#1090#1072#1097#1080#1090#1077' '#1084#1099#1096#1100#1102' '#1074' '#1087#1088#1086#1074#1086#1076#1085#1080#1082', '#1083#1080#1073#1086' '#1089#1082#1086#1087#1080#1088#1091#1081#1090#1077' '#1080#1093' '#1095#1077#1088#1077#1079' '#1087#1088 +
      #1072#1074#1091#1102' '#1082#1083#1072#1074#1080#1096#1091' '#1084#1099#1096#1080'.'
    WordWrap = True
  end
  object Label2: TLabel
    Left = 270
    Top = 8
    Width = 241
    Height = 57
    AutoSize = False
    Caption = 
      #1069#1090#1086' '#1080#1079#1086#1073#1088#1072#1078#1077#1085#1080#1077' '#1103#1074#1083#1103#1077#1090#1089#1103' '#1088#1077#1072#1083#1100#1085#1099#1084' '#1092#1072#1081#1083#1086#1084' '#1087#1088#1080#1089#1091#1090#1089#1090#1074#1091#1102#1097#1077#1084' '#1085#1072' '#1076#1080#1089#1082#1077 +
      ', '#1077#1075#1086' '#1087#1077#1088#1077#1085#1086#1089' '#1086#1089#1091#1097#1077#1089#1090#1074#1083#1103#1077#1090#1089#1103' '#1095#1077#1088#1077#1079' CF_HDROP '#1090'.'#1082'. '#1090#1086#1083#1100#1082#1086' '#1090#1072#1082#1086#1081' '#1090#1080 +
      #1087' '#1087#1077#1088#1077#1085#1086#1089#1072' '#1087#1086#1085#1080#1084#1072#1077#1090' MSPAINT.'
    WordWrap = True
  end
  object Image1: TImage
    Left = 272
    Top = 72
    Width = 239
    Height = 105
    Stretch = True
    OnMouseDown = Image1MouseDown
  end
  object CheckListBox1: TCheckListBox
    Left = 8
    Top = 71
    Width = 249
    Height = 106
    ItemHeight = 13
    Items.Strings = (
      'virtual_dropsource_demo.dpr'
      'virtual_dropsource_demo.dproj'
      'virtual_dropsource_demo.res'
      'virtual_dropsource_main.dfm'
      'virtual_dropsource_main.pas')
    PopupMenu = PopupMenu1
    TabOrder = 0
    OnMouseDown = CheckListBox1MouseDown
  end
  object PopupMenu1: TPopupMenu
    Left = 208
    Top = 80
    object N1: TMenuItem
      Caption = #1050#1086#1087#1080#1088#1086#1074#1072#1090#1100
      OnClick = N1Click
    end
  end
end
