object DropTargetMainForm: TDropTargetMainForm
  Left = 543
  Top = 293
  Caption = 'FWOleDragDrop Targert demo'
  ClientHeight = 354
  ClientWidth = 634
  Color = clBtnFace
  Constraints.MinWidth = 650
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  DesignSize = (
    634
    354)
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 8
    Width = 581
    Height = 13
    Caption = 
      #1055#1077#1088#1077#1090#1072#1097#1080#1090#1077' '#1092#1072#1081#1083#1099' '#1085#1072' '#1092#1086#1088#1084#1091' '#1080#1079' '#1083#1102#1073#1086#1075#1086' '#1080#1089#1090#1086#1095#1085#1080#1082#1072' ('#1087#1088#1086#1074#1086#1076#1085#1080#1082', '#1073#1088#1072#1091#1079#1077 +
      #1088', '#1076#1077#1084#1086#1087#1088#1080#1083#1086#1078#1077#1085#1080#1077' dropsource_demo.exe)'
  end
  object Memo1: TMemo
    Left = 8
    Top = 32
    Width = 616
    Height = 312
    Anchors = [akLeft, akTop, akRight, akBottom]
    ScrollBars = ssBoth
    TabOrder = 0
  end
end
