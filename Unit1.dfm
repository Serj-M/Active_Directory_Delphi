object Form1: TForm1
  Left = 196
  Top = 130
  BorderStyle = bsSingle
  Caption = 'Form1'
  ClientHeight = 552
  ClientWidth = 1040
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 360
    Top = 40
    Width = 149
    Height = 13
    Caption = #1042' '#1089#1087#1088#1072#1074#1086#1095#1085#1080#1082#1077' '#1085#1077' '#1085#1072#1081#1076#1077#1085#1085#1099' :'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label2: TLabel
    Left = 360
    Top = 296
    Width = 170
    Height = 13
    Caption = #1042' '#1089#1087#1088#1072#1074#1086#1095#1085#1080#1082#1077' '#1077#1089#1090#1100' '#1089#1086#1074#1087#1072#1076#1077#1085#1080#1103' :'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label3: TLabel
    Left = 8
    Top = 16
    Width = 150
    Height = 13
    Caption = #1040#1076#1088#1077#1089#1089' ('#1073#1091#1082#1074#1072') '#1089#1090#1086#1083#1073#1094#1072' :'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold, fsUnderline]
    ParentFont = False
  end
  object Label13: TLabel
    Left = 544
    Top = 8
    Width = 45
    Height = 13
    Caption = #1064#1072#1073#1083#1086#1085' :'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label15: TLabel
    Left = 544
    Top = 344
    Width = 58
    Height = 13
    Caption = #1056#1077#1079#1091#1083#1100#1090#1072#1090' :'
  end
  object Memo1: TMemo
    Left = 176
    Top = 40
    Width = 169
    Height = 497
    ScrollBars = ssBoth
    TabOrder = 0
  end
  object Button1: TButton
    Left = 176
    Top = 8
    Width = 89
    Height = 25
    Caption = #1048#1084#1087#1086#1088#1090' '#1080#1079' AD'
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 360
    Top = 8
    Width = 169
    Height = 25
    Caption = #1055#1086#1080#1089#1082' '#1087#1086' EXCEL-'#1089#1087#1088#1072#1074#1086#1095#1085#1080#1082#1091
    TabOrder = 2
    OnClick = Button2Click
  end
  object Memo2: TMemo
    Left = 360
    Top = 64
    Width = 169
    Height = 217
    ScrollBars = ssBoth
    TabOrder = 3
  end
  object Memo3: TMemo
    Left = 544
    Top = 24
    Width = 481
    Height = 313
    Lines.Strings = (
      'dn: cn=,ou=test,dc=serv01,dc=test,dc=ru'
      'changetype: modify'
      'replace: physicalDeliveryOfficeName'
      'physicalDeliveryOfficeName: '
      '-'
      'replace: telephoneNumber'
      'telephoneNumber: '
      '-'
      'replace: streetAddress'
      'streetAddress: '
      '-'
      'replace: l'
      'l: '
      '-'
      'replace: title'
      'title: '
      '-'
      'replace: department'
      'department: '
      '-'
      'replace: company'
      'company: '
      '-')
    TabOrder = 4
  end
  object Memo4: TMemo
    Left = 544
    Top = 360
    Width = 481
    Height = 177
    ScrollBars = ssVertical
    TabOrder = 5
  end
  object Memo5: TMemo
    Left = 360
    Top = 320
    Width = 169
    Height = 217
    ScrollBars = ssBoth
    TabOrder = 6
  end
  object Panel1: TPanel
    Left = 16
    Top = 64
    Width = 145
    Height = 241
    TabOrder = 7
    object Label4: TLabel
      Left = 8
      Top = 14
      Width = 64
      Height = 13
      Caption = #1044#1086#1083#1078#1085#1086#1089#1090#1100' :'
    end
    object Label5: TLabel
      Left = 8
      Top = 48
      Width = 75
      Height = 13
      Caption = #1044#1077#1087#1072#1088#1090#1072#1084#1077#1085#1090' :'
    end
    object Label6: TLabel
      Left = 8
      Top = 80
      Width = 57
      Height = 13
      Caption = #1050#1086#1084#1087#1072#1085#1080#1103' :'
    end
    object Label7: TLabel
      Left = 8
      Top = 112
      Width = 43
      Height = 13
      Caption = #1040#1076#1088#1077#1089#1089' :'
    end
    object Label8: TLabel
      Left = 8
      Top = 144
      Width = 36
      Height = 13
      Caption = #1043#1086#1088#1086#1076' :'
    end
    object Label9: TLabel
      Left = 8
      Top = 176
      Width = 51
      Height = 13
      Caption = #1058#1077#1083#1077#1092#1086#1085' :'
    end
    object Label10: TLabel
      Left = 8
      Top = 208
      Width = 48
      Height = 13
      Caption = #1050#1072#1073#1080#1085#1077#1090' :'
    end
    object Edit1: TEdit
      Left = 87
      Top = 10
      Width = 41
      Height = 21
      TabOrder = 0
    end
    object Edit2: TEdit
      Left = 87
      Top = 42
      Width = 41
      Height = 21
      TabOrder = 1
    end
    object Edit3: TEdit
      Left = 87
      Top = 77
      Width = 41
      Height = 21
      TabOrder = 2
    end
    object Edit4: TEdit
      Left = 86
      Top = 110
      Width = 41
      Height = 21
      TabOrder = 3
    end
    object Edit5: TEdit
      Left = 87
      Top = 142
      Width = 41
      Height = 21
      TabOrder = 4
    end
    object Edit6: TEdit
      Left = 87
      Top = 175
      Width = 41
      Height = 21
      TabOrder = 5
    end
    object Edit7: TEdit
      Left = 87
      Top = 204
      Width = 41
      Height = 21
      TabOrder = 6
    end
  end
  object Panel2: TPanel
    Left = 16
    Top = 40
    Width = 145
    Height = 25
    TabOrder = 8
    object Label11: TLabel
      Left = 4
      Top = 4
      Width = 69
      Height = 13
      Caption = #1055#1072#1088#1072#1084#1077#1090#1088' AD'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsUnderline]
      ParentFont = False
    end
    object Label12: TLabel
      Left = 91
      Top = 4
      Width = 30
      Height = 13
      Caption = #1041#1091#1082#1074#1072
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsUnderline]
      ParentFont = False
    end
  end
  object Panel3: TPanel
    Left = 16
    Top = 320
    Width = 145
    Height = 65
    TabOrder = 9
    object Label14: TLabel
      Left = 8
      Top = 8
      Width = 97
      Height = 13
      Caption = #1048#1084#1103' '#1054#1055' '#1080' '#1076#1086#1084#1077#1085#1072' :'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsUnderline]
      ParentFont = False
    end
    object Edit9: TEdit
      Left = 8
      Top = 32
      Width = 129
      Height = 21
      TabOrder = 0
      Text = 'OU=,DC=,DC='
    end
  end
end
