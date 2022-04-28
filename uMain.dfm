object frmMain: TfrmMain
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Sheet2DB'
  ClientHeight = 561
  ClientWidth = 1011
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object btnClose: TButton
    Left = 887
    Top = 8
    Width = 121
    Height = 33
    Caption = #54532#47196#44536#47016#51333#47308
    TabOrder = 0
    OnClick = btnCloseClick
  end
  object PageControl1: TPageControl
    Left = 8
    Top = 62
    Width = 1000
    Height = 491
    ActivePage = TabSheet2
    TabOrder = 1
    object TabSheet2: TTabSheet
      Caption = #50641#49472#49440#53469
      ImageIndex = 3
      object btnOpen: TButton
        Left = 8
        Top = 3
        Width = 121
        Height = 33
        Caption = #50641#49472#50676#44592
        TabOrder = 0
        OnClick = btnOpenClick
      end
      object StringGrid1: TStringGrid
        Left = 8
        Top = 40
        Width = 969
        Height = 409
        FixedCols = 0
        TabOrder = 1
      end
    end
    object TabSheet1: TTabSheet
      Caption = 'DB '#51217#49549
      object Label1: TLabel
        Left = 42
        Top = 16
        Width = 22
        Height = 13
        Caption = #49436#48260
      end
      object Label2: TLabel
        Left = 42
        Top = 40
        Width = 24
        Height = 13
        Caption = 'DB'#47749
      end
      object Label3: TLabel
        Left = 30
        Top = 64
        Width = 33
        Height = 13
        Caption = #50500#51060#46356
      end
      object Label4: TLabel
        Left = 18
        Top = 88
        Width = 44
        Height = 13
        Caption = #54056#49828#50892#46300
      end
      object Edit4: TEdit
        Left = 73
        Top = 85
        Width = 256
        Height = 21
        ImeName = 'Microsoft Office IME 2007'
        PasswordChar = '*'
        TabOrder = 0
      end
      object Edit5: TEdit
        Left = 73
        Top = 61
        Width = 256
        Height = 21
        ImeName = 'Microsoft Office IME 2007'
        TabOrder = 1
      end
      object Edit3: TEdit
        Left = 73
        Top = 37
        Width = 256
        Height = 21
        ImeName = 'Microsoft Office IME 2007'
        TabOrder = 2
      end
      object Edit1: TEdit
        Left = 73
        Top = 13
        Width = 256
        Height = 21
        ImeName = 'Microsoft Office IME 2007'
        TabOrder = 3
      end
      object Btn_DBConn: TButton
        Left = 72
        Top = 107
        Width = 121
        Height = 25
        Caption = 'DB'#50672#44208'/'#53580#51060#48660#47532#49828#53944
        TabOrder = 4
        OnClick = Btn_DBConnClick
      end
    end
    object TabSheet4: TTabSheet
      Caption = #51089#50629#49444#51221
      ImageIndex = 1
      object Label5: TLabel
        Left = 8
        Top = 21
        Width = 66
        Height = 13
        Caption = #53580#51060#48660#47532#49828#53944
      end
      object StringGrid2: TStringGrid
        Left = 8
        Top = 40
        Width = 265
        Height = 417
        TabOrder = 0
      end
      object btnColumn: TButton
        Left = 279
        Top = 3
        Width = 129
        Height = 33
        Caption = #53580#51060#48660#52972#47100#47749
        TabOrder = 1
        OnClick = btnColumnClick
      end
      object StringGrid3: TStringGrid
        Left = 279
        Top = 41
        Width = 282
        Height = 417
        TabOrder = 2
      end
      object StringGrid4: TStringGrid
        Left = 567
        Top = 40
        Width = 155
        Height = 417
        TabOrder = 3
      end
      object StringGrid5: TStringGrid
        Left = 823
        Top = 43
        Width = 154
        Height = 417
        ColCount = 2
        FixedCols = 0
        TabOrder = 4
      end
      object btnAdd: TButton
        Left = 735
        Top = 147
        Width = 82
        Height = 33
        Caption = #50672#44208
        TabOrder = 5
        OnClick = btnAddClick
      end
      object btnDel: TButton
        Left = 735
        Top = 186
        Width = 82
        Height = 33
        Caption = #52712#49548
        TabOrder = 6
        OnClick = btnDelClick
      end
    end
    object TabSheet3: TTabSheet
      Caption = #51089#50629#51652#54665
      ImageIndex = 2
      object Label6: TLabel
        Left = 1068
        Top = 156
        Width = 11
        Height = 13
        Caption = #44148
      end
      object ProgressBar1: TProgressBar
        Left = 3
        Top = 3
        Width = 858
        Height = 33
        TabOrder = 0
      end
      object btnSave: TButton
        Left = 867
        Top = 3
        Width = 122
        Height = 49
        Caption = 'DB'#51200#51109
        TabOrder = 1
        OnClick = btnSaveClick
      end
      object StringGrid6: TStringGrid
        Left = 3
        Top = 97
        Width = 858
        Height = 236
        FixedCols = 0
        TabOrder = 2
      end
      object btnDBRead: TButton
        Left = 867
        Top = 58
        Width = 122
        Height = 49
        Caption = #44208#44284#54869#51064
        TabOrder = 3
        OnClick = btnDBReadClick
      end
      object Edit2: TEdit
        Left = 867
        Top = 113
        Width = 122
        Height = 21
        ReadOnly = True
        TabOrder = 4
      end
      object Memo1: TMemo
        Left = 3
        Top = 339
        Width = 986
        Height = 121
        Lines.Strings = (
          '')
        TabOrder = 5
      end
    end
  end
  object ADOQuery1: TADOQuery
    Parameters = <>
    Left = 608
    Top = 8
  end
end
