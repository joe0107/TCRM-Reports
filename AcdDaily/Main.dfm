object fmMain: TfmMain
  Left = 417
  Top = 166
  Caption = 'ACD'#25509#32893#29575#26085#22577#34920
  ClientHeight = 490
  ClientWidth = 742
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -16
  Font.Name = #24494#36575#27491#40657#39636
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 20
  object ListBox1: TListBox
    Left = 0
    Top = 109
    Width = 742
    Height = 381
    Align = alClient
    ItemHeight = 20
    TabOrder = 0
    OnDblClick = ListBox1DblClick
  end
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 742
    Height = 109
    ActivePage = TabSheet1
    Align = alTop
    TabOrder = 1
    object TabSheet1: TTabSheet
      Caption = #32113#35336#65286#32000#37636#36039#26009
      object Label2: TLabel
        Left = 4
        Top = 13
        Width = 64
        Height = 20
        Caption = #32113#35336#26085#26399
      end
      object lbl1: TLabel
        Left = 222
        Top = 13
        Width = 16
        Height = 20
        Caption = #65374
      end
      object DateTimePicker2: TDateTimePicker
        Left = 80
        Top = 9
        Width = 129
        Height = 28
        Date = 42782.689854432870000000
        Time = 42782.689854432870000000
        TabOrder = 0
      end
      object DateTimePicker3: TDateTimePicker
        Left = 252
        Top = 9
        Width = 129
        Height = 28
        Date = 42782.689854432870000000
        Time = 42782.689854432870000000
        TabOrder = 1
      end
      object CheckBox_RecalcWICSIPH2: TCheckBox
        Left = 4
        Top = 44
        Width = 205
        Height = 17
        Caption = #37325#26032#35336#31639' WICSIPH2'
        TabOrder = 2
      end
      object btn_ACD1: TButton
        Left = 423
        Top = 12
        Width = 133
        Height = 50
        Caption = #35336#31639#65286#20786#23384#36039#26009
        TabOrder = 3
        OnClick = btn_ACD1Click
      end
    end
    object TabSheet5: TTabSheet
      Caption = #20841#24180#24230#27604#36611#34920
      ImageIndex = 4
      object Label5: TLabel
        Left = 4
        Top = 27
        Width = 64
        Height = 20
        Caption = #32113#35336#26085#26399
      end
      object DateTimePicker6: TDateTimePicker
        Left = 79
        Top = 23
        Width = 129
        Height = 28
        Date = 42782.689854432870000000
        Time = 42782.689854432870000000
        TabOrder = 0
      end
      object btnAcdSummary: TButton
        Left = 231
        Top = 12
        Width = 133
        Height = 50
        Caption = #22519#34892#22577#34920
        TabOrder = 1
        OnClick = btnAcdSummaryClick
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'ACD'#25509#32893#29575#26085#22577#34920
      ImageIndex = 1
      object Label1: TLabel
        Left = 4
        Top = 27
        Width = 64
        Height = 20
        Caption = #32113#35336#26085#26399
      end
      object DateTimePicker1: TDateTimePicker
        Left = 80
        Top = 23
        Width = 129
        Height = 28
        Date = 42782.689854432870000000
        Time = 42782.689854432870000000
        TabOrder = 0
      end
      object ComboBox1: TComboBox
        Left = 244
        Top = 23
        Width = 145
        Height = 28
        Style = csDropDownList
        ItemIndex = 0
        TabOrder = 1
        Text = #21488#21271
        Items.Strings = (
          #21488#21271
          #21271#21312
          #20013#21312
          #21335#21312)
      end
      object btnRunReport: TButton
        Left = 424
        Top = 12
        Width = 133
        Height = 50
        Caption = #22519#34892#22577#34920
        TabOrder = 2
        OnClick = btnRunReportClick
      end
    end
    object TabSheet3: TTabSheet
      Caption = #29151#26989#34389#22238#38651#25928#29575#32113#35336#34920
      ImageIndex = 2
      object Label3: TLabel
        Left = 4
        Top = 27
        Width = 64
        Height = 20
        Caption = #32113#35336#26085#26399
      end
      object DateTimePicker4: TDateTimePicker
        Left = 79
        Top = 23
        Width = 129
        Height = 28
        Date = 42782.689854432870000000
        Time = 42782.689854432870000000
        TabOrder = 0
      end
      object btnSitePhoneSummary: TButton
        Left = 231
        Top = 12
        Width = 133
        Height = 50
        Caption = #22519#34892#22577#34920
        TabOrder = 1
        OnClick = btnSitePhoneSummaryClick
      end
    end
    object TabSheet4: TTabSheet
      Caption = #35347#32244#20491#20154#22238#38651#25928#29575#32113#35336#34920
      ImageIndex = 3
      object Label4: TLabel
        Left = 4
        Top = 21
        Width = 64
        Height = 20
        Caption = #32113#35336#26085#26399
      end
      object DateTimePicker5: TDateTimePicker
        Left = 79
        Top = 17
        Width = 129
        Height = 28
        Date = 42782.689854432870000000
        Time = 42782.689854432870000000
        TabOrder = 0
      end
      object btnTePhoneSummary: TButton
        Left = 231
        Top = 12
        Width = 133
        Height = 50
        Caption = #22519#34892#22577#34920
        TabOrder = 1
        OnClick = btnTePhoneSummaryClick
      end
    end
  end
  object XLSDbRead51: TXLSDbRead5
    Column = 0
    ExcludeFields.Strings = (
      'DEAL'
      'PRESENTED'
      'CASE014')
    IncludeFieldnames = True
    IndentDetailTables = True
    ReadDetailTables = True
    FormatCells = False
    Row = 0
    Sheet = 0
    XLS = XLSReadWriteII51
    Left = 704
    Top = 96
  end
  object XLSReadWriteII51: TXLSReadWriteII5
    ComponentVersion = '5.20.67a'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 576
    Top = 128
  end
  object mdReport: TdxMemData
    Indexes = <>
    SortOptions = []
    OnFilterRecord = mdReportFilterRecord
    Left = 576
    Top = 96
  end
  object NetDrive1: TNetDrive
    Left = 672
    Top = 96
  end
  object JcVersionInfo1: TJcVersionInfo
    Left = 640
    Top = 96
  end
  object JcLog: TJcLog
    TimeStampFormat = 'yyyy/mm/dd hh:mm:ss'
    Left = 608
    Top = 96
  end
end
