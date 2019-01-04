object fmMain: TfmMain
  Left = 592
  Top = 92
  AutoScroll = False
  Caption = 'ACD'#25509#32893#29575#26376#22577#34920
  ClientHeight = 159
  ClientWidth = 405
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
  object pnl1: TPanel
    Left = 0
    Top = 0
    Width = 405
    Height = 57
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 16
      Top = 18
      Width = 64
      Height = 20
      Caption = #32113#35336#26085#26399
    end
    object Label2: TLabel
      Left = 232
      Top = 18
      Width = 16
      Height = 20
      Caption = #65374
    end
    object DateTimePickerBegin: TDateTimePicker
      Left = 90
      Top = 14
      Width = 129
      Height = 28
      Date = 42782.689854432870000000
      Time = 42782.689854432870000000
      TabOrder = 0
    end
    object DateTimePickerEnd: TDateTimePicker
      Left = 259
      Top = 14
      Width = 129
      Height = 28
      Date = 42782.689854432870000000
      Time = 42782.689854432870000000
      TabOrder = 1
    end
  end
  object btnRunReport: TButton
    Left = 16
    Top = 75
    Width = 372
    Height = 72
    Caption = #22519#34892#22577#34920
    TabOrder = 1
    OnClick = btnRunReportClick
  end
  object btnDebug: TButton
    Left = 16
    Top = 164
    Width = 372
    Height = 72
    Caption = 'Debug'
    TabOrder = 2
    OnClick = btnDebugClick
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
    Left = 149
    Top = 216
  end
  object XLSReadWriteII51: TXLSReadWriteII5
    ComponentVersion = '5.20.67a'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 21
    Top = 248
  end
  object JcVersionInfo1: TJcVersionInfo
    Left = 53
    Top = 216
  end
  object JcLog: TJcLog
    TimeStampFormat = 'yyyy/mm/dd hh:mm:ss'
    Left = 21
    Top = 216
  end
  object mtSiteSummary: TkbmMemTable
    DesignActivation = True
    AttachedAutoRefresh = True
    AttachMaxCount = 1
    FieldDefs = <>
    IndexDefs = <>
    SortOptions = []
    PersistentBackup = False
    ProgressFlags = [mtpcLoad, mtpcSave, mtpcCopy]
    LoadedCompletely = False
    SavedCompletely = False
    FilterOptions = []
    Version = '6.30'
    LanguageID = 0
    SortID = 0
    SubLanguageID = 1
    LocaleID = 1024
    Left = 85
    Top = 216
    object mtSiteSummaryYear: TIntegerField
      DisplayLabel = #24180#24230
      FieldName = 'Year'
    end
    object mtSiteSummaryMonth: TIntegerField
      DisplayLabel = #26376#20221
      FieldName = 'Month'
    end
    object mtSiteSummaryTaipei_YesDays: TIntegerField
      DisplayLabel = #36948#27161
      FieldName = 'Taipei_YesDays'
    end
    object mtSiteSummaryTaipei_NoDays: TIntegerField
      DisplayLabel = #26410#36948#27161
      FieldName = 'Taipei_NoDays'
    end
    object mtSiteSummaryTaipei_Score: TFloatField
      DisplayLabel = #27604#20363
      FieldName = 'Taipei_Score'
      DisplayFormat = '0.00'
    end
    object mtSiteSummaryTaoyuan_YesDays: TIntegerField
      DisplayLabel = #36948#27161
      FieldName = 'Taoyuan_YesDays'
    end
    object mtSiteSummaryTaoyuan_NoDays: TIntegerField
      DisplayLabel = #26410#36948#27161
      FieldName = 'Taoyuan_NoDays'
    end
    object mtSiteSummaryTaoyuan_Score: TFloatField
      DisplayLabel = #27604#20363
      FieldName = 'Taoyuan_Score'
      DisplayFormat = '0.00'
    end
    object mtSiteSummaryTaichung_YesDays: TIntegerField
      DisplayLabel = #36948#27161
      FieldName = 'Taichung_YesDays'
    end
    object mtSiteSummaryTaichung_NoDays: TIntegerField
      DisplayLabel = #26410#36948#27161
      FieldName = 'Taichung_NoDays'
    end
    object mtSiteSummaryTaichung_Score: TFloatField
      DisplayLabel = #27604#20363
      FieldName = 'Taichung_Score'
      DisplayFormat = '0.00'
    end
    object mtSiteSummaryTainan_YesDays: TIntegerField
      DisplayLabel = #36948#27161
      FieldName = 'Tainan_YesDays'
    end
    object mtSiteSummaryTainan_NoDays: TIntegerField
      DisplayLabel = #26410#36948#27161
      FieldName = 'Tainan_NoDays'
    end
    object mtSiteSummaryTainan_Score: TFloatField
      DisplayLabel = #27604#20363
      FieldName = 'Tainan_Score'
      DisplayFormat = '0.00'
    end
    object mtSiteSummaryWinton_YesDays: TIntegerField
      DisplayLabel = #36948#27161
      FieldName = 'Winton_YesDays'
    end
    object mtSiteSummaryWinton_NoDays: TIntegerField
      DisplayLabel = #26410#36948#27161
      FieldName = 'Winton_NoDays'
    end
    object mtSiteSummaryWInton_Score: TFloatField
      DisplayLabel = #27604#20363
      FieldName = 'WInton_Score'
      DisplayFormat = '0.00'
    end
  end
  object mtTeSummary: TkbmMemTable
    DesignActivation = True
    AttachedAutoRefresh = True
    AttachMaxCount = 1
    FieldDefs = <>
    IndexDefs = <>
    SortOptions = []
    PersistentBackup = False
    ProgressFlags = [mtpcLoad, mtpcSave, mtpcCopy]
    LoadedCompletely = False
    SavedCompletely = False
    FilterOptions = []
    Version = '6.30'
    LanguageID = 0
    SortID = 0
    SubLanguageID = 1
    LocaleID = 1024
    Left = 117
    Top = 216
    object mtTeSummaryEmpNo: TStringField
      DisplayLabel = #20195#34399
      FieldName = 'EmpNo'
      Size = 10
    end
    object mtTeSummaryEmpName: TStringField
      DisplayLabel = #22995#21517
      FieldName = 'EmpName'
    end
    object mtTeSummaryYesDays_1: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #36948#27161'_1'
      FieldName = 'YesDays_1'
    end
    object mtTeSummaryNoDays_1: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #26410#36948#27161'_1'
      FieldName = 'NoDays_1'
    end
    object mtTeSummaryScore_1: TFloatField
      DefaultExpression = '0'
      DisplayLabel = #27604#20363'_1'
      FieldName = 'Score_1'
      DisplayFormat = '0.00'
    end
    object mtTeSummaryYesDays_2: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #36948#27161'_2'
      FieldName = 'YesDays_2'
    end
    object mtTeSummaryNoDays_2: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #26410#36948#27161
      FieldName = 'NoDays_2'
    end
    object mtTeSummaryScore_2: TFloatField
      DefaultExpression = '0'
      DisplayLabel = #27604#20363'_2'
      FieldName = 'Score_2'
      DisplayFormat = '0.00'
    end
    object mtTeSummaryYesDays_3: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #36948#27161'_3'
      FieldName = 'YesDays_3'
    end
    object mtTeSummaryNoDays_3: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #26410#36948#27161'_3'
      FieldName = 'NoDays_3'
    end
    object mtTeSummaryScore_3: TFloatField
      DefaultExpression = '0'
      DisplayLabel = #27604#20363'_3'
      FieldName = 'Score_3'
      DisplayFormat = '0.00'
    end
    object mtTeSummaryYesDays_4: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #36948#27161'_4'
      FieldName = 'YesDays_4'
    end
    object mtTeSummaryNoDays_4: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #26410#36948#27161'_4'
      FieldName = 'NoDays_4'
    end
    object mtTeSummaryScore_4: TFloatField
      DefaultExpression = '0'
      DisplayLabel = #27604#20363'_4'
      FieldName = 'Score_4'
      DisplayFormat = '0.00'
    end
    object mtTeSummaryYesDays_5: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #36948#27161'_5'
      FieldName = 'YesDays_5'
    end
    object mtTeSummaryNoDays_5: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #26410#36948#27161'_5'
      FieldName = 'NoDays_5'
    end
    object mtTeSummaryScore_5: TFloatField
      DefaultExpression = '0'
      DisplayLabel = #27604#20363'_5'
      FieldName = 'Score_5'
      DisplayFormat = '0.00'
    end
    object mtTeSummaryYesDays_6: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #36948#27161'_6'
      FieldName = 'YesDays_6'
    end
    object mtTeSummaryNoDays_6: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #26410#36948#27161'_6'
      FieldName = 'NoDays_6'
    end
    object mtTeSummaryScore_6: TFloatField
      DefaultExpression = '0'
      DisplayLabel = #27604#20363'_6'
      FieldName = 'Score_6'
      DisplayFormat = '0.00'
    end
    object mtTeSummaryYesDays_7: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #36948#27161'_7'
      FieldName = 'YesDays_7'
    end
    object mtTeSummaryNoDays_7: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #26410#36948#27161'_7'
      FieldName = 'NoDays_7'
    end
    object mtTeSummaryScore_7: TFloatField
      DefaultExpression = '0'
      DisplayLabel = #27604#20363'_7'
      FieldName = 'Score_7'
      DisplayFormat = '0.00'
    end
    object mtTeSummaryYesDays_8: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #36948#27161'_8'
      FieldName = 'YesDays_8'
    end
    object mtTeSummaryNoDays_8: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #26410#36948#27161'_8'
      FieldName = 'NoDays_8'
    end
    object mtTeSummaryScore_8: TFloatField
      DefaultExpression = '0'
      DisplayLabel = #27604#20363'_8'
      FieldName = 'Score_8'
      DisplayFormat = '0.00'
    end
    object mtTeSummaryYesDays_9: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #36948#27161'_9'
      FieldName = 'YesDays_9'
    end
    object mtTeSummaryNoDays_9: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #26410#36948#27161'_9'
      FieldName = 'NoDays_9'
    end
    object mtTeSummaryScore_9: TFloatField
      DefaultExpression = '0'
      DisplayLabel = #27604#20363'_9'
      FieldName = 'Score_9'
      DisplayFormat = '0.00'
    end
    object mtTeSummaryYesDays_10: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #36948#27161'_10'
      FieldName = 'YesDays_10'
    end
    object mtTeSummaryNoDays_10: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #26410#36948#27161'_10'
      FieldName = 'NoDays_10'
    end
    object mtTeSummaryScore_10: TFloatField
      DefaultExpression = '0'
      DisplayLabel = #27604#20363'_10'
      FieldName = 'Score_10'
      DisplayFormat = '0.00'
    end
    object mtTeSummaryYesDays_11: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #36948#27161'_11'
      FieldName = 'YesDays_11'
    end
    object mtTeSummaryNoDays_11: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #26410#36948#27161'_11'
      FieldName = 'NoDays_11'
    end
    object mtTeSummaryScore_11: TFloatField
      DefaultExpression = '0'
      DisplayLabel = #27604#20363'_11'
      FieldName = 'Score_11'
      DisplayFormat = '0.00'
    end
    object mtTeSummaryYesDays_12: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #36948#27161'_12'
      FieldName = 'YesDays_12'
    end
    object mtTeSummaryNoDays_12: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #26410#36948#27161'_12'
      FieldName = 'NoDays_12'
    end
    object mtTeSummaryScore_12: TFloatField
      DefaultExpression = '0'
      DisplayLabel = #27604#20363'_12'
      FieldName = 'Score_12'
      DisplayFormat = '0.00'
    end
    object mtTeSummaryYesDays_Total: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #36948#27161'_'#21512#35336
      FieldName = 'YesDays_Total'
    end
    object mtTeSummaryNoDays_Total: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #26410#36948#27161'_'#21512#35336
      FieldName = 'NoDays_Total'
    end
    object mtTeSummaryScore_Total: TFloatField
      DefaultExpression = '0'
      DisplayLabel = #27604#20363'_'#21512#35336
      FieldName = 'Score_Total'
      DisplayFormat = '0.00'
    end
  end
end
