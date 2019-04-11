object dmAcdSummary: TdmAcdSummary
  OldCreateOrder = True
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  Height = 197
  Width = 351
  object qrAgentCount: TUniQuery
    Connection = connReport
    SQL.Strings = (
      'SELECT CHEM001, COUNT(*) AS AGENT_COUNT'
      'FROM WICSCHEM WITH(NOLOCK) '
      'WHERE '
      '(CHEM006='#39'22'#39'  OR CHEM006='#39'23'#39') '
      'AND (CHEM001 >= :CHEM001B AND CHEM001 <= :CHEM001E)'
      'GROUP BY CHEM001')
    Left = 112
    Top = 48
    ParamData = <
      item
        DataType = ftDateTime
        Name = 'CHEM001B'
      end
      item
        DataType = ftDateTime
        Name = 'CHEM001E'
      end>
  end
  object qryGetAcdInfo: TUniQuery
    Connection = dmReport.connTeleContact
    Left = 80
    Top = 80
  end
  object qrGetData_Src: TUniQuery
    LocalUpdate = True
    Connection = connReport
    SQL.Strings = (
      'SELECT'
      'T3.GUID, T1.GUID AS RPHE_GUID,'
      'T3.IPHE001, T1.RPHE001, T3.IPHE004, T3.IPHE005, T3.IPHE003,'
      'T1.RPHE003, T4.SALE002, T4.SALE003, T1.RPHE005, T1.RPHE006,'
      'T3.IPHE016, T3.IPHE017, T3.IPHE012, T3.IPHE008, T3.IPHE019,'
      'T1.RPHE011, T5.CLAS004, T6.CLAS004 AS CALL_KIND, T4.SALE024,'
      'T7.CUT1002, T7.FLAG_SW, T7.FLAG_HRS, T8.TK,'
      
        'T8.IPH2001, T8.IPH2002, ISNULL(T8.IPH2003, 0) AS IPH2003, T8.IPH' +
        '2004'
      'FROM WICSIPHE T3 WITH(NOLOCK)'
      'LEFT JOIN WICSRSCE T2 WITH(NOLOCK) ON T2.RSCE003 = T3.GUID'
      'LEFT JOIN WICSRPHE T1 WITH(NOLOCK) ON T1.GUID = T2.RSCE001'
      'LEFT JOIN WICSSALE T4 WITH(NOLOCK) ON T1.RPHE003 = T4.SALE001'
      
        'LEFT JOIN WICSCLAS T5 WITH(NOLOCK) ON T5.CLAS002 = IPHE012 AND T' +
        '5.CLAS001 = '#39'10'#39
      
        'LEFT JOIN WICSCLAS T6 WITH(NOLOCK) ON T6.CLAS002 = IPHE019 AND T' +
        '6.CLAS001 = '#39'01'#39
      'LEFT JOIN WICSCUT1 T7 WITH(NOLOCK) ON T7.CUT1001 = IPHE005'
      'LEFT JOIN WICSIPH2 T8 WITH(NOLOCK) ON T8.GUID = T3.GUID'
      'WHERE'
      '('
      ' (T1.RPHE005 >= :BDATE AND T1.RPHE005 <= :EDATE)'
      ' OR'
      
        ' (T1.RPHE001 IS NULL AND (T3.IPHE004 >= :BDATE AND T3.IPHE004 <=' +
        ' :EDATE))'
      ')'
      'ORDER BY IPHE001, RPHE001')
    FetchRows = 100
    Left = 48
    Top = 80
    ParamData = <
      item
        DataType = ftDateTime
        Name = 'BDATE'
        Value = 42214d
      end
      item
        DataType = ftDateTime
        Name = 'EDATE'
        Value = 42215d
      end>
    object qrGetData_SrcIPHE001: TIntegerField
      AutoGenerateValue = arAutoInc
      DisplayLabel = #20358#38651#32232#34399
      FieldName = 'IPHE001'
      ReadOnly = True
      Required = True
    end
    object qrGetData_SrcRPHE001: TIntegerField
      AutoGenerateValue = arAutoInc
      DisplayLabel = #22238#38651#32232#34399
      FieldName = 'RPHE001'
      ReadOnly = True
    end
    object qrGetData_SrcIPHE004: TDateTimeField
      DisplayLabel = #20358#38651#26178#38291
      FieldName = 'IPHE004'
      Required = True
    end
    object qrGetData_SrcIPHE005: TStringField
      DisplayLabel = #23458#25142#20195#34399
      FieldName = 'IPHE005'
      Required = True
      Size = 10
    end
    object qrGetData_SrcIPHE003: TStringField
      DisplayLabel = #20358#38651#27512#23660#37096#38272
      FieldName = 'IPHE003'
      Required = True
      Size = 10
    end
    object qrGetData_SrcRPHE003: TStringField
      DisplayLabel = #22238#38651#20154#20195#34399
      FieldName = 'RPHE003'
      ReadOnly = True
      Size = 10
    end
    object qrGetData_SrcSALE002: TStringField
      DisplayLabel = #22238#38651#20154
      FieldName = 'SALE002'
      ReadOnly = True
      Size = 10
    end
    object qrGetData_SrcSALE003: TStringField
      DisplayLabel = #22238#38651#37096#38272
      FieldName = 'SALE003'
      ReadOnly = True
      Size = 10
    end
    object qrGetData_SrcRPHE005: TDateTimeField
      DisplayLabel = #22238#38651#37096#38272
      FieldName = 'RPHE005'
      ReadOnly = True
    end
    object qrGetData_SrcRPHE006: TDateTimeField
      DisplayLabel = #23436#25104#26178#38291
      FieldName = 'RPHE006'
      ReadOnly = True
    end
    object qrGetData_SrcIPHE016: TDateTimeField
      DisplayLabel = #38283#22987#34389#29702#26178#38291
      FieldName = 'IPHE016'
    end
    object qrGetData_SrcIPHE017: TDateTimeField
      DisplayLabel = #32080#26696#26178#38291
      FieldName = 'IPHE017'
    end
    object qrGetData_SrcIPHE012: TStringField
      DisplayLabel = #31995#32113#21029
      FieldName = 'IPHE012'
      Size = 2
    end
    object qrGetData_SrcIPHE008: TStringField
      DisplayLabel = #30041#35328#20154
      FieldName = 'IPHE008'
      Size = 10
    end
    object qrGetData_SrcIPHE019: TStringField
      DisplayLabel = #20358#38651#21312#20998
      FieldName = 'IPHE019'
      Size = 2
    end
    object qrGetData_SrcCLAS004: TStringField
      DisplayLabel = #31995#32113#21029
      FieldName = 'CLAS004'
      ReadOnly = True
      Size = 40
    end
    object qrGetData_SrcCALL_KIND: TStringField
      DisplayLabel = #20358#38651#21312#20998
      FieldName = 'CALL_KIND'
      ReadOnly = True
      Size = 40
    end
    object qrGetData_SrcRPHE011: TStringField
      DisplayLabel = #28961#25928#22238#38651
      FieldName = 'RPHE011'
      ReadOnly = True
      Size = 1
    end
    object qrGetData_SrcCUT1002: TStringField
      DisplayLabel = #23458#25142#31777#31281
      FieldName = 'CUT1002'
      ReadOnly = True
      Size = 10
    end
    object qrGetData_SrcFLAG_SW: TStringField
      FieldName = 'FLAG_SW'
      ReadOnly = True
      Size = 2
    end
    object qrGetData_SrcFLAG_HRS: TStringField
      FieldName = 'FLAG_HRS'
      ReadOnly = True
      Size = 2
    end
    object qrGetData_SrcIPH2001: TStringField
      DisplayLabel = #21512#32004#36523#20998
      FieldName = 'IPH2001'
      ReadOnly = True
      Size = 2
    end
    object qrGetData_SrcIPH2002: TIntegerField
      DisplayLabel = #22238#38651#24310#36978#26178#38291
      FieldName = 'IPH2002'
      ReadOnly = True
    end
    object qrGetData_SrcIPH2003: TBooleanField
      DisplayLabel = #36926#26178#22238#38651
      FieldName = 'IPH2003'
      ReadOnly = True
      Required = True
    end
    object qrGetData_SrcGUID: TStringField
      FieldName = 'GUID'
      Required = True
      Size = 40
    end
    object qrGetData_SrcIPH2004: TStringField
      FieldName = 'IPH2004'
      ReadOnly = True
      Size = 40
    end
    object qrGetData_SrcRPHE_GUID: TStringField
      FieldName = 'RPHE_GUID'
      ReadOnly = True
      Size = 40
    end
    object qrGetData_SrcTK: TStringField
      DisplayLabel = #38651#35441#35672#21029#30908
      FieldName = 'TK'
      ReadOnly = True
    end
    object qrGetData_SrcSALE024: TIntegerField
      DisplayLabel = 'ACD'#26085#27161#28310
      FieldName = 'SALE024'
      ReadOnly = True
    end
  end
  object mdAcdTeDaily: TdxMemData
    Indexes = <>
    Persistent.Data = {
      5665728FC2F5285C8FFE3F110000000400000009000A0050686F6E6544617465
      000A00000001000600456D704964000A00000001000800456D704E616D650008
      0000000600050044617973000400000003000E004143445F416E735F546F7461
      6C000400000003000E004143445F416E735F56616C6964000400000003001100
      4143445F41737369676E5F546F74616C0004000000030011004143445F417373
      69676E5F56616C6964000400000003000E0043616C6C6F75745F546F74616C00
      0400000003000D004143445F496E5F546F74616C000400000003001300414344
      5F56616C69644F75745F546F74616C000800000006000A004143445F53636F72
      65000A000000010007004465707449640004000000030014004E6F7441636443
      616C6C6F75745F436F756E7400040000000300090054455F546F74616C001400
      000001000900536974654E616D65001400000001000E0050686F6E6544617465
      4465736300}
    SortOptions = []
    OnCalcFields = mdAcdTeDailyCalcFields
    Left = 48
    Top = 16
    object mdAcdTeDailyPhoneDate: TDateField
      DisplayLabel = #20358#38651#26085#26399
      FieldName = 'PhoneDate'
    end
    object mdAcdTeDailyEmpId: TStringField
      DisplayLabel = #20195#34399
      FieldName = 'EmpId'
      Size = 10
    end
    object mdAcdTeDailyEmpName: TStringField
      DisplayLabel = #35347#32244#24107
      FieldName = 'EmpName'
      Size = 10
    end
    object mdAcdTeDailyDays: TFloatField
      DefaultExpression = '0'
      DisplayLabel = #20540#27231#22825#25976
      FieldName = 'Days'
    end
    object mdAcdTeDailyACD_Ans_Total: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#36890#25976
      FieldName = 'ACD_Ans_Total'
      DisplayFormat = '#,0'
    end
    object mdAcdTeDailyACD_Ans_Valid: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#26377#25928#36890#25976
      FieldName = 'ACD_Ans_Valid'
    end
    object mdAcdTeDailyACD_Assign_Total: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#25351#23450#22238#38651#36890#25976
      FieldName = 'ACD_Assign_Total'
    end
    object mdAcdTeDailyACD_Assign_Valid: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#25351#23450#22238#38651#26377#25928#36890#25976
      FieldName = 'ACD_Assign_Valid'
    end
    object mdAcdTeDailyCallout_Total: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #32317#22238#38651#25976
      FieldName = 'Callout_Total'
      DisplayFormat = '#,0'
    end
    object mdAcdTeDailyACD_In_Total: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#32317#20358#38651#25976
      FieldName = 'ACD_In_Total'
    end
    object mdAcdTeDailyACD_ValidAns_Total: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#34389#29702#25976
      FieldName = 'ACD_ValidAns_Total'
    end
    object mdAcdTeDailyACD_Score: TFloatField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#25509#32893#29575
      FieldName = 'ACD_Score'
      DisplayFormat = '#.0 %'
    end
    object mdAcdTeDailyDeptId: TStringField
      DisplayLabel = #37096#38272#20195#34399
      FieldName = 'DeptId'
      Size = 10
    end
    object mdAcdTeDailyDeptName: TStringField
      DisplayLabel = #37096#38272
      FieldKind = fkCalculated
      FieldName = 'DeptName'
      Calculated = True
    end
    object mdAcdTeDailyNotAcdCallout_Count: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #38750#27966#36865#22238#38651#25976
      FieldName = 'NotAcdCallout_Count'
    end
    object mdAcdTeDailyTE_Total_C: TIntegerField
      DisplayLabel = 'TE'#21512#32004
      FieldName = 'TE_Total_C'
    end
    object mdAcdTeDailyTE_Total_NC: TIntegerField
      DisplayLabel = 'TE'#38750#21512#32004
      FieldName = 'TE_Total_NC'
    end
    object mdAcdTeDailySiteName: TStringField
      DisplayLabel = #29151#26989#34389
      FieldName = 'SiteName'
    end
    object mdAcdTeDailyPhoneDateDesc: TStringField
      DisplayLabel = #20358#38651#26085#26399
      FieldName = 'PhoneDateDesc'
    end
    object mdAcdTeDailyTimeOut_Count_C: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #21512#32004#36926#26178
      FieldName = 'TimeOut_Count_C'
    end
    object mdAcdTeDailyPhone_Count_C: TIntegerField
      DisplayLabel = #21512#32004#20358#38651
      FieldName = 'Phone_Count_C'
    end
    object mdAcdTeDailyPhone_Count_NC: TIntegerField
      DisplayLabel = #38750#21512#32004#20358#38651
      FieldName = 'Phone_Count_NC'
    end
    object mdAcdTeDailyTimeOut_Count_NC: TIntegerField
      DisplayLabel = #38750#21512#32004#36926#26178
      FieldName = 'TimeOut_Count_NC'
    end
    object mdAcdTeDailyPhoneOut_Count_C: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #21512#32004#22238#38651
      FieldName = 'PhoneOut_Count_C'
      DisplayFormat = '#,0'
    end
    object mdAcdTeDailyPhoneOut_Count_NC: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #38750#21512#32004#22238#38651
      FieldName = 'PhoneOut_Count_NC'
      DisplayFormat = '#,0'
    end
    object mdAcdTeDailyPhoneYear: TIntegerField
      DisplayLabel = #20358#38651#24180
      FieldName = 'PhoneYear'
    end
    object mdAcdTeDailyPhoneMonth: TIntegerField
      DisplayLabel = #20358#38651#26376
      FieldName = 'PhoneMonth'
    end
    object mdAcdTeDailyACD_DailyReqCount: TIntegerField
      DisplayLabel = 'ACD'#26085#27161#28310
      FieldName = 'ACD_DailyReqCount'
    end
    object mdAcdTeDailyDuty_AM: TStringField
      FieldName = 'Duty_AM'
    end
    object mdAcdTeDailyDuty_PM: TStringField
      FieldName = 'Duty_PM'
    end
  end
  object mdAcdSiteDaily: TdxMemData
    Indexes = <>
    SortOptions = []
    OnCalcFields = mdAcdSiteDailyCalcFields
    Left = 16
    Top = 16
    object mdAcdSiteDailyPhoneDate: TDateField
      DisplayLabel = #20358#38651#26085#26399
      FieldName = 'PhoneDate'
    end
    object mdAcdSiteDailySiteId: TStringField
      DisplayLabel = #29151#26989#34389
      FieldName = 'SiteId'
    end
    object mdAcdSiteDailyACD_Total: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#27966#36865#32317#25976
      FieldName = 'ACD_Total'
      DisplayFormat = '#,0'
    end
    object mdAcdSiteDailyACD_Assign_Invalid: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#25351#23450#22238#38651#28961#25928#25976
      FieldName = 'ACD_Assign_Invalid'
      DisplayFormat = '#,0'
    end
    object mdAcdSiteDailyACD_InvalidIn_Total: TIntegerField
      DisplayLabel = #38750#35347#32244#37096#38272#27966#36865#25976
      FieldName = 'ACD_InvalidIn_Total'
    end
    object mdAcdSiteDailyACD_ValidIn_Total: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#26377#25928#27966#36865#25976
      FieldName = 'ACD_ValidIn_Total'
      DisplayFormat = '#,0'
    end
    object mdAcdSiteDailyACD_ValidAns_Total: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#34389#29702#25976
      DisplayWidth = 10
      FieldName = 'ACD_ValidAns_Total'
      DisplayFormat = '#,0'
    end
    object mdAcdSiteDailyACD_Score: TFloatField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#25509#32893#29575
      FieldName = 'ACD_Score'
      DisplayFormat = '#.0 %'
    end
    object mdAcdSiteDailySiteName: TStringField
      DisplayLabel = #29151#26989#34389
      FieldKind = fkCalculated
      FieldName = 'SiteName'
      Calculated = True
    end
    object mdAcdSiteDailyNotAcdCallout_Count: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #38750#27966#36865#22238#38651#25976
      FieldName = 'NotAcdCallout_Count'
      DisplayFormat = '#,0'
    end
    object mdAcdSiteDailyTE_Total_C: TIntegerField
      DisplayLabel = 'TE'#21512#32004
      FieldName = 'TE_Total_C'
      DisplayFormat = '#,0'
    end
    object mdAcdSiteDailyTE_Total_NC: TIntegerField
      DisplayLabel = 'TE'#38750#21512#32004
      FieldName = 'TE_Total_NC'
      DisplayFormat = '#,0'
    end
    object mdAcdSiteDailyCallout_Total: TIntegerField
      DisplayLabel = #32317#22238#38651#25976
      FieldName = 'Callout_Total'
      DisplayFormat = '#,0'
    end
    object mdAcdSiteDailyDays: TFloatField
      DefaultExpression = '0'
      DisplayLabel = #20540#27231#20154#22825
      FieldName = 'Days'
    end
    object mdAcdSiteDailyPhone_Count_C: TIntegerField
      DisplayLabel = #21512#32004#20358#38651
      FieldName = 'Phone_Count_C'
      DisplayFormat = '#,0'
    end
    object mdAcdSiteDailyPhone_Count_NC: TIntegerField
      DisplayLabel = #38750#21512#32004#20358#38651
      FieldName = 'Phone_Count_NC'
      DisplayFormat = '#,0'
    end
    object mdAcdSiteDailyTimeOut_Count_C: TIntegerField
      DisplayLabel = #21512#32004#36926#26178
      FieldName = 'TimeOut_Count_C'
    end
    object mdAcdSiteDailyTimeOut_Rate_C: TFloatField
      DisplayLabel = #21512#32004#36926#26178#29575
      FieldName = 'TimeOut_Rate_C'
      OnGetText = mdAcdSiteDailyTimeOut_Rate_CGetText
    end
    object mdAcdSiteDailyTimeOut_Count_NC: TIntegerField
      DisplayLabel = #38750#21512#32004#36926#26178
      FieldName = 'TimeOut_Count_NC'
    end
    object mdAcdSiteDailyTimeOut_Rate_NC: TFloatField
      DisplayLabel = #38750#21512#32004#36926#26178#29575
      FieldName = 'TimeOut_Rate_NC'
      OnGetText = mdAcdSiteDailyTimeOut_Rate_CGetText
    end
    object mdAcdSiteDailyPhoneDateDesc: TStringField
      DisplayLabel = #20358#38651#26085#26399
      FieldName = 'PhoneDateDesc'
    end
    object mdAcdSiteDailyACD_Ans_Total: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#30452#25509#25509#32893#25976
      FieldName = 'ACD_Ans_Total'
      DisplayFormat = '#,0'
    end
    object mdAcdSiteDailyPhoneOut_Count_C: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #21512#32004#22238#38651
      FieldName = 'PhoneOut_Count_C'
      DisplayFormat = '#,0'
    end
    object mdAcdSiteDailyPhoneOut_Count_NC: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #38750#21512#32004#22238#38651
      FieldName = 'PhoneOut_Count_NC'
      DisplayFormat = '#,0'
    end
    object mdAcdSiteDailyNoAns_Count_C: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #21512#32004#26410#22238
      FieldName = 'NoAns_Count_C'
      DisplayFormat = '#,0'
    end
    object mdAcdSiteDailyNoAns_Count_NC: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = #38750#21512#32004#26410#22238
      FieldName = 'NoAns_Count_NC'
      DisplayFormat = '#,0'
    end
    object mdAcdSiteDailyPhoneYear: TIntegerField
      DisplayLabel = #20358#38651#24180
      FieldName = 'PhoneYear'
    end
    object mdAcdSiteDailyPhoneMonth: TIntegerField
      DisplayLabel = #20358#38651#26376
      FieldName = 'PhoneMonth'
    end
  end
  object mdAcdSrc: TkbmMemTable
    DesignActivation = True
    AttachedAutoRefresh = True
    AttachMaxCount = 1
    FieldDefs = <
      item
        Name = 'IPHE001'
        Attributes = [faReadonly, faRequired]
        DataType = ftInteger
      end
      item
        Name = 'RPHE001'
        Attributes = [faReadonly]
        DataType = ftInteger
      end
      item
        Name = 'IPHE004'
        Attributes = [faRequired]
        DataType = ftDateTime
      end
      item
        Name = 'IPHE005'
        Attributes = [faRequired]
        DataType = ftString
        Size = 10
      end
      item
        Name = 'IPHE003'
        Attributes = [faRequired]
        DataType = ftString
        Size = 10
      end
      item
        Name = 'RPHE003'
        Attributes = [faReadonly]
        DataType = ftString
        Size = 10
      end
      item
        Name = 'SALE002'
        Attributes = [faReadonly]
        DataType = ftString
        Size = 10
      end
      item
        Name = 'SALE003'
        Attributes = [faReadonly]
        DataType = ftString
        Size = 10
      end
      item
        Name = 'RPHE005'
        Attributes = [faReadonly]
        DataType = ftDateTime
      end
      item
        Name = 'RPHE006'
        Attributes = [faReadonly]
        DataType = ftDateTime
      end
      item
        Name = 'IPHE016'
        DataType = ftDateTime
      end
      item
        Name = 'IPHE017'
        DataType = ftDateTime
      end
      item
        Name = 'IPHE012'
        DataType = ftString
        Size = 2
      end
      item
        Name = 'IPHE008'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'IPHE019'
        DataType = ftString
        Size = 2
      end
      item
        Name = 'Valid'
        DataType = ftBoolean
      end
      item
        Name = 'CLAS004'
        DataType = ftString
        Size = 40
      end
      item
        Name = 'CALL_KIND'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'RPHE011'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'Except'
        DataType = ftBoolean
      end
      item
        Name = 'Remark'
        DataType = ftString
        Size = 200
      end
      item
        Name = 'CUT1002'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'IPH2001'
        DataType = ftString
        Size = 2
      end
      item
        Name = 'IPH2002'
        DataType = ftInteger
      end
      item
        Name = 'IPH2003'
        DataType = ftBoolean
      end
      item
        Name = 'FLAG_SW'
        DataType = ftString
        Size = 2
      end
      item
        Name = 'FLAG_HRS'
        DataType = ftString
        Size = 2
      end
      item
        Name = 'GUID'
        DataType = ftString
        Size = 40
      end
      item
        Name = 'IPH2004'
        DataType = ftString
        Size = 40
      end
      item
        Name = 'RPHE_GUID'
        DataType = ftString
        Size = 40
      end
      item
        Name = 'TK'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    SortOptions = []
    PersistentBackup = False
    ProgressFlags = [mtpcLoad, mtpcSave, mtpcCopy]
    LoadedCompletely = True
    SavedCompletely = False
    FilterOptions = []
    Version = '7.70.00 Professional Edition'
    LanguageID = 0
    SortID = 0
    SubLanguageID = 1
    LocaleID = 1024
    OnCalcFields = mdAcdSrcCalcFields
    Left = 144
    Top = 16
    object mdAcdSrcIPHE001: TIntegerField
      DisplayLabel = #20358#38651#32232#34399
      FieldName = 'IPHE001'
      ReadOnly = True
      Required = True
    end
    object mdAcdSrcRPHE001: TIntegerField
      DisplayLabel = #22238#38651#32232#34399
      FieldName = 'RPHE001'
      ReadOnly = True
    end
    object mdAcdSrcIPHE004: TDateTimeField
      DisplayLabel = #20358#38651#26178#38291
      FieldName = 'IPHE004'
      Required = True
    end
    object mdAcdSrcIPHE005: TStringField
      DisplayLabel = #23458#25142#20195#34399
      FieldName = 'IPHE005'
      Required = True
      Size = 10
    end
    object mdAcdSrcIPHE003: TStringField
      DisplayLabel = #20358#38651#27512#23660#37096#38272
      FieldName = 'IPHE003'
      Required = True
      Size = 10
    end
    object mdAcdSrcRPHE003: TStringField
      DisplayLabel = #22238#38651#20154#20195#34399
      FieldName = 'RPHE003'
      ReadOnly = True
      Size = 10
    end
    object mdAcdSrcSALE002: TStringField
      DisplayLabel = #22238#38651#20154
      FieldName = 'SALE002'
      ReadOnly = True
      Size = 10
    end
    object mdAcdSrcSALE003: TStringField
      DisplayLabel = #22238#38651#37096#38272
      FieldName = 'SALE003'
      ReadOnly = True
      Size = 10
    end
    object mdAcdSrcRPHE005: TDateTimeField
      DisplayLabel = #22238#38651#26178#38291
      FieldName = 'RPHE005'
      ReadOnly = True
    end
    object mdAcdSrcRPHE006: TDateTimeField
      DisplayLabel = #23436#25104#26178#38291
      FieldName = 'RPHE006'
      ReadOnly = True
    end
    object mdAcdSrcIPHE016: TDateTimeField
      DisplayLabel = #38283#22987#34389#29702#26178#38291
      FieldName = 'IPHE016'
    end
    object mdAcdSrcIPHE017: TDateTimeField
      DisplayLabel = #32080#26696#26178#38291
      FieldName = 'IPHE017'
    end
    object mdAcdSrcIPHE012: TStringField
      DisplayLabel = #31995#32113#21029
      FieldName = 'IPHE012'
      Size = 2
    end
    object mdAcdSrcIPHE008: TStringField
      DisplayLabel = #30041#35328#20154
      FieldName = 'IPHE008'
      Size = 10
    end
    object mdAcdSrcIPHE019: TStringField
      DisplayLabel = #20358#38651#21312#20998
      FieldName = 'IPHE019'
      Size = 2
    end
    object mdAcdSrcAnswerInTime: TBooleanField
      FieldKind = fkCalculated
      FieldName = 'AnswerInTime'
      Calculated = True
    end
    object mdAcdSrcValid: TBooleanField
      DefaultExpression = 'False'
      DisplayLabel = #26377#25928
      FieldName = 'Valid'
    end
    object mdAcdSrcCLAS004: TStringField
      DisplayLabel = #31995#32113#21029
      FieldName = 'CLAS004'
      Size = 40
    end
    object mdAcdSrcCALL_KIND: TStringField
      DisplayLabel = #20358#38651#21312#20998
      FieldName = 'CALL_KIND'
    end
    object mdAcdSrcRPHE011: TStringField
      DisplayLabel = #28961#25928#22238#38651
      FieldName = 'RPHE011'
      Size = 1
    end
    object mdAcdSrcExcept: TBooleanField
      DefaultExpression = 'False'
      DisplayLabel = #20363#22806
      FieldName = 'Except'
    end
    object mdAcdSrcRemark: TStringField
      DisplayLabel = #20633#35387
      FieldName = 'Remark'
      Size = 200
    end
    object mdAcdSrcRPHE011_REV: TBooleanField
      DisplayLabel = #26377#25928#22238#38651
      FieldKind = fkCalculated
      FieldName = 'RPHE011_REV'
      Calculated = True
    end
    object mdAcdSrcCUT1002: TStringField
      DisplayLabel = #23458#25142#31777#31281
      FieldName = 'CUT1002'
      Size = 10
    end
    object mdAcdSrcIPH2001: TStringField
      DisplayLabel = #21512#32004#36523#20998
      FieldName = 'IPH2001'
      Size = 2
    end
    object mdAcdSrcIPH2002: TIntegerField
      DisplayLabel = #22238#38651#24310#36978#26178#38291
      FieldName = 'IPH2002'
    end
    object mdAcdSrcIPH2003: TBooleanField
      DisplayLabel = #36926#26178#22238#38651
      FieldName = 'IPH2003'
    end
    object mdAcdSrcFLAG_SW: TStringField
      FieldName = 'FLAG_SW'
      Size = 2
    end
    object mdAcdSrcFLAG_HRS: TStringField
      FieldName = 'FLAG_HRS'
      Size = 2
    end
    object mdAcdSrcGUID: TStringField
      FieldName = 'GUID'
      Size = 40
    end
    object mdAcdSrcIPH2004: TStringField
      FieldName = 'IPH2004'
      Size = 40
    end
    object mdAcdSrcRPHE_GUID: TStringField
      FieldName = 'RPHE_GUID'
      Size = 40
    end
    object mdAcdSrcTK: TStringField
      DisplayLabel = #38651#35441#35672#21029#30908
      FieldName = 'TK'
    end
    object mdAcdSrcDUP: TBooleanField
      DefaultExpression = 'False'
      DisplayLabel = #35672#21029#30908#37325#35079
      FieldName = 'DUP'
    end
    object mdAcdSrcSALE024: TIntegerField
      DisplayLabel = 'ACD'#26085#27161#28310
      FieldName = 'SALE024'
    end
  end
  object qrGetData_Public: TUniQuery
    LocalUpdate = True
    Connection = connReport
    FetchRows = 100
    Left = 80
    Top = 48
  end
  object cmdInsIPH2: TUniSQL
    Connection = connReport
    SQL.Strings = (
      'INSERT INTO '
      'WICSIPH2(GUID, IPH2001, IPH2002, IPH2003, IPH2004)'
      'VALUES(:GUID, :IPH2001, :IPH2002, :IPH2003, :IPH2004)')
    Left = 144
    Top = 80
    ParamData = <
      item
        DataType = ftString
        Name = 'GUID'
      end
      item
        DataType = ftString
        Name = 'IPH2001'
      end
      item
        DataType = ftInteger
        Name = 'IPH2002'
      end
      item
        DataType = ftBoolean
        Name = 'IPH2003'
      end
      item
        DataType = ftString
        Name = 'IPH2004'
      end>
  end
  object mdAcdSwDaily: TdxMemData
    Indexes = <>
    SortOptions = []
    OnCalcFields = mdAcdSiteDailyCalcFields
    Left = 112
    Top = 16
    object mdAcdSwDailyPhoneDate: TDateField
      DisplayLabel = #20358#38651#26085#26399
      FieldName = 'PhoneDate'
    end
    object mdAcdSwDailySiteId: TStringField
      DisplayLabel = #29151#26989#34389
      FieldName = 'SiteId'
    end
    object mdAcdSwDailyACD_Total: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#27966#36865#32317#25976
      FieldName = 'ACD_Total'
      DisplayFormat = '#,0'
    end
    object mdAcdSwDailyACD_Assign_Invalid: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#25351#23450#22238#38651#28961#25928#25976
      FieldName = 'ACD_Assign_Invalid'
      DisplayFormat = '#,0'
    end
    object mdAcdSwDailyACD_InvalidIn_Total: TIntegerField
      DisplayLabel = #38750#35347#32244#37096#38272#27966#36865#25976
      FieldName = 'ACD_InvalidIn_Total'
    end
    object mdAcdSwDailyACD_ValidIn_Total: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#26377#25928#27966#36865#25976
      FieldName = 'ACD_ValidIn_Total'
      DisplayFormat = '#,0'
    end
    object mdAcdSwDailyACD_ValidAns_Total: TIntegerField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#34389#29702#25976
      FieldName = 'ACD_ValidAns_Total'
      DisplayFormat = '#,0'
    end
    object mdAcdSwDailyACD_Score: TFloatField
      DefaultExpression = '0'
      DisplayLabel = 'ACD'#25509#32893#29575
      FieldName = 'ACD_Score'
      DisplayFormat = '#.0 %'
    end
    object mdAcdSwDailySiteName: TStringField
      DisplayLabel = #29151#26989#34389
      FieldKind = fkCalculated
      FieldName = 'SiteName'
      Calculated = True
    end
    object mdAcdSwDailyTE_Total_C: TIntegerField
      DisplayLabel = 'TE'#21512#32004
      FieldName = 'TE_Total_C'
      DisplayFormat = '#,0'
    end
    object mdAcdSwDailyTE_Total_NC: TIntegerField
      DisplayLabel = 'TE'#38750#21512#32004
      FieldName = 'TE_Total_NC'
      DisplayFormat = '#,0'
    end
    object mdAcdSwDailyPhone_Count_C: TIntegerField
      DisplayLabel = #21512#32004#20358#38651
      FieldName = 'Phone_Count_C'
      DisplayFormat = '#,0'
    end
    object mdAcdSwDailyPhone_Count_NC: TIntegerField
      DisplayLabel = #38750#21512#32004#20358#38651
      FieldName = 'Phone_Count_NC'
      DisplayFormat = '#,0'
    end
    object mdAcdSwDailyTimeOut_Count_C: TIntegerField
      DisplayLabel = #21512#32004#36926#26178
      FieldName = 'TimeOut_Count_C'
    end
    object mdAcdSwDailyTimeOut_Rate_C: TFloatField
      DisplayLabel = #21512#32004#36926#26178#29575
      FieldName = 'TimeOut_Rate_C'
      OnGetText = mdAcdSiteDailyTimeOut_Rate_CGetText
    end
    object mdAcdSwDailyTimeOut_Count_NC: TIntegerField
      DisplayLabel = #38750#21512#32004#36926#26178
      FieldName = 'TimeOut_Count_NC'
    end
    object mdAcdSwDailyTimeOut_Rate_NC: TFloatField
      DisplayLabel = #38750#21512#32004#36926#26178#29575
      FieldName = 'TimeOut_Rate_NC'
      OnGetText = mdAcdSiteDailyTimeOut_Rate_CGetText
    end
    object mdAcdSwDailySw: TStringField
      DisplayLabel = #31995#32113
      FieldName = 'Sw'
    end
    object mdAcdSwDailyPhoneYear: TIntegerField
      DisplayLabel = #20358#38651#24180
      FieldName = 'PhoneYear'
    end
    object mdAcdSwDailyPhoneMonth: TIntegerField
      DisplayLabel = #20358#38651#26376
      FieldName = 'PhoneMonth'
    end
  end
  object cmdUpdIPH2: TUniSQL
    Connection = connReport
    SQL.Strings = (
      
        'UPDATE WICSIPH2 SET IPH2001 = :IPH2001, IPH2002 = :IPH2002, IPH2' +
        '003 = :IPH2003, IPH2004 = :IPH2004'
      'WHERE (GUID = :GUID)')
    Left = 112
    Top = 80
    ParamData = <
      item
        DataType = ftString
        Name = 'IPH2001'
      end
      item
        DataType = ftInteger
        Name = 'IPH2002'
      end
      item
        DataType = ftBoolean
        Name = 'IPH2003'
      end
      item
        DataType = ftString
        Name = 'IPH2004'
      end
      item
        DataType = ftString
        Name = 'GUID'
      end>
  end
  object qrGetData_PhoneOut: TUniQuery
    LocalUpdate = True
    Connection = connReport
    SQL.Strings = (
      'SELECT '
      
        'RPHE001, RPHE003, SALE002, DEPT001, DEPT002, RPHE005, RPHE011, I' +
        'PHE005, CUT1002, IPH2001'
      'FROM WICSRPHE R WITH(NOLOCK)'
      'LEFT JOIN WICSRSCE B WITH(NOLOCK) ON R.GUID = RSCE001'
      'LEFT JOIN WICSIPHE A WITH(NOLOCK) ON RSCE003 = A.GUID'
      'LEFT JOIN WICSIPH2 D WITH(NOLOCK) ON D.GUID = A.GUID'
      'LEFT JOIN WICSSALE S WITH(NOLOCK) ON SALE001 = RPHE003'
      'LEFT JOIN WICSCUT1 C WITH(NOLOCK) ON IPHE005 = CUT1001'
      'LEFT JOIN WICSDEPT T WITH(NOLOCK) ON DEPT001 = SALE003'
      'WHERE'
      '(RPHE005 >= :RPHE005B AND RPHE005 <= :RPHE005E)'
      
        'GROUP BY RPHE001, RPHE003, SALE002, DEPT001, DEPT002, RPHE005, R' +
        'PHE011, IPHE005, CUT1002, IPH2001'
      'ORDER BY RPHE001, IPH2001')
    FetchRows = 100
    Left = 144
    Top = 48
    ParamData = <
      item
        DataType = ftDateTime
        Name = 'RPHE005B'
        Value = 42228d
      end
      item
        DataType = ftDateTime
        Name = 'RPHE005E'
        Value = 42229d
      end>
    object qrGetData_PhoneOutRPHE001: TIntegerField
      AutoGenerateValue = arAutoInc
      FieldName = 'RPHE001'
      ReadOnly = True
      Required = True
    end
    object qrGetData_PhoneOutRPHE003: TStringField
      FieldName = 'RPHE003'
      Size = 10
    end
    object qrGetData_PhoneOutSALE002: TStringField
      FieldName = 'SALE002'
      ReadOnly = True
      Size = 10
    end
    object qrGetData_PhoneOutRPHE005: TDateTimeField
      FieldName = 'RPHE005'
    end
    object qrGetData_PhoneOutRPHE011: TStringField
      FieldName = 'RPHE011'
      Size = 1
    end
    object qrGetData_PhoneOutIPHE005: TStringField
      FieldName = 'IPHE005'
      ReadOnly = True
      Size = 10
    end
    object qrGetData_PhoneOutCUT1002: TStringField
      FieldName = 'CUT1002'
      ReadOnly = True
      Size = 10
    end
    object qrGetData_PhoneOutIPH2001: TStringField
      FieldName = 'IPH2001'
      ReadOnly = True
      Size = 2
    end
    object qrGetData_PhoneOutDEPT002: TStringField
      FieldName = 'DEPT002'
      ReadOnly = True
    end
    object qrGetData_PhoneOutDEPT001: TStringField
      FieldName = 'DEPT001'
      ReadOnly = True
      Size = 10
    end
  end
  object mdPhoneOutSrc: TkbmMemTable
    DesignActivation = True
    AttachedAutoRefresh = True
    AttachMaxCount = 1
    FieldDefs = <
      item
        Name = 'RPHE001'
        Attributes = [faReadonly, faRequired]
        DataType = ftInteger
      end
      item
        Name = 'RPHE003'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'SALE002'
        Attributes = [faReadonly]
        DataType = ftString
        Size = 10
      end
      item
        Name = 'RPHE005'
        DataType = ftDateTime
      end
      item
        Name = 'RPHE011'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'IPHE005'
        Attributes = [faReadonly]
        DataType = ftString
        Size = 10
      end
      item
        Name = 'CUT1002'
        Attributes = [faReadonly]
        DataType = ftString
        Size = 10
      end
      item
        Name = 'IPH2001'
        Attributes = [faReadonly]
        DataType = ftString
        Size = 2
      end
      item
        Name = 'Except'
        DataType = ftBoolean
      end>
    IndexDefs = <>
    SortOptions = []
    PersistentBackup = False
    ProgressFlags = [mtpcLoad, mtpcSave, mtpcCopy]
    LoadedCompletely = False
    SavedCompletely = False
    FilterOptions = []
    Version = '7.70.00 Professional Edition'
    LanguageID = 0
    SortID = 0
    SubLanguageID = 1
    LocaleID = 1024
    OnCalcFields = mdPhoneOutSrcCalcFields
    Left = 16
    Top = 48
    object mdPhoneOutSrcRPHE001: TIntegerField
      DisplayLabel = #22238#38651#32232#34399
      FieldName = 'RPHE001'
      ReadOnly = True
      Required = True
    end
    object mdPhoneOutSrcRPHE003: TStringField
      DisplayLabel = #22238#38651#20154#20195#34399
      FieldName = 'RPHE003'
      Size = 10
    end
    object mdPhoneOutSrcSALE002: TStringField
      DisplayLabel = #22238#38651#20154
      FieldName = 'SALE002'
      ReadOnly = True
      Size = 10
    end
    object mdPhoneOutSrcRPHE005: TDateTimeField
      DisplayLabel = #22238#38651#26178#38291
      FieldName = 'RPHE005'
    end
    object mdPhoneOutSrcRPHE011: TStringField
      DisplayLabel = #28961#25928#22238#38651
      FieldName = 'RPHE011'
      Size = 1
    end
    object mdPhoneOutSrcIPHE005: TStringField
      DisplayLabel = #23458#25142#20195#34399
      FieldName = 'IPHE005'
      ReadOnly = True
      Size = 10
    end
    object mdPhoneOutSrcCUT1002: TStringField
      DisplayLabel = #23458#25142#31777#31281
      FieldName = 'CUT1002'
      ReadOnly = True
      Size = 10
    end
    object mdPhoneOutSrcIPH2001: TStringField
      DisplayLabel = #21512#32004#36523#20998
      FieldName = 'IPH2001'
      ReadOnly = True
      Size = 2
    end
    object mdPhoneOutSrcInvalid: TBooleanField
      DefaultExpression = 'False'
      DisplayLabel = #20363#22806
      FieldName = 'Except'
    end
    object mdPhoneOutSrcRPHE011_REV: TBooleanField
      DisplayLabel = #26377#25928#22238#38651
      FieldKind = fkCalculated
      FieldName = 'RPHE011_REV'
      Calculated = True
    end
    object mdPhoneOutSrcDEPT002: TStringField
      DisplayLabel = #22238#38651#37096#38272
      FieldName = 'DEPT002'
    end
    object mdPhoneOutSrcDEPT001: TStringField
      DisplayLabel = #22238#38651#37096#38272#20195#34399
      FieldName = 'DEPT001'
      Size = 10
    end
  end
  object connReport: TUniConnection
    ProviderName = 'SQL Server'
    Database = 'WCRM'
    SpecificOptions.Strings = (
      'SQL Server.ConnectionTimeout=60')
    Username = 'tcrm'
    Server = '10.1.2.7'
    LoginPrompt = False
    Left = 48
    Top = 48
    EncryptedPassword = '97FF9AFF93FF93FF90FF8BFF9CFF8DFF92FF'
  end
  object mdDupCheck: TdxMemData
    Indexes = <>
    SortOptions = []
    Left = 80
    Top = 16
    object mdDupCheckTK: TStringField
      FieldName = 'TK'
      Size = 50
    end
    object mdDupCheckCount: TIntegerField
      FieldName = 'Count'
    end
  end
  object qrGetIPH2004: TUniQuery
    Connection = connReport
    SQL.Strings = (
      'SELECT TOP 1 R.GUID'
      'FROM WICSRPHE R WITH(NOLOCK) '
      'LEFT JOIN WICSRSCE B WITH(NOLOCK) ON RSCE001 = R.GUID'
      'LEFT JOIN WICSIPHE I WITH(NOLOCK) ON RSCE003 = I.GUID'
      'WHERE IPHE001 = :IPHE001'
      'ORDER BY RPHE005')
    Left = 16
    Top = 80
    ParamData = <
      item
        DataType = ftInteger
        Name = 'IPHE001'
      end>
    object qrGetIPH2004GUID: TStringField
      FieldName = 'GUID'
      Required = True
      Size = 40
    end
  end
  object XLSReadWriteII51: TXLSReadWriteII5
    ComponentVersion = '5.20.67a'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 176
    Top = 48
  end
  object qrWICSIPHH: TUniQuery
    Connection = dmReport.UniConnWinton
    SQL.Strings = (
      'SELECT * FROM WICSIPHH'
      'WHERE (IPHH001 >= :IPHH001B AND IPHH001 < :IPHH001E)')
    Left = 244
    Top = 20
    ParamData = <
      item
        DataType = ftDateTime
        Name = 'IPHH001B'
        ParamType = ptInput
      end
      item
        DataType = ftDateTime
        Name = 'IPHH001E'
        ParamType = ptInput
      end>
  end
  object qrPrecedingAcdTotal: TUniQuery
    Connection = dmReport.UniConnWinton
    SQL.Strings = (
      
        'SELECT IPH5003, DATEPART(year, IPH5002) AS _YEAR_, SUM(IPH5007) ' +
        'AS _IPH5007_SUM_'
      'FROM WICSIPH5'
      'WHERE IPH5002 >= :IPH5002B AND IPH5002 < :IPH5002E'
      'GROUP BY IPH5003, DATEPART(year, IPH5002)'
      'UNION'
      '('
      
        'SELECT IPH5003, DATEPART(year, IPH5002) AS _YEAR_, SUM(IPH5007) ' +
        'AS _IPH5007_SUM_'
      'FROM WICSIPH5'
      'WHERE IPH5002 >= :PREV_IPH5002B AND IPH5002 < :PREV_IPH5002E'
      'GROUP BY IPH5003, DATEPART(year, IPH5002)'
      ')')
    Left = 240
    Top = 90
    ParamData = <
      item
        DataType = ftDateTime
        Name = 'IPH5002B'
        ParamType = ptInput
      end
      item
        DataType = ftDateTime
        Name = 'IPH5002E'
        ParamType = ptInput
      end
      item
        DataType = ftDateTime
        Name = 'PREV_IPH5002B'
        ParamType = ptInput
      end
      item
        DataType = ftDateTime
        Name = 'PREV_IPH5002E'
        ParamType = ptInput
      end>
  end
end
