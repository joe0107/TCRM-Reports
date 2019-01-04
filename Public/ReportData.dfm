object dmReport: TdmReport
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Height = 394
  Width = 695
  object SQLServerUniProvider: TSQLServerUniProvider
    Left = 483
    Top = 16
  end
  object connTeleContact: TUniConnection
    ProviderName = 'SQL Server'
    Database = 'ReportDB'
    Username = 'telecontact'
    Server = '10.1.2.16'
    LoginPrompt = False
    Left = 37
    Top = 77
    EncryptedPassword = '8BFF9AFF93FF9AFF9CFF90FF91FF8BFF9EFF9CFF8BFF'
  end
  object UniConnTCRM: TUniConnection
    ProviderName = 'SQL Server'
    Database = 'WCRM'
    Options.KeepDesignConnected = False
    Username = 'tcrm'
    Server = '10.1.2.100'
    LoginPrompt = False
    Left = 33
    Top = 16
    EncryptedPassword = '97FF9AFF93FF93FF90FF8BFF9CFF8DFF92FF'
  end
  object UniConnWinton: TUniConnection
    ProviderName = 'SQL Server'
    Database = 'WCRM'
    Username = 'tcrm'
    Server = '10.1.1.212'
    LoginPrompt = False
    AfterConnect = DataModuleCreate
    Left = 258
    Top = 16
    EncryptedPassword = '97FF9AFF93FF93FF90FF8BFF9CFF8DFF92FF'
  end
  object connTcrmPublic: TUniConnection
    ProviderName = 'SQL Server'
    Database = 'WCRM'
    Username = 'tcrm'
    Server = '10.1.2.7'
    LoginPrompt = False
    Left = 149
    Top = 16
    EncryptedPassword = '97FF9AFF93FF93FF90FF8BFF9CFF8DFF92FF'
  end
  object qrCust: TUniQuery
    Connection = UniConnTCRM
    SQL.Strings = (
      'SELECT CUT1001, CUT1002, FLAG_SW, FLAG_HRS, FLAG_HW'
      'FROM WICSCUT1 WITH(NOLOCK)'
      'ORDER BY CUT1001')
    Left = 33
    Top = 138
    object qrCustCUT1001: TStringField
      DisplayLabel = #23458#20195
      FieldName = 'CUT1001'
      Required = True
      Size = 10
    end
    object qrCustCUT1002: TStringField
      DisplayLabel = #31777#31281
      FieldName = 'CUT1002'
      Size = 10
    end
    object qrCustFLAG_SW: TStringField
      DisplayLabel = #36575#32004
      FieldName = 'FLAG_SW'
      Size = 2
    end
    object qrCustFLAG_HRS: TStringField
      DisplayLabel = 'HRS'
      FieldName = 'FLAG_HRS'
      Size = 2
    end
    object qrCustFLAG_HW: TStringField
      DisplayLabel = #30828#32004
      FieldName = 'FLAG_HW'
      Size = 2
    end
  end
  object dsCust: TDataSource
    DataSet = qrCust
    Left = 62
    Top = 138
  end
  object qrTitle: TUniQuery
    Connection = UniConnTCRM
    SQL.Strings = (
      
        'SELECT CLAS002, CLAS004 FROM WICSCLAS WHERE CLAS001 = '#39'02'#39' ORDER' +
        ' BY CLAS002')
    Left = 145
    Top = 138
    object qrTitleCLAS002: TStringField
      DisplayLabel = #32232#34399
      FieldName = 'CLAS002'
      Required = True
      Size = 2
    end
    object qrTitleCLAS004: TStringField
      DisplayLabel = #32887#31281
      FieldName = 'CLAS004'
      Required = True
      Size = 40
    end
  end
  object dsTitle: TDataSource
    DataSet = qrTitle
    Left = 175
    Top = 138
  end
  object qrDbVersion: TUniQuery
    Connection = UniConnTCRM
    SQL.Strings = (
      
        'SELECT * FROM WICSOPTN WITH(NOLOCK) WHERE OPTN000 = '#39#36039#26009#24235#29256#26412#39' ORDE' +
        'R BY OPTN001')
    Left = 258
    Top = 138
    object qrDbVersionGUID: TStringField
      FieldName = 'GUID'
      Required = True
      Size = 40
    end
    object qrDbVersionOPTN000: TStringField
      FieldName = 'OPTN000'
      Required = True
      Size = 30
    end
    object qrDbVersionOPTN001: TStringField
      DisplayLabel = #20195#34399
      FieldName = 'OPTN001'
      Size = 10
    end
    object qrDbVersionOPTN002: TStringField
      DisplayLabel = #36039#26009#24235
      FieldName = 'OPTN002'
      Size = 30
    end
    object qrDbVersionOPTN003: TStringField
      DisplayLabel = #29256#26412
      FieldName = 'OPTN003'
      Size = 50
    end
    object qrDbVersionOPTN004: TStringField
      FieldName = 'OPTN004'
      Size = 250
    end
  end
  object dsDbVersion: TDataSource
    DataSet = qrDbVersion
    Left = 289
    Top = 138
  end
  object qrEmp: TUniQuery
    Connection = UniConnTCRM
    SQL.Strings = (
      'SELECT SALE001, SALE002'
      'FROM WICSSALE WITH(NOLOCK)'
      'ORDER BY SALE001')
    Left = 33
    Top = 199
    object qrEmpSALE001: TStringField
      DisplayLabel = #20195#34399
      FieldName = 'SALE001'
      Required = True
      Size = 10
    end
    object qrEmpSALE002: TStringField
      DisplayLabel = #22995#21517
      FieldName = 'SALE002'
      Size = 10
    end
    object qrEmpSALE_DESC: TStringField
      DisplayLabel = #20154#21729
      FieldKind = fkCalculated
      FieldName = 'SALE_DESC'
      Calculated = True
    end
  end
  object dsEmp: TDataSource
    DataSet = qrEmp
    Left = 62
    Top = 199
  end
  object dsDept: TDataSource
    DataSet = qrDept
    Left = 175
    Top = 199
  end
  object qrDept: TUniQuery
    Connection = UniConnTCRM
    SQL.Strings = (
      'SELECT DEPT001, DEPT002'
      'FROM WICSDEPT WITH(NOLOCK)'
      'ORDER BY DEPT001')
    Left = 145
    Top = 199
    object qrDeptDEPT001: TStringField
      DisplayLabel = #20195#34399
      FieldName = 'DEPT001'
      Required = True
      Size = 10
    end
    object qrDeptDEPT002: TStringField
      DisplayLabel = #37096#38272
      FieldName = 'DEPT002'
    end
  end
  object qrClass_A1: TUniQuery
    Connection = UniConnTCRM
    SQL.Strings = (
      'SELECT GUID, CLAS002, CLAS004 FROM WICSCLAS WITH(NOLOCK)'
      'WHERE CLAS001 = '#39'12'#39
      'ORDER BY CLAS002')
    Left = 370
    Top = 138
    object qrClass_A1GUID: TStringField
      FieldName = 'GUID'
      Required = True
      Size = 40
    end
    object qrClass_A1CLAS002: TStringField
      DisplayLabel = #20195#34399
      FieldName = 'CLAS002'
      Required = True
      Size = 2
    end
    object qrClass_A1CLAS004: TStringField
      DisplayLabel = #35498#26126
      FieldName = 'CLAS004'
      Required = True
      Size = 40
    end
  end
  object dsClass_A1: TDataSource
    DataSet = qrClass_A1
    Left = 400
    Top = 138
  end
  object qrClass_J0: TUniQuery
    Connection = UniConnTCRM
    SQL.Strings = (
      'SELECT GUID, CLAS002, CLAS004 FROM WICSCLAS WITH(NOLOCK)'
      'WHERE CLAS001 = '#39'J0'#39
      'ORDER BY CLAS002')
    Left = 370
    Top = 199
    object StringField1: TStringField
      FieldName = 'GUID'
      Required = True
      Size = 40
    end
    object StringField2: TStringField
      DisplayLabel = #20195#34399
      FieldName = 'CLAS002'
      Required = True
      Size = 2
    end
    object StringField3: TStringField
      DisplayLabel = #31995#32113
      FieldName = 'CLAS004'
      Required = True
      Size = 40
    end
  end
  object dsClass_J0: TDataSource
    DataSet = qrClass_J0
    Left = 400
    Top = 199
  end
  object qrClass_J1: TUniQuery
    Connection = UniConnTCRM
    SQL.Strings = (
      
        'SELECT GUID, CLAS002, CLAS003, CLAS004, CLAS005 FROM WICSCLAS WI' +
        'TH(NOLOCK)'
      'WHERE CLAS001 = '#39'J1'#39
      'ORDER BY CLAS002')
    Left = 258
    Top = 199
    object qrClass_J1GUID: TStringField
      FieldName = 'GUID'
      Required = True
      Size = 40
    end
    object qrClass_J1CLAS002: TStringField
      FieldName = 'CLAS002'
      Required = True
      Size = 2
    end
    object qrClass_J1CLAS003: TStringField
      FieldName = 'CLAS003'
      Size = 2
    end
    object qrClass_J1CLAS004: TStringField
      FieldName = 'CLAS004'
      Required = True
      Size = 40
    end
    object qrClass_J1CLAS005: TMemoField
      FieldName = 'CLAS005'
      BlobType = ftMemo
    end
  end
  object dsClass_J1: TDataSource
    DataSet = qrClass_J1
    Left = 289
    Top = 199
  end
  object UniConnLvUpd: TUniConnection
    ProviderName = 'SQL Server'
    Database = 'WintonLiveUpdate'
    Username = 'WintonLvUpd2015'
    Server = '10.1.0.203'
    LoginPrompt = False
    Left = 149
    Top = 76
    EncryptedPassword = 
      'CDFF8EFF88FFDCFF9AFF9BFF8DFFDAFF8BFF86FFC9FFD9FFC7FF8AFF95FF96FF' +
      'D7FF90FFD6FF'
  end
  object qrWICSCUT7: TUniQuery
    Connection = UniConnTCRM
    SQL.Strings = (
      'SELECT FLAG'
      ',CUT7001, CUT7002, CUT7003, CUT7004, CUT7005'
      ',CUT7006, CUT7007, CUT7008, CUT7009, CUT7010'
      ',CUT7013, CUT7016, CUT7017'
      'FROM WICSCUT7 WITH(NOLOCK)')
    Left = 33
    Top = 260
  end
  object SQLiteUniProvider: TSQLiteUniProvider
    Left = 483
    Top = 84
  end
  object UniConnReport: TUniConnection
    ProviderName = 'SQLite'
    Database = 'C:\Projects\TCRM-Group\TCRM-Reports\TcrmReport.db'
    LoginPrompt = False
    Left = 370
    Top = 16
  end
  object qrClass_10: TUniQuery
    Connection = UniConnTCRM
    SQL.Strings = (
      'SELECT GUID, CLAS002, CLAS004 FROM WICSCLAS WITH(NOLOCK)'
      'WHERE CLAS001 = '#39'10'#39
      'ORDER BY CLAS002')
    Left = 145
    Top = 260
    object qrClass_10GUID: TStringField
      FieldName = 'GUID'
      Required = True
      Size = 40
    end
    object qrClass_10CLAS002: TStringField
      DisplayLabel = #20195#34399
      FieldName = 'CLAS002'
      Required = True
      Size = 2
    end
    object qrClass_10CLAS004: TStringField
      DisplayLabel = #31995#32113#21029
      FieldName = 'CLAS004'
      Required = True
      Size = 40
    end
  end
  object dsClass_10: TDataSource
    DataSet = qrClass_10
    Left = 175
    Top = 260
  end
  object UniConnWcrm: TUniConnection
    ProviderName = 'SQL Server'
    Database = 'Winton_WCRM'
    Username = 'wcrm'
    Server = '10.1.1.1'
    LoginPrompt = False
    AfterConnect = DataModuleCreate
    Left = 262
    Top = 76
    EncryptedPassword = 'CFFFCFFFCFFFCFFF'
  end
  object qrTcrmConfig: TUniQuery
    Connection = UniConnReport
    SQL.Strings = (
      'SELECT * FROM TcrmConfig')
    Left = 492
    Top = 140
    object qrTcrmConfigSiteID: TStringField
      FieldName = 'SiteID'
      Size = 2
    end
    object qrTcrmConfigSite: TStringField
      FieldName = 'Site'
      Size = 10
    end
    object qrTcrmConfigBranch: TStringField
      FieldName = 'Branch'
      Size = 10
    end
    object qrTcrmConfigServer: TStringField
      FieldName = 'Server'
    end
    object qrTcrmConfigDatabase: TStringField
      FieldName = 'Database'
      Size = 255
    end
    object qrTcrmConfigTE_Admin_Email: TStringField
      FieldName = 'TE_Admin_Email'
      Size = 100
    end
    object qrTcrmConfigSite_Admin_Email: TStringField
      FieldName = 'Site_Admin_Email'
      Size = 100
    end
    object qrTcrmConfigTE_Leader_Email: TStringField
      FieldName = 'TE_Leader_Email'
      Size = 200
    end
  end
end
