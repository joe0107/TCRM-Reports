object dmSitePhoneSummary: TdmSitePhoneSummary
  OldCreateOrder = True
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  Height = 174
  Width = 257
  object qrGetData: TUniQuery
    LocalUpdate = True
    Connection = connReport
    SQL.Strings = (
      'SELECT'
      'RID, IPH4002, IPH4001, IPH4003, IPH4004, '
      'IPH4005, IPH4006, IPH4007, IPH4008, IPH4009,'
      'IPH4010, IPH4011, IPH4012, IPH4013, IPH4205, '
      'IPH4209, IPH4210, IPH4211, IPH4213, IPH4099'
      'FROM WICSIPH4 WITH(NOLOCK)'
      'WHERE (IPH4002 BETWEEN :IPH4002B AND :IPH4002E)'
      'AND (IPH4001 = :IPH4001)'
      'ORDER BY IPH4002')
    FetchRows = 100
    Left = 48
    Top = 16
    ParamData = <
      item
        DataType = ftDateTime
        Name = 'IPH4002B'
        ParamType = ptInput
      end
      item
        DataType = ftDateTime
        Name = 'IPH4002E'
        ParamType = ptInput
      end
      item
        DataType = ftString
        Name = 'IPH4001'
        ParamType = ptInput
      end>
  end
  object connReport: TUniConnection
    ProviderName = 'SQL Server'
    Database = 'WCRM'
    SpecificOptions.Strings = (
      'SQL Server.ConnectionTimeout=60')
    Username = 'tcrm'
    Server = '10.1.2.7'
    LoginPrompt = False
    Left = 16
    Top = 16
    EncryptedPassword = '97FF9AFF93FF93FF90FF8BFF9CFF8DFF92FF'
  end
  object XLSRW: TXLSReadWriteII5
    ComponentVersion = '5.20.67a'
    Version = xvExcel97
    DirectRead = False
    DirectWrite = False
    Left = 80
    Top = 16
  end
end
