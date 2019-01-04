unit PhoneAnalysis;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs, DB, StdCtrls,
  ADODB, ShellAPI, cxData, cxClasses, cxCustomData, cxDataStorage, cxDBData, JclStrings, MemDS, Uni,
  DBAccess, kbmMemTable, DateUtils, dxmdaset, XLSSheetData5, XLSReadWriteII5, XLSDbRead5, XLSNames5,
  IdEMailAddress, IdMessage, IdAttachmentFile;

type
  TdmPhoneAnalysis = class(TDataModule)
    XLSRW: TXLSReadWriteII5;
    procedure DataModuleCreate(Sender: TObject);
  protected
    procedure InitExecute;
    procedure BeginExecute;
    procedure EndExecute;
    procedure Log(AMsg: string);
    procedure LogLine;
  private
    FYear, FMonth: Word;
    FWICSIPH4, FWICSIPH5: TUniQuery;
    procedure GetData_IPH5(ABegTime, AEndTime: TDateTime);
    procedure PrepareData_IPH5_Y_1_M(Y, M: Word);
    procedure PrepareData_IPH5_Y_M(Y, M: Word);
    procedure GetData_IPH4(ABegTime, AEndTime: TDateTime);
    procedure PrepareData_IPH4_Y_1_M(Y, M: Word);
    procedure PrepareData_IPH4_Y_M(Y, M: Word);
  private
    FXlsFileName: string;
    procedure Init_XLS_Report;
    procedure WriteDataToXls_Titles(ASheet: TXLSWorksheet; Y, M: Word);
    procedure WriteDataToXls_PhoneCount(ASheet: TXLSWorksheet; AThisYear, AThisMonth: Boolean);
    procedure WriteDataToXls_Rate(ASheet: TXLSWorksheet; AThisYear, AThisMonth: Boolean);
    procedure WriteDataToXls_RateSummary(ASheet: TXLSWorksheet; AThisYear, AThisMonth: Boolean);
    function  GetColOffset_Sw(AThisYear, AThisMonth: Boolean): Integer;
    function  GetColOffset_Rate(AThisYear, AThisMonth: Boolean; AKind: Integer): Integer;
    function  GetRowOffset_Site(ASite: string): Integer;
    function  GetRowOffset_Sw(ASw: string): Integer;
    function  CalcColRow_Sw(ASite, ASw: string; AThisYear, AThisMonth: Boolean; var ACol, ARow: Word): Boolean;
    function  CalcColRow_Rate(ASite: string; AThisYear, AThisMonth: Boolean; AKind: Integer; var ACol, ARow: Word): Boolean;
    function  CopyReportFromTemplate(ADate: TDateTime): string;
    function  FindSheet(ASheetName: string): TXLSWorksheet;
  private
    procedure SendMail;
    function  MakeRecipients(AEmailAddrList: TIdEmailAddressList): Integer;
    procedure MakeCCList(AEmailAddrList: TIdEmailAddressList);
    function  MakeNotifyMessage: TIdMessage;
  public
    procedure Exec(Y, M: Word);
	end;

const
  MISC_ITEM = '(其他)';

var
  dmPhoneAnalysis: TdmPhoneAnalysis;

implementation

uses
  TcrmConstants, JcSysUtils, JcDateTimeUtils, JcNumUtils, JcDevExpressUtils, JcDataSetUtils, ReportData, Main;

{$R *.dfm}

{ TdmPhoneAnalysis }

procedure TdmPhoneAnalysis.InitExecute;
begin

end;

procedure TdmPhoneAnalysis.BeginExecute;
begin
  //nothing to do now
end;

procedure TdmPhoneAnalysis.EndExecute;
begin
  //nothing to do now
end;

procedure TdmPhoneAnalysis.Log(AMsg: string);
begin
	fmMain.Log(AMsg);
end;

procedure TdmPhoneAnalysis.LogLine;
begin
	fmMain.LogLine;
end;

function TdmPhoneAnalysis.CopyReportFromTemplate(ADate: TDateTime): string;
const
  XLS_FILE_TITLE = '電話分析月報表';
var
  aPath, aSrcFile, aDstFile: string;
begin
  Result := '';
  aPath := IncludeTrailingPathDelimiter(dmReport.GetTemplatePath);
  aSrcFile := Format('%s%s.xlsx', [aPath, XLS_FILE_TITLE]);
  aPath := IncludeTrailingPathDelimiter(dmReport.GetReportPath);
  aDstFile := Format('%s%s_%s.xlsx', [aPath, XLS_FILE_TITLE, FormatDateTime('yyyymm', ADate)]);
  FXlsFileName := aDstFile;

  if FileExists(aSrcFile) then
  begin
    Log(Format('複製XLS報表範本 %s -> %s', [aSrcFile, aDstFile]));

  	if CopyFile(PChar(aSrcFile), PChar(aDstFile), False) then
   		Result := aDstFile
    else
			Log(Format('!! 無法複製XLS報表範本 %s', [aDstFile]));
  end
  else
    Log(Format('!! 找不到XLS報表範本 %s', [aSrcFile]));
end;

procedure TdmPhoneAnalysis.WriteDataToXls_PhoneCount(ASheet: TXLSWorksheet; AThisYear, AThisMonth: Boolean);
var
  aSite, aSw: string;
  aCol, aRow: Word;
  aCount : Int64;
begin
  with FWICSIPH5 do
  begin
    First;

    while not Eof do
    begin
      aSite := FieldByName('IPH5001').AsString;
      aSw := FieldByName('IPH5003').AsString;
      aCount := FieldByName('IPH5006_SUM').AsInteger;

      if CalcColRow_Sw(aSite, aSw, AThisYear, AThisMonth, aCol, aRow) then
        ASheet.AsInteger[aCol, aRow] := aCount;

      Next;
    end;
  end;
end;

procedure TdmPhoneAnalysis.Exec(Y, M: Word);
var
  aSheet: TXLSWorksheet;
begin
	with dmPhoneAnalysis do
  begin
    dmReport.Calc_YM(Y, M);
    FYear := Y;
    FMonth:= M; 
    Init_XLS_Report;
    aSheet := FindSheet('產品來電');
    //寫入標題
    WriteDataToXls_Titles(aSheet, Y, M);
    //-------------------------------------------------------------------------
    //統計今年 1~M 月的ACD來電總數
    PrepareData_IPH5_Y_1_M(Y, M);
    WriteDataToXls_PhoneCount(aSheet, True, False);
    //統計去年 1~M 月的ACD來電總數
    PrepareData_IPH5_Y_1_M(Y-1, M);
    WriteDataToXls_PhoneCount(aSheet, False, False);
    //統計今年 M 月的ACD來電總數
    PrepareData_IPH5_Y_M(Y, M);
    WriteDataToXls_PhoneCount(aSheet, True, True);
    //統計去年 M 月的ACD來電總數
    PrepareData_IPH5_Y_M(Y-1, M);
    WriteDataToXls_PhoneCount(aSheet, False, True);
    //-------------------------------------------------------------------------
    //統計今年 1~M 月的逾時數
    PrepareData_IPH4_Y_1_M(Y, M);
    WriteDataToXls_Rate(aSheet, True, False);
    WriteDataToXls_RateSummary(aSheet, True, False);
    //統計去年 1~M 月的逾時數
    PrepareData_IPH4_Y_1_M(Y-1, M);
    WriteDataToXls_Rate(aSheet, False, False);
    WriteDataToXls_RateSummary(aSheet, False, False);
    //統計今年 M 月的逾時數
    PrepareData_IPH4_Y_M(Y, M);
    WriteDataToXls_Rate(aSheet, True, True);
    WriteDataToXls_RateSummary(aSheet, True, True);
    //統計去年 M 月的逾時數
    PrepareData_IPH4_Y_M(Y-1, M);
    WriteDataToXls_Rate(aSheet, False, True);
    WriteDataToXls_RateSummary(aSheet, False, True);
    //-------------------------------------------------------------------------
    XLSRW.Write;
    fmMain.CallExcelToSaveAs(FXlsFileName);

    FWICSIPH4.Close;
    FWICSIPH5.Close;
    
    if fmMain.NoMail then
      ShellExecute(0, 'open', PChar(FXlsFileName), nil, nil, SW_SHOWMAXIMIZED)
    else
      SendMail;
  end;
end;

procedure TdmPhoneAnalysis.MakeCCList(AEmailAddrList: TIdEmailAddressList);
var
  aList: TStringList;
  i: Integer;
begin
  with AEmailAddrList do
  begin
    if fmMain.DebugMode then
    begin
      Add.Address := 'f07@winton.com.tw';
      Add.Address := 'joe0107@gmail.com';
      Add.Address := 'wintonjoelee@gmail.com';
    end
    else
    begin
      aList := dmReport.GetAllEmail_Site_Admin;

      for i := 0 to aList.Count-1 do
        Add.Address := aList[i];

      Add.Address := 'Tony@winton.com.tw';
      Add.Address := 'trista62@winton.com.tw';
      Add.Address := 'sky@winton.com.tw';
      Add.Address := 'joelee@winton.com.tw';
      aList.Free;
    end;
  end;
end;

function TdmPhoneAnalysis.MakeNotifyMessage: TIdMessage;
var
  aDayOfWeek: string;
begin
  Result := TIdMessage.Create(Self);
  aDayOfWeek := GetChineseNumStr(DayOfWeek(Now) - 1);
  if (aDayOfWeek = '零') then aDayOfWeek := '日'; 

  with Result do
  begin
    //填入收件者
    MakeRecipients(Recipients);
    //填入副本
    MakeCCList(CCList);
    //填入郵件表頭資訊
    Subject := Format('電話分析月報表_%s(%s)', [FormatDateTime('yyyymmdd', Now), aDayOfWeek]);
    //寄件人地址
    From.Address := 'rdrepl@winton.com.tw';
    ContentType := 'multipart/mixed';
    //填入郵件內容
    Body.Text := '';
  end;

  if FileExists(FXlsFileName) then
  	TIdAttachmentFile.Create(Result.MessageParts, FXlsFileName);
end;

procedure TdmPhoneAnalysis.SendMail;
var
  aMsg: TIdMessage;
begin
  if not fmMain.NoMail then
  begin
    aMsg := MakeNotifyMessage;
    dmReport.SendNotofyMail_SSL(aMsg);
    Log('已透過郵件傳送報表');
  end;
end;

procedure TdmPhoneAnalysis.DataModuleCreate(Sender: TObject);
begin
  FWICSIPH4 := dmReport.GetQuery_WintonTcrm;
  FWICSIPH5 := dmReport.GetQuery_WintonTcrm;
end;

procedure TdmPhoneAnalysis.GetData_IPH5(ABegTime, AEndTime: TDateTime);
begin
  Log('讀取軟體產品的ACD接聽匯總資料(WICSIPH5)');

  with FWICSIPH5 do
  begin
    if Active then Close;
    LocalUpdate := True;
    SQL.Clear;
		SQL.Add('SELECT IPH5001, IPH5003, SUM(IPH5006) AS IPH5006_SUM');
    SQL.Add('FROM WICSIPH5 WITH (NOLOCK)');
    SQL.Add('WHERE IPH5002 BETWEEN :IPH5002B AND :IPH5002E');
    SQL.Add('GROUP BY IPH5001, IPH5003');
    SQL.Add('ORDER BY IPH5001, IPH5003');

    ParamByName('IPH5002B').AsDateTime := ABegTime;
    ParamByName('IPH5002E').AsDateTime := AEndTime;

    try
      Open;
      Log(Format('取得軟體產品的ACD接聽匯總資料(WICSIPH5)，記錄數 = %d', [RecordCount]));
    except
      on E: Exception do
        Log(Format('GetData_IPH5() failed, error = %s', [E.Message]));
    end;
  end;
end;

procedure TdmPhoneAnalysis.Init_XLS_Report;
begin
  CopyReportFromTemplate(EncodeDate(FYear, FMonth, 1));
  XLSRW.Filename := FXlsFileName;
  XLSRW.Read;
end;

procedure TdmPhoneAnalysis.PrepareData_IPH5_Y_1_M(Y, M: Word);
var
  aCalcBegTime, aCalcEndTime: TDateTime;
begin
  //統計 Y 年 1~M 月的ACD來電總數
  aCalcBegTime := EncodeDate(Y, 1, 1);
  aCalcEndTime := EndOfAMonth(Y, M);
  GetData_IPH5(aCalcBegTime, aCalcEndTime);
end;

procedure TdmPhoneAnalysis.PrepareData_IPH5_Y_M(Y, M: Word);
var
  aCalcBegTime, aCalcEndTime: TDateTime;
begin
  //統計 Y 年 M 月的ACD來電總數
  aCalcBegTime := EncodeDate(Y, M, 1);
  aCalcEndTime := EndOfAMonth(Y, M);
  GetData_IPH5(aCalcBegTime, aCalcEndTime);
end;

function TdmPhoneAnalysis.FindSheet(ASheetName: string): TXLSWorksheet;
begin
  with XLSRW do
  begin
    Result := SheetByName(ASheetName);

    if (Result = nil) then
    begin
      Log(Format('!! 找不到工作表[%s], 無法寫入工作表資料', [ASheetName]));
      Abort;
    end;
  end;
end;

function TdmPhoneAnalysis.CalcColRow_Sw(ASite, ASw: string; AThisYear, AThisMonth: Boolean; var ACol, ARow: Word): Boolean;
const
  SW_START_ROW = 3;
  SW_START_COL = 3;
var
  aColOffset, aRowOffset: Integer;
begin
  aColOffset := GetColOffset_Sw(AThisYear, AThisMonth);
  aRowOffset := GetRowOffset_Site(ASite) + GetRowOffset_Sw(ASw);

  if (aRowOffset >= 0) then
  begin
    ACol := SW_START_COL + aColOffset;
    ARow := SW_START_ROW + aRowOffset;
    Result := True;
  end
  else
    Result := False;
end;

procedure TdmPhoneAnalysis.WriteDataToXls_Titles(ASheet: TXLSWorksheet; Y, M: Word);
begin
  ASheet.AsString[3, 1] := Format('累計1-%d月份產品來電量', [M]);
  ASheet.AsString[7, 1] := Format('%d月份產品來電量', [M]);
  //---------------------------------------------------------------------------
  ASheet.AsInteger[3, 2] := Y - 1;
  ASheet.AsInteger[7, 2] := Y - 1;
  ASheet.AsInteger[4, 2] := Y;
  ASheet.AsInteger[8, 2] := Y;
  ASheet.AsString[5, 2] := Format('%d增加', [Y]);
  ASheet.AsString[9, 2] := ASheet.AsString[5, 2];
  //---------------------------------------------------------------------------
  ASheet.AsString[12, 0] := Format('累計1-%d月份單位逾時率', [M]);
  ASheet.AsString[16, 0] := Format('%d月份單位逾時率', [M]);
  //---------------------------------------------------------------------------
  ASheet.AsInteger[12, 1] := Y - 1;
  ASheet.AsInteger[16, 1] := Y - 1;
  ASheet.AsInteger[14, 1] := Y;
  ASheet.AsInteger[18, 1] := Y;
end;

procedure TdmPhoneAnalysis.GetData_IPH4(ABegTime, AEndTime: TDateTime);
begin
  Log('讀取營業處的ACD資料(WICSIPH4)');

  with FWICSIPH4 do
  begin
    if Active then Close;
    LocalUpdate := True;
    SQL.Clear;
		SQL.Add('SELECT IPH4001, SUM(IPH4007) AS IPH4007_SUM, SUM(IPH4008) AS IPH4008_SUM,');
    SQL.Add('SUM(IPH4010) AS IPH4010_SUM, SUM(IPH4011) AS IPH4011_SUM');
    SQL.Add('FROM WICSIPH4 WITH(NOLOCK)');
    SQL.Add('WHERE IPH4002 BETWEEN :IPH4002B AND :IPH4002E');
    SQL.Add('GROUP BY IPH4001');
    SQL.Add('ORDER BY IPH4001');

    ParamByName('IPH4002B').AsDateTime := ABegTime;
    ParamByName('IPH4002E').AsDateTime := AEndTime;

    try
      Open;
      Log(Format('取得營業處的ACD資料(WICSIPH4)，記錄數 = %d', [RecordCount]));
    except
      on E: Exception do
        Log(Format('GetData_IPH4() failed, error = %s', [E.Message]));
    end;
  end;
end;

procedure TdmPhoneAnalysis.PrepareData_IPH4_Y_1_M(Y, M: Word);
var
  aCalcBegTime, aCalcEndTime: TDateTime;
begin
  //統計 Y 年 1~M 月的逾時數
  aCalcBegTime := EncodeDate(Y, 1, 1);
  aCalcEndTime := EndOfAMonth(Y, M);
  GetData_IPH4(aCalcBegTime, aCalcEndTime);
end;

procedure TdmPhoneAnalysis.PrepareData_IPH4_Y_M(Y, M: Word);
var
  aCalcBegTime, aCalcEndTime: TDateTime;
begin
  //統計 Y 年 M 月的逾時數
  aCalcBegTime := EncodeDate(Y, M, 1);
  aCalcEndTime := EndOfAMonth(Y, M);
  GetData_IPH4(aCalcBegTime, aCalcEndTime);
end;

procedure TdmPhoneAnalysis.WriteDataToXls_Rate(ASheet: TXLSWorksheet; AThisYear, AThisMonth: Boolean);
var
  aSite: string;
  aRate: Extended;
  aCol, aRow: Word;
begin
  with FWICSIPH4 do
  begin
    First;

    while not Eof do
    begin
      aSite := FieldByName('IPH4001').AsString;
      //-----------------------------------------------------------------------
      aRate := FieldByName('IPH4008_SUM').AsInteger / FieldByName('IPH4007_SUM').AsInteger;

      if CalcColRow_Rate(aSite, AThisYear, AThisMonth, 0, aCol, aRow) then
        ASheet.AsFloat[aCol, aRow] := aRate;
      //-----------------------------------------------------------------------
      aRate := FieldByName('IPH4010_SUM').AsInteger / FieldByName('IPH4011_SUM').AsInteger;

      if CalcColRow_Rate(aSite, AThisYear, AThisMonth, 1, aCol, aRow) then
        ASheet.AsFloat[aCol, aRow] := aRate;
      //-----------------------------------------------------------------------
      Next;
    end;
  end;
end;

function TdmPhoneAnalysis.GetColOffset_Sw(AThisYear, AThisMonth: Boolean): Integer;
begin
  if (not AThisYear) and (not AThisMonth) then
    Result := 0
  else if AThisYear and (not AThisMonth) then
    Result := 1
  else if (not AThisYear) and AThisMonth then
    Result := 4
  else if AThisYear and AThisMonth then
    Result := 5
  else
    Result := 0;
end;

function TdmPhoneAnalysis.GetRowOffset_Site(ASite: string): Integer;
begin
  if Pos('台北', ASite) > 0 then
    Result := 0
  else if Pos('北區', ASite) > 0 then
    Result := 7
  else if Pos('中區', ASite) > 0 then
    Result := 14
  else if Pos('南區', ASite) > 0 then
    Result := 21
  else
  begin
    Log(Format('!! GetRowOffset_Site() error, Site = %s', [ASite]));
    Result := -99999;
  end;
end;

function TdmPhoneAnalysis.GetRowOffset_Sw(ASw: string): Integer;
begin
  if Pos('WSTP', ASw) > 0 then
    Result := 0
  else if Pos('WBEC', ASw) > 0 then
    Result := 1
  else if Pos('TF', ASw) > 0 then
    Result := 2
  else if Pos('CPA', ASw) > 0 then
    Result := 3
  else if Pos('MERP', ASw) > 0 then
    Result := 4
  else if Pos('HR', ASw) > 0 then
    Result := 5
  else
  begin
    Log(Format('!! GetRowOffset_Sw() error, Sw = %s', [ASw]));
    Result := -99999;
  end;
end;

function TdmPhoneAnalysis.CalcColRow_Rate(ASite: string; AThisYear, AThisMonth: Boolean; AKind: Integer; var ACol, ARow: Word): Boolean;
const
  SW_START_ROW = 3;
  SW_START_COL = 12;
var
  aColOffset, aRowOffset: Integer;
begin
  aColOffset := GetColOffset_Rate(AThisYear, AThisMonth, AKind);
  aRowOffset := GetRowOffset_Site(ASite);

  if (aRowOffset >= 0) then
  begin
    ACol := SW_START_COL + aColOffset;
    ARow := SW_START_ROW + aRowOffset;
    Result := True;
  end
  else
    Result := False;
end;

function TdmPhoneAnalysis.GetColOffset_Rate(AThisYear, AThisMonth: Boolean; AKind: Integer): Integer;
begin
  if (not AThisYear) and (not AThisMonth) then
    Result := 0
  else if AThisYear and (not AThisMonth) then
    Result := 2
  else if (not AThisYear) and AThisMonth then
    Result := 4
  else if AThisYear and AThisMonth then
    Result := 6
  else
    Result := 0;

  Result := Result + AKind; 
end;

procedure TdmPhoneAnalysis.WriteDataToXls_RateSummary(ASheet: TXLSWorksheet; AThisYear, AThisMonth: Boolean);
const
  SW_START_COL  = 12;
  SUMMARY_ROW   = 32;
var
  aRate: Extended;
  aCol: Word;
  aIPH4007, aIPH4008, aIPH4010, aIPH4011: Integer;
begin
  aIPH4007 := 0;
  aIPH4008 := 0;
  aIPH4010 := 0;
  aIPH4011 := 0;

  with FWICSIPH4 do
  begin
    First;

    while not Eof do
    begin
      aIPH4007 := aIPH4007 + FieldByName('IPH4007_SUM').AsInteger;
      aIPH4008 := aIPH4008 + FieldByName('IPH4008_SUM').AsInteger;
      aIPH4010 := aIPH4010 + FieldByName('IPH4010_SUM').AsInteger;
      aIPH4011 := aIPH4011 + FieldByName('IPH4011_SUM').AsInteger;
      Next;
    end;
    //-------------------------------------------------------------------------
    aCol := GetColOffset_Rate(AThisYear, AThisMonth, 0);
    aRate := aIPH4008 /aIPH4007;
    ASheet.AsFloat[SW_START_COL+aCol, SUMMARY_ROW] := aRate;
    //-------------------------------------------------------------------------
    aCol := GetColOffset_Rate(AThisYear, AThisMonth, 1);
    aRate := aIPH4010 /aIPH4011;
    ASheet.AsFloat[SW_START_COL+aCol, SUMMARY_ROW] := aRate;
  end;
end;

function TdmPhoneAnalysis.MakeRecipients(AEmailAddrList: TIdEmailAddressList): Integer;
var
  aList1, aList2: TStringList;
  i: Integer;
  aText: string;
begin
  with AEmailAddrList do
  begin
    if not fmMain.DebugMode then
    begin
      aList1 := dmReport.GetAllEmail_TE_Admin;
      aList2 := dmReport.GetAllEmail_TE_Leader;
      Result := aList1.Count + aList2.Count;
      aText := '';

      for i := 0 to aList1.Count-1 do
        aText := aText + aList1[i] + ',';

      for i := 0 to aList2.Count-1 do
        aText := aText + aList2[i] + ',';

      if Length(aText) > 0 then
        System.Delete(aText, Length(aText), 1);

      EMailAddresses := aText;
      aList1.Free;
      aList2.Free;
    end
    else
    begin
      Result := 2;
      Add.Address := 'joelee@winton.com.tw';
      Add.Address := 'joe0107@gmail.com';
    end;
  end;
end;

end.
