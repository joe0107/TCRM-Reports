unit SitePhoneSummary;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs, DB, ADODB, StdCtrls, ShellAPI,
  cxData, cxClasses, cxCustomData, cxDataStorage, cxDBData, JclStrings, MemDS, DBAccess, Uni, DateUtils, XLSSheetData5,
  dxmdaset, XLSReadWriteII5, XLSDbRead5, XLSNames5, IdEMailAddress, IdMessage, IdAttachmentFile, CodeSiteLogging,
  XLSCmdFormat5, Xc12DataStyleSheet5, Xc12Utils5, XLSFormattedObj5;

type
  TdmSitePhoneSummary = class(TDataModule)
    qrGetData: TUniQuery;
    connReport: TUniConnection;
    XLSRW: TXLSReadWriteII5;
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  protected
    FInCalcData: Boolean;
    FCalcBegTime, FCalcEndTime: TDateTime;

    procedure InitExecute;
    procedure BeginExecute;
    procedure EndExecute;
    //切換取得資料來源的公用連線
    procedure InitReportConn;
    procedure InitData;
    procedure Log(AMsg: string);
    procedure LogLine;
  private
    procedure GetData(ASite: string; ABegTime, AEndTime: TDateTime);
    procedure GetData_Summary(ASite: string; ABegTime, AEndTime: TDateTime);
    //取得指定日期的值機總人天
    function  GetOnDutyTotal(ADate: TDateTime): Extended;
  private
    //FNewRptCount: Integer;
    FXlsFileName: string;
    function  GetReportFileName(ADate: TDateTime): string;
    procedure XLS_Init(AFileName: string);
    function  XLS_WriteReport(ASite: string; ADataSet: TDataSet): Boolean;
    procedure XLS_WriteTitle(AXLS: TXLSReadWriteII5; ASheet: TXLSWorksheet);
    procedure XLS_WriteData(AXLS: TXLSReadWriteII5; ASheet: TXLSWorksheet; ADataSet: TDataSet);
    procedure UpdateWorkSheet_Summary(AYear: Integer);
	protected
    procedure SendMail;
    function  MakeRecipients(AEmailAddrList: TIdEmailAddressList): Integer;
    procedure MakeCCList(AEmailAddrList: TIdEmailAddressList);
    function  MakeNotifyMessage: TIdMessage;
    function  MakeAdminNotifyMessage: TIdMessage;
  public
    function  PrintReport(ADate: TDateTime): Boolean;
    // 產生指定日期的 XLS 報表
    class procedure Exec(ADate: TDateTime);
	end;

var
  dmSitePhoneSummary: TdmSitePhoneSummary;

implementation

uses
  TcrmConstants, JcDateTimeUtils, JcNumUtils, JcDevExpressUtils, JcDataSetUtils, ReportData, Main;

const
  RPT_FILE_TITLE = '營業處回電效率統計表';
  RPT_START_ROW = 0;
  RPT_START_COL = 0;

{$R *.dfm}

{ TdmSitePhoneSummary }

procedure TdmSitePhoneSummary.FormDestroy(Sender: TObject);
begin
  inherited;
  dmSitePhoneSummary := nil;
end;

procedure TdmSitePhoneSummary.FormCreate(Sender: TObject);
begin
  inherited;
  Log(Format('開始執行[%s]', [RPT_FILE_TITLE]));
  FInCalcData := False;
  InitReportConn;
end;

procedure TdmSitePhoneSummary.InitExecute;
begin
  InitData;
end;

procedure TdmSitePhoneSummary.BeginExecute;
begin
  //nothing to do now
end;

procedure TdmSitePhoneSummary.EndExecute;
begin
  //nothing to do now
end;

procedure TdmSitePhoneSummary.InitReportConn;
begin
  with dmReport do
  begin
    // 從文中資料庫取得統計數據
  	SetUniConn_TCRM(connReport, GetSiteIp(SITE_NAME_Winton_TC));
  end;
end;

procedure TdmSitePhoneSummary.InitData;
begin
  // nothing to do now
end;

procedure TdmSitePhoneSummary.Log(AMsg: string);
begin
	fmMain.Log(AMsg);
end;

procedure TdmSitePhoneSummary.LogLine;
begin
	fmMain.LogLine;
end;

function TdmSitePhoneSummary.XLS_WriteReport(ASite: string; ADataSet: TDataSet): Boolean;
var
  aRC: Integer;
  aSheet: TXLSWorksheet;
begin
  // Added by Joe 2019/01/02 14:42:21
  aRC := 0;

  with ADataSet do
  begin
    // 檢查有無[ACD處理數]
    First;
    while not Eof do
    begin
      if (ADataSet.FieldByName('IPH4008').AsInteger > 0) then
        Inc(aRC);
      Next;
    end;
  end;

  if (aRC = 0) then
  begin
    Result := False;
    Log(Format('無ACD處理數資料，不產生報表[%s]', [ASite]));
    Exit;
  end;
  //---------------------------------------------------------------------------
  Log(Format('產生XLS報表[%s]', [ASite]));

  with XLSRW do
  begin
    // 加入指定營業處的工作表
    if (ASite <> SITE_NAME_Winton_TC) then
      aSheet := Add
    else
      aSheet := Sheets[0];
      
    aSheet.Name := ASite;
    XLS_WriteTitle(XLSRW, aSheet);
    XLS_WriteData(XLSRW, aSheet, ADataSet);
  end;

  Result := True;
end;

procedure TdmSitePhoneSummary.UpdateWorkSheet_Summary(AYear: Integer);
const
  WORK_SHEET_NAME = '統計表';
var
  aSheet: TXLSWorksheet;
  aText: string;
begin
  with XLSRW do
  begin
    aSheet := SheetByName(WORK_SHEET_NAME);

    if (aSheet = nil) then
    begin
      Log(Format('!! 找不到工作表[%s], 無法更新統計表資料', [WORK_SHEET_NAME]));
      Exit;
    end;
    aText := Format('%d', [AYear]);
    aSheet.AsString[0, 0] := StringReplace(aSheet.AsString[0, 0], '[yyyy]', aText, [rfReplaceAll, rfIgnoreCase]);
    aSheet.AsString[1, 1] := StringReplace(aSheet.AsString[1, 1], '[yyyy]', aText, [rfReplaceAll, rfIgnoreCase]);
    aText := Format('%d', [AYear-1]);
    aSheet.AsString[1, 1] := StringReplace(aSheet.AsString[1, 1], '[yyyy-prev]', aText, [rfReplaceAll, rfIgnoreCase]);
  end;
end;

function TdmSitePhoneSummary.PrintReport(ADate: TDateTime): Boolean;
var
  aRptCount: Integer;
begin
  Result := False;
  FCalcBegTime := StartOfTheMonth(ADate);
  FCalcEndTime := EndOfTheDay(ADate);
  FXlsFileName := GetReportFileName(ADate);
  aRptCount := 0; // Added by Joe 2019/01/02 14:46:16

  if GetOnDutyTotal(ADate) = 0 then
  begin
    Log('本日無人值機，不產生報表');
    Exit;
  end;

  with XLSRW do
  begin
    XLS_Init(FXlsFileName);
    //--------------------------------------------------------------------------
    GetData_Summary(SITE_NAME_Winton_TC, FCalcBegTime, FCalcEndTime);

    if XLS_WriteReport(SITE_NAME_Winton_TC, qrGetData) then
      Inc(aRptCount);
    //--------------------------------------------------------------------------
    GetData(SITE_DESC_Taipei_TC, FCalcBegTime, FCalcEndTime);

    if XLS_WriteReport(SITE_DESC_Taipei_TC, qrGetData) then
      Inc(aRptCount);
    //--------------------------------------------------------------------------
    GetData(SITE_DESC_Taoyuan_TC, FCalcBegTime, FCalcEndTime);

    if XLS_WriteReport(SITE_DESC_Taoyuan_TC, qrGetData) then
      Inc(aRptCount);
    //--------------------------------------------------------------------------
    GetData(SITE_DESC_Taichung_TC, FCalcBegTime, FCalcEndTime);

    if XLS_WriteReport(SITE_DESC_Taichung_TC, qrGetData) then
      Inc(aRptCount);
    //--------------------------------------------------------------------------
    GetData(SITE_DESC_Tainan_TC, FCalcBegTime, FCalcEndTime);

    if XLS_WriteReport(SITE_DESC_Tainan_TC, qrGetData) then
      Inc(aRptCount);
    //--------------------------------------------------------------------------
    if (aRptCount > 0) then
    begin
      Write;
      Result := True;
    end;
  end;
end;

class procedure TdmSitePhoneSummary.Exec(ADate: TDateTime);
begin
	if not Assigned(dmSitePhoneSummary) then
  	Application.CreateForm(TdmSitePhoneSummary, dmSitePhoneSummary);

	with dmSitePhoneSummary do
  begin
    if PrintReport(ADate) then
    begin
      fmMain.CallExcelToSaveAs(FXlsFileName);

      if not fmMain.MailMode then
        ShellExecute(0, 'open', PChar(FXlsFileName), nil, nil, SW_SHOWMAXIMIZED)
      else
        SendMail;
    end;
  	Free;
  end;
end;

function TdmSitePhoneSummary.MakeAdminNotifyMessage: TIdMessage;
begin
  Result := nil;
end;

procedure TdmSitePhoneSummary.MakeCCList(AEmailAddrList: TIdEmailAddressList);
var
  aList: TStringList;
  i: Integer;
begin
  with AEmailAddrList do
  begin
    if fmMain.DebugMode then
    begin
      //Add.Address := 'f07@winton.com.tw';
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

function TdmSitePhoneSummary.MakeNotifyMessage: TIdMessage;
var
  aDayOfWeek: string;
begin
  Result := TIdMessage.Create(Self);
  aDayOfWeek := GetChineseNumStr(DayOfWeek(FCalcEndTime) - 1);
  if (aDayOfWeek = '零') then aDayOfWeek := '日'; 

  with Result do
  begin
    //填入收件者
    MakeRecipients(Recipients);
    //填入副本
    MakeCCList(CCList);
    //填入郵件表頭資訊
    Subject := Format('%s_%s(%s)', [RPT_FILE_TITLE, FormatDateTime('yyyymmdd', FCalcEndTime), aDayOfWeek]);
    //寄件人地址
    From.Address := 'rdrepl@winton.com.tw';
    ContentType := 'multipart/mixed';
    //填入郵件內容
    Body.Text := '';
  end;

  if FileExists(FXlsFileName) then
  	TIdAttachmentFile.Create(Result.MessageParts, FXlsFileName);
end;

procedure TdmSitePhoneSummary.SendMail;
var
  aMsg: TIdMessage;
begin
  if fmMain.MailMode then
  begin
    aMsg := MakeNotifyMessage;
    dmReport.SendNotofyMail_SSL(aMsg);
    Log('已透過郵件傳送報表');
  end;
end;

function TdmSitePhoneSummary.GetReportFileName(ADate: TDateTime): string;
var
  aPath: string;
begin
  aPath := IncludeTrailingPathDelimiter(dmReport.GetReportPath);
  Result := Format('%s%s_%s.xls', [aPath, RPT_FILE_TITLE, FormatDateTime('yyyymmdd', ADate)]);
end;

procedure TdmSitePhoneSummary.XLS_WriteTitle(AXLS: TXLSReadWriteII5; ASheet: TXLSWorksheet);
var
  aCmdFormat: TXLSCmdFormat;
  i: Integer;

  procedure SetFormat(ACol, ARow: Integer);
  var
    aCell: TXLSCell;
  begin
    aCell := ASheet.Cell[ACol, ARow];

    with aCell do
    begin
      if (ACol >= 11) and (ACol <= 15) then
        FillPatternForeColor := TXc12IndexColor(25)
      else
        FillPatternForeColor := TXc12IndexColor(30);

      FontColor := clWhite;
    end;
  end;
begin
  with ASheet do
  begin
    MergeCells(RPT_START_COL, RPT_START_ROW, RPT_START_COL, RPT_START_ROW+1);
    AsString[RPT_START_COL, RPT_START_ROW]      := '日期';
    MergeCells(RPT_START_COL+1, RPT_START_ROW, RPT_START_COL+1, RPT_START_ROW+1);
    AsString[RPT_START_COL+1, RPT_START_ROW]    := '值機'+#10+'人天';
    //--------------------------------------------------------------------------
    MergeCells(RPT_START_COL+2, RPT_START_ROW, RPT_START_COL+10, RPT_START_ROW);
    AsString[RPT_START_COL+2, RPT_START_ROW]    := '合約';
    AsString[RPT_START_COL+2, RPT_START_ROW+1]  := '來電數';
    AsString[RPT_START_COL+3, RPT_START_ROW+1]  := '回電數';
    AsString[RPT_START_COL+4, RPT_START_ROW+1]  := 'ACD'+#10+'有效派送';
    AsString[RPT_START_COL+5, RPT_START_ROW+1]  := 'ACD'+#10+'直接數';
    AsString[RPT_START_COL+6, RPT_START_ROW+1]  := 'ACD'+#10+'處理數';
    AsString[RPT_START_COL+7, RPT_START_ROW+1]  := 'ACD'+#10+'接聽率';
    AsString[RPT_START_COL+8, RPT_START_ROW+1]  := '逾時數';
    AsString[RPT_START_COL+9, RPT_START_ROW+1]  := '逾時率';
    AsString[RPT_START_COL+10, RPT_START_ROW+1] := '未回數';
    //--------------------------------------------------------------------------
    MergeCells(RPT_START_COL+11, RPT_START_ROW, RPT_START_COL+15, RPT_START_ROW);
    AsString[RPT_START_COL+11, RPT_START_ROW]   := '非合約';
    AsString[RPT_START_COL+11, RPT_START_ROW+1] := '來電數';
    AsString[RPT_START_COL+12, RPT_START_ROW+1] := '回電數';
    AsString[RPT_START_COL+13, RPT_START_ROW+1] := '逾時數';
    AsString[RPT_START_COL+14, RPT_START_ROW+1] := '逾時率';
    AsString[RPT_START_COL+15, RPT_START_ROW+1] := '未回數';
    //--------------------------------------------------------------------------
    MergeCells(RPT_START_COL+16, RPT_START_ROW, RPT_START_COL+18, RPT_START_ROW);
    AsString[RPT_START_COL+16, RPT_START_ROW]   := '平均';
    AsString[RPT_START_COL+16, RPT_START_ROW+1] := AsString[RPT_START_COL+5, RPT_START_ROW+1];
    AsString[RPT_START_COL+17, RPT_START_ROW+1] := AsString[RPT_START_COL+6, RPT_START_ROW+1];
    AsString[RPT_START_COL+18, RPT_START_ROW+1] := '回電數';
  end;
  // 設定欄寬
  with ASheet do
  begin
    // 日期
    Columns[RPT_START_COL].CharWidth := 10;
    // 值機人天
    Columns[RPT_START_COL+1].CharWidth := 8;
  end;
  //============================================================================
  aCmdFormat := AXLS.CmdFormat;
  ASheet.CalcDimensions;

  with aCmdFormat do
  begin
    //畫出外框，標題置中，標題顏色
    BeginEdit(ASheet);
    Alignment.Horizontal := chaCenter;
    Alignment.Vertical := cvaCenter;
    Border.Style := cbsThin;
    Border.Preset(cbspOutline);

    for i := 0 to ASheet.LastCol do
    begin
      Apply(i, RPT_START_ROW, i, RPT_START_ROW);
      Apply(i, RPT_START_ROW+1, i, RPT_START_ROW+1);
      SetFormat(i, RPT_START_ROW);
      SetFormat(i, RPT_START_ROW+1);
    end;
  end;
end;

procedure TdmSitePhoneSummary.GetData(ASite: string; ABegTime, AEndTime: TDateTime);
begin
  Log(Format('取得營業處來回電資料[%s]', [ASite]));

  with qrGetData do
  begin
    if Active then Close;
    SQL.Clear;
    SQL.Add('SELECT');
    SQL.Add('RID, IPH4002, IPH4001, IPH4003, IPH4004,');
    SQL.Add('IPH4005, IPH4006, IPH4007, IPH4008, IPH4009,');
    SQL.Add('IPH4010, IPH4011, IPH4012, IPH4013, IPH4205,');
    SQL.Add('IPH4209, IPH4210, IPH4211, IPH4213');
    SQL.Add('FROM WICSIPH4 WITH(NOLOCK)');
    SQL.Add('ORDER BY IPH4002');
    AddWhere('(IPH4002 BETWEEN :IPH4002B AND :IPH4002E)');
    AddWhere('(IPH4001 = :IPH4001)');
    ParamByName('IPH4001').AsString := ASite;
    ParamByName('IPH4002B').AsDateTime := ABegTime;
    ParamByName('IPH4002E').AsDateTime := AEndTime;
    Open;
  end;
end;

procedure TdmSitePhoneSummary.XLS_WriteData(AXLS: TXLSReadWriteII5; ASheet: TXLSWorksheet; ADataSet: TDataSet);
const
  INT_FMT = '#,0';
var
  aRow, aWeekDay, i, j, aDataCount: Integer;
  aYear, aMonth, aDay: Word;
  aText, aCellRef1, aCellRef2, aCellRef3: string;
  aCmdFormat: TXLSCmdFormat;
begin
  if not JcDataSetIsValid(ADataSet) then Exit;
  //============================================================================
  with ADataSet do
  begin
    First;
    aRow := RPT_START_ROW+2;
    aDataCount := 0;  // Added by Joe 2019/01/02 13:57:11

    while not Eof do
    begin
      DecodeDate(FieldByName('IPH4002').AsDateTime, aYear, aMonth, aDay);
      aWeekDay := DayOfWeek(FieldByName('IPH4002').AsDateTime)-1;
      // 略過週末(日)的資料
      // 略過當日接聽數為0的資料(當天非上班日)
      //if ((aWeekDay = 0) or (aWeekDay = 6)) or (FieldByName('IPH4008').AsInteger = 0) then
      if (FieldByName('IPH4008').AsInteger = 0) then
      begin
        Next;
        Continue;
      end;
      //------------------------------------------------------------------------
      aText := GetChineseDayNumStr(aWeekDay);
      aText := Format('%.2d/%.2d(%s)', [aMonth, aDay, aText]);

      with ASheet do
      begin
        // 日期
        AsString[RPT_START_COL, aRow] := aText;
        // 值機人天
        AsFloat[RPT_START_COL+1, aRow] := FieldByName('IPH4003').AsFloat;
        Cell[RPT_START_COL+1, aRow].NumberFormat := '###0.0';
        //--------------------------------------------------------------------------
        // 來電通數
        AsInteger[RPT_START_COL+2, aRow] := FieldByName('IPH4011').AsInteger;
        Cell[RPT_START_COL+2, aRow].NumberFormat := INT_FMT;
        // 回電通數
        AsInteger[RPT_START_COL+3, aRow] := FieldByName('IPH4009').AsInteger;
        Cell[RPT_START_COL+3, aRow].NumberFormat := INT_FMT;
        // ACD有效派送
        AsInteger[RPT_START_COL+4, aRow] := FieldByName('IPH4007').AsInteger;
        Cell[RPT_START_COL+4, aRow].NumberFormat := INT_FMT;
        // ACD直接接聽通數
        AsInteger[RPT_START_COL+5, aRow] := FieldByName('IPH4012').AsInteger;
        Cell[RPT_START_COL+5, aRow].NumberFormat := INT_FMT;
        // ACD處理通數
        AsInteger[RPT_START_COL+6, aRow] := FieldByName('IPH4008').AsInteger;
        Cell[RPT_START_COL+6, aRow].NumberFormat := INT_FMT;
        // ACD接聽率
        //AsFloat[RPT_START_COL+7, aRow] := FieldByName('IPH4004').AsFloat/100;
        AsFloat[RPT_START_COL+7, aRow] := JcDivide(FieldByName('IPH4008').AsInteger, FieldByName('IPH4007').AsInteger);
        Cell[RPT_START_COL+7, aRow].NumberFormat := '##0.0 %';
        // 逾時通數
        AsInteger[RPT_START_COL+8, aRow] := FieldByName('IPH4010').AsInteger;
        Cell[RPT_START_COL+8, aRow].NumberFormat := INT_FMT;
        // 逾時率
        //AsFloat[RPT_START_COL+9, aRow] := FieldByName('IPH4005').AsFloat/100;
        AsFloat[RPT_START_COL+9, aRow] := JcDivide(FieldByName('IPH4010').AsInteger, FieldByName('IPH4011').AsInteger);
        Cell[RPT_START_COL+9, aRow].NumberFormat := '##0.0 %';
        // 未回通數
        AsInteger[RPT_START_COL+10, aRow] := FieldByName('IPH4013').AsInteger;
        Cell[RPT_START_COL+10, aRow].NumberFormat := INT_FMT;
        //--------------------------------------------------------------------------
        // 來電通數(非合約)
        AsInteger[RPT_START_COL+11, aRow] := FieldByName('IPH4211').AsInteger;
        Cell[RPT_START_COL+11, aRow].NumberFormat := INT_FMT;
        // 回電通數(非合約)
        AsInteger[RPT_START_COL+12, aRow] := FieldByName('IPH4209').AsInteger;
        Cell[RPT_START_COL+12, aRow].NumberFormat := INT_FMT;
        // 逾時通數(非合約)
        AsInteger[RPT_START_COL+13, aRow] := FieldByName('IPH4210').AsInteger;
        Cell[RPT_START_COL+13, aRow].NumberFormat := INT_FMT;
        // 逾時率(非合約)
        //AsFloat[RPT_START_COL+14, aRow] := FieldByName('IPH4205').AsFloat/100;
        AsFloat[RPT_START_COL+14, aRow] := JcDivide(FieldByName('IPH4210').AsInteger, FieldByName('IPH4211').AsInteger);
        Cell[RPT_START_COL+14, aRow].NumberFormat := '##0.0 %';
        // 未回通數(非合約)
        AsInteger[RPT_START_COL+15, aRow] := FieldByName('IPH4213').AsInteger;
        Cell[RPT_START_COL+15, aRow].NumberFormat := INT_FMT;
        //--------------------------------------------------------------------------
        // ACD直接接聽通數(平均)
        aCellRef1 := ColRowToRefStr(RPT_START_COL+5, aRow);
        aCellRef2 := ColRowToRefStr(RPT_START_COL+1, aRow);
        AsFormula[RPT_START_COL+16, aRow] := Format('%s/%s', [aCellRef1, aCellRef2]);
        Cell[RPT_START_COL+16, aRow].NumberFormat := '###.0';
        // ACD處理通數(平均)
        aCellRef1 := ColRowToRefStr(RPT_START_COL+6, aRow);
        aCellRef2 := ColRowToRefStr(RPT_START_COL+1, aRow);
        AsFormula[RPT_START_COL+17, aRow] := Format('%s/%s', [aCellRef1, aCellRef2]);
        Cell[RPT_START_COL+17, aRow].NumberFormat := '###.0';
        // 回電通數(平均)
        aCellRef1 := ColRowToRefStr(RPT_START_COL+3, aRow);
        aCellRef2 := ColRowToRefStr(RPT_START_COL+12, aRow);
        aCellRef3 := ColRowToRefStr(RPT_START_COL+1, aRow);
        AsFormula[RPT_START_COL+18, aRow] := Format('(%s+%s)/%s', [aCellRef1, aCellRef2, aCellRef3]);
        Cell[RPT_START_COL+18, aRow].NumberFormat := '###.0';
      end;
      Inc(aRow);
      Inc(aDataCount);  // Added by Joe 2019/01/02 13:57:30
      Next;
    end;
  end;
  // Added by Joe 2019/01/02 14:19:11
  if (aDataCount = 0) then
    Exit;
  //============================================================================
  // 產生合計列
  //============================================================================
  with ASheet do
  begin
    CalcDimensions;
    AsString[RPT_START_COL, aRow] := '合計';
    Cell[RPT_START_COL, aRow].FillPatternForeColor := TXc12IndexColor(43);
    // 除平均值欄外,先全部產生再個別調整
    for i := RPT_START_COL+1 to LastCol do
    begin
      AsFormula[i, aRow] := Format('SUM(%s)', [AreaToRefStr(i, RPT_START_ROW+2, i, LastRow)]);
      Cell[i, aRow].NumberFormat := '#,0';
      Cell[i, aRow].FillPatternForeColor := TXc12IndexColor(43);
    end;
    // 合計-ACD接聽率
    aCellRef1 := ColRowToRefStr(RPT_START_COL+6, aRow);
    aCellRef2 := ColRowToRefStr(RPT_START_COL+4, aRow);
    AsFormula[RPT_START_COL+7, aRow] := Format('%s/%s', [aCellRef1, aCellRef2]);
    // 合計-逾時率
    aCellRef1 := ColRowToRefStr(RPT_START_COL+8, aRow);
    aCellRef2 := ColRowToRefStr(RPT_START_COL+2, aRow);
    AsFormula[RPT_START_COL+9, aRow] := Format('%s/%s', [aCellRef1, aCellRef2]);
    // 合計-逾時率(非合約)
    aCellRef1 := ColRowToRefStr(RPT_START_COL+13, aRow);
    aCellRef2 := ColRowToRefStr(RPT_START_COL+11, aRow);
    AsFormula[RPT_START_COL+14, aRow] := Format('%s/%s', [aCellRef1, aCellRef2]);
    // ACD直接接聽通數(平均)
    aCellRef1 := ColRowToRefStr(RPT_START_COL+5, aRow);
    aCellRef2 := ColRowToRefStr(RPT_START_COL+1, aRow);
    AsFormula[RPT_START_COL+16, aRow] := Format('%s/%s', [aCellRef1, aCellRef2]);
    Cell[RPT_START_COL+16, aRow].NumberFormat := '###.0';
    // ACD處理通數(平均)
    aCellRef1 := ColRowToRefStr(RPT_START_COL+6, aRow);
    aCellRef2 := ColRowToRefStr(RPT_START_COL+1, aRow);
    AsFormula[RPT_START_COL+17, aRow] := Format('%s/%s', [aCellRef1, aCellRef2]);
    Cell[RPT_START_COL+17, aRow].NumberFormat := '###.0';
    // 回電通數(平均)
    aCellRef1 := ColRowToRefStr(RPT_START_COL+3, aRow);
    aCellRef2 := ColRowToRefStr(RPT_START_COL+12, aRow);
    aCellRef3 := ColRowToRefStr(RPT_START_COL+1, aRow);
    AsFormula[RPT_START_COL+18, aRow] := Format('(%s+%s)/%s', [aCellRef1, aCellRef2, aCellRef3]);
    Cell[RPT_START_COL+18, aRow].NumberFormat := '###.0';
    //--------------------------------------------------------------------------
    // 調整顯示格式-值機人天
    Cell[RPT_START_COL+1, aRow].NumberFormat := '#,#.0';
    // 調整顯示格式-ACD接聽率
    Cell[RPT_START_COL+7, aRow].NumberFormat := '##0.0 %';
    // 調整顯示格式-逾時率
    Cell[RPT_START_COL+9, aRow].NumberFormat := '##0.0 %';
    // 調整顯示格式-逾時率(非合約)
    Cell[RPT_START_COL+14, aRow].NumberFormat := '##0.0 %';
  end;

  with AXLS do
  begin
    CmdFormat.BeginEdit(ASheet);
    CmdFormat.Font.Bold := True;
    CmdFormat.Apply(0, aRow, ASheet.LastCol, aRow);
  end;
  //============================================================================
  // 畫出外框
  ASheet.CalcDimensions;
  aCmdFormat := AXLS.CmdFormat;

  with aCmdFormat do
  begin
    BeginEdit(ASheet);
    Border.Style := cbsThin;
    Border.Preset(cbspOutline);

    for i := RPT_START_ROW+2 to ASheet.LastRow do
    begin
      for j := 0 to ASheet.LastCol do
        aCmdFormat.Apply(j, i, j, i);
    end;
  end;
end;

procedure TdmSitePhoneSummary.XLS_Init(AFileName: string);
var
  aDefFmt: TXLSDefaultFormat;
begin
	with XLSRW do
  begin
    Filename := AFileName;
    // 建立預設的儲存格格式
    CmdFormat.BeginEdit(Nil);
    CmdFormat.Font.Name := '微軟正黑體';
    CmdFormat.Font.Size := 10;
    CmdFormat.Alignment.Vertical := cvaCenter;
    CmdFormat.Alignment.WrapText := True; //自動折行
    aDefFmt := CmdFormat.AddAsDefault('DefFormat');
    DefaultFormat := aDefFmt;
  end;
end;

procedure TdmSitePhoneSummary.GetData_Summary(ASite: string; ABegTime, AEndTime: TDateTime);
begin
  Log('取得[文中]的來回電匯總資料');
  
  with qrGetData do
  begin
    if Active then Close;
    SQL.Clear;
    SQL.Add('SELECT IPH4002,');
    SQL.Add('SUM(IPH4003) AS IPH4003, SUM(IPH4006) AS IPH4006,');
    SQL.Add('SUM(IPH4007) AS IPH4007, SUM(IPH4008) AS IPH4008, SUM(IPH4009) AS IPH4009,');
    SQL.Add('SUM(IPH4010) AS IPH4010, SUM(IPH4011) AS IPH4011, SUM(IPH4012) AS IPH4012,');
    SQL.Add('SUM(IPH4013) AS IPH4013, SUM(IPH4209) AS IPH4209,');
    SQL.Add('SUM(IPH4210) AS IPH4210, SUM(IPH4211) AS IPH4211, SUM(IPH4213) AS IPH4213');
    SQL.Add('FROM WICSIPH4 WITH(NOLOCK)');
    SQL.Add('GROUP BY IPH4002');
    SQL.Add('ORDER BY IPH4002');
    AddWhere('(IPH4002 BETWEEN :IPH4002B AND :IPH4002E)');
    ParamByName('IPH4002B').AsDateTime := ABegTime;
    ParamByName('IPH4002E').AsDateTime := AEndTime;
    Open;
  end;
end;

function TdmSitePhoneSummary.MakeRecipients(AEmailAddrList: TIdEmailAddressList): Integer;
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

function TdmSitePhoneSummary.GetOnDutyTotal(ADate: TDateTime): Extended;
begin
  with qrGetData do
  begin
    if Active then Close;
    SQL.Clear;
    SQL.Add('SELECT SUM(IPH4003) AS _TOTAL_');
    SQL.Add('FROM WICSIPH4 WITH(NOLOCK)');
    AddWhere('(IPH4002 = :IPH4002)');
    ParamByName('IPH4002').AsDateTime := DateOf(FCalcEndTime);
    Open;
    Result := FieldByName('_TOTAL_').AsFloat;
    Close;
  end;
end;

end.
