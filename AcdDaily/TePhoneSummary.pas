unit TePhoneSummary;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs, DB, ADODB, StdCtrls, ShellAPI,
  cxData, cxClasses, cxCustomData, cxDataStorage, cxDBData, JclStrings, MemDS, DBAccess, Uni, DateUtils, XLSSheetData5,
  dxmdaset, XLSReadWriteII5, XLSDbRead5, XLSNames5, IdEMailAddress, IdMessage, IdAttachmentFile, CodeSiteLogging,
  XLSCmdFormat5, Xc12DataStyleSheet5, Xc12Utils5, XLSFormattedObj5;

type
  TdmTePhoneSummary = class(TDataModule)
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
    //取得指定日期對應的ROW
    function  GetDatRow(ASheet: TXLSWorksheet; ADate: TDateTime): Integer;
    //取得指定日期的值機總人天
    function  GetOnDutyTotal(ADate: TDateTime): Extended;
  private
    //FNewRptCount: Integer;
    FXlsFileName: string;
    function  GetReportFileName(ADate: TDateTime): string;
    procedure XLS_Init(AFileName: string);
    procedure XLS_WriteReport(ASite: string; ADataSet: TDataSet);
    procedure XLS_WriteTitle(AXLS: TXLSReadWriteII5; ASheet: TXLSWorksheet);
    procedure XLS_WriteData(AXLS: TXLSReadWriteII5; ASheet: TXLSWorksheet; ADataSet: TDataSet);
    procedure XLS_WriteHeaderText(ASheet: TXLSWorksheet; ACol, ARow: Integer; AText: string);
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
  dmTePhoneSummary: TdmTePhoneSummary;

implementation

uses
  TcrmConstants, JcDateTimeUtils, JcNumUtils, JcDevExpressUtils, JcDataSetUtils, ReportData, Main;

const
  RPT_FILE_TITLE = '個人回電效率統計表';
  RPT_START_ROW = 0;
  RPT_START_COL = 0;

{$R *.dfm}

{ TdmTePhoneSummary }

procedure TdmTePhoneSummary.FormDestroy(Sender: TObject);
begin
  inherited;
  dmTePhoneSummary := nil;
end;

procedure TdmTePhoneSummary.FormCreate(Sender: TObject);
begin
  inherited;
  Log(Format(STR_BEGIN_TO_EXE+'[%s]', [RPT_FILE_TITLE]));
  FInCalcData := False;
  InitReportConn;
end;

procedure TdmTePhoneSummary.InitExecute;
begin
  InitData;
end;

procedure TdmTePhoneSummary.BeginExecute;
begin
  //nothing to do now
end;

procedure TdmTePhoneSummary.EndExecute;
begin
  //nothing to do now
end;

procedure TdmTePhoneSummary.InitReportConn;
begin
  with dmReport do
  begin
    // 從文中資料庫取得統計數據
  	SetUniConn_TCRM(connReport, GetSiteIp(SITE_NAME_Winton_TC));
  end;
end;

procedure TdmTePhoneSummary.InitData;
begin
  // nothing to do now
end;

procedure TdmTePhoneSummary.Log(AMsg: string);
begin
	fmMain.Log(AMsg);
end;

procedure TdmTePhoneSummary.LogLine;
begin
	fmMain.LogLine;
end;

procedure TdmTePhoneSummary.XLS_WriteReport(ASite: string; ADataSet: TDataSet);
var
  aSheet: TXLSWorksheet;
begin
  Log(Format(STR_BUILD_XLS_RPT+'[%s]', [ASite]));

  with XLSRW do
  begin
    // 加入指定營業處的工作表
    if (ASite <> SITE_DESC_Taipei_TC) then
      aSheet := Add
    else
      aSheet := Sheets[0];
      
    aSheet.Name := ASite;
    XLS_WriteTitle(XLSRW, aSheet);
    XLS_WriteData(XLSRW, aSheet, ADataSet);
  end;
end;

function TdmTePhoneSummary.PrintReport(ADate: TDateTime): Boolean;
begin
  Result := False;
  FCalcBegTime := StartOfTheMonth(ADate);
  FCalcEndTime := EndOfTheDay(ADate);
  FXlsFileName := GetReportFileName(ADate);

  if GetOnDutyTotal(ADate) = 0 then
  begin
    Log('本日無人值機，'+STR_NO_BUILD_RPT);
    Exit;
  end;

  with XLSRW do
  begin
    XLS_Init(FXlsFileName);
    //--------------------------------------------------------------------------
    GetData(SITE_DESC_Taipei_TC, FCalcBegTime, FCalcEndTime);
    XLS_WriteReport(SITE_DESC_Taipei_TC, qrGetData);
    //--------------------------------------------------------------------------
    GetData(SITE_DESC_Taoyuan_TC, FCalcBegTime, FCalcEndTime);
    XLS_WriteReport(SITE_DESC_Taoyuan_TC, qrGetData);
    //--------------------------------------------------------------------------
    GetData(SITE_DESC_Taichung_TC, FCalcBegTime, FCalcEndTime);
    XLS_WriteReport(SITE_DESC_Taichung_TC, qrGetData);
    //--------------------------------------------------------------------------
    GetData(SITE_DESC_Tainan_TC, FCalcBegTime, FCalcEndTime);
    XLS_WriteReport(SITE_DESC_Tainan_TC, qrGetData);
    //--------------------------------------------------------------------------    
    Write;
  end;

  Result := True;
end;

class procedure TdmTePhoneSummary.Exec(ADate: TDateTime);
begin
	if not Assigned(dmTePhoneSummary) then
  	Application.CreateForm(TdmTePhoneSummary, dmTePhoneSummary);

	with dmTePhoneSummary do
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

function TdmTePhoneSummary.MakeAdminNotifyMessage: TIdMessage;
begin
  Result := nil;
end;

procedure TdmTePhoneSummary.MakeCCList(AEmailAddrList: TIdEmailAddressList);
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

function TdmTePhoneSummary.MakeNotifyMessage: TIdMessage;
var
  aDayOfWeek: string;
begin
  Result := TIdMessage.Create(Self);
   dmReport.Init_IdMessage(Result);    //設定郵件屬性 Added by Joe Lee 2017/11/20 09:47:24
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
    //填入郵件內容
    Body.Text := '';
  end;

  if FileExists(FXlsFileName) then
  	TIdAttachmentFile.Create(Result.MessageParts, FXlsFileName);
end;

procedure TdmTePhoneSummary.SendMail;
var
  aMsg: TIdMessage;
begin
  if fmMain.MailMode then
  begin
    aMsg := MakeNotifyMessage;
    dmReport.SendNotofyMail_SSL(aMsg);
    Log(STR_SND_RPT_BY_MAIL);
  end;
end;

function TdmTePhoneSummary.GetReportFileName(ADate: TDateTime): string;
var
  aPath: string;
begin
  aPath := IncludeTrailingPathDelimiter(dmReport.GetReportPath);
  Result := Format('%s%s_%s.xls', [aPath, RPT_FILE_TITLE, FormatDateTime('yyyymmdd', ADate)]);
end;

procedure TdmTePhoneSummary.XLS_WriteTitle(AXLS: TXLSReadWriteII5; ASheet: TXLSWorksheet);
var
  aCmdFormat: TXLSCmdFormat;
  i, aCol, aRow: Integer;
  aData: TUniQuery;
  aText: string;

  procedure SetCellText(ACol, ARow: Integer; AText: string);
  var
    aCell: TXLSCell;
  begin
    ASheet.AsString[ACol, ARow] := AText;
    aCell := ASheet.Cell[ACol, ARow];
    aCell.FillPatternForeColor := TXc12IndexColor(30);
    aCell.FontColor := clWhite;
  end;

  procedure GetDateList;
  begin
    with aData do
    begin
      SQL.Add('SELECT IPH3002 FROM WICSIPH3 A WITH(NOLOCK)');
      SQL.Add('GROUP BY IPH3002');
      SQL.Add('ORDER BY IPH3002');
      AddWhere('(IPH3002 BETWEEN :IPH3002B AND :IPH3002E)');
      AddWhere('(ISNULL(IPH3003, 0) > 0)');
      AddWhere('(IPH3007 > 0)');
      ParamByName('IPH3002B').AsDateTime := FCalcBegTime;
      ParamByName('IPH3002E').AsDateTime := FCalcEndTime;
      Open;
    end;
  end;
begin
  aData := dmReport.GetQuery(connReport);
  GetDateList;

  try
    with ASheet do
    begin
      MergeCells(RPT_START_COL, RPT_START_ROW, RPT_START_COL, RPT_START_ROW+1);
      SetCellText(RPT_START_COL, RPT_START_ROW,  '日期');
      //------------------------------------------------------------------------
      i := 0;
      aData.First;

      while not aData.Eof do
      begin
        // 日期
        aText := dmReport.GetShortDateDesc(aData.FieldByName('IPH3002').AsDateTime);
        aCol := RPT_START_COL;
        aRow := RPT_START_ROW+2+i;
        SetCellText(aCol, aRow, aText);
        aData.Next;
        Inc(i);
      end;
      CalcDimensions;
      SetCellText(aCol, LastRow+1, '合計');
      SetCellText(aCol, LastRow+2, '平均');
      // 設定欄寬
      Columns[RPT_START_COL].CharWidth := 10;   // 日期
      // 凍結首行
      FreezePanes(1, 0);
    end;
  finally
    aData.Free;
  end;

  with ASheet do
  begin
    aCmdFormat := AXLS.CmdFormat;
    CalcDimensions;

    with aCmdFormat do
    begin
      //畫出外框，標題置中，標題顏色
      BeginEdit(ASheet);
      Alignment.Horizontal := chaCenter;
      Alignment.Vertical := cvaCenter;
      Border.Style := cbsThin;
      Border.Preset(cbspOutline);
      Apply(RPT_START_COL, RPT_START_ROW, RPT_START_COL, LastRow);
    end;
  end;
end;

procedure TdmTePhoneSummary.GetData(ASite: string; ABegTime, AEndTime: TDateTime);
begin
  Log(Format('取得個人來回電資料[%s]', [ASite]));

  with qrGetData do
  begin
    if Active then Close;
    SQL.Clear;
    SQL.Add('SELECT');
    SQL.Add('A.IPH3002, A.IPH3001, B.SALE002, C.DEPT001, C.DEPT002,');
    SQL.Add('A.IPH3006, A.IPH3007*0.5 AS IPH3007, A.IPH3008, A.IPH3010,');
    SQL.Add('A.IPH3011, A.IPH3012, A.IPH3013, A.IPH3017*0.5 AS IPH3017');
    SQL.Add('FROM WICSIPH3 A WITH(NOLOCK)');
    SQL.Add('LEFT JOIN WICSSALE B WITH(NOLOCK) ON SALE001 = IPH3001');
    SQL.Add('LEFT JOIN WICSDEPT C WITH(NOLOCK) ON DEPT001 = SALE003');
    SQL.Add('ORDER BY DEPT001, IPH3001, IPH3002');

    AddWhere('(IPH3002 >= :IPH3002B AND IPH3002 <= :IPH3002E)');
    AddWhere('(IPH3007 > 0)');

    if (ASite = SITE_DESC_Taipei_TC) then
      AddWhere('(DEPT001 LIKE ''02%'')')
    else if (ASite = SITE_DESC_Taoyuan_TC) then
      AddWhere('(DEPT001 = ''052'' OR DEPT001 = ''062'')')
    else if (ASite = SITE_DESC_Taichung_TC) then
      AddWhere('(DEPT001 LIKE ''07%'')')
    else if (ASite = SITE_DESC_Tainan_TC) then
      AddWhere('(DEPT001 = ''082'' OR DEPT001 = ''092'')');

    ParamByName('IPH3002B').AsDateTime := ABegTime;
    ParamByName('IPH3002E').AsDateTime := AEndTime;
    Open;
  end;
end;

function TdmTePhoneSummary.GetDatRow(ASheet: TXLSWorksheet; ADate: TDateTime): Integer;
var
  aText: string;
  i: Integer;
begin
  Result := -1;
  aText := dmReport.GetShortDateDesc(ADate);

  with ASheet do
  begin
    for i := 0 to LastRow do
    begin
      if AsString[RPT_START_COL, i] = aText then
      begin
        Result := i;
        CodeSite.SendFmtMsg('Date = %s, Row = %d', [DateTimeToStr(ADate), Result]);
        Exit;
      end;
    end;
  end;
end;

procedure TdmTePhoneSummary.XLS_WriteData(AXLS: TXLSReadWriteII5; ASheet: TXLSWorksheet; ADataSet: TDataSet);
const
  NUM_FMT_INT     = '#,#';
  NUM_FMT_FLOAT_1 = '###0.0';
var
  aCol, aRow, aNdx, i, j, aIPH3006, aIPH3013_Total: Integer;
  aTE, aDept, aText: string;
  aCmdFormat: TXLSCmdFormat;
  aNdxColor: Txc12IndexColor;
  aIPH3007_Total, aIPH3017_Total: Extended;

  procedure WriteStaffHeader(ASheet: TXLSWorksheet; ACol: Integer; IPH3006, IPH3013: Integer);
  begin
    with ASheet do
    begin
      XLS_WriteHeaderText(ASheet, ACol, RPT_START_ROW, Format('%s (%d-%d)', [aTE, IPH3006, IPH3013]));
      MergeCells(ACol, RPT_START_ROW, ACol+2, RPT_START_ROW);
    end;
  end;

  procedure WriteOnDutySummary(ASheet: TXLSWorksheet; ACol, ARow, ACount: Integer; IPH3007, IPH3017: Extended);
  begin
    with ASheet do
    begin
      // 值機天數的合計值
      AsString[ACol, ARow-1] := Format('%2.1f / %2.1f', [IPH3007, IPH3017]);
      Cell[ACol, ARow-1].FillPatternForeColor := TXc12IndexColor(43);
      // 值機天數的平均值
      AsString[ACol, ARow] := Format('%2.1f / %2.1f', [IPH3007/ACount, IPH3017/ACount]);
      Cell[ACol, ARow].FillPatternForeColor := TXc12IndexColor(43);
      // ACD處理數的平均值
      AsFormula[ACol+1, ARow] := Format('%s/%f', [ColRowToRefStr(ACol+1, ARow-1), IPH3007]);
      Cell[ACol+1, ARow].FillPatternForeColor := TXc12IndexColor(43);
      Cell[ACol+1, ARow].NumberFormat := NUM_FMT_FLOAT_1;
    end;
  end;
begin
  if not JcDataSetIsValid(ADataSet) then Exit;
  //============================================================================
  with ADataSet do
  begin
    First;
    aNdx := 0;
    aTE := '';
    aDept := '';

    while not Eof do
    begin
      with ASheet do
      begin
        // 紀錄目前正在處理的人員
        aText := Format('%s_%s', [FieldByName('IPH3001').AsString, FieldByName('SALE002').AsString]);
        // 開始產生新出現人員的資料
        if (aTE <> aText) then
        begin
          // 先寫入前一位人員的[值機天數]的合計及平均值
          if (aNdx > 0) then
          begin
            WriteOnDutySummary(ASheet, aCol, LastRow, LastRow-3,  aIPH3007_Total, aIPH3017_Total);
            WriteStaffHeader(ASheet, aCol, aIPH3006, aIPH3013_Total);
          end;
          //--------------------------------------------------------------------
          aTE := aText;
          aCol := aNdx*3+1;
          aRow := RPT_START_ROW;
          Inc(aNdx);
          aIPH3006 := FieldByName('IPH3006').AsInteger;
          // 寫入標題 ----------------------------------------------------------
          // 訓練師
//          XLS_WriteHeaderText(ASheet, aCol, aRow, Format('%s (%d)', [aTE, FieldByName('IPH3006').AsInteger]));
//          MergeCells(aCol, aRow, aCol+2, aRow);
          // 值機天數
          XLS_WriteHeaderText(ASheet, aCol, aRow+1, '值機'+#10+'天數');
          Columns[aCol].CharWidth := 11;  // 設定欄寬
          // ACD處理數
          XLS_WriteHeaderText(ASheet, aCol+1, aRow+1, 'ACD'+#10+'處理數');
          // 非派送回電數
          XLS_WriteHeaderText(ASheet, aCol+2, aRow+1, '非派送'+#10+'回電數');
          // 設定群組交錯顏色 --------------------------------------------------
          if (aDept = '') then
          begin
            aNdxColor := TXc12IndexColor(27);
            aDept := FieldByName('DEPT002').AsString;
          end
          else if (aDept <> FieldByName('DEPT002').AsString) then
          begin
            if (aNdxColor = TXc12IndexColor(27)) then
              aNdxColor := TXc12IndexColor(31)
            else
              aNdxColor := TXc12IndexColor(27);

            aDept := FieldByName('DEPT002').AsString;
          end;
          // 值機天數統計值歸零
          aIPH3007_Total := 0;
          aIPH3017_Total := 0;
          aIPH3013_Total := 0;
          //--------------------------------------------------------------------
        end;

        aRow := GetDatRow(ASheet, FieldByName('IPH3002').AsDateTime);
        //CodeSite.SendFmtMsg('ROW = %d', [aRow]);
        // 寫入資料
        if (aRow <> -1) then
        begin
          aIPH3007_Total := aIPH3007_Total + FieldByName('IPH3007').AsFloat;
          aIPH3017_Total := aIPH3017_Total + FieldByName('IPH3017').AsFloat;

          if not FieldByName('IPH3013').AsBoolean then
            Inc(aIPH3013_Total);
          //--------------------------------------------------------------------
          AsString[aCol, aRow] := Format('%2.1f / %2.1f', [FieldByName('IPH3007').AsFloat, FieldByName('IPH3017').AsFloat]);
          //AsFloat[aCol, aRow] := FieldByName('IPH3007').AsFloat;
          AsInteger[aCol+1, aRow] := FieldByName('IPH3008').AsInteger;
          AsInteger[aCol+2, aRow] := FieldByName('IPH3010').AsInteger;
          // 設定群組顏色
          for j := 0 to 2 do
            Cell[aCol+j, aRow].FillPatternForeColor := aNdxColor;

          if not FieldByName('IPH3013').AsBoolean then
            Cell[aCol+1, aRow].FontColor := clBlue  // !! 這樣顯示出才會來是紅色
        end
      end;
      Next;
    end;
    // 寫入最後一位人員的[值機天數]的合計及平均值
    WriteOnDutySummary(ASheet, aCol, ASheet.LastRow, ASheet.LastRow-3,  aIPH3007_Total, aIPH3017_Total);
    WriteStaffHeader(ASheet, aCol, FieldByName('IPH3006').AsInteger, aIPH3013_Total);
  end;
  //============================================================================
  // 寫入 合計/平均 公式
  with ASheet do
  begin
    CalcDimensions;

    for aCol := 1 to LastCol do
    begin
      if (aCol mod 3) <> 1 then
      begin
        AsFormula[aCol, LastRow-1] := Format('SUM(%s)', [AreaToRefStr(aCol, RPT_START_ROW+2, aCol, LastRow-2)]);
        Cell[aCol, LastRow-1].FillPatternForeColor := TXc12IndexColor(43);
        Cell[aCol, LastRow-1].NumberFormat := NUM_FMT_FLOAT_1;
      end;

      if (aCol mod 3) = 0 then
      begin
        AsFormula[aCol, LastRow] := Format('AVERAGE(%s)', [AreaToRefStr(aCol, RPT_START_ROW+2, aCol, LastRow-2)]);
        Cell[aCol, LastRow].FillPatternForeColor := TXc12IndexColor(43);
        Cell[aCol, LastRow].NumberFormat := NUM_FMT_FLOAT_1;
      end
    end;
  end;
  //============================================================================
  // 畫出外框
  ASheet.CalcDimensions;

  with AXLS.CmdFormat do
  begin
    BeginEdit(ASheet);
    Border.Style := cbsThin;
    Border.Preset(cbspOutline);

    for i := RPT_START_ROW to ASheet.LastRow do
    begin
      for j := 0 to ASheet.LastCol do
        Apply(j, i, j, i);
    end;
  end;
  // 標題置中
  with AXLS.CmdFormat do
  begin
    BeginEdit(ASheet);
    Alignment.Horizontal := chaCenter;
    // 將人員的標題置中
    for i := RPT_START_ROW to ASheet.LastRow do
    begin
      for j := 0 to ASheet.LastCol do
        Apply(j, RPT_START_ROW, j, RPT_START_ROW+1);
    end;

    i := 0;
    while True do
    begin
      aCol := RPT_START_COL+3*i+1;
      if (aCol > ASheet.LastCol) then Break;
      Apply(aCol, RPT_START_ROW+2, aCol, ASheet.LastRow);
      Inc(i);
    end;
  end;
  // 將[合計/平均]設成租體字
  with AXLS.CmdFormat do
  begin
    BeginEdit(ASheet);
    Font.Bold := True;
    Apply(RPT_START_COL+1, ASheet.LastRow-1, ASheet.LastCol, ASheet.LastRow);
  end;
end;

procedure TdmTePhoneSummary.XLS_WriteHeaderText(ASheet: TXLSWorksheet; ACol, ARow: Integer; AText: string);
var
  aCell: TXLSCell;
begin
  ASheet.AsString[ACol, ARow] := AText;
  aCell := ASheet.Cell[ACol, ARow];
  aCell.FillPatternForeColor := TXc12IndexColor(30);
  aCell.FontColor := clWhite;
end;

procedure TdmTePhoneSummary.XLS_Init(AFileName: string);
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

function TdmTePhoneSummary.MakeRecipients(AEmailAddrList: TIdEmailAddressList): Integer;
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

function TdmTePhoneSummary.GetOnDutyTotal(ADate: TDateTime): Extended;
begin
  with qrGetData do
  begin
    if Active then Close;
    SQL.Clear;
    SQL.Add('SELECT SUM(IPH3007) AS _TOTAL_');
    SQL.Add('FROM WICSIPH3 WITH(NOLOCK)');
    AddWhere('(IPH3002 = :IPH3002)');
    ParamByName('IPH3002').AsDateTime := DateOf(FCalcEndTime);
    Open;
    Result := FieldByName('_TOTAL_').AsFloat;
    Close;
  end;
end;

end.
