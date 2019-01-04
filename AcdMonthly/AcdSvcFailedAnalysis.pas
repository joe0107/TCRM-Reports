unit AcdSvcFailedAnalysis;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs, DB, ADODB, StdCtrls, ShellAPI,
  cxData, cxClasses, cxCustomData, cxDataStorage, cxDBData, JclStrings, MemDS, DBAccess, Uni, kbmMemTable, DateUtils,
  dxmdaset, XLSSheetData5, XLSReadWriteII5, XLSDbRead5, XLSNames5, IdEMailAddress, IdMessage, IdAttachmentFile, Math,
  XLSCmdFormat5, Xc12DataStyleSheet5, Xc12Utils5, TcrmConstants;

type
  TdmAcdSvcFailedAnalysis = class(TDataModule)
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
    FStartCol: array [TWtnSiteNdx] of Integer;
    FWICSIPH4: TUniQuery;
    procedure GetData_IPH4(ABegTime, AEndTime: TDateTime);
    procedure PrepareData_IPH4_Y_M(Y, M: Word);
  private
    FXlsFileName: string;
    procedure InitSiteStartCol;
    procedure Init_XLS_Report(Y, M: Word);
    procedure WriteToXls_Title(AXLS: TXLSReadWriteII5; ASheet: TXLSWorksheet);
    procedure WriteToXls_Data(ASheet: TXLSWorksheet);
    procedure WriteToXls_SummaryData(AXLS: TXLSReadWriteII5; ASheet: TXLSWorksheet);
    procedure WriteToXls_Format(AXLS: TXLSReadWriteII5; ASheet: TXLSWorksheet);
    function  FindSheet(ASheetName: string): TXLSWorksheet;
    function  GetXlsFileName(Y, M: Word): string;
    function  GetColOffset_Site(ASite: string): Integer;
  private
    procedure SendMail;
    function  MakeRecipients(AEmailAddrList: TIdEmailAddressList): Integer;
    procedure MakeCCList(AEmailAddrList: TIdEmailAddressList);
    function  MakeNotifyMessage: TIdMessage;
  public
    procedure Exec(AYear, AMonth: Word);
	end;

const
  ACD_FACTOR          = 0.8;
  ROW_START_DATA      = 2;
  COL_START_Taipei    = 1;
  COL_START_Taoyuan   = 8;
  COL_START_Taichung  = 15;
  COL_START_Tainan    = 22;
  MISC_ITEM           = '(其他)';

var
  dmAcdSvcFailedAnalysis: TdmAcdSvcFailedAnalysis;

implementation

uses
  JcSysUtils, JcDateTimeUtils, JcNumUtils, JcDevExpressUtils, JcDataSetUtils, ReportData, Main;

{$R *.dfm}

{ TdmAcdSvcFailedAnalysis }

procedure TdmAcdSvcFailedAnalysis.InitExecute;
begin

end;

procedure TdmAcdSvcFailedAnalysis.BeginExecute;
begin
  //nothing to do now
end;

procedure TdmAcdSvcFailedAnalysis.EndExecute;
begin
  //nothing to do now
end;

procedure TdmAcdSvcFailedAnalysis.Log(AMsg: string);
begin
	fmMain.Log(AMsg);
end;

procedure TdmAcdSvcFailedAnalysis.LogLine;
begin
	fmMain.LogLine;
end;

procedure TdmAcdSvcFailedAnalysis.Exec(AYear, AMonth: Word);
var
  aSheet: TXLSWorksheet;
begin
	with dmAcdSvcFailedAnalysis do
  begin
    dmReport.Calc_YM(AYear, AMonth);
    FYear  := AYear;
    FMonth := AMonth;
    //取得指定年月中接聽率未達80%的資料
    PrepareData_IPH4_Y_M(AYear, AMonth);
    //-------------------------------------------------------------------------
    Init_XLS_Report(AYear, AMonth);
    aSheet := XLSRW.Sheets[0];
    aSheet.Name := Format('%d月', [AMonth]);
    //寫入標題
    WriteToXls_Title(XLSRW, aSheet);
    WriteToXls_Data(aSheet);
    WriteToXls_Format(XLSRW, aSheet);
    WriteToXls_SummaryData(XLSRW, aSheet);    
    XLSRW.Write;
    FWICSIPH4.Close;

    if fmMain.NoMail then
      ShellExecute(0, 'open', PChar(FXlsFileName), nil, nil, SW_SHOWMAXIMIZED)
    else
      SendMail;
  end;
end;

procedure TdmAcdSvcFailedAnalysis.MakeCCList(AEmailAddrList: TIdEmailAddressList);
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

function TdmAcdSvcFailedAnalysis.MakeNotifyMessage: TIdMessage;
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
    Subject := Format('電話未達狀況分析月報表_%s(%s)', [FormatDateTime('yyyymmdd', Now), aDayOfWeek]);
    //寄件人地址
    From.Address := 'rdrepl@winton.com.tw';
    ContentType := 'multipart/mixed';
    //填入郵件內容
    Body.Text := '';
  end;

  if FileExists(FXlsFileName) then
  	TIdAttachmentFile.Create(Result.MessageParts, FXlsFileName);
end;

procedure TdmAcdSvcFailedAnalysis.SendMail;
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

procedure TdmAcdSvcFailedAnalysis.DataModuleCreate(Sender: TObject);
begin
  FWICSIPH4 := dmReport.GetQuery_WintonTcrm;
  InitSiteStartCol;  
end;

procedure TdmAcdSvcFailedAnalysis.Init_XLS_Report(Y, M: Word);
var
  aDefFmt: TXLSDefaultFormat;
begin
  FXlsFileName := GetXlsFileName(Y, M);
  XLSRW.Filename := FXlsFileName;
  // 建立預設的儲存格格式
	with XLSRW do
  begin
    CmdFormat.BeginEdit(Nil);
    CmdFormat.Font.Name := '微軟正黑體';
    CmdFormat.Font.Size := 12;
    CmdFormat.Alignment.Vertical := cvaCenter;
    CmdFormat.Alignment.WrapText := True; //自動折行
    aDefFmt := CmdFormat.AddAsDefault('DefFormat');
    DefaultFormat := aDefFmt;
  end;
end;

function TdmAcdSvcFailedAnalysis.FindSheet(ASheetName: string): TXLSWorksheet;
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

procedure TdmAcdSvcFailedAnalysis.WriteToXls_Title(AXLS: TXLSReadWriteII5; ASheet: TXLSWorksheet);
var
  i: Integer;
begin
  with ASheet do
  begin
    AsString[0, 1] := '日期';

    MergeCells(COL_START_Taipei, 0, COL_START_Taipei+6, 0);
    AsString[COL_START_Taipei, 0] := SITE_NAME_Taipei_TC;

    MergeCells(COL_START_Taoyuan, 0, COL_START_Taoyuan+6, 0);
    AsString[COL_START_Taoyuan, 0] := SITE_NAME_Taoyuan_TC;

    MergeCells(COL_START_Taichung, 0, COL_START_Taichung+6, 0);
    AsString[COL_START_Taichung, 0] := SITE_NAME_Taichung_TC;

    MergeCells(COL_START_Tainan, 0, COL_START_Tainan+6, 0);
    AsString[COL_START_Tainan, 0] := SITE_NAME_Tainan_TC;
    //---------------------------------------------------------------------------
    for i := 0 to 3 do
    begin
      AsString[1+7*i, 1] := 'ACD派送數';
      AsString[2+7*i, 1] := 'ACD處理數';
      AsString[3+7*i, 1] := '80%指標差異通數';
      AsString[4+7*i, 1] := '值機'+#13+'人力';
      AsString[5+7*i, 1] := '直接率';
      AsString[6+7*i, 1] := '平均接聽數';
      AsString[7+7*i, 1] := '每人應多接';
    end;
    //設定標題顏色
    for i := 0 to COL_START_Tainan+6 do
    begin
      with Cell[i, 1] do
      begin
        FillPatternForeColor := TXc12IndexColor(30);
        FontColor := clWhite;
      end;
    end;
  end;

	with XLSRW do
  begin
    with CmdFormat do
    begin
      BeginEdit(ASheet);
      //標題置中
      Clear;
      Alignment.Horizontal := chaCenter;
      Apply(1, 1, 28, 1);
    end;
  end;
end;

procedure TdmAcdSvcFailedAnalysis.GetData_IPH4(ABegTime, AEndTime: TDateTime);
const
  DATA_DESC = '營業處的ACD接聽率未達80%%資料(WICSIPH4)';
begin
  Log('讀取' + DATA_DESC);

  with FWICSIPH4 do
  begin
    if Active then Close;
    LocalUpdate := True;
    SQL.Clear;
		SQL.Add('SELECT *, CAST(IPH4008 AS FLOAT) / CAST(IPH4007 AS FLOAT) AS SCORE');
    SQL.Add('FROM WICSIPH4 WITH(NOLOCK)');
    SQL.Add('WHERE (IPH4002 BETWEEN :IPH4002B AND :IPH4002E)');
    SQL.Add('AND (IPH4006 <> 0)');
    SQL.Add('AND (IPH4003 > 0)');
    SQL.Add('AND (CAST(IPH4008 AS FLOAT) / CAST(IPH4007 AS FLOAT)) < 0.799');
    SQL.Add('ORDER BY IPH4002, IPH4001');

    ParamByName('IPH4002B').AsDateTime := ABegTime;
    ParamByName('IPH4002E').AsDateTime := AEndTime;

    try
      Open;
      Log(Format('取得' +  DATA_DESC + ', 記錄數 = %d', [RecordCount]));
    except
      on E: Exception do
        Log(Format('GetData_IPH4() failed, error = %s', [E.Message]));
    end;
  end;
end;

procedure TdmAcdSvcFailedAnalysis.PrepareData_IPH4_Y_M(Y, M: Word);
var
  aCalcBegTime, aCalcEndTime: TDateTime;
begin
  //統計 Y 年 M 月的逾時數
  aCalcBegTime := EncodeDate(Y, M, 1);
  aCalcEndTime := IncDay(IncMonth(aCalcBegTime, 1), -1);
  GetData_IPH4(aCalcBegTime, aCalcEndTime);
end;

function TdmAcdSvcFailedAnalysis.MakeRecipients(AEmailAddrList: TIdEmailAddressList): Integer;
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

function TdmAcdSvcFailedAnalysis.GetXlsFileName(Y, M: Word): string;
var
  aReportFolder: string;
begin
  aReportFolder := IncludeTrailingPathDelimiter(dmReport.GetReportSockFolder);
  Result := Format('電話未達狀況分析_%s.xlsx', [FormatDateTime('yyyymm', EncodeDate(Y, M, 1))]);
  Result := aReportFolder + Result;
end;

function TdmAcdSvcFailedAnalysis.GetColOffset_Site(ASite: string): Integer;
begin
  if Pos('台北', ASite) > 0 then
    Result := COL_START_Taipei
  else if Pos('北區', ASite) > 0 then
    Result := COL_START_Taoyuan
  else if Pos('中區', ASite) > 0 then
    Result := COL_START_Taichung
  else if Pos('南區', ASite) > 0 then
    Result := COL_START_Tainan
  else
  begin
    Log(Format('!! GetColOffset_Site() error, Site = %s', [ASite]));
    Result := -99999;
  end;
end;

procedure TdmAcdSvcFailedAnalysis.WriteToXls_Data(ASheet: TXLSWorksheet);
var
  aCol, aRow: Word;
  aLastDate: TDateTime;
  aFloat: Extended;
  aSiteNdx: TWtnSiteNdx;
  i, aInt, aDataLastRow: Integer;
begin
  if not JcDataSetIsValid(FWICSIPH4) then Exit;

  with FWICSIPH4 do
  begin
    First;
    aLastDate := FieldByName('IPH4002').AsDateTime;
    aRow := 2;

    while not Eof do
    begin
      aCol := GetColOffset_Site(FieldByName('IPH4001').AsString);

      if (FieldByName('IPH4002').AsDateTime <> aLastDate) then
      begin
        Inc(aRow);
        aLastDate := FieldByName('IPH4002').AsDateTime;
      end;
      //日期
      ASheet.AsDateTime[0, aRow] := FieldByName('IPH4002').AsDateTime;
      //ACD派送數
      ASheet.AsInteger[aCol, aRow] := FieldByName('IPH4007').AsInteger;
      //ACD處理數
      ASheet.AsInteger[aCol+1, aRow] := FieldByName('IPH4008').AsInteger;
      //80%指標差異通數
      ASheet.AsInteger[aCol+2, aRow] := Trunc(SimpleRoundTo(FieldByName('IPH4007').AsInteger * ACD_FACTOR, 0)) - FieldByName('IPH4008').AsInteger;
      //值機人力
      ASheet.AsFloat[aCol+3, aRow] := FieldByName('IPH4003').AsFloat;
      //直接率
      if FieldByName('IPH4007').AsFloat > 0 then
        aFloat := FieldByName('IPH4008').AsFloat / FieldByName('IPH4007').AsFloat
      else
        aFloat := 0;

      ASheet.AsFloat[aCol+4, aRow] := aFloat;
      //平均接聽數
      if (FieldByName('IPH4003').AsFloat > 0) then
        aFloat := FieldByName('IPH4007').AsInteger / FieldByName('IPH4003').AsFloat
      else
        aFloat := 0;

      ASheet.AsFloat[aCol+5, aRow] := aFloat;
      //每人應多接
      if FieldByName('IPH4003').AsFloat > 0 then
        aInt := Ceil(ASheet.AsInteger[aCol+2, aRow] / FieldByName('IPH4003').AsFloat)
      else
        aInt := 0;

      ASheet.AsInteger[aCol+6, aRow] := aInt;
      //-----------------------------------------------------------------------
      Next;
    end;
  end;
  //計算[未達平均數]的各項數值
  with ASheet do
  begin
    CalcDimensions;
    aDataLastRow := LastRow;
    aRow := aDataLastRow+1;
    //寫入標題
    AsString[0, aRow] := '未達之平均數';
    Cell[0, aRow].FillPatternForeColor := TXc12IndexColor(45);
    //寫入公式
    for aSiteNdx := Low(TWtnSiteNdx) to High(TWtnSiteNdx) do
    begin
      for i := 0 to 6 do
      begin
        aCol := FStartCol[aSiteNdx]+i;
        AsFormula[aCol, aRow] := Format('AVERAGE(%s)', [AreaToRefStr(aCol, ROW_START_DATA, aCol, aDataLastRow)]);
        Cell[aCol, aRow].FontColor := clFuchsia;
      end;
    end;
  end;
end;

procedure TdmAcdSvcFailedAnalysis.WriteToXls_Format(AXLS: TXLSReadWriteII5; ASheet: TXLSWorksheet);
var
  i, j, aCol: Integer;
  aSiteNdx: TWtnSiteNdx;
begin
	with AXLS do
  begin
    //先計算工作表使用範圍,否則第一次會取到 LastRow = -1, LastCol = -1
    ASheet.CalcDimensions;
    //畫出儲存格邊框
    CmdFormat.BeginEdit(ASheet);
    CmdFormat.Clear;
    CmdFormat.Border.Style := cbsThin;
    CmdFormat.Border.Preset(cbspOutline);

    for i := 0 to ASheet.LastRow do
    begin
      for j := 0 to ASheet.LastCol do
        CmdFormat.Apply(j , i, j, i);
    end;
    //塗上營業處間隔顏色
    ASheet.Cell[FStartCol[wsnTaipei], 0].FillPatternForeColor := TXc12IndexColor(26);
    ASheet.Cell[FStartCol[wsnTaichung], 0].FillPatternForeColor := TXc12IndexColor(26);

    for i := FStartCol[wsnTaipei] to FStartCol[wsnTaipei]+6 do
    begin
      for j := ROW_START_DATA to ASheet.LastRow do
      begin
        with ASheet.Cell[i, j] do
        begin
          FillPatternForeColor := TXc12IndexColor(26);
        end;
      end;
    end;

    for i := FStartCol[wsnTaichung] to FStartCol[wsnTaichung]+6 do
    begin
      for j := ROW_START_DATA to ASheet.LastRow do
      begin
        with ASheet.Cell[i, j] do
        begin
          FillPatternForeColor := TXc12IndexColor(26);
        end;
      end;
    end;
    //調整行寬
    with ASheet do
    begin
      for i := 1 to 28 do
        Columns[i].CharWidth := 8;
      (*
      for aSiteNdx := Low(TWtnSiteNdx) to High(TWtnSiteNdx) do
      begin
        aCol := FStartCol[aSiteNdx]+2;
        Columns[aCol].CharWidth := 11;
      end;
      *)
      Columns[0].CharWidth  := 10;
    end;
    //設定日期格式
    CmdFormat.Clear;
    CmdFormat.Alignment.Horizontal := chaCenter;
    CmdFormat.Number.Format := 'mm/dd';
    CmdFormat.Apply(0, 2, 0, ASheet.LastRow);
    //設定直接率顯示格式
    CmdFormat.Number.Format := '0.0%';
    CmdFormat.Apply(COL_START_Taipei+4, 2, COL_START_Taipei+4, ASheet.LastRow);
    CmdFormat.Apply(COL_START_Taoyuan+4, 2, COL_START_Taoyuan+4, ASheet.LastRow);
    CmdFormat.Apply(COL_START_Taichung+4, 2, COL_START_Taichung+4, ASheet.LastRow);
    CmdFormat.Apply(COL_START_Tainan+4, 2, COL_START_Tainan+4, ASheet.LastRow);
  end;
end;

procedure TdmAcdSvcFailedAnalysis.WriteToXls_SummaryData(AXLS: TXLSReadWriteII5; ASheet: TXLSWorksheet);
var
  i, aCol, aRow, aDataLastRow: Integer;
  aSiteNdx: TWtnSiteNdx;
  aCmdFormat: TXLSCmdFormat;
  aSiteName, aCellAddr1, aCellAddr2: string;
begin
  aCmdFormat := AXLS.CmdFormat;

  with ASheet do
  begin
    CalcDimensions;
    aDataLastRow := LastRow;
    //-------------------------------------------------------------------------
    MergeCells(0, aDataLastRow+2, 0, aDataLastRow+3);
    AsString[0, aDataLastRow+2] := '統計'+#13+'資訊';

    with Cell[0, aDataLastRow+2] do
    begin
      FillPatternForeColor := TXc12IndexColor(30);
      FontColor := clWhite;
    end;
    //寫入公式
    for aSiteNdx := Low(TWtnSiteNdx) to High(TWtnSiteNdx) do
    begin
      aRow := aDataLastRow+2;
      aSiteName := dmReport.GetSiteName(aSiteNdx);
      dmReport.SetUniConn_TCRM(aSiteNdx);
      //寫入[訓練人數]
      aCol := FStartCol[aSiteNdx];
      AsString[aCol, aRow] := '訓練'+#13+'人數';
      AsInteger[aCol, aRow+1] := dmReport.GetTeCount(aSiteName)-1;
      //寫入[工作天數]
      aCol := FStartCol[aSiteNdx]+1;
      AsString[aCol, aRow] := '工作'+#13+'天數';
      AsInteger[aCol, aRow+1] := dmReport.GetOnDutyDays(FYear, FMonth);
      //寫入[未達天數]的標題及公式
      aCol := FStartCol[aSiteNdx]+2;
      AsString[aCol, aRow] := '未達'+#13+'天數';
      AsFormula[aCol, aRow+1] := Format('COUNTA(%s)', [AreaToRefStr(aCol, ROW_START_DATA, aCol, aDataLastRow-1)]);
      //寫入[達成比例]的標題及公式
      aCol := FStartCol[aSiteNdx]+3;
      AsString[aCol, aRow] := '達成'+#13+'比例';
      aCellAddr1 := ColRowToRefStr(aCol-2, aRow+1);
      aCellAddr2 := ColRowToRefStr(aCol-1, aRow+1);
      AsFormula[aCol, aRow+1] := Format('(%s-%s)/%s', [aCellAddr1, aCellAddr2, aCellAddr1]);
      //寫入[未達天之投入人力比例]的標題及公式
      aCol := FStartCol[aSiteNdx]+4;
      MergeCells(aCol, aRow, aCol+1, aRow);
      AsString[aCol, aRow] := '未達天之投'+#13+'入人力比例';
      MergeCells(aCol, aRow+1, aCol+1, aRow+1);
      aCellAddr1 := ColRowToRefStr(FStartCol[aSiteNdx]+3, aDataLastRow);
      aCellAddr2 := ColRowToRefStr(FStartCol[aSiteNdx], aRow+1);
      AsFormula[aCol, aRow+1] := Format('%s/%s', [aCellAddr1, aCellAddr2]);
    end;
  end;
  
  with aCmdFormat do
  begin
    //畫出外框，標題置中，標題顏色
    BeginEdit(ASheet);
    Alignment.Horizontal := chaCenter;
    Alignment.Vertical := cvaCenter;
    Border.Style := cbsThin;
    Border.Preset(cbspOutline);

    for i := 0 to COL_START_Tainan+6 do
    begin
      Apply(i, aDataLastRow+2, i, aDataLastRow+2);
      Apply(i, aDataLastRow+3, i, aDataLastRow+3);

      with ASheet.Cell[i, aDataLastRow+2] do
      begin
        FillPatternForeColor := TXc12IndexColor(30);
        FontColor := clWhite;
      end;
    end;
    //設定 % 格式
    Number.Format := '0.0%';

    for aSiteNdx := Low(TWtnSiteNdx) to High(TWtnSiteNdx) do
      Apply(FStartCol[aSiteNdx]+3, aDataLastRow+3, FStartCol[aSiteNdx]+4, aDataLastRow+3);
    // Added by Joe 2017/10/24 11:06:05
    // 寫入統計說明
    with ASheet do
    begin
      aRow := aDataLastRow+5;
      MergeCells(0, aRow, 10, aRow);
      AsString[0, aRow] := '＊訓練人數不包含部門主管';
      //-----------------------------------------------------------------------
      aRow := aDataLastRow+6;
      MergeCells(0, aRow, 10, aRow);
      AsString[0, aRow] := '＊未達天之投入人力比例 = 值機人力的未達平均值 / 訓練人數';
    end;
  end;
end;

procedure TdmAcdSvcFailedAnalysis.InitSiteStartCol;
begin
  FStartCol[wsnTaipei] := 1;
  FStartCol[wsnTaoyuan] := 8;
  FStartCol[wsnTaichung] := 15;
  FStartCol[wsnTainan] := 22;
end;

end.
