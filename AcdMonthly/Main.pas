unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs, DB, DateUtils, dxmdaset, JcLog,
  XLSSheetData5, XLSReadWriteII5, XLSDbRead5, Xc12DataStyleSheet5, XLSCmdFormat5, Xc12Utils5, XLSComment5, XLSDrawing5,
  IdEMailAddress, IdMessage, IdBaseComponent, Grids, DBGrids, IdComponent, Uni, ShellAPI, IdAttachmentFile, StdCtrls,
  ComCtrls, ExtCtrls, JcVersionInfo, Math, JcNumUtils, kbmMemTable, MemDS, DBAccess, ComObj;

type
  TfmMain = class(TForm)
    XLSDbRead51: TXLSDbRead5;
    XLSReadWriteII51: TXLSReadWriteII5;
    pnl1: TPanel;
    Label1: TLabel;
    DateTimePickerBegin: TDateTimePicker;
    JcVersionInfo1: TJcVersionInfo;
    JcLog: TJcLog;
    mtSiteSummary: TkbmMemTable;
    mtSiteSummaryYear: TIntegerField;
    mtSiteSummaryMonth: TIntegerField;
    mtSiteSummaryTaipei_YesDays: TIntegerField;
    mtSiteSummaryTaipei_NoDays: TIntegerField;
    mtSiteSummaryTaipei_Score: TFloatField;
    mtSiteSummaryTaoyuan_YesDays: TIntegerField;
    mtSiteSummaryTaoyuan_NoDays: TIntegerField;
    mtSiteSummaryTaoyuan_Score: TFloatField;
    mtSiteSummaryTaichung_YesDays: TIntegerField;
    mtSiteSummaryTaichung_NoDays: TIntegerField;
    mtSiteSummaryTaichung_Score: TFloatField;
    mtSiteSummaryTainan_YesDays: TIntegerField;
    mtSiteSummaryTainan_NoDays: TIntegerField;
    mtSiteSummaryTainan_Score: TFloatField;
    mtSiteSummaryWinton_YesDays: TIntegerField;
    mtSiteSummaryWinton_NoDays: TIntegerField;
    mtSiteSummaryWInton_Score: TFloatField;
    mtTeSummary: TkbmMemTable;
    mtTeSummaryYesDays_1: TIntegerField;
    mtTeSummaryNoDays_1: TIntegerField;
    mtTeSummaryScore_1: TFloatField;
    mtTeSummaryYesDays_2: TIntegerField;
    mtTeSummaryNoDays_2: TIntegerField;
    mtTeSummaryScore_2: TFloatField;
    mtTeSummaryYesDays_3: TIntegerField;
    mtTeSummaryNoDays_3: TIntegerField;
    mtTeSummaryScore_3: TFloatField;
    mtTeSummaryYesDays_4: TIntegerField;
    mtTeSummaryNoDays_4: TIntegerField;
    mtTeSummaryScore_4: TFloatField;
    mtTeSummaryYesDays_5: TIntegerField;
    mtTeSummaryNoDays_5: TIntegerField;
    mtTeSummaryScore_5: TFloatField;
    mtTeSummaryEmpNo: TStringField;
    mtTeSummaryEmpName: TStringField;
    mtTeSummaryYesDays_6: TIntegerField;
    mtTeSummaryNoDays_6: TIntegerField;
    mtTeSummaryScore_6: TFloatField;
    mtTeSummaryYesDays_7: TIntegerField;
    mtTeSummaryYesDays_8: TIntegerField;
    mtTeSummaryYesDays_9: TIntegerField;
    mtTeSummaryYesDays_10: TIntegerField;
    mtTeSummaryYesDays_11: TIntegerField;
    mtTeSummaryYesDays_12: TIntegerField;
    mtTeSummaryNoDays_7: TIntegerField;
    mtTeSummaryNoDays_8: TIntegerField;
    mtTeSummaryNoDays_9: TIntegerField;
    mtTeSummaryNoDays_10: TIntegerField;
    mtTeSummaryNoDays_11: TIntegerField;
    mtTeSummaryNoDays_12: TIntegerField;
    mtTeSummaryScore_7: TFloatField;
    mtTeSummaryScore_8: TFloatField;
    mtTeSummaryScore_9: TFloatField;
    mtTeSummaryScore_10: TFloatField;
    mtTeSummaryScore_11: TFloatField;
    mtTeSummaryScore_12: TFloatField;
    mtTeSummaryYesDays_Total: TIntegerField;
    mtTeSummaryNoDays_Total: TIntegerField;
    mtTeSummaryScore_Total: TFloatField;
    DateTimePickerEnd: TDateTimePicker;
    Label2: TLabel;
    btnRunReport: TButton;
    btnDebug: TButton;
    procedure FormCreate(Sender: TObject);
    procedure btnRunReportClick(Sender: TObject);
    procedure btnDebugClick(Sender: TObject);
  private
    FAutoMode: Boolean;
    FDebugMode: Boolean;
    FNoMail: Boolean;
    // 檢查執行模式
    procedure CheckExeMode;
  private
    FReportName: string;
    procedure InitSysParams;
    //產生ACD接聽率月報表
    procedure DoReport;
  private
    // ACD接聽率月報表=========================================================
    FXlsFileName: string;
  	// 報表統計起訖時間
    FCalcBeginTime, FCalcEndTime: TDateTime;
    FWICSIPH3, FWICSIPH4: TUniQuery;
    // 產生報表資料
    function	PrepareData_IPH3(ABeginTime, AEndTime: TDateTime): Integer;
    function	PrepareData_IPH4(ABeginTime, AEndTime: TDateTime): Integer;
		// 取得指定區間的訓練師ACD接聽資料
    procedure GetData_IPH3(ABeginTime, AEndTime: TDateTime);
    procedure FixData_IPH3(ADataSet: TDataSet);
    procedure UpdSummaryData_IPH3(AEmpNo, AEmpName: string; ADataNdx, APhoneCount, ARequired, AJobDays: Integer);
    procedure CalcSummaryData_IPH3;
		// 取得指定區間的營業處ACD接聽資料
    procedure GetData_IPH4(ABeginTime, AEndTime: TDateTime);
    procedure UpdSummaryData_IPH4(ASite: string; ADate: TDateTime; AScore: Extended);
    procedure CalcSummaryData_IPH4;

    procedure PrepareXLS;
    // 寫入來源資料清單
    procedure WriteXls_SiteSummary;
    procedure WriteXls_SiteData(ADataSet: TDataSet);
    procedure WriteXls_TeSummary;
    procedure WriteXls_TeData(ADataSet: TDataSet);
	protected
    procedure SendMail;
    procedure MakeCCList(AEmailAddrList: TIdEmailAddressList);
    function  MakeNotifyMessage: TIdMessage;
    // ------------------------------------------------------------------------
    procedure SendAdminMail;
    function  MakeAdminNotifyMessage: TIdMessage;
    // ------------------------------------------------------------------------
    procedure ErrorMsg(AMsg: string);
  public
		procedure Exec;
    procedure Log(AMsg: string);
    procedure LogLine;
    procedure CallExcelToSaveAs(AFileName: string);
  public
    property AutoMode: Boolean read FAutoMode;
    property DebugMode: Boolean read FDebugMode;
    property NoMail: Boolean read FNoMail;
  end;

var
  fmMain: TfmMain;

implementation

uses ReportData, TcrmConstants, PhoneAnalysis, AcdSvcFailedAnalysis;

{$R *.dfm}

procedure TfmMain.WriteXls_SiteSummary;
const
  RPT_TITLE = 'ACD接聽率月報表_營業處';
var
  i: Integer;
  aSheet: TXLSWorksheet;
  aDefFmt: TXLSDefaultFormat;
begin
 	if (not mtSiteSummary.Active) or mtSiteSummary.IsEmpty then
  begin
    JcLog.Write(Format('沒有[%s]資料，不產生XLS工作表。', [RPT_TITLE]));
  	Exit;
  end;

  aSheet := XLSReadWriteII51.Sheets[0];
  //aSheet := XLSReadWriteII51.Add;
  // 凍結標題橫列
  aSheet.FreezePanes(0, 1);
  // 建立預設的儲存格格式
	with XLSReadWriteII51 do
  begin
    CmdFormat.BeginEdit(Nil);
    CmdFormat.Border.Style := cbsThin;
    CmdFormat.Border.Preset(cbspOutline);
    aDefFmt := CmdFormat.AddAsDefault('Format1');
    DefaultFormat := aDefFmt;
  end;
	// 將資料讀入工作表中
  with XLSDbRead51 do
  begin
    Sheet := 0;
    //Sheet := aSheet.Index;
    Dataset := mtSiteSummary;
    Read;
  end;
  // 調整格式
  with XLSReadWriteII51 do
  begin
    with aSheet do
    begin
      Name := RPT_TITLE;
      //-----------------------------------------------------------------------
      for i := 2 to 16 do
        Columns[i].CharWidth := 8;
      //-----------------------------------------------------------------------
      for i := 1 to 12 do
      begin
        Cell[4, i].NumberFormat := '0%';
        Cell[7, i].NumberFormat := '0%';
        Cell[10, i].NumberFormat := '0%';
        Cell[13, i].NumberFormat := '0%';
        Cell[16, i].NumberFormat := '0%';
      end;
      // 塗上標條列底色
      for i := 0 to 16 do
      begin
        if Cell[i, 0] <> nil then
        begin
          Cell[i, 0].FillPatternForeColorRGB := clSilver;
          Cell[i, 0].FontStyle := [xfsBold];
        end;
      end;
      // 為不同營業處塗上交錯底色
      CmdFormat.Mode := xcfmMerge;
      CmdFormat.BeginEdit(aSheet);
      CmdFormat.Fill.BackgroundColor.RGB := $00C6FFFF;

      for i := 0 to 2 do
        CmdFormat.Apply(2+3*2*i, 1, 4+3*2*i, LastRow);
      //-----------------------------------------------------------------------
      InsertRows(0, 1);

      AsString[2, 0] := '台北';
      AsString[5, 0] := '北區';
      AsString[8, 0] := '中區';
      AsString[11, 0] := '南區';
      AsString[14, 0] := '文中';

      MergeCells(2, 0, 4, 0);
      MergeCells(5, 0, 7, 0);
      MergeCells(8, 0, 10, 0);
      MergeCells(11, 0, 13, 0);
      MergeCells(14, 0, 16, 0);
      // 塗上標條列底色
      for i := 0 to 16 do
      begin
        if Cell[i, 0] <> nil then
        begin
          Cell[i, 0].FillPatternForeColorRGB := clSilver;
          Cell[i, 0].FontStyle := [xfsBold];
        end;
      end;
    end;
  end;
  JcLog.Write(Format('已產生[%s]工作表', [RPT_TITLE]));
end;

procedure TfmMain.FormCreate(Sender: TObject);
begin
  // 檢查是否為自動執行模式
  CheckExeMode;
  InitSysParams;

  if FAutoMode then
  begin
    try
      Exec;
      dmPhoneAnalysis.Exec(0, 0);
      // Added by Joe 2017/10/24 14:00:52
      // 產生電話未達分析報告月報表
      dmAcdSvcFailedAnalysis.Exec(0, 0);
      //-----------------------------------------------------------------------
      try
        SendAdminMail;
      except
        on E: Exception do
          JcLog.Write(Format('傳送管理報表發生異常，Err = %s', [E.Message]));
      end;
    finally
      Application.Terminate;
    end;
  end;
end;

function TfmMain.MakeNotifyMessage: TIdMessage;
var
  aList: TStringList;
  i: Integer;
begin
  Result := TIdMessage.Create(Self);

  with Result do
  begin
    if not DebugMode then
    begin
      //填入收件者
      aList := dmReport.GetAllEmail_TE_Admin;

      try
        for i := 0 to aList.Count - 1 do
        begin
          Recipients.Add.Address := aList[i];
          JcLog.Write(Format('收件人 = %s', [aList[i]]));
        end;
      finally
        aList.Free;
      end;
      //填入副本
      aList := dmReport.GetAllEmail_Site_Admin;

      try
        for i := 0 to aList.Count - 1 do
        begin
          CCList.Add.Address := aList[i];
          JcLog.Write(Format('副本 = %s', [aList[i]]));
        end;
        MakeCCList(CCList);        
      finally
        aList.Free;
      end;
    end
    else
      Recipients.Add.Address := 'joelee@winton.com.tw';
    //填入郵件表頭資訊
    Subject := ExtractFileName(FXlsFileName);
    //寄件人地址
    From.Address := 'rdrepl@winton.com.tw';
    ContentType := 'multipart/mixed';
    //填入郵件內容
    Body.Text := '';
  end;

  if FileExists(FXlsFileName) then
  	TIdAttachmentFile.Create(Result.MessageParts, FXlsFileName);
end;

procedure TfmMain.PrepareXLS;
var
  aReportFolder: string;
begin
  aReportFolder := ExtractFilePath(Application.ExeName) + 'ReportStock';
  ForceDirectories(aReportFolder);
  FXlsFileName := aReportFolder + Format('\%s.xlsx', [FReportName]);

  with XLSReadWriteII51 do
  begin
    Filename := FXlsFileName;
  end;

  WriteXls_SiteSummary;
  WriteXls_TeSummary;
  WriteXls_SiteData(FWICSIPH4);
  WriteXls_TeData(FWICSIPH3);

  XLSReadWriteII51.Write;
  JcLog.Write(Format('產生XLS檔案 = %s', [FXlsFileName]));

  if not FAutoMode then
 	  ShellExecute(0, 'open', PChar(FXlsFileName), nil, nil, SW_SHOWMAXIMIZED);
end;

procedure TfmMain.Exec;
begin
  FCalcBeginTime := DateTimePickerBegin.Date;
  FCalcEndTime := DateTimePickerEnd.Date;
  //產生ACD接聽率月報表
  DoReport;
  //傳送報表
  SendMail;
end;

procedure TfmMain.GetData_IPH4(ABeginTime, AEndTime: TDateTime);
begin
  JcLog.Write('讀取營業處ACD接聽資料(WICSIPH4)');

  with FWICSIPH4 do
  begin
    if Active then Close;
    LocalUpdate := True;
    SQL.Clear;
		SQL.Add('SELECT RID, IPH4001, IPH4002, IPH4004 FROM WICSIPH4 WITH(NOLOCK)');
    SQL.Add('WHERE (IPH4004 > 0) AND (IPH4002 BETWEEN :IPH4002B AND :IPH4002E)');
    SQL.Add('ORDER BY IPH4001, IPH4002');
    ParamByName('IPH4002B').AsDateTime := ABeginTime;
    ParamByName('IPH4002E').AsDateTime := AEndTime;

    try
      Open;
      JcLog.Write(Format('取得營業處ACD接聽資料(WICSIPH4)，記錄數 = %d', [RecordCount]));
    except
      on E: Exception do
        JcLog.Write(Format('GetData_IPH4() failed, error = %s', [E.Message]));
    end;
  end;
end;

procedure TfmMain.btnRunReportClick(Sender: TObject);
begin
  Exec;
  dmPhoneAnalysis.Exec(0, 0);
  //產生電話未達狀況分析月報表
  dmAcdSvcFailedAnalysis.Exec(0, 0);
  ShowMessage('Done');
end;

function TfmMain.PrepareData_IPH4(ABeginTime, AEndTime: TDateTime): Integer;
var
  aSite, aHost, aText: string;
  aDate: TDateTime;
  aScore: Extended;
begin
  Result := 0;
  
  JcLog.Write(Format('開始整理營業處ACD月報表資料, %s ~ %s',
    [FormatDateTime('yyyy/mm/dd', ABeginTime), FormatDateTime('yyyy/mm/dd', AEndTime)]));

  if mtSiteSummary.Active then mtSiteSummary.Close;

	with dmReport do
  begin
    aHost := GetBranchHostIp(BRANCH_NAME_Winton_TC);

    try
      if (ABeginTime > AEndTime) then
      begin
        aText := 'PrepareData_IPH4() error, 起始時間不可以大於結束時間';
        ErrorMsg(aText);
        Exit;
      end;

      if not SetConn_Tcrm(aHost) then
      begin
        aText := Format('PrepareData_IPH4() error, 無法連線到主機 %s', [aHost]);
        ErrorMsg(aText);
        Exit;
      end;
			// 取得營業處ACD接聽資料
      GetData_IPH4(ABeginTime, AEndTime);
      Result := FWICSIPH4.RecordCount;
      mtSiteSummary.DisableControls;

      with FWICSIPH4 do
      begin
        First;

        while not Eof do
        begin
          aSite  := FieldByName('IPH4001').AsString;
          aDate  := FieldByName('IPH4002').AsDateTime;
          aScore := FieldByName('IPH4004').AsFloat;
          UpdSummaryData_IPH4(aSite, aDate, aScore);
          Next;
          Application.ProcessMessages;
        end;
      end;
      CalcSummaryData_IPH4;
      JcLog.Write(Format('已完成營業處ACD月報表資料整理(%d)', [Result]));
		finally
      mtSiteSummary.EnableControls;
    end;
  end;
end;

procedure TfmMain.MakeCCList(AEmailAddrList: TIdEmailAddressList);
begin
  with AEmailAddrList do
  begin
    if FDebugMode then
    begin
      //Add.Address := 'f07@winton.com.tw';
      Add.Address := 'joe0107@gmail.com';
      //Add.Address := 'wintonjoelee@gmail.com';
      //Add.Address := 'c45@winton.com.tw';
    end
    else
    begin
      Add.Address := 'orderchen@winton.com.tw';
      Add.Address := 'ericl@winton.com.tw';
      Add.Address := 'Tony@winton.com.tw';
      Add.Address := 'trista62@winton.com.tw';
      Add.Address := 'sky@winton.com.tw';
      Add.Address := 'joelee@winton.com.tw';
    end;
  end;
end;

procedure TfmMain.SendMail;
var
  aMsg: TIdMessage;
begin
  aMsg := MakeNotifyMessage;

  if not NoMail then
  begin
    dmReport.SendNotofyMail_SSL(aMsg);
    JcLog.Write('已透過郵件傳送報表');
  end;
end;

procedure TfmMain.DoReport;
begin
  PrepareData_IPH4(FCalcBeginTime, FCalcEndTime);
  PrepareData_IPH3(FCalcBeginTime, FCalcEndTime);
  PrepareXLS;

  FWICSIPH3.Close;
  FWICSIPH4.Close;
  mtSiteSummary.Close;
  mtTeSummary.Close;
end;

procedure TfmMain.SendAdminMail;
var
  aMsg: TIdMessage;
begin
  aMsg := MakeAdminNotifyMessage;

  if not NoMail then
    dmReport.SendNotofyMail_SSL(aMsg);
end;

function TfmMain.MakeAdminNotifyMessage: TIdMessage;
var
  aDayOfWeek: string;
begin
  Result := TIdMessage.Create(Self);
  aDayOfWeek := GetChineseNumStr(DayOfWeek(FCalcBeginTime) - 1);

  with Result do
  begin
    //填入收件者
    Recipients.EMailAddresses := dmReport.AdminEmail;
    //填入郵件表頭資訊
    Subject := Format('ACD月報表執行報告_%s', [FormatDateTime('yyyymmdd', Now)]);
    //寄件人地址
    From.Address := 'rdrepl@winton.com.tw';
    ContentType := 'multipart/mixed';
    //填入郵件內容
    Body.Text := '';
  end;

  if FileExists(JcLog.FileName) then
  	TIdAttachmentFile.Create(Result.MessageParts, JcLog.FileName);
end;

procedure TfmMain.CalcSummaryData_IPH4;
var
  aTotalDays: Integer;

  procedure _CalcSummaryData_IPH4(AYesDays, ANoDays: TIntegerField; aScore: TFloatField);
  var
    aTotalDays: Integer;
  begin
    aTotalDays := AYesDays.AsInteger + ANoDays.AsInteger;

    if (aTotalDays = 0) then
      aScore.AsFloat :=  0
    else
      aScore.AsFloat := AYesDays.AsInteger / aTotalDays;
  end;
begin
  with mtSiteSummary do
  begin
    if (not Active) or IsEmpty then
      Exit;
    First;

    while not Eof do
    begin
      Edit;
      // 台北
      _CalcSummaryData_IPH4(mtSiteSummaryTaipei_YesDays, mtSiteSummaryTaipei_NoDays, mtSiteSummaryTaipei_Score);
      // 北區
      _CalcSummaryData_IPH4(mtSiteSummaryTaoyuan_YesDays, mtSiteSummaryTaoyuan_NoDays, mtSiteSummaryTaoyuan_Score);
      // 中區
      _CalcSummaryData_IPH4(mtSiteSummaryTaichung_YesDays, mtSiteSummaryTaichung_NoDays, mtSiteSummaryTaichung_Score);
      // 南區
      _CalcSummaryData_IPH4(mtSiteSummaryTainan_YesDays, mtSiteSummaryTainan_NoDays, mtSiteSummaryTainan_Score);
      // 文中
      mtSiteSummaryWinton_YesDays.AsInteger := mtSiteSummaryTaipei_YesDays.AsInteger +
        mtSiteSummaryTaoyuan_YesDays.AsInteger + mtSiteSummaryTaichung_YesDays.AsInteger +
        mtSiteSummaryTainan_YesDays.AsInteger;

      mtSiteSummaryWinton_NoDays.AsInteger := mtSiteSummaryTaipei_NoDays.AsInteger +
        mtSiteSummaryTaoyuan_NoDays.AsInteger + mtSiteSummaryTaichung_NoDays.AsInteger +
        mtSiteSummaryTainan_NoDays.AsInteger;

      aTotalDays := mtSiteSummaryWinton_YesDays.AsInteger + mtSiteSummaryWinton_NoDays.AsInteger;

      if (aTotalDays = 0) then
        mtSiteSummaryWinton_Score.AsFloat := 0
      else
        mtSiteSummaryWinton_Score.AsFloat := mtSiteSummaryWinton_YesDays.AsInteger / aTotalDays;

      Post;
      Next;
    end;
  end;
end;

procedure TfmMain.UpdSummaryData_IPH4(ASite: string; ADate: TDateTime; AScore: Extended);
var
  aYesField, aNoField: TIntegerField;
  //aScoreField: TFloatField;
  aYear, aMonth: Integer;
  aText: string;
begin
  if (ASite = SITE_DESC_Taipei_TC) then
  begin
    aYesField := mtSiteSummaryTaipei_YesDays;
    aNoField  := mtSiteSummaryTaipei_NoDays;
    //aScoreField := mtSiteSummaryTaipei_Score;
  end
  else if (ASite = SITE_DESC_Taoyuan_TC) then
  begin
    aYesField := mtSiteSummaryTaoyuan_YesDays;
    aNoField  := mtSiteSummaryTaoyuan_NoDays;
    //aScoreField := mtSiteSummaryTaoyuan_Score;
  end
  else if (ASite = SITE_DESC_Taichung_TC) then
  begin
    aYesField := mtSiteSummaryTaichung_YesDays;
    aNoField  := mtSiteSummaryTaichung_NoDays;
    //aScoreField := mtSiteSummaryTaichung_Score;
  end
  else if (ASite = SITE_DESC_Tainan_TC) then
  begin
    aYesField := mtSiteSummaryTainan_YesDays;
    aNoField  := mtSiteSummaryTainan_NoDays;
    //aScoreField := mtSiteSummaryTainan_Score;
  end
  else
  begin
    if not FAutoMode then
    begin
      aText := Format('FillData_IPH4() error, 未定義的營業處 %s', [ASite]);
      Application.MessageBox(PChar(aText), PChar(Application.Title), MB_OK + MB_ICONWARNING);
    end;
    JcLog.Write(aText);
    Exit;
  end;

  aYear := YearOf(ADate);
  aMonth := MonthOf(ADate);

  with mtSiteSummary do
  begin
    if not Active then Open;
    if not Locate('Year;Month', VarArrayOf([aYear, aMonth]), []) then
    begin
      Append;
      mtSiteSummaryYear.AsInteger := AYear;
      mtSiteSummaryMonth.AsInteger := AMonth;
    end
    else
      Edit;

    if (AScore >= 79.99) then
      aYesField.AsInteger := aYesField.AsInteger + 1
    else
      aNoField.AsInteger := aNoField.AsInteger + 1;

    Post;
  end;
end;

procedure TfmMain.ErrorMsg(AMsg: string);
begin
  if not FAutoMode then
    MessageBox(Handle, PChar(AMsg), PChar(Application.Title), MB_OK + MB_ICONSTOP);

  JcLog.Write(AMsg);
end;

procedure TfmMain.WriteXls_SiteData(ADataSet: TDataSet);
var
  i: Integer;
  aSheet: TXLSWorksheet;
  aDefFmt: TXLSDefaultFormat;
begin
  aSheet := XLSReadWriteII51.Add;
  // 凍結標題橫列
  aSheet.FreezePanes(0, 1);
  // 建立預設的儲存格格式
	with XLSReadWriteII51 do
  begin
    CmdFormat.BeginEdit(Nil);
    CmdFormat.Border.Style := cbsThin;
    CmdFormat.Border.Preset(cbspOutline);
    aDefFmt := CmdFormat.AddAsDefault('Format1');
    DefaultFormat := aDefFmt;
  end;
	// 將資料讀入工作表中
  with XLSDbRead51 do
  begin
    Sheet := aSheet.Index;
    Dataset := ADataSet;
    ExcludeFields.Clear;
    ExcludeFields.Add('RID');
    Read;
  end;
  // 調整格式
  with XLSReadWriteII51 do
  begin
    with aSheet do
    begin
      Name := 'ACD接聽率月報表_營業處_明細';
      //-----------------------------------------------------------------------
      Columns[1].CharWidth := 12;
      Columns[2].CharWidth := 12;
      //-----------------------------------------------------------------------
      for i := 1 to LastRow do
      begin
        Cell[1, i].NumberFormat := ShortDateFormat;
        Cell[2, i].NumberFormat := '0.00';
      end;

      AsString[0, 0] := '營業處';
      AsString[1, 0] := '統計日期';
      AsString[2, 0] := 'ACD接聽率';

      for i := 0 to 2 do
      begin
        if Cell[i, 0] <> nil then
        begin
          Cell[i, 0].FillPatternForeColorRGB := clSilver;
          Cell[i, 0].FontStyle := [xfsBold];
        end;
      end;
    end;
  end;
  JcLog.Write('已產生[營業處明細]工作表');
end;

procedure TfmMain.GetData_IPH3(ABeginTime, AEndTime: TDateTime);
begin
  JcLog.Write('讀取訓練師ACD接聽資料(WICSIPH3)');

  with FWICSIPH3 do
  begin
    if Active then Close;
    LocalUpdate := True;
    SQL.Clear;
		SQL.Add('SELECT IPH3001, SALE002, IPH3002,');
    //SQL.Add('DATEPART(YEAR, IPH3002) AS YEAR, DATEPART(MONTH, IPH3002) AS MONTH,');
    SQL.Add('IPH3003, IPH3007, IPH3006,');
    SQL.Add('DATEDIFF(month, :INDEX, IPH3002) AS DATA_INDEX');
    SQL.Add('FROM WICSIPH3 WITH(NOLOCK)');
    SQL.Add('LEFT JOIN WICSSALE WITH(NOLOCK) ON SALE001 = IPH3001');
    SQL.Add('WHERE (IPH3002 BETWEEN :IPH3002B AND :IPH3002E)');
    SQL.Add('AND (IPH3001 <> ''0-(其他)'')');
    SQL.Add('ORDER BY IPH3001, IPH3002');

    ParamByName('INDEX').AsDateTime := IncMonth(ABeginTime, -1);
    ParamByName('IPH3002B').AsDateTime := ABeginTime;
    ParamByName('IPH3002E').AsDateTime := AEndTime;

    try
      Open;
      JcLog.Write(Format('取得訓練師ACD接聽資料(WICSIPH3)，記錄數 = %d', [RecordCount]));
    except
      on E: Exception do
        JcLog.Write(Format('GetData_IPH3() failed, error = %s', [E.Message]));
    end;
  end;
end;

procedure TfmMain.UpdSummaryData_IPH3(AEmpNo, AEmpName: string; ADataNdx, APhoneCount, ARequired, AJobDays: Integer);
var
  aYesField, aNoField: TIntegerField;
begin
  aYesField := mtTeSummary.FindField(Format('YesDays_%d', [ADataNdx])) as TIntegerField;
  aNoField  := mtTeSummary.FindField(Format('NoDays_%d', [ADataNdx])) as TIntegerField;

  with mtTeSummary do
  begin
    if not Active then Open;
    if not Locate('EmpNo', AEmpNo, []) then
    begin
      Append;
      mtTeSummaryEmpNo.AsString := AEmpNo;
      mtTeSummaryEmpName.AsString := AEmpName;
    end
    else
      Edit;

    if (APhoneCount >= (ARequired*AJobDays/2)) then
      aYesField.AsInteger := aYesField.AsInteger + 1
    else
      aNoField.AsInteger := aNoField.AsInteger + 1;

    Post;
  end;
end;

procedure TfmMain.CalcSummaryData_IPH3;
var
  i, aTotalDays: Integer;
  aYesField, aNoField: TIntegerField;
  aScoreField: TFloatField;

  procedure _CalcSummaryData_IPH3(AYesDays, ANoDays: TIntegerField; aScore: TFloatField);
  var
    aTotalDays: Integer;
  begin
    aTotalDays := AYesDays.AsInteger + ANoDays.AsInteger;

    if (aTotalDays = 0) then
      aScore.AsFloat :=  0
    else
      aScore.AsFloat := AYesDays.AsInteger / aTotalDays;
  end;
begin
  with mtTeSummary do
  begin
    if (not Active) or IsEmpty then Exit;
    First;

    while not Eof do
    begin
      Edit;

      for i := 1 to 12 do
      begin
        aYesField := FindField(Format('YesDays_%d', [i])) as TIntegerField;
        aNoField  := FindField(Format('NoDays_%d', [i])) as TIntegerField;
        aScoreField := FindField(Format('Score_%d', [i])) as TFloatField;
        aTotalDays := aYesField.AsInteger + aNoField.AsInteger;

        if (aTotalDays = 0) then
          aScoreField.AsFloat :=  0
        else
          aScoreField.AsFloat := aYesField.AsInteger / aTotalDays;

        if (i = 1) then
        begin
          mtTeSummaryYesDays_Total.AsInteger := aYesField.AsInteger;
          mtTeSummaryNoDays_Total.AsInteger := aNoField.AsInteger;
        end
        else
        begin
          mtTeSummaryYesDays_Total.AsInteger := mtTeSummaryYesDays_Total.AsInteger + aYesField.AsInteger;
          mtTeSummaryNoDays_Total.AsInteger := mtTeSummaryNoDays_Total.AsInteger + aNoField.AsInteger;
        end
      end;

      aTotalDays := mtTeSummaryYesDays_Total.AsInteger + mtTeSummaryNoDays_Total.AsInteger;

      if (aTotalDays = 0) then
        mtTeSummaryScore_Total.AsFloat :=  0
      else
        mtTeSummaryScore_Total.AsFloat := mtTeSummaryYesDays_Total.AsInteger / aTotalDays;

      Post;
      Next;
    end;
  end;
end;

function TfmMain.PrepareData_IPH3(ABeginTime, AEndTime: TDateTime): Integer;
var
  aEmpNo, aEmpName, aHost, aText: string;
  aNdx, aPhoneCount, aARequired, aJobDays: Integer;
begin
  Result := 0;

  JcLog.Write(Format('開始整理訓練師ACD月報表資料, %s ~ %s',
    [FormatDateTime('yyyy/mm/dd', ABeginTime), FormatDateTime('yyyy/mm/dd', AEndTime)]));

  if mtTeSummary.Active then mtTeSummary.Close;

	with dmReport do
  begin
    aHost := GetBranchHostIp(BRANCH_NAME_Winton_TC);

    try
      if (ABeginTime > AEndTime) then
      begin
        aText := 'PrepareData_IPH3() error, 起時時間不可以大於結束時間';
        ErrorMsg(aText);
        Exit;
      end;

      if not SetConn_Tcrm(aHost) then
      begin
        aText := Format('PrepareData_IPH3() error, 無法連線到主機 %s', [aHost]);
        ErrorMsg(aText);
        Exit;
      end;
			// 取得營業處ACD接聽資料
      GetData_IPH3(ABeginTime, AEndTime);
      FixData_IPH3(FWICSIPH3);
      Result := FWICSIPH3.RecordCount;
      mtTeSummary.DisableControls;

      with FWICSIPH3 do
      begin
        First;

        while not Eof do
        begin
          aEmpNo  := FieldByName('IPH3001').AsString;
          aEmpName := FieldByName('SALE002').AsString;
          aNdx := FieldByName('DATA_INDEX').AsInteger;
          aPhoneCount := FieldByName('IPH3003').AsInteger;
          aARequired := FieldByName('IPH3006').AsInteger;
          aJobDays := FieldByName('IPH3007').AsInteger;
          UpdSummaryData_IPH3(aEmpNo, aEmpName, aNdx, aPhoneCount, aARequired, aJobDays);
          Next;
          Application.ProcessMessages;
        end;
      end;
      CalcSummaryData_IPH3;
      JcLog.Write(Format('已完成訓練師ACD月報表資料整理(%d)', [Result]));
		finally
      mtTeSummary.EnableControls;
    end;
  end;
end;

procedure TfmMain.WriteXls_TeData(ADataSet: TDataSet);
var
  i: Integer;
  aSheet: TXLSWorksheet;
  aDefFmt: TXLSDefaultFormat;
begin
  aSheet := XLSReadWriteII51.Add;
  // 凍結標題橫列
  aSheet.FreezePanes(0, 1);
  // 建立預設的儲存格格式
	with XLSReadWriteII51 do
  begin
    CmdFormat.BeginEdit(Nil);
    CmdFormat.Border.Style := cbsThin;
    CmdFormat.Border.Preset(cbspOutline);
    aDefFmt := CmdFormat.AddAsDefault('Format1');
    DefaultFormat := aDefFmt;
  end;
	// 將資料讀入工作表中
  with XLSDbRead51 do
  begin
    Sheet := aSheet.Index;
    Dataset := ADataSet;
    ExcludeFields.Clear;
    ExcludeFields.Add('DATA_INDEX');
    Read;
  end;
  // 調整格式
  with XLSReadWriteII51 do
  begin
    with aSheet do
    begin
      Name := 'ACD接聽率月報表_訓練師_明細';
      //-----------------------------------------------------------------------
      Columns[1].CharWidth := 12;
      Columns[2].CharWidth := 12;
      Columns[5].CharWidth := 12;
      //-----------------------------------------------------------------------
      for i := 1 to LastRow do
      begin
        Cell[2, i].NumberFormat := ShortDateFormat;
        AsFloat[4, i] := AsFloat[4, i] * 0.5;
      end;

      AsString[0, 0] := '代號';
      AsString[1, 0] := '姓名';
      AsString[2, 0] := '日期';
      AsString[3, 0] := 'ACD通數';
      AsString[4, 0] := '值機天數';
      AsString[5, 0] := 'ACD日標準';

      for i := 0 to 5 do
      begin
        if Cell[i, 0] <> nil then
        begin
          Cell[i, 0].FillPatternForeColorRGB := clSilver;
          Cell[i, 0].FontStyle := [xfsBold];
        end;
      end;
    end;
  end;
  JcLog.Write('已產生[訓練師明細]工作表');
end;

procedure TfmMain.FixData_IPH3(ADataSet: TDataSet);
begin
  with ADataSet do
  begin
    First;

    while not Eof do
    begin
      if FieldByName('IPH3007').AsInteger = 0 then
      begin
        Delete;
        Continue;
      end;

      Next;
    end;
  end;
end;

procedure TfmMain.WriteXls_TeSummary;
const
  RPT_TITLE = 'ACD接聽率月報表_訓練師';
var
  i, j: Integer;
  aSheet: TXLSWorksheet;
  aDefFmt: TXLSDefaultFormat;
  aDate: TDateTime;
begin
 	if (not mtTeSummary.Active) or mtTeSummary.IsEmpty then
  begin
    JcLog.Write(Format('沒有[%s]資料，不產生XLS工作表。', [RPT_TITLE]));
  	Exit;
  end;

  aSheet := XLSReadWriteII51.Add;
  // 凍結標題橫列
  //aSheet.FreezePanes(0, 1);
  // 建立預設的儲存格格式
	with XLSReadWriteII51 do
  begin
    CmdFormat.BeginEdit(Nil);
    CmdFormat.Border.Style := cbsThin;
    CmdFormat.Border.Preset(cbspOutline);
    aDefFmt := CmdFormat.AddAsDefault('Format1');
    DefaultFormat := aDefFmt;
  end;
	// 將資料讀入工作表中
  with XLSDbRead51 do
  begin
    Sheet := aSheet.Index;
    Dataset := mtTeSummary;
    Read;
  end;
  // 調整格式
  with XLSReadWriteII51 do
  begin
    with aSheet do
    begin
      Name := RPT_TITLE;
      //-----------------------------------------------------------------------
      for i := 1 to LastRow do
      begin
        for j := 0 to 12 do
          Cell[4+3*j, i].NumberFormat := '0%';
      end;
      // 設定欄位標題
      for j := 0 to 12 do
      begin
        AsString[2+3*j, 0] := '達標';
        AsString[3+3*j, 0] := '未達';
        AsString[4+3*j, 0] := '比例';
      end;
      // 塗上標條列底色
      for i := 0 to 41 do
      begin
        if Cell[i, 0] <> nil then
        begin
          Cell[i, 0].FillPatternForeColorRGB := clSilver;
          Cell[i, 0].FontStyle := [xfsBold];
        end;
      end;
      // 為不同月份塗上交錯底色
      CmdFormat.Mode := xcfmMerge;
      CmdFormat.BeginEdit(aSheet);
      CmdFormat.Fill.BackgroundColor.RGB := $00C6FFFF;

      for i := 0 to 6 do
        CmdFormat.Apply(2+3*2*i, 1, 4+3*2*i, LastRow);
      //-----------------------------------------------------------------------
      InsertRows(0, 1);

      for i := 0 to 11 do
      begin
        aDate := IncMonth(FCalcBeginTime, i);
        AsString[2+3*i, 0] := FormatDateTime('yyyy/mm', aDate);
        MergeCells(2+3*i, 0, 4+3*i, 0);
      end;

      AsString[39, 0] := '年度合計';
      MergeCells(38, 0, 40, 0);

      for i := 0 to 41 do
      begin
        if Cell[i, 0] <> nil then
        begin
          Cell[i, 0].FillPatternForeColorRGB := clSilver;
          Cell[i, 0].FontStyle := [xfsBold];
        end;
      end;
      // 凍結標題橫列
      aSheet.FreezePanes(2, 2);
      // 自動調整欄位最適寬度
      AutoWidthCols(0, 40);
    end;
  end;
  JcLog.Write(Format('已產生[%s]工作表', [RPT_TITLE]));
end;

procedure TfmMain.InitSysParams;
var
  aDate: TDateTime;
begin
  JcVersionInfo1.FileName := Application.ExeName;
	Self.Caption := Format('%s - V %s', [Self.Caption, JcVersionInfo1.FileVersion]);

  JcLog.FileName := JcLog.GetExeNameTimeSerialStr('Log') + '.log';
  JcLog.Active := True;
  // 計算預設的統計起訖日期
  aDate := EncodeDate(YearOf(Date), MonthOf(Date), 1);
  FCalcBeginTime := IncMonth(aDate, -12);
  FCalcEndTime := IncDay(aDate, -1);

  DateTimePickerBegin.Date := FCalcBeginTime;
  DateTimePickerEnd.Date := FCalcEndTime;

  FReportName := Format('ACD接聽率月報表_%s_%s',
    [FormatDateTime('yyyymm', FCalcBeginTime), FormatDateTime('yyyymm', FCalcEndTime)]);

  FWICSIPH3 := dmReport.GetQuery_WintonTcrm;
  FWICSIPH4 := dmReport.GetQuery_WintonTcrm;
end;

procedure TfmMain.Log(AMsg: string);
begin
  JcLog.Write(AMsg);
end;

procedure TfmMain.LogLine;
begin
  JcLog.Line();
end;

procedure TfmMain.CheckExeMode;
var
  i: Integer;
  aText: string;
begin
  FAutoMode := False;
  FDebugMode := False;
  FNoMail := False;
  // 檢查啟動模式
  for i := 1 to ParamCount do
  begin
    aText := UpperCase(ParamStr(i));

    if (aText = '/AUTO') then
      FAutoMode := True
    else if (aText = '/DEBUG') then
      FDebugMode := True
    else if (aText = '/NOMAIL') then
      FNoMail := True;
  end;
end;

procedure TfmMain.btnDebugClick(Sender: TObject);
begin
  //dmPhoneAnalysis.Exec(2017, 8);
  dmAcdSvcFailedAnalysis.Exec(2017, 7);
end;

procedure TfmMain.CallExcelToSaveAs(AFileName: string);
var
	aExcelApp: Variant;
begin
	aExcelApp := CreateOleObject('Excel.Application');

  try
    aExcelApp.WorkBooks.Open(AFileName);
    aExcelApp.Application.DisplayAlerts := False;
    aExcelApp.ActiveWorkbook.SaveAs(AFileName);
  finally
  	aExcelApp.Quit;
  end;
end;

end.

