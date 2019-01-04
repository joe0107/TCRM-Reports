unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs, DB, DateUtils, dxmdaset,
  XLSSheetData5, XLSReadWriteII5, XLSDbRead5, Xc12DataStyleSheet5, XLSCmdFormat5, Xc12Utils5, XLSComment5, XLSDrawing5,
  IdEMailAddress, IdMessage, IdBaseComponent, IdComponent, Uni, ShellAPI, IdAttachmentFile, Grids, DBGrids, StdCtrls,
  ComCtrls, ExtCtrls, NetDrive, JcVersionInfo, Math, JcNumUtils, JcLog, ComObj, CodeSiteLogging, IdGlobal;

type
  TfmMain = class(TForm)
    XLSDbRead51: TXLSDbRead5;
    XLSReadWriteII51: TXLSReadWriteII5;
    mdReport: TdxMemData;
    NetDrive1: TNetDrive;
    JcVersionInfo1: TJcVersionInfo;
    JcLog: TJcLog;
    ListBox1: TListBox;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    Label2: TLabel;
    DateTimePicker2: TDateTimePicker;
    lbl1: TLabel;
    DateTimePicker3: TDateTimePicker;
    CheckBox_RecalcWICSIPH2: TCheckBox;
    btn_ACD1: TButton;
    TabSheet2: TTabSheet;
    Label1: TLabel;
    DateTimePicker1: TDateTimePicker;
    ComboBox1: TComboBox;
    btnRunReport: TButton;
    TabSheet3: TTabSheet;
    Label3: TLabel;
    DateTimePicker4: TDateTimePicker;
    btnSitePhoneSummary: TButton;
    TabSheet4: TTabSheet;
    Label4: TLabel;
    DateTimePicker5: TDateTimePicker;
    btnTePhoneSummary: TButton;
    TabSheet5: TTabSheet;
    Label5: TLabel;
    DateTimePicker6: TDateTimePicker;
    btnAcdSummary: TButton;
    procedure FormCreate(Sender: TObject);
    procedure btnRunReportClick(Sender: TObject);
    procedure mdReportFilterRecord(DataSet: TDataSet; var Accept: Boolean);
    procedure btn_ACD1Click(Sender: TObject);
    procedure ListBox1DblClick(Sender: TObject);
    procedure btnSitePhoneSummaryClick(Sender: TObject);
    procedure btnTePhoneSummaryClick(Sender: TObject);
    procedure btnAcdSummaryClick(Sender: TObject);
  private
    FAutoMode: Boolean;
    FDebugMode: Boolean;
    FMailMode: Boolean;
    FCodeSiteMode: Boolean;
    // 檢查執行模式
    procedure CheckExeMode;
  private
  	// 報表統計基準日
    FCalcDate: TDateTime;
    FSiteScore, FSiteTimeOutScore, FSiteAcdDays: Extended;

    FAcdAnsCount: Integer;			// 直接總通數
    FAcdTeCount: Extended; 			// 值機人數
    FAvgAnsCount: Extended;			// 每人直接平均通數
    FAcdTotalCount: Extended;   // 派送總通數
    FAvgGoodAnsCount: Extended;	// 可達標的每人平均直接通數
    // 產生報表資料
    function	PrepareData(ASite: string): Boolean;
		// 取得ACD接聽資料
    function  GetData_IPH3: TUniQuery;
    // 取得ACD值機的排班資料
    function  GetData_CHEM: TUniQuery;
    // 取得回電計數資料
    function  GetData_RPHE: TUniQuery;
    // 計算營業處的總值機人力
    function  CalcSiteAcdDays(ADataSet: TDataSet): Extended;
    // 檢查營業處的直接率是否達到 80%
    function  CheckSiteAcdScore(AData: TDataSet): Boolean;
    // 依據值機日數計算個人當日的日標準結果
    procedure CalcPassData(AIPH3, ACHEM, ARPHE: TDataSet);
    // 計算可達標的平均數據
    procedure CalcData(ADataSet: TDataSet);
    // 取得ACD派送總通數
    function  ACD_GetSiteCount(ASite: string; ABeginTime, AEndTime: TDateTime; AAnswerOnly: Boolean): Integer;
    // 從 WICSIPH4 取得以統計完成的資訊 Added by Joe 2017/10/24 14:38:23
    procedure GetAcdSummaryInfo(ASite: string; var AcdTotalCount, SiteScore, SiteTimeOutScore, SiteAcdDays: Extended);
	private
    FXlsFileName: string;
    FNetDrvRootDir, FNetDrvUser, FNetDrvPwd: string;

    procedure PrepareXLS(ASite: string);
    // 寫入來源資料清單
    procedure WriteDataToXls;
    // 將檔案複製到營業處
    procedure CopyXlsToSite(ASite: string);
	protected
    procedure SendMail(ASite: string);
    procedure MakeCCList(AEmailAddrList: TIdEmailAddressList);
    function  MakeNotifyMessage(ASite: string): TIdMessage;
    // ------------------------------------------------------------------------
    procedure SendAdminMail;
    function  MakeAdminNotifyMessage: TIdMessage;
  public
    procedure DoAutoJobs;
		procedure Exec;
    procedure Log(AMsg: string);
    procedure LogLine(CH: Char = '-');
    procedure CallExcelToSaveAs(AFileName: string);
    // Added by Joe 2017/11/01 15:55:59
    procedure InitCodeSite;
  public
    property AutoMode: Boolean read FAutoMode;
    property DebugMode: Boolean read FDebugMode;
    property MailMode: Boolean read FMailMode;
    property CodeSiteMode: Boolean read FCodeSiteMode;
  end;

var
  fmMain: TfmMain;

implementation

uses ReportData, TcrmConstants, AcdSummary, SitePhoneSummary, TePhoneSummary;

{$R *.dfm}

procedure TfmMain.WriteDataToXls;
var
  i: Integer;
  aSheet: TXLSWorksheet;
  aDefFmt: TXLSDefaultFormat;
begin
  aSheet := XLSReadWriteII51.Sheets[0];
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
    Dataset := mdReport;
    ExcludeFields.Clear;
    ExcludeFields.Add('RecId');
    ExcludeFields.Add('IPH3001');
    ExcludeFields.Add('IPH3004');
    ExcludeFields.Add('IPH3005');
    ExcludeFields.Add('PASS');
    Read;
  end;
  // 調整格式
  with XLSReadWriteII51 do
  begin
    with aSheet do
    begin
      Name := 'ACD接聽率';
      //-----------------------------------------------------------------------
      Columns[0].CharWidth := 15; //訓練組別
      Columns[2].CharWidth := 10; //日標準
      Columns[3].CharWidth := 10; //值機天數
      Columns[4].CharWidth := 10; //ACD直接
      Columns[5].CharWidth := 12; //總回電數
      Columns[6].CharWidth := 50; //原因
      Columns[7].CharWidth := 50; //改善作法
      //-----------------------------------------------------------------------
      AsString[0, 0] := '訓練組別';
      AsString[1, 0] := '姓名';
      AsString[2, 0] := '日標準';
      AsString[3, 0] := '值機天數';
      AsString[4, 0] := '直接通數';
      AsString[5, 0] := '總回電數';
      AsString[6, 0] := '原因';
      AsString[7, 0] := '改善作法';
      //-----------------------------------------------------------------------
      for i := 0 to 7 do
        Cell[i, 0].FillPatternForeColorRGB := clSilver;
    end;

    // 將未達標人員以紅字標示
    for i := 1 to aSheet.LastRow do
    begin
      // 如果沒有[ACD日標準]通數,不要計算
      if Trim(aSheet.AsString[2, i]) <> '' then // Added by Joe 2017/09/28 16:02:55
      begin
        try
          if (aSheet.AsFloat[2, i] * aSheet.AsFloat[3, i]) > aSheet.AsFloat[4, i] then
          begin
            CmdFormat.BeginEdit(aSheet);
            CmdFormat.Font.Color.RGB := $FF0000;
            CmdFormat.ApplyRows(i, i);
          end;
        except
          Log(Format('標示未達標人員發生錯誤 = %s', [aSheet.AsString[1, i]]));
        end;
      end;
    end;
  end;
end;

procedure TfmMain.FormCreate(Sender: TObject);
begin
  JcVersionInfo1.FileName := Application.ExeName;
	Self.Caption := Format('%s - V %s', [Self.Caption, JcVersionInfo1.FileVersion]);
  JcLog.FileName := JcLog.GetExeNameTimeSerialStr('Log') + '.log';
  JcLog.Active := True;
  // 計算前一天的ACD接聽資料
  FCalcDate := IncDay(Date,  -1);
  DateTimePicker1.Date := FCalcDate;
  DateTimePicker2.Date := FCalcDate;
  DateTimePicker3.Date := FCalcDate;
//  DateTimePicker2.Date := EncodeDate(2017, 11, 7);
//  DateTimePicker3.Date := EncodeDate(2017, 11, 7);
  DateTimePicker4.Date := FCalcDate;
  DateTimePicker5.Date := FCalcDate;
  DateTimePicker6.Date := FCalcDate;
  // 檢查是否為自動執行模式
  CheckExeMode;
  InitCodeSite; // Added by Joe 2017/11/01 15:57:09
  
  if FAutoMode then
  begin
    try
      DoAutoJobs;
    finally
      Application.Terminate;
    end;
  end;
end;

function TfmMain.MakeNotifyMessage(ASite: string): TIdMessage;
var
  aText, aDayOfWeek: string;
  aCount: Extended;
begin
  Result := TIdMessage.Create(Self);
  dmReport.Init_IdMessage(Result);    //設定郵件屬性 Added by Joe Lee 2017/11/20 09:47:24
  aDayOfWeek := GetChineseNumStr(DayOfWeek(FCalcDate) - 1);
  if (aDayOfWeek = '零') then aDayOfWeek := '日';

  with Result do
  begin
    //填入收件者
    if FDebugMode then
    begin
      Recipients.Add.Address := 'joelee@winton.com.tw';
    end
    else
    begin
      aText := dmReport.GetEmail_TE_Admin(ASite);
      Recipients.Add.Address := aText;
      //填入副本
      aText := dmReport.GetEmail_Site_Admin(ASite);
      CCList.Add.Address := aText;
    end;

    MakeCCList(CCList);
    //填入郵件表頭資訊
    Subject := Format('電話效率未達通知_%s(%s)_%s', [FormatDateTime('yyyymmdd', FCalcDate), aDayOfWeek, ASite]);
    //填入郵件內容
	  aCount := FAvgGoodAnsCount - FAvgAnsCount;
    aText := Format('[%s] %s(%s)', 	[ASite, FormatDateTime('yyyy/mm/dd', FCalcDate), aDayOfWeek]);
    aText := aText + CHR(13) + Format('ACD派送總數 %.0f，電話直接率 %.2f%%，逾時率 %.1f%%，值機人力 %.1f，未達標準。',
    	[FAcdTotalCount, FSiteScore, FSiteTimeOutScore*100, FSiteAcdDays]);

		aText := aText + CHR(13) + Format('要達成直接率 80%% 的標準，平均每位值機人員要再多接 %d 通ACD派送電話。', [Ceil(aCount)]);
    aText := aText + CHR(13) + '個人未達日標準清單請參閱附件中的紅字標示資料。';
    Body.Text :=aText;
  end;

  if FileExists(FXlsFileName) then
  	TIdAttachmentFile.Create(Result.MessageParts, FXlsFileName);
end;

procedure TfmMain.PrepareXLS(ASite: string);
var
  aReportFolder: string;
begin
 	if (not mdReport.Active) or mdReport.IsEmpty then
  begin
    JcLog.Write(Format('沒有報表資料，不產生XLS檔案。', [ASite]));
  	Exit;
  end;

  aReportFolder := ExtractFilePath(Application.ExeName) + 'ReportStock';
  ForceDirectories(aReportFolder);
  FXlsFileName := aReportFolder + Format('\電話效率未達報告_%s_%s.xlsx', [FormatDateTime('yyyymmdd', FCalcDate), ASite]);

  with XLSReadWriteII51 do
  begin
    Clear;
    Filename := FXlsFileName;
  end;

   WriteDataToXls;
   XLSReadWriteII51.Write;
   JcLog.Write(Format('產生XLS檔案 = %s', [FXlsFileName]));

  if FileExists(FXlsFileName) then
  	CopyXlsToSite(ASite);

//  if not FAutoMode then
//  	ShellExecute(0, 'open', PChar(FXlsFileName), nil, nil, SW_SHOWMAXIMIZED);
end;

procedure TfmMain.Exec;
begin
  JcLog.Write('開始執行[ACD電話效率未達報告]');

  try
    if PrepareData(SITE_NAME_Taipei_TC) then
    begin
      PrepareXLS(SITE_NAME_Taipei_TC);
      SendMail(SITE_NAME_Taipei_TC);
    end;
  finally

  end;

  try
    if PrepareData(SITE_NAME_Taoyuan_TC) then
    begin
      PrepareXLS(SITE_NAME_Taoyuan_TC);
      SendMail(SITE_NAME_Taoyuan_TC);
    end;
  finally

  end;

  try
    if PrepareData(SITE_NAME_Taichung_TC) then
    begin
      PrepareXLS(SITE_NAME_Taichung_TC);
      SendMail(SITE_NAME_Taichung_TC);
    end;
  finally

  end;

  try
    if PrepareData(SITE_NAME_Tainan_TC) then
    begin
      PrepareXLS(SITE_NAME_Tainan_TC);
      SendMail(SITE_NAME_Tainan_TC);
    end;
  finally

  end;

  JcLog.Line('=');
end;

function TfmMain.GetData_IPH3: TUniQuery;
begin
  JcLog.Write('讀取ACD接聽資料');
  Result := dmReport.GetQuery_Tcrm;

  with Result do
  begin
    LocalUpdate := True;

		SQL.Add('SELECT C.DEPT002, A.IPH3001, B.SALE002, B.SALE024, 1.0 AS ACD_DAY, A.IPH3003, 0 AS CallOutCount');
		SQL.Add(', A.IPH3004, A.IPH3005, ''Y'' AS PASS');
    SQL.Add(', '''' AS REASON, '''' AS IMPROVE');
    SQL.Add('FROM WICSIPH3 A WITH(NOLOCK)');
    SQL.Add('LEFT JOIN WICSSALE B WITH(NOLOCK) ON SALE001 = IPH3001');
    SQL.Add('LEFT JOIN WICSDEPT C WITH(NOLOCK) ON SALE003 = DEPT001');
    SQL.Add('WHERE (IPH3002 = :IPH3002)');
    SQL.Add('ORDER BY DEPT001, IPH3001');
    ParamByName('IPH3002').AsDateTime := FCalcDate;

    try
      Open;
      JcLog.Write(Format('取得ACD接聽資料，記錄數 = %d', [RecordCount]));
    except
      on E: Exception do
        JcLog.Write(Format('GetData_IPH3() failed, error = %s', [E.Message]));
    end;
  end;
end;

function TfmMain.CheckSiteAcdScore(AData: TDataSet): Boolean;
begin
  Result := True;

  if (AData = nil) or (not AData.Active) or AData.IsEmpty then
  	Exit;

  with AData do
  begin
    First;
    FSiteScore := FieldByName('IPH3004').AsFloat;
    // Added by Joe 2017/04/14 13:53:46
    if (FSiteScore > 0.799) and (FSiteScore < 0.8) then
      FSiteScore := 0.8;
    //-------------------------------------------------------------------------
    FSiteTimeOutScore := FieldByName('IPH3005').AsFloat;

    if FSiteScore < 0.8 then
    	Result := False;
  end;
end;

function TfmMain.GetData_CHEM: TUniQuery;
begin
  JcLog.Write('讀取ACD值機的排班資料');
  Result := dmReport.GetQuery_Tcrm;

  with Result do
  begin
		SQL.Add('SELECT CHEM004, SALE002, COUNT(*) AS ACD_DAY');
    SQL.Add('FROM WICSCHEM A WITH(NOLOCK)');
    SQL.Add('LEFT JOIN WICSSALE S WITH(NOLOCK) ON SALE001 = CHEM004');
    SQL.Add('LEFT JOIN WICSSTM2 T WITH(NOLOCK) ON T.STM2001 = CHEM005 AND T.STM2002 = CHEM006');
    SQL.Add('WHERE (CHEM001 = :CHEM001)');
    SQL.Add('AND (STM2004 = ''Y'')');
    SQL.Add('AND (SALE003 IN(''021'',''022'',''023'',''026'',''028'',''052'',''062'',''075'',''076'',''082'',''092''))');
    SQL.Add('GROUP BY CHEM004, SALE002');
    ParamByName('CHEM001').AsDateTime := FCalcDate;

    try
      Open;
      JcLog.Write(Format('取得ACD值機的排班資料，記錄數 = %d', [RecordCount]));
    except
      on E: Exception do
        JcLog.Write(Format('GetData_CHEM() failed, error = %s', [E.Message]));
    end;
  end;
end;

procedure TfmMain.CalcPassData(AIPH3, ACHEM, ARPHE: TDataSet);
var
  aAcdDay: Extended;
  aCallOutCount: Integer;
begin
  JcLog.Write('依據值機日數計算個人當日的日標準結果');

	with AIPH3 do
  begin
    Filtered := False;
    First;

    while not Eof do
    begin
      // 如果找到ACD值機的排班資料，換算成標準日額比例
      if ACHEM.Locate('CHEM004', FieldByName('IPH3001').AsString, []) then
        aAcdDay := 0.5 * ACHEM.FieldByName('ACD_DAY').AsInteger
      else
      	aAcdDay := 0;

      if ARPHE.Locate('RPHE003', FieldByName('IPH3001').AsString, []) then
        aCallOutCount := ARPHE.FieldByName('_COUNT_').AsInteger
      else
      	aCallOutCount := 0;

      Edit;
      FieldByName('ACD_DAY').AsFloat := aAcdDay;
      FieldByName('CallOutCount').AsInteger := aCallOutCount;

      if (FieldByName('IPH3003').AsFloat < (aAcdDay * FieldByName('SALE024').AsFloat)) and (aAcdDay > 0) then
        FieldByName('PASS').AsString := 'N';

      Post;
      Next;
    end;
  end;
end;

procedure TfmMain.btnAcdSummaryClick(Sender: TObject);
begin
  TdmAcdSummary.Exec(DateTimePicker6.Date);
end;

procedure TfmMain.btnRunReportClick(Sender: TObject);
var
  aSite: string;
begin
  FCalcDate := DateOf(DateTimePicker1.DateTime);
  aSite := ComboBox1.Text;
  
  if (aSite = SITE_NAME_Taipei_TC) then
  begin
    if PrepareData(SITE_NAME_Taipei_TC) then
    begin
      PrepareXLS(SITE_NAME_Taipei_TC);
		  SendMail(SITE_NAME_Taipei_TC);
    end;
  end
  else if (aSite = SITE_NAME_Taoyuan_TC) then
  begin
    if PrepareData(SITE_NAME_Taoyuan_TC) then
    begin
      PrepareXLS(SITE_NAME_Taoyuan_TC);
		  SendMail(SITE_NAME_Taoyuan_TC);
    end;
  end
  else if (aSite = SITE_NAME_Taichung_TC) then
  begin
    if PrepareData(SITE_NAME_Taichung_TC) then
    begin
      PrepareXLS(SITE_NAME_Taichung_TC);
      SendMail(SITE_NAME_Taichung_TC);
    end;
  end
  else if (aSite = SITE_NAME_Tainan_TC) then
  begin
    if PrepareData(SITE_NAME_Tainan_TC) then
    begin
      PrepareXLS(SITE_NAME_Tainan_TC);
      SendMail(SITE_NAME_Tainan_TC);
    end;
  end;
  JcLog.Line;
  ShowMessage(Format('Score = %.2f%%', [FSiteScore]));
  //ShowMessage(Format('Score = %.1f%%', [FSiteScore*100]));
end;

function TfmMain.PrepareData(ASite: string): Boolean;
var
  aHost, aText: string;
  aWICSIPH3, aWICSCHEM, aWICSRPHE: TUniQuery;
begin
  Result := False;
  JcLog.Line('-');
  JcLog.Write(Format('開始整理報表資料[%s]', [ASite]));
  // Added by Joe 2017/10/24 14:57:11
  // 直接從 10.1.1.212 WICSIPH4 取得統計資訊
  GetAcdSummaryInfo(ASite, FAcdTotalCount, FSiteScore, FSiteTimeOutScore, FSiteAcdDays);
  // 如果ACD接聽率超過80%，不需要繼續處理
  if (FSiteScore >= 80) then
  begin
    JcLog.Write(Format('ACD接聽率 %.2f%% 超過80%%，結束報表[%s]', [FSiteScore, ASite]));
    Exit;
  end
  else
    JcLog.Write(Format('ACD接聽率 %.2f%% 未達80%%，開始產生報表[%s]', [FSiteScore, ASite]));
  //---------------------------------------------------------------------------
  aWICSIPH3 := nil;
  aWICSCHEM := nil;
  aWICSRPHE := nil;
  if mdReport.Active then mdReport.Close;

	with dmReport do
  begin
    //aHost := GetBranchHostIp(ASite);
    aHost := GetSiteIp(ASite);

    try
      if not SetConn_Tcrm(aHost) then
      begin
        if not FAutoMode then
        begin
          aText := Format('無法連線到主機 %s', [aHost]);
          Application.MessageBox(PChar(aText), PChar(Application.Title), MB_OK + MB_ICONWARNING);
        end;
        JcLog.Write(Format('Error, failed to connect %s', [aHost]));
        Exit;
      end;
			// 取得ACD接聽資料
      aWICSIPH3 := GetData_IPH3;
      // 如果沒有ACD接聽資料，不需要繼續處理
      if (aWICSIPH3.IsEmpty) then
      begin
        JcLog.Write(Format('沒有ACD接聽資料，結束報表[%s]', [ASite]));
        Exit;
      end;
      (**
      if CheckSiteAcdScore(aWICSIPH3) then
      begin
        JcLog.Write(Format('ACD接聽率 %.1f%% 超過80%%，結束報表[%s]', [FSiteScore*100, ABranch]));
      	Exit;
      end
      else
        JcLog.Write(Format('ACD接聽率 %.1f%% 未達80%%，開始產生報表[%s]', [FSiteScore*100, ABranch]));
      **)
      // 取得ACD值機的排班資料
			aWICSCHEM := GetData_CHEM;
    	// 取得回電計數資料
      aWICSRPHE := GetData_RPHE;
    	// 計算營業處的總值機人力
    	//FSiteAcdDays := CalcSiteAcdDays(aWICSCHEM);
      // 依據值機日數計算個人當日的日標準結果
      mdReport.LoadFromDataSet(aWICSIPH3);
			CalcPassData(mdReport, aWICSCHEM, aWICSRPHE);
			// 過濾未值機的資料
      JcLog.Write(Format('過濾未值機的資料[%s]', [ASite]));
    	mdReport.Filtered := True;
      // 取得ACD派送總數 Added by Joe 2017/03/29 11:32:00
      //FAcdTotalCount := ACD_GetSiteCount(ABranch, FCalcDate, IncMilliSecond(FCalcDate+1, -1), False);
			// 計算要達標的各項分析數據
      CalcData(mdReport);

      if FAcdAnsCount > 0 then
        Result := True;

      JcLog.Write(Format('已完成報表資料整理[%s]', [ASite]));
		finally
			CloseAndFree(aWICSIPH3);
      CloseAndFree(aWICSCHEM);
      CloseAndFree(aWICSRPHE);
    end;
  end;
end;

procedure TfmMain.mdReportFilterRecord(DataSet: TDataSet; var Accept: Boolean);
begin
	//Accept := (DataSet.FieldByName('PASS').AsString = 'N');
  Accept := (DataSet.FieldByName('ACD_DAY').AsFloat > 0);
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

procedure TfmMain.SendMail(ASite: string);
var
  aMsg: TIdMessage;
begin
  if FMailMode then
  begin
    aMsg := MakeNotifyMessage(ASite);
    dmReport.SendNotofyMail_SSL(aMsg);
    JcLog.Write(Format('已透過郵件傳送報表[%s]', [ASite]));
  end;
end;

procedure TfmMain.CheckExeMode;
var
  i: Integer;
  aText: string;
begin
  FAutoMode := False;
  FDebugMode := False;
  FMailMode := False;
  // 檢查啟動模式
  for i := 1 to ParamCount do
  begin
    aText := UpperCase(ParamStr(i));

    if (aText = '/AUTO') then
      FAutoMode := True
    else if (aText = '/DEBUG') then
      FDebugMode := True
    else if (aText = '/MAIL') then
      FMailMode := True
    else if (aText = '/CODESITE') then
      FCodeSiteMode := True;
  end;
end;

function TfmMain.GetData_RPHE: TUniQuery;
begin
  JcLog.Write('讀取回電計數資料');
  Result := dmReport.GetQuery_Tcrm;

  with Result do
  begin
		SQL.Add('SELECT RPHE003, COUNT(*) AS _COUNT_');
    SQL.Add('FROM WICSRPHE WITH(NOLOCK)');
    SQL.Add('WHERE (RPHE005 >= :RPHE005B) AND (RPHE005 < :RPHE005E)');
    SQL.Add('GROUP BY RPHE003');
    ParamByName('RPHE005B').AsDateTime := FCalcDate;
    ParamByName('RPHE005E').AsDateTime := IncDay(FCalcDate);

    try
      Open;
      JcLog.Write(Format('取得回電計數資料，記錄數 = %d', [RecordCount]));
    except
      on E: Exception do
        JcLog.Write(Format('GetData_RPHE() failed, error = %s', [E.Message]));
    end;
  end;
end;

function TfmMain.CalcSiteAcdDays(ADataSet: TDataSet): Extended;
begin
  JcLog.Write('計算營業處的總值機人力');
  Result := 0;

	with ADataSet do
  begin
    while not Eof do
    begin
  		Result := Result + FieldByName('ACD_DAY').AsInteger;
      Next;
    end;
  end;

  Result := Result * 0.5;
  JcLog.Write(Format('營業處的總值機人力 = %.1f', [Result]));
end;

procedure TfmMain.CopyXlsToSite(ASite: string);
var
  aDstFileName: string;
begin
  JcLog.Write(Format('複製XLS檔案到營業處[%s]', [ASite]));

  if (aSite = SITE_NAME_Taipei_TC) then
  	FNetDrvRootDir := '\\wtp4\Winnan\訓練專區\台北電話效率未達分析表'
  else if (aSite = SITE_NAME_Taoyuan_TC) then
  	FNetDrvRootDir := '\\10.3.1.45\public\wty2\訓練部\電話效率未達分析表'
  else if (aSite = SITE_NAME_Taichung_TC) then
  	FNetDrvRootDir := '\\10.5.1.4\TE\TCRM統計資料\回電統計\中區電話效率未達分析表'
  else if (aSite = SITE_NAME_Tainan_TC) then
    	FNetDrvRootDir := '\\10.6.1.66\未達分析表';

	FNetDrvUser := 'winton\rdrepl';
  FNetDrvPwd := 'Wint0n2k';

  try
    aDstFileName := Format('%s\%s', [FNetDrvRootDir, ExtractFileName(FXlsFileName)]);
    JcLog.Write(Format('目標XLS檔案 = %s', [aDstFileName]));
    NetDrive1.Connect(FNetDrvRootDir, FNetDrvUser, FNetDrvPwd);
    JcLog.Write('連線網路磁碟機');

    if NetDrive1.Connected or (NetDrive1.ErrorCode = 1219) then
    begin
      if FMailMode then
      begin
        CopyFile(PChar(FXlsFileName), PChar(aDstFileName), False);
        JcLog.Write('已複製XLS檔案到營業處');
      end
    end
    else
      JcLog.Write('注意!!無法複製XLS檔案到營業處');
  finally
  	NetDrive1.Disconnect;
  end;
end;

procedure TfmMain.CalcData(ADataSet: TDataSet);
var
  aAcdCount: Extended;	  // 有效派送總通數
begin
  JcLog.Write('計算要達標的各項分析數據');
  FAcdAnsCount := 0;			// 直接總通數
  FAcdTeCount  := 0;			// 值機人數
  FAvgAnsCount := 0;			// 每人直接平均通數
  FAvgGoodAnsCount := 0;	// 可達標的每人平均直接通數

	with ADataSet do
  begin
    First;
    while not Eof do
    begin
      FAcdTeCount := FAcdTeCount + FieldByName('ACD_DAY').AsFloat;
    	FAcdAnsCount := FAcdAnsCount + FieldByName('IPH3003').AsInteger;
    	Next;
    end;
    // 每人直接平均通數
    FAvgAnsCount := JcDivide(FAcdAnsCount, FAcdTeCount);
    JcLog.Write(Format('每人直接平均通數 = %.1f', [FAvgAnsCount]));
    // 派送總通數
    aAcdCount := JcDivide(FAcdAnsCount, FieldByName('IPH3004').AsFloat);
    JcLog.Write(Format('派送總通數 = %.1f', [aAcdCount]));
    // 可達標的每人平均直接通數
    FAvgGoodAnsCount := JcDivide((aAcdCount * 0.8), FAcdTeCount);
    JcLog.Write(Format('可達標的每人平均直接通數 = %.1f', [FAvgGoodAnsCount]));
  end;
end;

function TfmMain.ACD_GetSiteCount(ASite: string; ABeginTime, AEndTime: TDateTime; AAnswerOnly: Boolean): Integer;
var
  aData: TUniQuery;
begin
  JcLog.Write(Format('讀取ACD派送總數[%s]', [ASite]));

  if not Assigned(dmReport) then
  begin
    if not FAutoMode then
      raise Exception.Create(ERR_DMMSSQL_MISSING)
    else
      JcLog.Write(ERR_DMMSSQL_MISSING);
  end;

  dmReport.SetConn_TeleContact(ASite);
  aData := dmReport.GetQuery_TeleContact;

  try
    with aData do
    begin
      if Active then Close;
      SQL.Clear;
      SQL.Add('SELECT DATEPART(year, ITIME) AS _YEAR_,');
      SQL.Add('DATEPART(month, ITIME) AS _MONTH_, DATEPART(day, ITIME) AS _DAY_,');
      SQL.Add('COUNT(*) AS PHONE_COUNT FROM CALL_LOG_AGG WITH(NOLOCK)');
      SQL.Add('WHERE ((CTIME >= :CTIME1) AND (CTIME <= :CTIME2))');
      SQL.Add('AND (RTRIM(PID) <> '''')');
      SQL.Add('AND ((PID <> ''2007'') AND (AID NOT LIKE ''12%''))');  // Added by Joe 2015/07/24 16:33:27
      //-------------------------------------------------------------------------
      if AAnswerOnly then
        SQL.Add('AND (SCODE = 1 )');

      SQL.Add('GROUP BY DATEPART(year, ITIME), DATEPART(month, ITIME), DATEPART(day, ITIME)');

      Params.ParamValues['CTIME1'] := ABeginTime;
      Params.ParamValues['CTIME2'] := AEndTime;
      //-------------------------------------------------------------------------
      try
        Open;
        JcLog.Write(Format('取得ACD派送總數，記錄數 = %d', [RecordCount]));
      except
        on E: Exception do
          JcLog.Write(Format('ACD_GetSiteCount(%s) failed, error = %s', [ASite, E.Message]));
      end;

      First;
      Result := FieldByName('PHONE_COUNT').AsInteger;
      JcLog.Write(Format('取得ACD派送總數 = %d', [Result]));
    end;
  finally
    aData.Close;
    aData.Free;
  end;
end;

procedure TfmMain.SendAdminMail;
var
  aMsg: TIdMessage;
begin
  aMsg := MakeAdminNotifyMessage;
  dmReport.SendNotofyMail_SSL(aMsg);
end;

function TfmMain.MakeAdminNotifyMessage: TIdMessage;
var
  aDayOfWeek: string;
begin
  Result := TIdMessage.Create(Self);
  aDayOfWeek := GetChineseNumStr(DayOfWeek(FCalcDate) - 1);
  if (aDayOfWeek = '零') then aDayOfWeek := '日';  

  with Result do
  begin
    //填入收件者
    Recipients.EMailAddresses := dmReport.AdminEmail;
    //填入郵件表頭資訊
    Subject := Format('ACD日報表_%s(%s)_執行報告', [FormatDateTime('yyyymmdd', FCalcDate), aDayOfWeek]);
    //寄件人地址
    From.Address := 'rdrepl@winton.com.tw';
    ContentType := 'multipart/mixed';
    //填入郵件內容
    Body.Text := '';
  end;

  if FileExists(JcLog.FileName) then
  	TIdAttachmentFile.Create(Result.MessageParts, JcLog.FileName);
end;

procedure TfmMain.btn_ACD1Click(Sender: TObject);
var
  aBegDate, aEndDate: TDateTime;
  aCalcInBegTime, aCalcInEndTime, aCalcOutBegTime, aCalcOutEndTime: TDateTime;
begin
  aBegDate := DateOf(DateTimePicker2.Date);
  aEndDate := DateOf(DateTimePicker3.Date);

  while aBegDate <= aEndDate do
  begin
    aCalcInBegTime := DateOf(aBegDate);
    aCalcInEndTime := EndOfTheDay(aBegDate);
    aCalcOutBegTime:= aCalcInBegTime;
    aCalcOutEndTime:= aCalcInEndTime;
		TdmAcdSummary.Exec_CalcData(aCalcInBegTime, aCalcInEndTime, aCalcOutBegTime, aCalcOutEndTime,
      CheckBox_RecalcWICSIPH2.Checked);
    aBegDate := IncDay(aBegDate);
  end;

  ShowMessage('Done');
end;

procedure TfmMain.Log(AMsg: string);
begin
	JcLog.Write(AMsg);
  ListBox1.Items.Add(AMsg);
  ListBox1.ItemIndex := ListBox1.Items.Count - 1;
end;

procedure TfmMain.ListBox1DblClick(Sender: TObject);
begin
	ListBox1.Clear;
end;

procedure TfmMain.LogLine(CH: Char);
begin
	JcLog.Line(CH);
  ListBox1.Items.Add(StringOfChar(CH, 60));
  ListBox1.ItemIndex := ListBox1.Items.Count - 1;
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

procedure TfmMain.DoAutoJobs;
var
  aCalcDate: TDateTime;
begin
  aCalcDate := IncDay(Date, -1);
  // 兩年度來電量比較表
  // 必須先執行這個作業來產生彙整的統計數據
	TdmAcdSummary.Exec(aCalcDate);
  // ACD接聽率日報表
  Exec;
  //營業處回電效率統計表 Added by Joe 2017/11/09 15:02:02
  TdmSitePhoneSummary.Exec(aCalcDate);
  //訓練個人回電效率統計表 Added by Joe Lee 2017/11/17 16:43:08
  TdmTePhoneSummary.Exec(aCalcDate);
  // Added by Joe 2017/04/21 09:07:38
  try
    if MailMode then
      SendAdminMail;
  except
    on E: Exception do
      JcLog.Write(Format('傳送管理報表發生異常，Err = %s', [E.Message]));
  end;
end;

procedure TfmMain.GetAcdSummaryInfo(ASite: string; var AcdTotalCount, SiteScore, SiteTimeOutScore, SiteAcdDays: Extended);
var
  aQr: TUniQuery;
begin
  aQr := dmReport.GetQuery_WintonTcrm;

  try
    with aQr do
    begin
      SQL.Add('SELECT IPH4003, IPH4004, IPH4006, IPH4010, IPH4011 FROM WICSIPH4 WITH(NOLOCK)');
      AddWhere('(IPH4001 = :IPH4001)');
      AddWhere('(IPH4002 = :IPH4002)');
      ParamByName('IPH4001').Value := dmReport.GetDescOfSite(ASite);
      ParamByName('IPH4002').Value := FCalcDate;
      Open;
      //ACD派送總數
      AcdTotalCount := FieldByName('IPH4006').AsInteger;
      //電話直接率
      SiteScore := FieldByName('IPH4004').AsFloat;
      //值機人力
      SiteAcdDays := FieldByName('IPH4003').AsFloat;
      //逾時率
      if (FieldByName('IPH4011').AsInteger <> 0) then
        SiteTimeOutScore := FieldByName('IPH4010').AsInteger / FieldByName('IPH4011').AsInteger
      else
        SiteTimeOutScore := 0;
    end;
  finally
    dmReport.CloseAndFree(aQr);
  end;
end;

procedure TfmMain.InitCodeSite;
begin
  CodeSite.Enabled := FCodeSiteMode;
end;

procedure TfmMain.btnSitePhoneSummaryClick(Sender: TObject);
begin
  TdmSitePhoneSummary.Exec(DateTimePicker4.Date);
end;

procedure TfmMain.btnTePhoneSummaryClick(Sender: TObject);
begin
  TdmTePhoneSummary.Exec(DateTimePicker5.Date);
end;

end.
