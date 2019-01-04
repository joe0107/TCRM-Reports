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
    // �ˬd����Ҧ�
    procedure CheckExeMode;
  private
  	// ����έp��Ǥ�
    FCalcDate: TDateTime;
    FSiteScore, FSiteTimeOutScore, FSiteAcdDays: Extended;

    FAcdAnsCount: Integer;			// �����`�q��
    FAcdTeCount: Extended; 			// �Ⱦ��H��
    FAvgAnsCount: Extended;			// �C�H���������q��
    FAcdTotalCount: Extended;   // ���e�`�q��
    FAvgGoodAnsCount: Extended;	// �i�F�Ъ��C�H���������q��
    // ���ͳ�����
    function	PrepareData(ASite: string): Boolean;
		// ���oACD��ť���
    function  GetData_IPH3: TUniQuery;
    // ���oACD�Ⱦ����ƯZ���
    function  GetData_CHEM: TUniQuery;
    // ���o�^�q�p�Ƹ��
    function  GetData_RPHE: TUniQuery;
    // �p����~�B���`�Ⱦ��H�O
    function  CalcSiteAcdDays(ADataSet: TDataSet): Extended;
    // �ˬd��~�B�������v�O�_�F�� 80%
    function  CheckSiteAcdScore(AData: TDataSet): Boolean;
    // �̾ڭȾ���ƭp��ӤH��骺��зǵ��G
    procedure CalcPassData(AIPH3, ACHEM, ARPHE: TDataSet);
    // �p��i�F�Ъ������ƾ�
    procedure CalcData(ADataSet: TDataSet);
    // ���oACD���e�`�q��
    function  ACD_GetSiteCount(ASite: string; ABeginTime, AEndTime: TDateTime; AAnswerOnly: Boolean): Integer;
    // �q WICSIPH4 ���o�H�έp��������T Added by Joe 2017/10/24 14:38:23
    procedure GetAcdSummaryInfo(ASite: string; var AcdTotalCount, SiteScore, SiteTimeOutScore, SiteAcdDays: Extended);
	private
    FXlsFileName: string;
    FNetDrvRootDir, FNetDrvUser, FNetDrvPwd: string;

    procedure PrepareXLS(ASite: string);
    // �g�J�ӷ���ƲM��
    procedure WriteDataToXls;
    // �N�ɮ׽ƻs����~�B
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
  // �ᵲ���D��C
  aSheet.FreezePanes(0, 1);
  // �إ߹w�]���x�s��榡
	with XLSReadWriteII51 do
  begin
    CmdFormat.BeginEdit(Nil);
    CmdFormat.Border.Style := cbsThin;
    CmdFormat.Border.Preset(cbspOutline);
    aDefFmt := CmdFormat.AddAsDefault('Format1');
    DefaultFormat := aDefFmt;
  end;
	// �N���Ū�J�u�@��
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
  // �վ�榡
  with XLSReadWriteII51 do
  begin
    with aSheet do
    begin
      Name := 'ACD��ť�v';
      //-----------------------------------------------------------------------
      Columns[0].CharWidth := 15; //�V�m�էO
      Columns[2].CharWidth := 10; //��з�
      Columns[3].CharWidth := 10; //�Ⱦ��Ѽ�
      Columns[4].CharWidth := 10; //ACD����
      Columns[5].CharWidth := 12; //�`�^�q��
      Columns[6].CharWidth := 50; //��]
      Columns[7].CharWidth := 50; //�ﵽ�@�k
      //-----------------------------------------------------------------------
      AsString[0, 0] := '�V�m�էO';
      AsString[1, 0] := '�m�W';
      AsString[2, 0] := '��з�';
      AsString[3, 0] := '�Ⱦ��Ѽ�';
      AsString[4, 0] := '�����q��';
      AsString[5, 0] := '�`�^�q��';
      AsString[6, 0] := '��]';
      AsString[7, 0] := '�ﵽ�@�k';
      //-----------------------------------------------------------------------
      for i := 0 to 7 do
        Cell[i, 0].FillPatternForeColorRGB := clSilver;
    end;

    // �N���F�ФH���H���r�Х�
    for i := 1 to aSheet.LastRow do
    begin
      // �p�G�S��[ACD��з�]�q��,���n�p��
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
          Log(Format('�Хܥ��F�ФH���o�Ϳ��~ = %s', [aSheet.AsString[1, i]]));
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
  // �p��e�@�Ѫ�ACD��ť���
  FCalcDate := IncDay(Date,  -1);
  DateTimePicker1.Date := FCalcDate;
  DateTimePicker2.Date := FCalcDate;
  DateTimePicker3.Date := FCalcDate;
//  DateTimePicker2.Date := EncodeDate(2017, 11, 7);
//  DateTimePicker3.Date := EncodeDate(2017, 11, 7);
  DateTimePicker4.Date := FCalcDate;
  DateTimePicker5.Date := FCalcDate;
  DateTimePicker6.Date := FCalcDate;
  // �ˬd�O�_���۰ʰ���Ҧ�
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
  dmReport.Init_IdMessage(Result);    //�]�w�l���ݩ� Added by Joe Lee 2017/11/20 09:47:24
  aDayOfWeek := GetChineseNumStr(DayOfWeek(FCalcDate) - 1);
  if (aDayOfWeek = '�s') then aDayOfWeek := '��';

  with Result do
  begin
    //��J�����
    if FDebugMode then
    begin
      Recipients.Add.Address := 'joelee@winton.com.tw';
    end
    else
    begin
      aText := dmReport.GetEmail_TE_Admin(ASite);
      Recipients.Add.Address := aText;
      //��J�ƥ�
      aText := dmReport.GetEmail_Site_Admin(ASite);
      CCList.Add.Address := aText;
    end;

    MakeCCList(CCList);
    //��J�l����Y��T
    Subject := Format('�q�ܮĲv���F�q��_%s(%s)_%s', [FormatDateTime('yyyymmdd', FCalcDate), aDayOfWeek, ASite]);
    //��J�l�󤺮e
	  aCount := FAvgGoodAnsCount - FAvgAnsCount;
    aText := Format('[%s] %s(%s)', 	[ASite, FormatDateTime('yyyy/mm/dd', FCalcDate), aDayOfWeek]);
    aText := aText + CHR(13) + Format('ACD���e�`�� %.0f�A�q�ܪ����v %.2f%%�A�O�ɲv %.1f%%�A�Ⱦ��H�O %.1f�A���F�зǡC',
    	[FAcdTotalCount, FSiteScore, FSiteTimeOutScore*100, FSiteAcdDays]);

		aText := aText + CHR(13) + Format('�n�F�������v 80%% ���зǡA�����C��Ⱦ��H���n�A�h�� %d �qACD���e�q�ܡC', [Ceil(aCount)]);
    aText := aText + CHR(13) + '�ӤH���F��зǲM��аѾ\���󤤪����r�Хܸ�ơC';
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
    JcLog.Write(Format('�S�������ơA������XLS�ɮסC', [ASite]));
  	Exit;
  end;

  aReportFolder := ExtractFilePath(Application.ExeName) + 'ReportStock';
  ForceDirectories(aReportFolder);
  FXlsFileName := aReportFolder + Format('\�q�ܮĲv���F���i_%s_%s.xlsx', [FormatDateTime('yyyymmdd', FCalcDate), ASite]);

  with XLSReadWriteII51 do
  begin
    Clear;
    Filename := FXlsFileName;
  end;

   WriteDataToXls;
   XLSReadWriteII51.Write;
   JcLog.Write(Format('����XLS�ɮ� = %s', [FXlsFileName]));

  if FileExists(FXlsFileName) then
  	CopyXlsToSite(ASite);

//  if not FAutoMode then
//  	ShellExecute(0, 'open', PChar(FXlsFileName), nil, nil, SW_SHOWMAXIMIZED);
end;

procedure TfmMain.Exec;
begin
  JcLog.Write('�}�l����[ACD�q�ܮĲv���F���i]');

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
  JcLog.Write('Ū��ACD��ť���');
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
      JcLog.Write(Format('���oACD��ť��ơA�O���� = %d', [RecordCount]));
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
  JcLog.Write('Ū��ACD�Ⱦ����ƯZ���');
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
      JcLog.Write(Format('���oACD�Ⱦ����ƯZ��ơA�O���� = %d', [RecordCount]));
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
  JcLog.Write('�̾ڭȾ���ƭp��ӤH��骺��зǵ��G');

	with AIPH3 do
  begin
    Filtered := False;
    First;

    while not Eof do
    begin
      // �p�G���ACD�Ⱦ����ƯZ��ơA���⦨�зǤ��B���
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
  JcLog.Write(Format('�}�l��z������[%s]', [ASite]));
  // Added by Joe 2017/10/24 14:57:11
  // �����q 10.1.1.212 WICSIPH4 ���o�έp��T
  GetAcdSummaryInfo(ASite, FAcdTotalCount, FSiteScore, FSiteTimeOutScore, FSiteAcdDays);
  // �p�GACD��ť�v�W�L80%�A���ݭn�~��B�z
  if (FSiteScore >= 80) then
  begin
    JcLog.Write(Format('ACD��ť�v %.2f%% �W�L80%%�A��������[%s]', [FSiteScore, ASite]));
    Exit;
  end
  else
    JcLog.Write(Format('ACD��ť�v %.2f%% ���F80%%�A�}�l���ͳ���[%s]', [FSiteScore, ASite]));
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
          aText := Format('�L�k�s�u��D�� %s', [aHost]);
          Application.MessageBox(PChar(aText), PChar(Application.Title), MB_OK + MB_ICONWARNING);
        end;
        JcLog.Write(Format('Error, failed to connect %s', [aHost]));
        Exit;
      end;
			// ���oACD��ť���
      aWICSIPH3 := GetData_IPH3;
      // �p�G�S��ACD��ť��ơA���ݭn�~��B�z
      if (aWICSIPH3.IsEmpty) then
      begin
        JcLog.Write(Format('�S��ACD��ť��ơA��������[%s]', [ASite]));
        Exit;
      end;
      (**
      if CheckSiteAcdScore(aWICSIPH3) then
      begin
        JcLog.Write(Format('ACD��ť�v %.1f%% �W�L80%%�A��������[%s]', [FSiteScore*100, ABranch]));
      	Exit;
      end
      else
        JcLog.Write(Format('ACD��ť�v %.1f%% ���F80%%�A�}�l���ͳ���[%s]', [FSiteScore*100, ABranch]));
      **)
      // ���oACD�Ⱦ����ƯZ���
			aWICSCHEM := GetData_CHEM;
    	// ���o�^�q�p�Ƹ��
      aWICSRPHE := GetData_RPHE;
    	// �p����~�B���`�Ⱦ��H�O
    	//FSiteAcdDays := CalcSiteAcdDays(aWICSCHEM);
      // �̾ڭȾ���ƭp��ӤH��骺��зǵ��G
      mdReport.LoadFromDataSet(aWICSIPH3);
			CalcPassData(mdReport, aWICSCHEM, aWICSRPHE);
			// �L�o���Ⱦ������
      JcLog.Write(Format('�L�o���Ⱦ������[%s]', [ASite]));
    	mdReport.Filtered := True;
      // ���oACD���e�`�� Added by Joe 2017/03/29 11:32:00
      //FAcdTotalCount := ACD_GetSiteCount(ABranch, FCalcDate, IncMilliSecond(FCalcDate+1, -1), False);
			// �p��n�F�Ъ��U�����R�ƾ�
      CalcData(mdReport);

      if FAcdAnsCount > 0 then
        Result := True;

      JcLog.Write(Format('�w���������ƾ�z[%s]', [ASite]));
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
    JcLog.Write(Format('�w�z�L�l��ǰe����[%s]', [ASite]));
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
  // �ˬd�ҰʼҦ�
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
  JcLog.Write('Ū���^�q�p�Ƹ��');
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
      JcLog.Write(Format('���o�^�q�p�Ƹ�ơA�O���� = %d', [RecordCount]));
    except
      on E: Exception do
        JcLog.Write(Format('GetData_RPHE() failed, error = %s', [E.Message]));
    end;
  end;
end;

function TfmMain.CalcSiteAcdDays(ADataSet: TDataSet): Extended;
begin
  JcLog.Write('�p����~�B���`�Ⱦ��H�O');
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
  JcLog.Write(Format('��~�B���`�Ⱦ��H�O = %.1f', [Result]));
end;

procedure TfmMain.CopyXlsToSite(ASite: string);
var
  aDstFileName: string;
begin
  JcLog.Write(Format('�ƻsXLS�ɮר���~�B[%s]', [ASite]));

  if (aSite = SITE_NAME_Taipei_TC) then
  	FNetDrvRootDir := '\\wtp4\Winnan\�V�m�M��\�x�_�q�ܮĲv���F���R��'
  else if (aSite = SITE_NAME_Taoyuan_TC) then
  	FNetDrvRootDir := '\\10.3.1.45\public\wty2\�V�m��\�q�ܮĲv���F���R��'
  else if (aSite = SITE_NAME_Taichung_TC) then
  	FNetDrvRootDir := '\\10.5.1.4\TE\TCRM�έp���\�^�q�έp\���Ϲq�ܮĲv���F���R��'
  else if (aSite = SITE_NAME_Tainan_TC) then
    	FNetDrvRootDir := '\\10.6.1.66\���F���R��';

	FNetDrvUser := 'winton\rdrepl';
  FNetDrvPwd := 'Wint0n2k';

  try
    aDstFileName := Format('%s\%s', [FNetDrvRootDir, ExtractFileName(FXlsFileName)]);
    JcLog.Write(Format('�ؼ�XLS�ɮ� = %s', [aDstFileName]));
    NetDrive1.Connect(FNetDrvRootDir, FNetDrvUser, FNetDrvPwd);
    JcLog.Write('�s�u�����Ϻо�');

    if NetDrive1.Connected or (NetDrive1.ErrorCode = 1219) then
    begin
      if FMailMode then
      begin
        CopyFile(PChar(FXlsFileName), PChar(aDstFileName), False);
        JcLog.Write('�w�ƻsXLS�ɮר���~�B');
      end
    end
    else
      JcLog.Write('�`�N!!�L�k�ƻsXLS�ɮר���~�B');
  finally
  	NetDrive1.Disconnect;
  end;
end;

procedure TfmMain.CalcData(ADataSet: TDataSet);
var
  aAcdCount: Extended;	  // ���Ĭ��e�`�q��
begin
  JcLog.Write('�p��n�F�Ъ��U�����R�ƾ�');
  FAcdAnsCount := 0;			// �����`�q��
  FAcdTeCount  := 0;			// �Ⱦ��H��
  FAvgAnsCount := 0;			// �C�H���������q��
  FAvgGoodAnsCount := 0;	// �i�F�Ъ��C�H���������q��

	with ADataSet do
  begin
    First;
    while not Eof do
    begin
      FAcdTeCount := FAcdTeCount + FieldByName('ACD_DAY').AsFloat;
    	FAcdAnsCount := FAcdAnsCount + FieldByName('IPH3003').AsInteger;
    	Next;
    end;
    // �C�H���������q��
    FAvgAnsCount := JcDivide(FAcdAnsCount, FAcdTeCount);
    JcLog.Write(Format('�C�H���������q�� = %.1f', [FAvgAnsCount]));
    // ���e�`�q��
    aAcdCount := JcDivide(FAcdAnsCount, FieldByName('IPH3004').AsFloat);
    JcLog.Write(Format('���e�`�q�� = %.1f', [aAcdCount]));
    // �i�F�Ъ��C�H���������q��
    FAvgGoodAnsCount := JcDivide((aAcdCount * 0.8), FAcdTeCount);
    JcLog.Write(Format('�i�F�Ъ��C�H���������q�� = %.1f', [FAvgGoodAnsCount]));
  end;
end;

function TfmMain.ACD_GetSiteCount(ASite: string; ABeginTime, AEndTime: TDateTime; AAnswerOnly: Boolean): Integer;
var
  aData: TUniQuery;
begin
  JcLog.Write(Format('Ū��ACD���e�`��[%s]', [ASite]));

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
        JcLog.Write(Format('���oACD���e�`�ơA�O���� = %d', [RecordCount]));
      except
        on E: Exception do
          JcLog.Write(Format('ACD_GetSiteCount(%s) failed, error = %s', [ASite, E.Message]));
      end;

      First;
      Result := FieldByName('PHONE_COUNT').AsInteger;
      JcLog.Write(Format('���oACD���e�`�� = %d', [Result]));
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
  if (aDayOfWeek = '�s') then aDayOfWeek := '��';  

  with Result do
  begin
    //��J�����
    Recipients.EMailAddresses := dmReport.AdminEmail;
    //��J�l����Y��T
    Subject := Format('ACD�����_%s(%s)_������i', [FormatDateTime('yyyymmdd', FCalcDate), aDayOfWeek]);
    //�H��H�a�}
    From.Address := 'rdrepl@winton.com.tw';
    ContentType := 'multipart/mixed';
    //��J�l�󤺮e
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
  // ��~�רӹq�q�����
  // ����������o�ӧ@�~�Ӳ��ͷJ�㪺�έp�ƾ�
	TdmAcdSummary.Exec(aCalcDate);
  // ACD��ť�v�����
  Exec;
  //��~�B�^�q�Ĳv�έp�� Added by Joe 2017/11/09 15:02:02
  TdmSitePhoneSummary.Exec(aCalcDate);
  //�V�m�ӤH�^�q�Ĳv�έp�� Added by Joe Lee 2017/11/17 16:43:08
  TdmTePhoneSummary.Exec(aCalcDate);
  // Added by Joe 2017/04/21 09:07:38
  try
    if MailMode then
      SendAdminMail;
  except
    on E: Exception do
      JcLog.Write(Format('�ǰe�޲z����o�Ͳ��`�AErr = %s', [E.Message]));
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
      //ACD���e�`��
      AcdTotalCount := FieldByName('IPH4006').AsInteger;
      //�q�ܪ����v
      SiteScore := FieldByName('IPH4004').AsFloat;
      //�Ⱦ��H�O
      SiteAcdDays := FieldByName('IPH4003').AsFloat;
      //�O�ɲv
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
