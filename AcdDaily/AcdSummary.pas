unit AcdSummary;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs, DB, ADODB, StdCtrls, ShellAPI,
  cxData, cxClasses, cxCustomData, cxDataStorage, cxDBData, JclStrings, MemDS, DBAccess, Uni, kbmMemTable, DateUtils,
  dxmdaset, XLSSheetData5, XLSReadWriteII5, XLSDbRead5, XLSNames5, Xc12Utils5, IdEMailAddress, IdAttachmentFile,
  IdMessage, CodeSiteLogging;

type
  TdmAcdSummary = class(TDataModule)
    {$REGION 'RAD'}
    qrAgentCount: TUniQuery;
    qryGetAcdInfo: TUniQuery;
    qrGetData_Src: TUniQuery;
    qrGetData_SrcIPHE001: TIntegerField;
    qrGetData_SrcRPHE001: TIntegerField;
    qrGetData_SrcIPHE004: TDateTimeField;
    qrGetData_SrcIPHE005: TStringField;
    qrGetData_SrcIPHE003: TStringField;
    qrGetData_SrcRPHE003: TStringField;
    qrGetData_SrcSALE002: TStringField;
    qrGetData_SrcSALE003: TStringField;
    qrGetData_SrcRPHE005: TDateTimeField;
    qrGetData_SrcRPHE006: TDateTimeField;
    qrGetData_SrcIPHE016: TDateTimeField;
    qrGetData_SrcIPHE017: TDateTimeField;
    qrGetData_SrcIPHE012: TStringField;
    qrGetData_SrcIPHE008: TStringField;
    qrGetData_SrcIPHE019: TStringField;
    mdAcdTeDaily: TdxMemData;
    mdAcdTeDailyPhoneDate: TDateField;
    mdAcdTeDailyEmpId: TStringField;
    mdAcdTeDailyEmpName: TStringField;
    mdAcdTeDailyDays: TFloatField;
    mdAcdTeDailyACD_Ans_Total: TIntegerField;
    mdAcdTeDailyACD_Ans_Valid: TIntegerField;
    mdAcdTeDailyACD_Assign_Total: TIntegerField;
    mdAcdTeDailyACD_Assign_Valid: TIntegerField;
    mdAcdTeDailyCallout_Total: TIntegerField;
    mdAcdTeDailyACD_In_Total: TIntegerField;
    mdAcdTeDailyACD_ValidAns_Total: TIntegerField;
    mdAcdTeDailyACD_Score: TFloatField;
    mdAcdSiteDaily: TdxMemData;
    mdAcdSiteDailyPhoneDate: TDateField;
    mdAcdSiteDailySiteId: TStringField;
    mdAcdSiteDailyACD_Total: TIntegerField;
    mdAcdSiteDailyACD_Assign_Invalid: TIntegerField;
    mdAcdSiteDailyACD_ValidIn_Total: TIntegerField;
    mdAcdSiteDailyACD_ValidAns_Total: TIntegerField;
    mdAcdSiteDailyACD_Score: TFloatField;
    mdAcdSrc: TkbmMemTable;
    mdAcdSrcIPHE001: TIntegerField;
    mdAcdSrcRPHE001: TIntegerField;
    mdAcdSrcIPHE004: TDateTimeField;
    mdAcdSrcIPHE005: TStringField;
    mdAcdSrcIPHE003: TStringField;
    mdAcdSrcRPHE003: TStringField;
    mdAcdSrcSALE002: TStringField;
    mdAcdSrcSALE003: TStringField;
    mdAcdSrcRPHE005: TDateTimeField;
    mdAcdSrcRPHE006: TDateTimeField;
    mdAcdSrcIPHE016: TDateTimeField;
    mdAcdSrcIPHE017: TDateTimeField;
    mdAcdSrcIPHE012: TStringField;
    mdAcdSrcIPHE008: TStringField;
    mdAcdSrcIPHE019: TStringField;
    mdAcdSrcAnswerInTime: TBooleanField;
    mdAcdTeDailyDeptId: TStringField;
    mdAcdSiteDailySiteName: TStringField;
    mdAcdTeDailyDeptName: TStringField;
    mdAcdSrcValid: TBooleanField;
    qrGetData_Public: TUniQuery;
    mdAcdTeDailyNotAcdCallout_Count: TIntegerField;
    mdAcdSiteDailyNotAcdCallout_Count: TIntegerField;
    mdAcdSrcCLAS004: TStringField;
    mdAcdSrcCALL_KIND: TStringField;
    mdAcdSrcRPHE011: TStringField;
    mdAcdSrcExcept: TBooleanField;
    mdAcdSrcRemark: TStringField;
    mdAcdSrcRPHE011_REV: TBooleanField;
    mdAcdSiteDailyTE_Total_C: TIntegerField;
    mdAcdSiteDailyCallout_Total: TIntegerField;
    mdAcdSrcCUT1002: TStringField;
    qrGetData_SrcCUT1002: TStringField;
    mdAcdSiteDailyDays: TFloatField;
    mdAcdTeDailyTE_Total_C: TIntegerField;
    mdAcdTeDailySiteName: TStringField;
    mdAcdTeDailyPhoneDateDesc: TStringField;
    qrGetData_SrcFLAG_SW: TStringField;
    qrGetData_SrcFLAG_HRS: TStringField;
    qrGetData_SrcIPH2001: TStringField;
    qrGetData_SrcIPH2002: TIntegerField;
    qrGetData_SrcIPH2003: TBooleanField;
    mdAcdSrcIPH2001: TStringField;
    mdAcdSrcIPH2002: TIntegerField;
    mdAcdSrcIPH2003: TBooleanField;
    mdAcdSrcFLAG_SW: TStringField;
    mdAcdSrcFLAG_HRS: TStringField;
    mdAcdTeDailyTimeOut_Count_C: TIntegerField;
    mdAcdTeDailyPhone_Count_C: TIntegerField;
    mdAcdTeDailyPhone_Count_NC: TIntegerField;
    mdAcdSiteDailyPhone_Count_C: TIntegerField;
    mdAcdSiteDailyPhone_Count_NC: TIntegerField;
    mdAcdSiteDailyTimeOut_Count_C: TIntegerField;
    mdAcdSiteDailyTimeOut_Rate_C: TFloatField;
    mdAcdTeDailyTimeOut_Count_NC: TIntegerField;
    mdAcdSiteDailyTimeOut_Count_NC: TIntegerField;
    mdAcdSiteDailyTimeOut_Rate_NC: TFloatField;
    mdAcdSiteDailyPhoneDateDesc: TStringField;
    mdAcdSiteDailyACD_InvalidIn_Total: TIntegerField;
    cmdInsIPH2: TUniSQL;
    qrGetData_SrcGUID: TStringField;
    mdAcdSrcGUID: TStringField;
    mdAcdSwDaily: TdxMemData;
    mdAcdSwDailyPhoneDate: TDateField;
    mdAcdSwDailySiteId: TStringField;
    mdAcdSwDailyACD_Total: TIntegerField;
    mdAcdSwDailyACD_Assign_Invalid: TIntegerField;
    mdAcdSwDailyACD_InvalidIn_Total: TIntegerField;
    mdAcdSwDailyACD_ValidIn_Total: TIntegerField;
    mdAcdSwDailyACD_ValidAns_Total: TIntegerField;
    mdAcdSwDailyACD_Score: TFloatField;
    mdAcdSwDailySiteName: TStringField;
    mdAcdSwDailyTE_Total_C: TIntegerField;
    mdAcdSwDailyPhone_Count_C: TIntegerField;
    mdAcdSwDailyPhone_Count_NC: TIntegerField;
    mdAcdSwDailyTimeOut_Count_C: TIntegerField;
    mdAcdSwDailyTimeOut_Rate_C: TFloatField;
    mdAcdSwDailyTimeOut_Count_NC: TIntegerField;
    mdAcdSwDailyTimeOut_Rate_NC: TFloatField;
    mdAcdSwDailySw: TStringField;
    mdAcdSiteDailyACD_Ans_Total: TIntegerField;
    mdAcdSiteDailyPhoneOut_Count_C: TIntegerField;
    mdAcdSiteDailyPhoneOut_Count_NC: TIntegerField;
    cmdUpdIPH2: TUniSQL;
    qrGetData_SrcIPH2004: TStringField;
    qrGetData_SrcRPHE_GUID: TStringField;
    mdAcdSrcIPH2004: TStringField;
    mdAcdSrcRPHE_GUID: TStringField;
    qrGetData_PhoneOut: TUniQuery;
    mdPhoneOutSrc: TkbmMemTable;
    mdPhoneOutSrcRPHE001: TIntegerField;
    mdPhoneOutSrcRPHE003: TStringField;
    mdPhoneOutSrcSALE002: TStringField;
    mdPhoneOutSrcRPHE005: TDateTimeField;
    mdPhoneOutSrcRPHE011: TStringField;
    mdPhoneOutSrcIPHE005: TStringField;
    mdPhoneOutSrcCUT1002: TStringField;
    mdPhoneOutSrcIPH2001: TStringField;
    mdPhoneOutSrcInvalid: TBooleanField;
    qrGetData_PhoneOutRPHE001: TIntegerField;
    qrGetData_PhoneOutRPHE003: TStringField;
    qrGetData_PhoneOutSALE002: TStringField;
    qrGetData_PhoneOutRPHE005: TDateTimeField;
    qrGetData_PhoneOutRPHE011: TStringField;
    qrGetData_PhoneOutIPHE005: TStringField;
    qrGetData_PhoneOutCUT1002: TStringField;
    qrGetData_PhoneOutIPH2001: TStringField;
    mdPhoneOutSrcRPHE011_REV: TBooleanField;
    qrGetData_PhoneOutDEPT002: TStringField;
    mdPhoneOutSrcDEPT002: TStringField;
    qrGetData_PhoneOutDEPT001: TStringField;
    mdPhoneOutSrcDEPT001: TStringField;
    mdAcdSiteDailyNoAns_Count_C: TIntegerField;
    mdAcdSiteDailyNoAns_Count_NC: TIntegerField;
    mdAcdTeDailyPhoneOut_Count_C: TIntegerField;
    mdAcdTeDailyPhoneOut_Count_NC: TIntegerField;
    qrGetData_SrcTK: TStringField;
    mdAcdSrcTK: TStringField;
    mdAcdTeDailyTE_Total_NC: TIntegerField;
    mdAcdSiteDailyTE_Total_NC: TIntegerField;
    mdAcdSwDailyTE_Total_NC: TIntegerField;
    connReport: TUniConnection;
    mdDupCheck: TdxMemData;
    mdDupCheckTK: TStringField;
    mdDupCheckCount: TIntegerField;
    mdAcdSrcDUP: TBooleanField;
    mdAcdSwDailyPhoneYear: TIntegerField;
    mdAcdSwDailyPhoneMonth: TIntegerField;
    mdAcdSiteDailyPhoneYear: TIntegerField;
    mdAcdSiteDailyPhoneMonth: TIntegerField;
    mdAcdTeDailyPhoneYear: TIntegerField;
    mdAcdTeDailyPhoneMonth: TIntegerField;
    qrGetIPH2004: TUniQuery;
    qrGetIPH2004GUID: TStringField;
    mdAcdTeDailyACD_DailyReqCount: TIntegerField;
    qrGetData_SrcSALE024: TIntegerField;
    mdAcdSrcSALE024: TIntegerField;
    XLSReadWriteII51: TXLSReadWriteII5;
    qrWICSIPHH: TUniQuery;
    mdAcdTeDailyDuty_AM: TStringField;
    mdAcdTeDailyDuty_PM: TStringField;
    qrPrecedingAcdTotal: TUniQuery;
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure mdAcdSrcCalcFields(DataSet: TDataSet);
    procedure mdAcdSiteDailyCalcFields(DataSet: TDataSet);
    procedure mdAcdTeDailyCalcFields(DataSet: TDataSet);
    procedure mdAcdSiteDailyTimeOut_Rate_CGetText(Sender: TField; var Text: String; DisplayText: Boolean);
    procedure mdPhoneOutSrcCalcFields(DataSet: TDataSet);
    {$ENDREGION}
  protected
    FInCalcData: Boolean;
    FCalcInBegTime, FCalcInEndTime: TDateTime;
    FCalcOutBegTime, FCalcOutEndTime: TDateTime;
    FForceCalcData: Boolean;

    procedure InitExecute;
    procedure BeginExecute;
    procedure EndExecute;
    //�������o��ƨӷ������γs�u
    procedure InitReportConn(ASiteId: string);
    procedure InitData;
    procedure Log(AMsg: string);
    procedure LogLine(Ch: Char = '-');
  private
    function  CheckTeRec(ADeptId: string): Boolean;
    function  CheckTeDailyRec(ADataSet: TDataSet): Boolean; overload;
    function  CheckTeDailyRec(ADataSet: TDataSet; ASiteId, AEmpID, AEmpName, ADeptId: string; ADate: TDateTime): Boolean; overload;
    function  CheckSiteDailyRec(ASiteId: string; ADate: TDateTime): Boolean;
    function  CheckSwDailyRec(ASiteId, ASw: string; ADate: TDateTime): Boolean; overload;
    function  GetDeptFilterStr(ASiteId: string): string;
  private
    function  GetSwByAgentGroup(AAgentGroup: string): string;
    //�qTeleContact���o���w���I��ACD���e�q��
    //AAnswerOnly = False�A�Ҧ��z�LACD���e���q��
    //AAnswerOnly = True�A�Ҧ��z�LACD���e�ӥB���Q��ť���q��
    procedure ACD_GetSiteCount(ASite: string; ABeginTime, AEndTime: TDateTime; AAnswerOnly: Boolean = False);
    procedure ACD_GetGroupCount(ASite: string; ABeginTime, AEndTime: TDateTime; AAnswerOnly: Boolean);
    //�p��ϻ���ACD��T
    procedure ACD_CalcData_Site(ASiteId: string; AAnswerOnly: Boolean = False);
    procedure ACD_CalcData_Group(ASiteId: string);
  private
    //�z��[�V�m�v]ACD�ȯZ���
    procedure GetTeOnDutyCount(ASiteId: string; ABeginTime, AEndTime: TDateTime);
    //�έp[�V�m�v]ACD�ȯZ���
    procedure FillData_TeOnDutyCount(ASiteId: string; ABeginTime, AEndTime: TDateTime);
    //Added by Joe 2017/11/13 14:41:16
    //���o[�V�m�v]ACD�Ⱦ����ظ��
    function  GetTeDutyItem(ASiteId: string; ABegTime, AEndTime: TDateTime): TUniQuery;
    //��J[�V�m�v]ACD�Ⱦ����ظ��
    procedure FillData_TeDutyItem(ASiteId: string; ABegTime, AEndTime: TDateTime);
  private
    //Ū�����w��~�B���ӷ����
    procedure PrepareData_PhoneIn(ASiteId: string; ABeginTime, AEndTime: TDateTime);
    procedure PrepareData_PhoneOut(ASiteId: string; ABeginTime, AEndTime: TDateTime);
    procedure GetData_PhoneIn(ASiteId: string; ABeginTime, AEndTime: TDateTime);
    procedure GetData_PhoneOut(ASiteId: string; ABeginTime, AEndTime: TDateTime);
    //�M�z�ӹq��ơA�u�O�d�����ӹq���Ĥ@���^�q����
    procedure CleanData_PhoneIn(ASiteId: string; ADataSet: TDataSet);
    //�M�z�^�q���
    procedure CleanData_PhoneOut(ADataSet: TDataSet);
    //���O�n�������ҥ~�ӹq
    procedure AddOneExceptPhone(ASiteId, ASw: string; ADate: TDateTime; AKind: Integer);
    //�N�ӷ���Ƶ��O���L�Ĩӹq
    procedure MarkSrcPhoneAsValid;
    //�N�ӷ���Ƶ��O���ҥ~�ӹq
    procedure MarkSrcPhoneAsExcept(ARemark: string);
    //�P�_�O�_���X���Ȥ�
    function  IsContracted(AFlag: string): Boolean;
    //�P�_�Y���h�ݭn�M�ήɶ��L�o����
    function  UseTimeFilter: Boolean;
    procedure PrepreData_WICSIPHH(ABegTime, AEndTime: TDateTime); // Added by Joe 2017/11/10 15:51:04
    //�P�_�O�_���]���ȯZ��
    function  IsNightShift(ADate: TDateTime): Boolean;
    //�P�_�O�_����w���� Added by Joe 2019/04/11 10:10:19
    function  IsNationalHoliday(ADate: TDateTime): Boolean;
  private
    //�p��έp�ӷ����
    procedure CalcData_PhoneIn(ASiteId: string);
    procedure CalcData_PhoneOut(ASiteId: string);
    procedure CalcOneData_IPH2;
    //�p��V�m�v����έp���
    procedure CalcOneData_Te_Daily(ASiteId: string; ADate: TDateTime);
    procedure CalcData_Te_Daily;
    procedure CalcOneData_Te_PhoneCount(ASiteId: string; ADate: TDateTime);
    //-------------------------------------------------------------------------
    //�p����~�B����έp���
    procedure CalcOneData_SiteDaily_Count(ASiteId: string);
    //�p����~�B����έp���
    procedure CalcData_Site_Daily(ASiteId: string);
    //-------------------------------------------------------------------------
    //�p��t�ΧO����έp���
    procedure CalcOneData_SwDaily(ASiteId: string);
    //�p��t�ΧO����έp���
    procedure CalcData_Sw_Daily;
    //�ˬd���ƪ��q���ѧO�X 2015.12.02
    procedure CheckDuplicateTK;
   	//-------------------------------------------------------------------------
   	// Added by Joe 2016/07/11 10:48:23
   	function  Get_IPH2004(IPHE001: Integer): string;
    // Added by Joe 2017/11/09 16:11:57
    // �P�_TE�פJ�ӥB�w�W�L�W�Z�ɶ����ӹq
    function  Is_OffDuty_TE: Boolean;
  private
   	// Added by Joe 2017/04/25 14:58:41
   	// �N�έp����x�s����v��Ʈw��
   	procedure SaveAcdSiteDaily;
   	procedure SaveAcdTeDaily;
    procedure SaveAcdSwDaily;
    procedure CalcAcdData(ASiteId: string);
    procedure PrepareData_ACD(ASiteId: string);
  private
    FXlsFileName: string;
    FNewRptCount: Integer;
    function  CopyReportFromTemplate(ADate: TDateTime): string;
    function  PrepareData_SwACD(ASw: string; AYear, AMonth: Integer): TUniQuery;
    procedure WriteDataToXls(ASw: string; AYear, AMonth, ADay: Word);
    procedure UpdateWorkSheet_Summary(AYear, ADay: Integer);
    // �ǳƨ�~�ת��~�׫e��ACD�q�ƦX�p��� // Added by Joe 2018/05/21 11:27:44
    procedure PrepareData_PrecedingAcdTotal;
    // ���o���w�t�ΧO����~�׫e��ACD�q�ƦX�p // Added by Joe 2018/05/21 11:43:46
    procedure GetPrecedingAcdTotal(IPH5003: string; var ATotal, APrevTotal: Integer);
	protected
    procedure SendMail;
    procedure MakeCCList(AEmailAddrList: TIdEmailAddressList);
    function  MakeNotifyMessage: TIdMessage;
    function  MakeAdminNotifyMessage: TIdMessage;
  public
    procedure PrintReport(ADate: TDateTime);
    procedure CalcReportData(ACalcInBegTime, ACalcInEndTime, ACalcOutBegTime, ACalcOutEndTime: TDateTime);
    // �p����x�s���w�϶���ACD��ť�v�έp���
    class procedure Exec_CalcData(ACalcInBegTime, ACalcInEndTime, ACalcOutBegTime, ACalcOutEndTime: TDateTime;
      AForceReCalc: Boolean = False);
    // ���ͫ��w����� XLS ����(�����s�p��,�������Ȳ��ͳ���)
    class procedure Exec_PrintReport(ADate: TDateTime);
    // ���ͫ��w����� XLS ����(���p����,�M�Უ�ͳ���)    
    class procedure Exec(ADate: TDateTime);
	end;

var
  dmAcdSummary: TdmAcdSummary;

implementation

uses
  TcrmConstants, JcDateTimeUtils, JcNumUtils, JcDevExpressUtils, JcDataSetUtils, ReportData, Main;

{$R *.dfm}

{ TfmAcdSummary }

procedure TdmAcdSummary.FormDestroy(Sender: TObject);
begin
  inherited;
  dmAcdSummary := nil;
end;

procedure TdmAcdSummary.FormCreate(Sender: TObject);
begin
  inherited;
  FInCalcData := False;
  FForceCalcData := False;
  dmReport.InitLookup_Class_10;
end;

procedure TdmAcdSummary.GetTeOnDutyCount(ASiteId: string; ABeginTime, AEndTime: TDateTime);
begin
  with qrAgentCount do
  begin
    if Active then Close;
    SQL.Clear;
    SQL.Add('SELECT CHEM001, CHEM004, SALE002, DEPT001, COUNT(*) AS AGENT_COUNT, SALE024');
    SQL.Add('FROM WICSCHEM WITH(NOLOCK)');
    SQL.Add('LEFT JOIN WICSSALE WITH(NOLOCK) ON SALE001 = CHEM004');
    SQL.Add('LEFT JOIN WICSDEPT WITH(NOLOCK) ON DEPT001 = SALE003');
    SQL.Add('LEFT JOIN WICSSTM2 WITH(NOLOCK) ON STM2001 = CHEM005 AND STM2002 = CHEM006');
    SQL.Add('WHERE (STM2004 = ''Y'')');
    //SQL.Add('WHERE (CHEM006=''19'' OR CHEM006=''22'' OR CHEM006=''23'' OR CHEM006=''24'')');
    SQL.Add('AND (CHEM001 >= :CHEM001B AND CHEM001 <= :CHEM001E)');
    //-----------------------------------------------------------------------
    SQL.Add('GROUP BY CHEM001, CHEM004, SALE002, DEPT001, SALE024');

    ParamByName('CHEM001B').Value := DateOf(ABeginTime);
    ParamByName('CHEM001E').Value := DateOf(AEndTime);
    Open;
  end;
end;

procedure TdmAcdSummary.FillData_TeOnDutyCount(ASiteID: string; ABeginTime, AEndTime: TDateTime);
var
  aDate: TDateTime;
  aBranch, aTe, aTeName: string;
begin
  GetTeOnDutyCount(ASiteID, ABeginTime, AEndTime);

  with qrAgentCount do
  begin
    First;

    while not Eof do
    begin
      aBranch := FieldByName('DEPT001').AsString;
      aBranch := Copy(aBranch, 1, 2);

      if (aBranch = BRANCH_ID_Hsinchu) then
        aBranch := BRANCH_ID_Taoyuan
      else if (aBranch = BRANCH_ID_Kaohsiung) then
        aBranch := BRANCH_ID_Tainan;

      if (aBranch = ASiteID) then
      begin
        aDate := DateOf(FieldByName('CHEM001').AsDateTime);
        aTe   := FieldByName('CHEM004').AsString;
        aTeName := FieldByName('SALE002').AsString;
        aBranch := FieldByName('DEPT001').AsString;
        //��J�Ⱦ��H��
        if CheckTeDailyRec(mdAcdTeDaily, ASiteId, aTe, aTeName, aBranch, aDate) then
        begin
          mdAcdTeDaily.Edit;
          mdAcdTeDailyDays.AsFloat := FieldByName('AGENT_COUNT').AsInteger / 2;
          // Added by Joe 2017/02/23 14:53:16
          mdAcdTeDailyACD_DailyReqCount.AsInteger := FieldByName('SALE024').AsInteger;
          // ------------------------------------------------------------------
          mdAcdTeDaily.Post;
        end;
      end;
      Next;
    end;
  end;
end;

procedure TdmAcdSummary.ACD_GetSiteCount(ASite: string; ABeginTime, AEndTime: TDateTime; AAnswerOnly: Boolean);
begin
  if not Assigned(dmReport) then
    raise Exception.Create(ERR_DMMSSQL_MISSING);

  dmReport.SetConn_TeleContact(ASite);

  with qryGetAcdInfo do
  begin
    if Active then Close;
    SQL.Clear;
    SQL.Add('SELECT DATEPART(year, ITIME) AS _YEAR_,');
    SQL.Add('DATEPART(month, ITIME) AS _MONTH_, DATEPART(day, ITIME) AS _DAY_,');
    SQL.Add('COUNT(*) AS PHONE_COUNT FROM CALL_LOG_AGG WITH(NOLOCK)');
    SQL.Add('WHERE ((CTIME >= :CTIME1) AND (CTIME <= :CTIME2))');
    SQL.Add('AND (RTRIM(PID) <> '''')');
    SQL.Add('AND ((PID <> ''2007'') AND (AID NOT LIKE ''12%''))');  // Added by Joe 2015/07/24 16:33:27
    //�p�G�O���έp�A���[�W�ɶ��L�o����
    if UseTimeFilter then
    begin
      AddWhere('(CAST(CONVERT(varchar(8), ITIME, 14) AS TIME) >= ''0:0:0'')');
      AddWhere('(CAST(CONVERT(varchar(8), ITIME, 14) AS TIME) <= ''23:59:59'')');
    end;
    //-------------------------------------------------------------------------
    if AAnswerOnly then
      SQL.Add('AND (SCODE = 1 )');

    SQL.Add('GROUP BY DATEPART(year, ITIME), DATEPART(month, ITIME), DATEPART(day, ITIME)');

    Params.ParamValues['CTIME1'] := ABeginTime;
    Params.ParamValues['CTIME2'] := AEndTime;
    //-------------------------------------------------------------------------
    Open;
  end;
end;

procedure TdmAcdSummary.InitExecute;
begin
  InitData;
end;

procedure TdmAcdSummary.BeginExecute;
begin
  //nothing to do now
end;

procedure TdmAcdSummary.EndExecute;
begin
  //nothing to do now
end;

procedure TdmAcdSummary.GetData_PhoneIn(ASiteId: string; ABeginTime, AEndTime: TDateTime);
begin
  with qrGetData_Src do
  begin
    if Active then Close;

    with SQL do
    begin
      Clear;
      Add('SELECT');
      Add('T3.GUID, T1.GUID AS RPHE_GUID,');
      Add('T3.IPHE001, T1.RPHE001, T3.IPHE004, T3.IPHE005, T3.IPHE003,');
      Add('T1.RPHE003, T4.SALE002, T4.SALE003, T1.RPHE005, T1.RPHE006,');
      Add('T3.IPHE016, T3.IPHE017, T3.IPHE012, T3.IPHE008, T3.IPHE019,');
      Add('T1.RPHE011, T5.CLAS004, T6.CLAS004 AS CALL_KIND, T4.SALE024,');
      Add('T7.CUT1002, T7.FLAG_SW, T7.FLAG_HRS, T8.TK,');
      Add('T8.IPH2001, T8.IPH2002, ISNULL(T8.IPH2003, 0) AS IPH2003, T8.IPH2004');
      Add('FROM WICSIPHE T3 WITH(NOLOCK)');
      Add('LEFT JOIN WICSIPH2 T8 WITH(NOLOCK) ON T8.GUID = T3.GUID');
      Add('LEFT JOIN WICSRPHE T1 WITH(NOLOCK) ON T1.GUID = T8.IPH2004');
      Add('LEFT JOIN WICSSALE T4 WITH(NOLOCK) ON T1.RPHE003 = T4.SALE001');
      Add('LEFT JOIN WICSCLAS T5 WITH(NOLOCK) ON T5.CLAS002 = IPHE012 AND T5.CLAS001 = ''10''');
      Add('LEFT JOIN WICSCLAS T6 WITH(NOLOCK) ON T6.CLAS002 = IPHE019 AND T6.CLAS001 = ''01''');
      Add('LEFT JOIN WICSCUT1 T7 WITH(NOLOCK) ON T7.CUT1001 = IPHE005');
      Add('WHERE');
      Add('(T3.IPHE004 >= :BDATE AND T3.IPHE004 <= :EDATE)');
      //�p�G�O���έp�A���[�W�ɶ��L�o����
      if UseTimeFilter then
      begin
        AddWhere('(CAST(CONVERT(varchar(8), IPHE004, 14) AS TIME) >= ''0:0:0'')');
        AddWhere('(CAST(CONVERT(varchar(8), IPHE004, 14) AS TIME) <= ''23:59:59'')');
      end;
    end;

    ParamByName('BDATE').AsDateTime := ABeginTime;
    ParamByName('EDATE').AsDateTime := AEndTime;
    //-------------------------------------------------------------------------
    try
      Open;
    except
      on E: Exception do
        Log(Format('GetData_PhoneIn() error = %s', [E.Message]));
    end;
  end;
end;

procedure TdmAcdSummary.GetData_PhoneOut(ASiteId: string; ABeginTime, AEndTime: TDateTime);
begin
  with qrGetData_PhoneOut do
  begin
    if Active then Close;

    with SQL do
    begin
      Clear;
      Add('SELECT');
      Add('RPHE001, RPHE003, SALE002, DEPT001, DEPT002, RPHE005, RPHE011, IPHE005, CUT1002, IPH2001');
      Add('FROM WICSRPHE R WITH(NOLOCK)');
      Add('LEFT JOIN WICSRSCE B WITH(NOLOCK) ON R.GUID = RSCE001');
      Add('LEFT JOIN WICSIPHE A WITH(NOLOCK) ON RSCE003 = A.GUID');
      Add('LEFT JOIN WICSIPH2 D WITH(NOLOCK) ON D.GUID = A.GUID');
      Add('LEFT JOIN WICSSALE S WITH(NOLOCK) ON SALE001 = RPHE003');
      Add('LEFT JOIN WICSCUT1 C WITH(NOLOCK) ON IPHE005 = CUT1001');
      Add('LEFT JOIN WICSDEPT T WITH(NOLOCK) ON DEPT001 = SALE003');
      Add('WHERE (RPHE005 >= :RPHE005B AND RPHE005 <= :RPHE005E)');
      //�p�G�O���έp�A���[�W�ɶ��L�o����
      if UseTimeFilter then
      begin
        AddWhere('(CAST(CONVERT(varchar(8), RPHE005, 14) AS TIME) >= ''0:0:0'')');
        AddWhere('(CAST(CONVERT(varchar(8), RPHE005, 14) AS TIME) <= ''23:59:59'')');
      end;
      //-----------------------------------------------------------------------
      Add('GROUP BY RPHE001, RPHE003, SALE002, DEPT001, DEPT002, RPHE005, RPHE011, IPHE005, CUT1002, IPH2001');
      Add('ORDER BY RPHE001, IPH2001');
    end;

    ParamByName('RPHE005B').AsDateTime := ABeginTime;
    ParamByName('RPHE005E').AsDateTime := AEndTime;
    //-------------------------------------------------------------------------
    Open;
  end;
end;

procedure TdmAcdSummary.CleanData_PhoneIn(ASiteId: string; ADataSet: TDataSet);
var
  aBranchList: string;
  aIPHE001, aRPHE001: Integer;
  aFd_IPHE001, aFd_RPHE001: TIntegerField;
  aFd_IPHE004: TDateTimeField;
  aFd_IPHE003, aFd_IPHE012, aFd_SALE003: TStringField;
  aTime, aBegTime, aEndTime: TDateTime;
begin
  aBranchList := ASiteId;
  // Added by Joe 2016/07/02 10:49:31
  aBegTime := EncodeTime(0, 0, 0, 0);
  aEndTime := EncodeTime(23, 59, 59, 0);
  //---------------------------------------------------------------------------
  if (aBranchList = SITE_ID_Taoyuan) then
    aBranchList := aBranchList + ';' + BRANCH_ID_Hsinchu
  else if (aBranchList = SITE_ID_Tainan) then
    aBranchList := aBranchList + ';' + BRANCH_ID_Kaohsiung;

  aIPHE001 := -1;
  aRPHE001 := -1;
  //�u�d�U�����ӹq���Ĥ@���^�q����
  with ADataSet do
  begin
    First;
    aFd_RPHE001 := FindField('RPHE001') as TIntegerField;
    aFd_IPHE001 := FindField('IPHE001') as TIntegerField;
    aFd_IPHE003 := FindField('IPHE003') as TStringField;
    aFd_IPHE004 := FindField('IPHE004') as TDateTimeField;
    aFd_IPHE012 := FindField('IPHE012') as TStringField;
    aFd_SALE003 := FindField('SALE003') as TStringField;

    while not Eof do
    begin
      // Added by Joe 2016/07/05 11:53:59
      // �ˬd��ƬO�_���b���Ī��ɶ��_���϶�
      aTime := TimeOf(aFd_IPHE004.AsDateTime);

      if (aTime < aBegTime) or (aTime > aEndTime) then
      begin
        Delete;
        Continue;
      end;
      //-----------------------------------------------------------------------
      //�D���w�����q����ơA���έp
      if (Pos(Copy(aFd_IPHE003.AsString, 1, 2), aBranchList) = 0) and
         (Trim(aFd_SALE003.AsString) <> '') and
         (Pos(Copy(aFd_SALE003.AsString, 1, 2), aBranchList) = 0) then
      begin
        Next;
        Continue;
      end;
      // Added by Joe 2019/04/11 10:12:42
      // ��w�����Ƥ��C�J�έp
      if IsNationalHoliday(DateOf(aFd_IPHE004.AsDateTime)) then
      begin
        Delete;
        Continue;
      end;
      //�u�O�d�C�q�ӹq���Ĥ@���^�q��T
      if (aIPHE001 <> aFd_IPHE001.AsInteger) then
      begin
        aIPHE001 := aFd_IPHE001.AsInteger;
        aRPHE001 := aFd_RPHE001.AsInteger;
      end
      else if (aFd_RPHE001.AsInteger > aRPHE001) then
      begin
        Delete;
        Continue;
      end;
      //�p��P�Ҧ��Ӧ^�q��������T===========================================
      CalcOneData_IPH2;
      //�簣���O�V�m�����Ϋ��w�^�q����������
      if (aFd_IPHE004.AsDateTime < FCalcInBegTime) or (aFd_IPHE004.AsDateTime > FCalcInEndTime) or
         ((Pos(aFd_IPHE003.AsString + ';', TE_DEPT_LIST) = 0) and (Pos(aFd_SALE003.AsString + ';', TE_DEPT_LIST) = 0)) then
      begin
        Delete;
        Continue;
      end;
      //�ˬd�t�ΧO
      if Pos(aFd_IPHE012.AsString, '10^20^22^30^40^50^61^70') = 0 then
      begin
        Delete;
        Continue;
      end;
      Next;
    end;
  end;
end;

procedure TdmAcdSummary.CleanData_PhoneOut(ADataSet: TDataSet);
var
  aRPHE001: Integer;
  aDEPT001: string;
  aFd_RPHE001: TIntegerField;
  aFd_DEPT001: TStringField;
  aFd_Except: TBooleanField;
begin
  with ADataSet do
  begin
    aFd_RPHE001 := FindField('RPHE001') as TIntegerField;
    aFd_DEPT001 := FindField('DEPT001') as TStringField;
    aFd_Except  := FindField('Except') as TBooleanField;
    // Added by Joe Lee 2017/11/10 15:08:25
    // ���簣���O�V�m�������^�q���
    First;

    while not Eof do
    begin
      if (Pos(aFd_DEPT001.AsString + ';', TE_DEPT_LIST) = 0) then
      begin
        Delete;
        Continue;
      end
      else
        Next;
    end;
    //�u�d�U�����ӹq���Ĥ@���^�q����
    First;
    aRPHE001 := -1;
    aDEPT001 := '';

    while not Eof do
    begin
      //��H�@�q�^�q�Ӧ^�h�Өӹq�ɡA�Y�ӹq���V�X�F�X���P�D�X���t�ΡA�h�u���ĭp�X������
      if (aDEPT001 <> aFd_DEPT001.AsString) then
      begin
        aRPHE001 := aFd_RPHE001.AsInteger;
        aDEPT001 := aFd_DEPT001.AsString;
      end
      else if (aFd_RPHE001.AsInteger = aRPHE001) then
      begin
        //���ƪ����P�X�������ӹq�A���p�J�έp
        Edit;
        aFd_Except.AsBoolean := True;
        Post;
      end
      else
        aRPHE001 := aFd_RPHE001.AsInteger;

      Next;
    end;
  end;
end;

procedure TdmAcdSummary.PrepareData_PhoneIn(ASiteId: string; ABeginTime, AEndTime: TDateTime);
begin
  //���o�ӹq���
	Log(Format('Ū���ӹq���(%s)', [ASiteId]));
  GetData_PhoneIn(ASiteId, ABeginTime, AEndTime);
  Log(Format('���`�ӹq���(%s)', [ASiteId]));
  mdAcdSrc.LoadFromDataSet(qrGetData_Src, [mtcpoAppend]);
  Log(Format('�ƧǨӹq���(%s)', [ASiteId]));
  mdAcdSrc.SortOn('IPHE001;RPHE001', []);
  Log('�M�z�ӹq���');
  CleanData_PhoneIn(ASiteId, mdAcdSrc);
  qrGetData_Src.Close;
end;

procedure TdmAcdSummary.PrepareData_PhoneOut(ASiteId: string; ABeginTime, AEndTime: TDateTime);
begin
  //���o�^�q���
  Log('Ū���^�q���');
  GetData_PhoneOut(ASiteId, ABeginTime, AEndTime);
  mdPhoneOutSrc.LoadFromDataSet(qrGetData_PhoneOut, [mtcpoAppend]);
  qrGetData_PhoneOut.Close;
end;

procedure TdmAcdSummary.PrepareData_PrecedingAcdTotal;
begin
  with qrPrecedingAcdTotal do
  begin
    ParamByName('IPH5002B').Value := StartOfAMonth(YearOf(Now), 1);
    ParamByName('IPH5002E').Value := StartOfAMonth(YearOf(Now), MonthOf(Now));
    ParamByName('PREV_IPH5002B').Value := StartOfAMonth(YearOf(Now)-1, 1);
    ParamByName('PREV_IPH5002E').Value := StartOfAMonth(YearOf(Now)-1, MonthOf(Now));
    if Active then Refresh else Open;
    Locate('IPH5003;_YEAR_', VarArrayOf(['WSTP2000', 2018]), []);
  end;
end;

procedure TdmAcdSummary.CalcData_PhoneIn(ASiteId: string);
var
  aBranchList, aDeptList, aCallKind: string;
  aPhoneDate: TDateTime;
  aValidAcdPhone: Boolean;
begin
  aDeptList := GetDeptFilterStr(ASiteId);
  mdAcdTeDaily.DisableControls;
  mdAcdSrc.DisableControls;
  aBranchList := ASiteId;

  if (aBranchList = SITE_ID_Taoyuan) then
    aBranchList := aBranchList + ';' + BRANCH_ID_Hsinchu
  else if (aBranchList = SITE_ID_Tainan) then
    aBranchList := aBranchList + ';' + BRANCH_ID_Kaohsiung;

  try
    with mdAcdSrc do
    begin
      if not Active then Exit;
      First;

      while not Eof do
      begin
        //�D���w�����q����ơA���έp
        if (Pos(Copy(mdAcdSrcSALE003.AsString, 1, 2), aBranchList) = 0) and
           (Pos(Copy(mdAcdSrcIPHE003.AsString, 1, 2), aBranchList) = 0) then
        begin
          Next;
          Continue;
        end;
        //---------------------------------------------------------------------
        aPhoneDate := DateOf(mdAcdSrcIPHE004.AsDateTime);
        aCallKind := mdAcdSrcIPHE019.AsString;
        //���q�ܻP�^�q�ܬO�P�@�ӤH�A�B�Ӧ^�q���ġA�o�ˤ~��O�@�q����ACD�ӹq�B�z
        aValidAcdPhone := (mdAcdSrcRPHE003.AsString = mdAcdSrcIPHE008.AsString) and (mdAcdSrcRPHE011.AsString = 'N');

        if mdAcdSrcIPHE019.AsString = CALLIN_KIND_ACD then              //ACD�q��
        begin
          if aValidAcdPhone and mdAcdSrcAnswerInTime.AsBoolean then     //ACD���ĳq��
            MarkSrcPhoneAsValid;
        end
        else if mdAcdSrcIPHE019.AsString = CALLIN_KIND_ACD_ASSIGN then  //ACD���w�^�q�q��
        begin
          if aValidAcdPhone and mdAcdSrcAnswerInTime.AsBoolean then     //ACD���w�^�q���ĳq��
            MarkSrcPhoneAsValid
          else
          begin
            MarkSrcPhoneAsExcept('�O�ɪ�ACD���w�^�q');
            AddOneExceptPhone(ASiteId, mdAcdSrcCLAS004.AsString, aPhoneDate, 2);
          end;
        end;
        //�֭p��~�B��TE�d���`��
        CalcOneData_SiteDaily_Count(ASiteId);
        CalcOneData_Te_PhoneCount(ASiteId, aPhoneDate);
        CalcOneData_SwDaily(ASiteId);
        //�p��ACD��������T====================================================
        //�D���w�V�m�������ӹq�A�qACD�ӹq�`�Ƥ�����
        if (aCallKind = CALLIN_KIND_ACD) or (aCallKind = CALLIN_KIND_ACD_ASSIGN) or (aCallKind = CALLIN_KIND_TE) then
        begin
          if (Trim(mdAcdSrcSALE003.AsString) <> '') and (Pos(mdAcdSrcSALE003.AsString, aDeptList) = 0) then
          begin
            MarkSrcPhoneAsExcept('�D�V�m�������ӹq');
            AddOneExceptPhone(ASiteId, mdAcdSrcCLAS004.AsString, aPhoneDate, 1);
            Next;
            Continue;
          end
          else if (Pos(mdAcdSrcIPHE003.AsString, aDeptList) = 0) then
          begin
            //�^�q�����O�V�m�����A���ӹq�w���k�ݰV�m�����A���ӹq�w�g�L��L�B�z
            if mdAcdSrcRPHE011.AsString <> 'N' then
            begin
              MarkSrcPhoneAsExcept('�ӹq�w��D�V�m�����A�B���L�Ħ^�q');
              AddOneExceptPhone(ASiteId, mdAcdSrcCLAS004.AsString, aPhoneDate, 1);
              Next;
              Continue;
            end;
          end;
        end;
        (*
        //���^���ӹq�A���C�J�ӤH�Z�Ĳέp
        if mdAcdSrcRPHE001.IsNull then
        begin
          Next;
          Continue;
        end;
        *)
        CalcOneData_Te_Daily(ASiteId, aPhoneDate);
        Next;
      end;
    end;
  finally
    mdAcdTeDaily.EnableControls;
    mdAcdSrc.EnableControls;
  end;
end;

procedure TdmAcdSummary.CalcData_PhoneOut(ASiteId: string);
var
  //aBranchList, aDeptList, aTe, aTeName, aDeptId: string;
  aBranchList, aTe, aTeName, aDeptId: string;
  aPhoneDate: TDateTime;
begin
  //aDeptList := GetDeptFilterStr(ASiteId);
  mdAcdTeDaily.DisableControls;
  mdPhoneOutSrc.DisableControls;
  aBranchList := ASiteId;

  if (aBranchList = SITE_ID_Taoyuan) then
    aBranchList := aBranchList + ';' + BRANCH_ID_Hsinchu
  else if (aBranchList = SITE_ID_Tainan) then
    aBranchList := aBranchList + ';' + BRANCH_ID_Kaohsiung;

  try
    with mdPhoneOutSrc do
    begin
      if not Active then Exit;
      First;

      while not Eof do
      begin
        //�D���w�����q�εL�Ī���ơA���έp
        if (Pos(Copy(mdPhoneOutSrcDEPT001.AsString, 1, 2), aBranchList) = 0) or mdPhoneOutSrcInvalid.AsBoolean then
        begin
          Next;
          Continue;
        end;
        //�έp�ӤH�`�^�q�� Added by Joe 2015/08/13 17:07:46
        aPhoneDate := DateOf(mdPhoneOutSrcRPHE005.AsDateTime);
        aTe := mdPhoneOutSrcRPHE003.AsString;
        aTeName := mdPhoneOutSrcSALE002.AsString;
        aDeptId := mdPhoneOutSrcDEPT001.AsString;

        if CheckTeDailyRec(mdAcdTeDaily, ASiteId, aTe, aTeName, aDeptId, aPhoneDate) then
        begin
          mdAcdTeDaily.Edit;
          
          mdAcdTeDailyCallout_Total.AsInteger := mdAcdTeDailyCallout_Total.AsInteger + 1;
          //�X���^�q
          if (mdPhoneOutSrcIPH2001.AsString = 'S') or (mdPhoneOutSrcIPH2001.AsString = 'N') then
            mdAcdTeDailyPhoneOut_Count_C.AsInteger := mdAcdTeDailyPhoneOut_Count_C.AsInteger + 1
          else  //�D�X���^�q
            mdAcdTeDailyPhoneOut_Count_NC.AsInteger := mdAcdTeDailyPhoneOut_Count_NC.AsInteger + 1;

          mdAcdTeDaily.Post;
        end;
        //---------------------------------------------------------------------
        if CheckSiteDailyRec(ASiteId, aPhoneDate) then
        begin
          mdAcdSiteDaily.Edit;
          //�X���^�q
          if (mdPhoneOutSrcIPH2001.AsString = 'S') or (mdPhoneOutSrcIPH2001.AsString = 'N') then
            mdAcdSiteDailyPhoneOut_Count_C.AsInteger := mdAcdSiteDailyPhoneOut_Count_C.AsInteger + 1
          else  //�D�X���^�q
            mdAcdSiteDailyPhoneOut_Count_NC.AsInteger := mdAcdSiteDailyPhoneOut_Count_NC.AsInteger + 1;

          mdAcdSiteDaily.Post;
        end;

        Next;
      end;
    end;
  finally
    mdAcdTeDaily.EnableControls;
    mdPhoneOutSrc.EnableControls;
  end;
end;

procedure TdmAcdSummary.AddOneExceptPhone(ASiteId, ASw: string; ADate: TDateTime; AKind: Integer);
begin
  if not CheckSiteDailyRec(ASiteId, ADate) then
    Exit;

  with mdAcdSiteDaily do
  begin
    Edit;

    if AKind = 1 then
      mdAcdSiteDailyACD_InvalidIn_Total.AsInteger := mdAcdSiteDailyACD_InvalidIn_Total.AsInteger + 1
    else if AKind = 2 then
      mdAcdSiteDailyACD_Assign_Invalid.AsInteger := mdAcdSiteDailyACD_Assign_Invalid.AsInteger + 1;

    Post;
  end;

  CheckSwDailyRec(ASiteId, ASw, ADate);

  with mdAcdSwDaily do
  begin
    Edit;

    if AKind = 1 then
      mdAcdSwDailyACD_InvalidIn_Total.AsInteger := mdAcdSwDailyACD_InvalidIn_Total.AsInteger + 1
    else if AKind = 2 then
      mdAcdSwDailyACD_Assign_Invalid.AsInteger := mdAcdSwDailyACD_Assign_Invalid.AsInteger + 1;

    Post;
  end;
end;

function TdmAcdSummary.GetDeptFilterStr(ASiteId: string): string;
begin
  if (ASiteId = SITE_ID_Taipei) then
    Result := Format('%s;%s;%s;%s', [TE_DEPT_021, TE_DEPT_022, TE_DEPT_023, TE_DEPT_026])
  else if (ASiteId = SITE_ID_Taoyuan) then
    Result := Format('%s;%s', [TE_DEPT_052, TE_DEPT_062])
  else if (ASiteId = SITE_ID_Taichung) then
    Result := Format('%s;%s', [TE_DEPT_075, TE_DEPT_076])
  else if (ASiteId = SITE_ID_Tainan) then
    Result := Format('%s;%s', [TE_DEPT_082, TE_DEPT_092])
  else
    Result := '';
end;

procedure TdmAcdSummary.GetPrecedingAcdTotal(IPH5003: string; var ATotal, APrevTotal: Integer);
begin
  ATotal := 0;
  APrevTotal := 0;

  if not JcDataSetIsValid(qrPrecedingAcdTotal) then
    Exit;

  with qrPrecedingAcdTotal do
  begin
    if Locate('IPH5003;_YEAR_', VarArrayOf([IPH5003, YearOf(Now)]), []) then
      ATotal := FieldByName('_IPH5007_SUM_').AsInteger;

    if Locate('IPH5003;_YEAR_', VarArrayOf([IPH5003, YearOf(Now)-1]), []) then
      APrevTotal := FieldByName('_IPH5007_SUM_').AsInteger;
  end;
end;

procedure TdmAcdSummary.ACD_CalcData_Site(ASiteId: string; AAnswerOnly: Boolean);
var
  aDate: TDateTime;
begin
  with qryGetAcdInfo do
  begin
    if (not Active) or IsEmpty then Exit;
    First;

    while not Eof do
    begin
      aDate := EncodeDate(FieldByName('_YEAR_').AsInteger, FieldByName('_MONTH_').AsInteger, FieldByName('_DAY_').AsInteger);

      if CheckSiteDailyRec(ASiteId, aDate) then
      begin
        mdAcdSiteDaily.Edit;
        //---------------------------------------------------------------------
        if not AAnswerOnly then
        begin
          //ACD���e�`��
          mdAcdSiteDailyACD_Total.AsInteger := FieldByName('PHONE_COUNT').AsInteger;
          //ACD���Ĭ��e��
          mdAcdSiteDailyACD_ValidIn_Total.AsInteger := mdAcdSiteDailyACD_Total.AsInteger
            - mdAcdSiteDailyACD_InvalidIn_Total.AsInteger - mdAcdSiteDailyACD_Assign_Invalid.AsInteger;
          CodeSite.SendFmtMsg(AnsiToUtf8('ACD���Ĭ��e��[%s] = %d-%d-%d'), [ASiteId, mdAcdSiteDailyACD_Total.AsInteger,
            mdAcdSiteDailyACD_InvalidIn_Total.AsInteger, mdAcdSiteDailyACD_Assign_Invalid.AsInteger]);
          //ACD��ť�v
          if mdAcdSiteDailyACD_ValidIn_Total.AsInteger = 0 then
            mdAcdSiteDailyACD_Score.AsFloat := 0
          else
            mdAcdSiteDailyACD_Score.AsFloat := 100 * mdAcdSiteDailyACD_ValidAns_Total.AsInteger
              / mdAcdSiteDailyACD_ValidIn_Total.AsInteger;

//          if mdAcdSiteDailyACD_Score.AsFloat > 100 then
//            mdAcdSiteDailyACD_Score.AsFloat := 100;
        end
        else  //ACD������ť��
          mdAcdSiteDailyACD_Ans_Total.AsInteger := FieldByName('PHONE_COUNT').AsInteger;;
        //---------------------------------------------------------------------
        mdAcdSiteDaily.Post;
      end;

      Next;
    end;
  end;
end;

procedure TdmAcdSummary.mdAcdSrcCalcFields(DataSet: TDataSet);
begin
  inherited;
  mdAcdSrcAnswerInTime.AsBoolean :=
    SecondsBetween(mdAcdSrcIPHE004.AsDateTime, mdAcdSrcRPHE005.AsDateTime) <= ACD_VALID_SECONDS;
  mdAcdSrcRPHE011_REV.AsBoolean := (mdAcdSrcRPHE011.AsString = 'N');
end;

procedure TdmAcdSummary.mdAcdSiteDailyCalcFields(DataSet: TDataSet);
begin
  inherited;
  DataSet['SiteName'] := dmReport.GetBranchOfDept(VarToStr(DataSet['SiteId']), True);
end;

procedure TdmAcdSummary.mdAcdTeDailyCalcFields(DataSet: TDataSet);
begin
  inherited;
  DataSet['DeptName'] := dmReport.Get_Dept_Name(VarToStr(DataSet['DeptId']));
end;

procedure TdmAcdSummary.InitReportConn(ASiteId: string);
begin
  with dmReport do
  begin
  	SetUniConn_TCRM(connReport, GetSiteIp(ASiteId));
  end;
end;

procedure TdmAcdSummary.CalcData_Site_Daily(ASiteId: string);
var
  aBranchList, aDeptList: string;
begin
  mdAcdTeDaily.DisableControls;
  aDeptList := GetDeptFilterStr(ASiteId);
  aBranchList := ASiteId;

  if (aBranchList = SITE_ID_Taoyuan) then
    aBranchList := aBranchList + ';' + BRANCH_ID_Hsinchu
  else if (aBranchList = SITE_ID_Tainan) then
    aBranchList := aBranchList + ';' + BRANCH_ID_Kaohsiung;

  try
    with mdAcdTeDaily do
    begin
      if not Active then Exit;
      First;

      while not Eof do
      begin
        //�D���w�����q����ơA�άO�D���w�V�m�������ӹq�A���έp
        if (Pos(Copy(mdAcdTeDailyDeptId.AsString, 1, 2), aBranchList) = 0) or
           (Pos(mdAcdTeDailyDeptId.AsString, aDeptList) = 0) then
        begin
          Next;
          Continue;
        end;

        if CheckSiteDailyRec(ASiteId, mdAcdTeDailyPhoneDate.AsDateTime) then
        begin
          mdAcdSiteDaily.Edit;
          //=====================================================================
          //ACD�B�z��
          mdAcdSiteDailyACD_ValidAns_Total.AsInteger :=
            mdAcdSiteDailyACD_ValidAns_Total.AsInteger + mdAcdTeDailyACD_ValidAns_Total.AsInteger;
          //�D���e�^�q��
          mdAcdSiteDailyNotAcdCallout_Count.AsInteger :=
            mdAcdSiteDailyNotAcdCallout_Count.AsInteger + mdAcdTeDailyNotAcdCallout_Count.AsInteger;
          //�`�^�q�q��
          mdAcdSiteDailyCallout_Total.AsInteger :=
            mdAcdSiteDailyCallout_Total.AsInteger + mdAcdTeDailyCallout_Total.AsInteger;
          //�Ⱦ��H��
          mdAcdSiteDailyDays.AsFloat := mdAcdSiteDailyDays.AsFloat + mdAcdTeDailyDays.AsFloat;
          //�X���O�ɲv
          if mdAcdSiteDailyPhone_Count_C.AsInteger <> 0 then
            mdAcdSiteDailyTimeOut_Rate_C.AsFloat :=
              mdAcdSiteDailyTimeOut_Count_C.AsInteger / mdAcdSiteDailyPhone_Count_C.AsInteger
          else
            mdAcdSiteDailyTimeOut_Rate_C.AsFloat := 0;
          //�D�X���O�ɲv
          if mdAcdSiteDailyPhone_Count_NC.AsInteger <> 0 then
            mdAcdSiteDailyTimeOut_Rate_NC.AsFloat :=
              mdAcdSiteDailyTimeOut_Count_NC.AsInteger / mdAcdSiteDailyPhone_Count_NC.AsInteger
          else
            mdAcdSiteDailyTimeOut_Rate_C.AsFloat := 0;
          //=====================================================================
          mdAcdSiteDaily.Post;
        end;
        Next;
      end;
    end;
  finally
    mdAcdTeDaily.EnableControls;
  end;
end;

procedure TdmAcdSummary.CalcOneData_Te_Daily(ASiteId: string; ADate: TDateTime);
var
  aPhoneInTotal: Integer;
begin
  if not CheckTeDailyRec(mdAcdTeDaily) then Exit;

  mdAcdTeDaily.Edit;
  //===========================================================================
  if mdAcdSrcIPHE019.AsString = CALLIN_KIND_ACD then        //ACD�q��
  begin
    mdAcdTeDailyACD_Ans_Total.AsInteger := mdAcdTeDailyACD_Ans_Total.AsInteger + 1;

    if mdAcdSrcValid.AsBoolean then
      mdAcdTeDailyACD_Ans_Valid.AsInteger := mdAcdTeDailyACD_Ans_Valid.AsInteger + 1;
  end
  else if mdAcdSrcIPHE019.AsString = CALLIN_KIND_ACD_ASSIGN then   //ACD���w�^�q�q��
  begin
    mdAcdTeDailyACD_Assign_Total.AsInteger := mdAcdTeDailyACD_Assign_Total.AsInteger + 1;

    if mdAcdSrcValid.AsBoolean then
      mdAcdTeDailyACD_Assign_Valid.AsInteger := mdAcdTeDailyACD_Assign_Valid.AsInteger + 1;
  end
  else if mdAcdSrcIPHE019.AsString = CALLIN_KIND_TE then   //TE�d����
  begin
    if IsContracted(mdAcdSrcIPH2001.AsString) then
      mdAcdTeDailyTE_Total_C.AsInteger := mdAcdTeDailyTE_Total_C.AsInteger + 1
    else
      mdAcdTeDailyTE_Total_NC.AsInteger := mdAcdTeDailyTE_Total_NC.AsInteger + 1;
  end;
  //ACD�`�ӹq��
  mdAcdTeDailyACD_In_Total.AsInteger := mdAcdTeDailyACD_Ans_Total.AsInteger +  mdAcdTeDailyACD_Assign_Total.AsInteger;
  //ACD���Ħ^�q�`��
  mdAcdTeDailyACD_ValidAns_Total.AsInteger :=
    mdAcdTeDailyACD_Ans_Valid.AsInteger + mdAcdTeDailyACD_Assign_Valid.AsInteger;
  //ACD�Z��
  aPhoneInTotal := mdAcdTeDailyACD_Ans_Total.AsInteger + mdAcdTeDailyACD_Assign_Valid.AsInteger;

  if aPhoneInTotal = 0 then
    mdAcdTeDailyACD_Score.AsFloat := 0
  else
    mdAcdTeDailyACD_Score.AsFloat := mdAcdTeDailyACD_ValidAns_Total.AsInteger / aPhoneInTotal * 100;
  //===========================================================================
  mdAcdTeDaily.Post;
end;

procedure TdmAcdSummary.MarkSrcPhoneAsExcept(ARemark: string);
begin
  with mdAcdSrc do
  begin
    Edit;
    mdAcdSrcExcept.AsBoolean := True;
    mdAcdSrcRemark.AsString := ARemark;
    Post;
  end;
end;

procedure TdmAcdSummary.MarkSrcPhoneAsValid;
begin
  with mdAcdSrc do
  begin
    Edit;
    mdAcdSrcValid.AsBoolean := True;
    Post;
  end;
end;

function TdmAcdSummary.CheckTeDailyRec(ADataSet: TDataSet): Boolean;
var
  A: Variant;
  aIsTeDept: Boolean;
  aRPHE003, aSALE002, aSALE003, aSiteName: string;
begin
  Result := False;

  //�D�V�m�����ӹq�A���έp Added by Joe 2016/07/04 10:33:41
  if Pos(mdAcdSrcIPHE003.AsString, TE_DEPT_LIST) = 0 then
    Exit;
  //---------------------------------------------------------------------------

  with ADataSet do
  begin
    if not Active then Exit;
    aIsTeDept := CheckTeRec(mdAcdSrcSALE003.AsString);
    aSiteName := dmReport.GetBranchOfDept(mdAcdSrcSALE003.AsString);
    // Added by Joe 2016/07/04 13:50:51
    // �p�G�|�����^�q�H�A����ӹq�k�ݳ������ݪ����
    if (aSiteName = '') then
      aSiteName := dmReport.GetBranchOfDept(mdAcdSrcIPHE003.AsString);
    //-------------------------------------------------------------------------
    if aIsTeDept then
    begin
      aRPHE003 := mdAcdSrcRPHE003.AsString;
      aSALE002 := mdAcdSrcSALE002.AsString;
      aSALE003 := mdAcdSrcSALE003.AsString;
    end
    else
    begin
      aRPHE003 := '0-' + STR_ITEM_MISC;
      aSALE002 := '0-' + STR_ITEM_MISC;
      aSALE003 := '0-' + STR_ITEM_MISC;
    end;

    A := VarArrayOf([aRPHE003, mdAcdSrcIPHE004.AsDateTime, aSiteName]);

    if not Locate('EmpId;PhoneDate;SiteName', A, []) then
    begin
      Append;
      FieldByName('EmpId').AsString := aRPHE003;
      FieldByName('EmpName').AsString := aSALE002;
      FieldByName('DeptId').AsString := aSALE003;
      FieldByName('SiteName').AsString := aSiteName;
      FieldByName('PhoneDate').AsDateTime := mdAcdSrcIPHE004.AsDateTime;
      FieldByName('PhoneDateDesc').AsString := dmReport.GetDateDispTextDOW(mdAcdSrcIPHE004.AsDateTime);
      // Added by Joe 2016/07/02 11:57:24
      FieldByName('PhoneYear').AsInteger := YearOf(mdAcdSrcIPHE004.AsDateTime);
      FieldByName('PhoneMonth').AsInteger := MonthOf(mdAcdSrcIPHE004.AsDateTime);
      // Added by Joe 2017/02/23 14:26:42
      FieldByName('ACD_DailyReqCount').AsInteger := mdAcdSrcSALE024.AsInteger;
      //-----------------------------------------------------------------------
      Post;
    end;
    Result := True;
  end;
end;

function TdmAcdSummary.CheckSiteDailyRec(ASiteId: string; ADate: TDateTime): Boolean;
var
  A: Variant;
begin
  Result := False;
  // Added by Joe 2019/04/11 10:38:48
  if IsNationalHoliday(DateOf(ADate)) then
    Exit;
  //---------------------------------------------------------------------------
  A := VarArrayOf([ASiteId, ADate]);

  try
    if not mdAcdSiteDaily.Locate('SiteId;PhoneDate', A, []) then
    begin
      mdAcdSiteDaily.Append;
      mdAcdSiteDailySiteId.AsString := ASiteId;
      mdAcdSiteDailyPhoneDate.AsDateTime := ADate;
      mdAcdSiteDailyPhoneDateDesc.AsString := dmReport.GetDateDispTextDOW(ADate);
      // Added by Joe 2016/07/02 11:57:24
      mdAcdSiteDailyPhoneYear.AsInteger := YearOf(ADate);
      mdAcdSiteDailyPhoneMonth.AsInteger := MonthOf(ADate);
      //-----------------------------------------------------------------------
      mdAcdSiteDaily.Post;
    end;
    Result := True;
  except
    Result := False;
  end;
end;

procedure TdmAcdSummary.CalcOneData_SiteDaily_Count(ASiteId: string);
var
  aPhoneDate: TDateTime;
begin
  //�D�V�m�����ӹq�A���έp Added by Joe 2015/09/01 09:26:09
  if Pos(mdAcdSrcIPHE003.AsString, TE_DEPT_LIST) = 0 then
    Exit;
  //---------------------------------------------------------------------------
  aPhoneDate := DateOf(mdAcdSrcIPHE004.AsDateTime);

  if CheckSiteDailyRec(ASiteId, aPhoneDate) then
  begin
    mdAcdSiteDaily.Edit;
    //�֭pTE�d����
    if (mdAcdSrcIPHE019.AsString = CALLIN_KIND_TE) then
    begin
      if IsContracted(mdAcdSrcIPH2001.AsString) then
        mdAcdSiteDailyTE_Total_C.AsInteger := mdAcdSiteDailyTE_Total_C.AsInteger + 1
      else
        mdAcdSiteDailyTE_Total_NC.AsInteger := mdAcdSiteDailyTE_Total_NC.AsInteger + 1;
    end;
    //�֭p�`�ӹq�ƻP�O�ɼ�
    if IsContracted(mdAcdSrcIPH2001.AsString) then
    begin
      //�ӹq��
      mdAcdSiteDailyPhone_Count_C.AsInteger := mdAcdSiteDailyPhone_Count_C.AsInteger + 1;
      //CodeSite.SendFmtMsg(AnsiToUtf8('�ӹq��+1, IPH1001 = %d'), [mdAcdSrcIPHE001.AsInteger]);
      //�O��
      if mdAcdSrcIPH2003.AsBoolean then
      begin
        mdAcdSiteDailyTimeOut_Count_C.AsInteger := mdAcdSiteDailyTimeOut_Count_C.AsInteger + 1;
        Log(Format('! �O�ɨӹq(%s) = %d', [ASiteId, mdAcdSrcIPHE001.AsInteger]));
      end;
      //���^�q��
      //�p�G�OTE�פJ�ӥB�w�W�L�W�Z�ɶ�,�@�ߤ��@���O�ɨӭp��
      if not Is_OffDuty_TE then // Added by Joe 2017/11/09 16:16:46
      begin
        // Modified by Administrator 2017/11/14 16:14:00
        //if (mdAcdSrcRPHE001.AsInteger = 0) and mdAcdSrcIPHE017.IsNull then
        if (mdAcdSrcRPHE001.AsInteger = 0) or
           (DateOf(mdAcdSrcIPHE016.AsDateTime) > DateOf(mdAcdSrcIPHE004.AsDateTime)) then
        begin
          mdAcdSiteDailyNoAns_Count_C.AsInteger := mdAcdSiteDailyNoAns_Count_C.AsInteger + 1;
          Log(Format('! ���^�ӹq(%s) = %d', [ASiteId, mdAcdSrcIPHE001.AsInteger]));          
        end;
      end
    end
    else
    begin
      //�ӹq��
      mdAcdSiteDailyPhone_Count_NC.AsInteger := mdAcdSiteDailyPhone_Count_NC.AsInteger + 1;
      //�O��
      if mdAcdSrcIPH2003.AsBoolean and mdAcdSrcIPHE017.IsNull then
        mdAcdSiteDailyTimeOut_Count_NC.AsInteger := mdAcdSiteDailyTimeOut_Count_NC.AsInteger + 1;
      //���^�q��
      // Modified by Administrator 2017/11/14 16:14:00
      //if (mdAcdSrcRPHE001.AsInteger = 0) and mdAcdSrcIPHE017.IsNull then
      if (mdAcdSrcRPHE001.AsInteger = 0) or
         (DateOf(mdAcdSrcIPHE016.AsDateTime) > DateOf(mdAcdSrcIPHE004.AsDateTime)) then
        mdAcdSiteDailyNoAns_Count_NC.AsInteger := mdAcdSiteDailyNoAns_Count_NC.AsInteger + 1;
    end;
    //---------------------------------------------------------------------------
    mdAcdSiteDaily.CheckBrowseMode;
  end;
end;

function TdmAcdSummary.CheckTeDailyRec(ADataSet: TDataSet; ASiteId, AEmpId, AEmpName, ADeptId: string; ADate: TDateTime): Boolean;
var
  A: Variant;
begin
  Result := False;

  with ADataSet do
  begin
    if not Active then Exit;
    if not CheckTeRec(ADeptId) then Exit;
    A := VarArrayOf([AEmpId, ADate, ADeptId]);

    if not Locate('EmpId;PhoneDate;DeptId', A, []) then
    begin
      Append;
      FieldByName('EmpId').AsString := AEmpId;
      FieldByName('EmpName').AsString := AEmpName;
      FieldByName('DeptId').AsString := ADeptId;
      FieldByName('SiteName').AsString := dmReport.GetBranchOfDept(ASiteID, True);
      FieldByName('PhoneDate').AsDateTime := ADate;
      Post;
    end;
    Result := True;
  end;
end;

procedure TdmAcdSummary.CalcOneData_IPH2;
const
  INI_SYS_CALLINTIMEOUT = 60;
  INI_SYS_NONEAGREETIMEOUT = 150;
var
  aIPH2001, aIPH2004: string;
  aIPH2002, aChangeCount: Integer;
  aIPH2003, aReply: Boolean;
  aReplyTime: TDateTime;

  function InsertIPH2: Boolean;
  begin
    Result := False;

    try
      with cmdInsIPH2 do
      begin
        ParamByName('GUID').AsString    := mdAcdSrcGUID.AsString;
        ParamByName('IPH2001').AsString := mdAcdSrcIPH2001.AsString;
        ParamByName('IPH2002').AsInteger:= mdAcdSrcIPH2002.AsInteger;
        ParamByName('IPH2003').AsBoolean:= mdAcdSrcIPH2003.AsBoolean;
        ParamByName('IPH2004').AsString := mdAcdSrcIPH2004.AsString;
        Execute;
      end;
      Result := True;
    except

    end;
  end;

  function UpdateIPH2: Integer;
  begin
    try
      with cmdUpdIPH2 do
      begin
        ParamByName('GUID').AsString    := mdAcdSrcGUID.AsString;
        ParamByName('IPH2001').AsString := mdAcdSrcIPH2001.AsString;
        ParamByName('IPH2002').AsInteger:= mdAcdSrcIPH2002.AsInteger;
        ParamByName('IPH2003').AsBoolean:= mdAcdSrcIPH2003.AsBoolean;
        ParamByName('IPH2004').AsString := mdAcdSrcIPH2004.AsString;
        Execute;
        Result := RowsAffected;
      end;
    except
      Result := -1;
    end;
  end;
begin
  //�p�G�w�g���ȤF�A���ݭn�A���s�p��
  if (Pos(mdAcdSrcIPHE019.AsString, '10^11^20^30') > 0) and (mdAcdSrcIPH2004.AsString <> '') and
     (not FForceCalcData) then
    Exit;

  with mdAcdSrc do
  begin
    Edit;
    aChangeCount := 0;
    //�p��ӹq�ɪ��X������
    if mdAcdSrcIPHE012.AsString = '30' then
      aIPH2001 := mdAcdSrcFLAG_HRS.AsString
    else
      aIPH2001 := mdAcdSrcFLAG_SW.AsString;

    if (aIPH2001 <> mdAcdSrcIPH2001.AsString) then
    begin
      Inc(aChangeCount);
      mdAcdSrcIPH2001.AsString := aIPH2001;
    end;
    //�p��^�q�O�ɸ�T
    if mdAcdSrcIPHE019.AsString = '30' then   //�p�G�O�H�u��ť�A���ݭn�p��O�ɸ��
    begin
      Inc(aChangeCount);
      mdAcdSrcIPH2002.AsInteger := 0;
      mdAcdSrcIPH2003.AsBoolean := False;
    end
    else
    begin
      // �p�G�OTE�פJ�ӥB�w�W�L�W�Z�ɶ�,�@�ߤ��p�J�O�ɦ^�q
      if Is_OffDuty_TE then // Added by Joe 2017/11/09 16:16:46
        aReplyTime := mdAcdSrcIPHE004.AsDateTime
      else
      begin
        aReply := not mdAcdSrcRPHE001.IsNull;
        //�p�G�w�^�q�A�H�^�q�ɶ��ӭp��C�p�G�|���^�q�A�H��ɪ��ɶ��ӭp��C
        if aReply then
          aReplyTime := mdAcdSrcIPHE016.AsDateTime
        else
        begin
          //�p�G�|���^�q�A�ӬO�������סA�N�^�q�ɶ��]���ӹq�ɶ��A�Ϧ^�q����0�A���O��
          if not mdAcdSrcIPHE017.IsNull then
            aReplyTime := mdAcdSrcIPHE004.AsDateTime
          else
            aReplyTime := Now;
        end;
      end;

      if dmReport.Calc_CalloutDelay(mdAcdSrcIPHE004.AsDateTime, aReplyTime, aIPH2002) then
      begin
        if (aIPH2002 <> mdAcdSrcIPH2002.AsInteger) then
        begin
          Inc(aChangeCount);
          mdAcdSrcIPH2002.AsInteger := aIPH2002;
        end;

        if Pos(mdAcdSrcIPHE019.AsString, '10^11^20^30') > 0 then
        begin
          if (aIPH2001 = 'S') or (aIPH2001 = 'N') then
            aIPH2003 := (aIPH2002 > INI_SYS_CALLINTIMEOUT)
          else
            aIPH2003 := (aIPH2002 > INI_SYS_NONEAGREETIMEOUT);
        end
        else
          aIPH2003 := False;
        // Added by Joe 2017/11/09 17:31:29
        //if aIPH2003 then
        //  CodeSite.SendFmtMsg('�O�ɨӹq = %d', [mdAcdSrcIPHE001.AsInteger]);
        //----------------------------------------------------------------------          
        if (aIPH2003 <> mdAcdSrcIPH2003.AsBoolean) then
        begin
          Inc(aChangeCount);
          mdAcdSrcIPH2003.AsBoolean := aIPH2003;
        end;
      end;
    end;
    //�^�qGUID
    aIPH2004 := mdAcdSrcRPHE_GUID.AsString;
    // Added by Joe 2016/07/11 10:53:17
    if (aIPH2004 = '') then
      aIPH2004 := Get_IPH2004(mdAcdSrcIPHE001.AsInteger);
    //-------------------------------------------------------------------------
    if (aIPH2004 <> mdAcdSrcIPH2004.AsString) then
    begin
      Inc(aChangeCount);
      mdAcdSrcIPH2004.AsString := aIPH2004;
    end;
    //-------------------------------------------------------------------------
    CheckBrowseMode;
  end;

  //�g�J�ӹq�X�R��T
  if (aChangeCount = 0) then Exit;

  if UpdateIPH2 <= 0 then
    InsertIPH2;
end;

procedure TdmAcdSummary.mdAcdSiteDailyTimeOut_Rate_CGetText(Sender: TField; var Text: String; DisplayText: Boolean);
begin
  inherited;
  Text := Format('%.1n %%', [Sender.AsFloat * 100]);
end;

function TdmAcdSummary.CheckTeRec(ADeptId: string): Boolean;
begin
  ADeptId := Trim(ADeptId);

  if (ADeptId = '') or (Pos(ADeptId + ';', TE_DEPT_LIST) = 0) then
    Result := False
  else
    Result := True;
end;

procedure TdmAcdSummary.CalcOneData_Te_PhoneCount(ASiteId: string; ADate: TDateTime);
begin
  if CheckTeDailyRec(mdAcdTeDaily) then
  begin
    mdAcdTeDaily.Edit;
    //---------------------------------------------------------------------------
    //�֭p�`�ӹq�� Added by Joe 2015/07/31 10:20:24
    if (mdAcdSrcIPH2001.AsString = 'S') or (mdAcdSrcIPH2001.AsString = 'N') then
      mdAcdTeDailyPhone_Count_C.AsInteger := mdAcdTeDailyPhone_Count_C.AsInteger + 1
    else
      mdAcdTeDailyPhone_Count_NC.AsInteger := mdAcdTeDailyPhone_Count_NC.AsInteger + 1;
    //�O�ɳq�� Added by Joe 2015/07/31 10:05:31
    if mdAcdSrcIPH2003.AsBoolean then
    begin
      if (mdAcdSrcIPH2001.AsString = 'S') or (mdAcdSrcIPH2001.AsString = 'N') then
        mdAcdTeDailyTimeOut_Count_C.AsInteger := mdAcdTeDailyTimeOut_Count_C.AsInteger + 1
      else
        mdAcdTeDailyTimeOut_Count_NC.AsInteger := mdAcdTeDailyTimeOut_Count_NC.AsInteger + 1;
    end;
    //---------------------------------------------------------------------------
    mdAcdTeDaily.Post;
  end;
end;

function TdmAcdSummary.CheckSwDailyRec(ASiteId, ASw: string; ADate: TDateTime): Boolean;
var
  A: Variant;
begin
  Result := False;
  // Added by Joe 2019/04/11 10:38:48
  if IsNationalHoliday(DateOf(ADate)) then
    Exit;
  //---------------------------------------------------------------------------
  with mdAcdSwDaily do
  begin
    if not Active then Exit;
    A := VarArrayOf([ASiteId, ASw, ADate]);

    if not Locate('SiteId;Sw;PhoneDate', A, []) then
    begin
      Append;
      FieldByName('SiteId').AsString := ASiteId;
      FieldByName('SiteName').AsString := dmReport.GetBranchOfDept(ASiteID, True);
      FieldByName('Sw').AsString := ASw;
      FieldByName('PhoneDate').AsDateTime := ADate;
      // Added by Joe 2016/07/02 11:57:24
      FieldByName('PhoneYear').AsInteger := YearOf(ADate);
      FieldByName('PhoneMonth').AsInteger := MonthOf(ADate);
      //-----------------------------------------------------------------------
      Post;
    end;
    Result := True;
  end;
end;

procedure TdmAcdSummary.CalcOneData_SwDaily(ASiteId: string);
var
  aSw: string;
  aPhoneDate: TDateTime;
begin
  //�D�V�m�����ӹq�A���έp Added by Joe 2016/07/04 09:52:27
  if Pos(mdAcdSrcIPHE003.AsString, TE_DEPT_LIST) = 0 then
    Exit;
  //---------------------------------------------------------------------------
  aSw := Trim(mdAcdSrcCLAS004.AsString);
  aPhoneDate := DateOf(mdAcdSrcIPHE004.AsDateTime);

  if not CheckSwDailyRec(ASiteId, aSw, aPhoneDate) then
    Exit;

  mdAcdSwDaily.Edit;
  //�֭pTE�d����
  if (mdAcdSrcIPHE019.AsString = CALLIN_KIND_TE) then
  begin
    if IsContracted(mdAcdSrcIPH2001.AsString) then
      mdAcdSwDailyTE_Total_C.AsInteger := mdAcdSwDailyTE_Total_C.AsInteger + 1
    else
      mdAcdSwDailyTE_Total_NC.AsInteger := mdAcdSwDailyTE_Total_NC.AsInteger + 1          
  end;

  //�֭p�`�ӹq�ƻP�O�ɼ�
  if IsContracted(mdAcdSrcIPH2001.AsString) then
  begin
    mdAcdSwDailyPhone_Count_C.AsInteger := mdAcdSwDailyPhone_Count_C.AsInteger + 1;

    if mdAcdSrcIPH2003.AsBoolean then
      mdAcdSwDailyTimeOut_Count_C.AsInteger := mdAcdSwDailyTimeOut_Count_C.AsInteger + 1;
  end
  else
  begin
    mdAcdSwDailyPhone_Count_NC.AsInteger := mdAcdSwDailyPhone_Count_NC.AsInteger + 1;

    if mdAcdSrcIPH2003.AsBoolean then
      mdAcdSwDailyTimeOut_Count_NC.AsInteger := mdAcdSwDailyTimeOut_Count_NC.AsInteger + 1;
  end;
  //ACD�B�z��
  if mdAcdSrcValid.AsBoolean then
    mdAcdSwDailyACD_ValidAns_Total.AsInteger := mdAcdSwDailyACD_ValidAns_Total.AsInteger + 1;
  //---------------------------------------------------------------------------
  mdAcdSwDaily.CheckBrowseMode;
end;

procedure TdmAcdSummary.ACD_GetGroupCount(ASite: string; ABeginTime, AEndTime: TDateTime; AAnswerOnly: Boolean);
begin
  if not Assigned(dmReport) then
    raise Exception.Create(ERR_DMMSSQL_MISSING);

  dmReport.SetConn_TeleContact(ASite);

  with qryGetAcdInfo do
  begin
    if Active then Close;
    SQL.Clear;
    SQL.Add('SELECT CONVERT(varchar(10), ITIME, 111) AS _DATE_');
    SQL.Add(',PID ,COUNT(*) AS PHONE_COUNT');
    SQL.Add('FROM CALL_LOG_AGG WITH(NOLOCK)');
    SQL.Add('WHERE ((CTIME >= :CTIME1) AND (CTIME <= :CTIME2))');
    SQL.Add('AND (RTRIM(PID) <> '''')');
    //�p�G�O���έp�A���[�W�ɶ��L�o����
    if UseTimeFilter then
    begin
      AddWhere('(CAST(CONVERT(varchar(8), ITIME, 14) AS TIME) >= ''0:0:0'')');
      AddWhere('(CAST(CONVERT(varchar(8), ITIME, 14) AS TIME) <= ''23:59:59'')');
    end;
    //-------------------------------------------------------------------------
    if AAnswerOnly then
      SQL.Add('AND (SCODE =1 )');

    SQL.Add('GROUP BY CONVERT(varchar(10), ITIME, 111), PID');

    Params.ParamValues['CTIME1'] := ABeginTime;
    Params.ParamValues['CTIME2'] := AEndTime;
    //-------------------------------------------------------------------------
    Open;
  end;
end;

procedure TdmAcdSummary.ACD_CalcData_Group(ASiteId: string);
var
  aSw: string;
  aDate: TDateTime;
  A: Variant;
begin
  with qryGetAcdInfo do
  begin
    if (not Active) or IsEmpty then Exit;
    First;

    while not Eof do
    begin
      aSw := GetSwByAgentGroup(FieldByName('PID').AsString);
      aDate := FieldByName('_DATE_').AsDateTime;
      A := VarArrayOf([ASiteId, aSw, aDate]);

      //if mdAcdSwDaily.Locate('SiteId;Sw;PhoneDate', A, []) then
      if CheckSwDailyRec(ASiteId, aSw, aDate) then
      begin
        mdAcdSwDaily.Edit;
        //---------------------------------------------------------------------
        mdAcdSwDailyACD_Total.AsInteger := FieldByName('PHONE_COUNT').AsInteger;
        
        mdAcdSwDailyACD_ValidIn_Total.AsInteger := mdAcdSwDailyACD_Total.AsInteger
          - mdAcdSwDailyACD_InvalidIn_Total.AsInteger - mdAcdSwDailyACD_Assign_Invalid.AsInteger;

        if mdAcdSwDailyACD_ValidIn_Total.AsInteger = 0 then
          mdAcdSwDailyACD_Score.AsFloat := 0
        else
          mdAcdSwDailyACD_Score.AsFloat := 100 * mdAcdSwDailyACD_ValidAns_Total.AsInteger
            / mdAcdSwDailyACD_ValidIn_Total.AsInteger;

//        if mdAcdSwDailyACD_Score.AsFloat > 100 then
//          mdAcdSwDailyACD_Score.AsFloat := 100;
        //---------------------------------------------------------------------
        mdAcdSwDaily.Post;
      end;
      Next;
    end;
  end;
end;

function TdmAcdSummary.GetSwByAgentGroup(AAgentGroup: string): string;
begin
  AAgentGroup := Trim(AAgentGroup);

	with dmReport do
  begin
    if Pos(AAgentGroup, '2001^3901^5914^6905') > 0 then
      Result := GetSwName_10('20') //WSTP
    else if Pos(AAgentGroup, '2002^3902^5915^6906') > 0 then
      Result := GetSwName_10('10') //MERP
    else if Pos(AAgentGroup, '2003^3903^5916^6907') > 0 then
      Result := GetSwName_10('30') //HRS
    else if Pos(AAgentGroup, '2004^3904^5917^6909') > 0 then
      Result := GetSwName_10('50') //WBEC
    else if Pos(AAgentGroup, '2005^3905^5918^6908') > 0 then
      Result := GetSwName_10('40') //WSTF
    else
      Result := STR_ITEM_MISC;  //unknown
	end;
end;

procedure TdmAcdSummary.CalcData_Sw_Daily;
begin
  with mdAcdSwDaily do
  begin
    DisableControls;
    First;

    try
      while not Eof do
      begin
        Edit;
        //�X���O�ɲv
        if mdAcdSwDailyPhone_Count_C.AsInteger <> 0 then
          mdAcdSwDailyTimeOut_Rate_C.AsFloat := mdAcdSwDailyTimeOut_Count_C.AsInteger / mdAcdSwDailyPhone_Count_C.AsInteger
        else
          mdAcdSwDailyTimeOut_Rate_C.AsFloat := 0;
        //�D�X���O�ɲv
        if mdAcdSwDailyPhone_Count_NC.AsInteger <> 0 then
          mdAcdSwDailyTimeOut_Rate_NC.AsFloat := mdAcdSwDailyTimeOut_Count_NC.AsInteger / mdAcdSwDailyPhone_Count_NC.AsInteger
        else
          mdAcdSwDailyTimeOut_Rate_NC.AsFloat := 0;
        //---------------------------------------------------------------------
        CheckBrowseMode;
        Next;
      end;
    finally
      EnableControls;
    end;
  end;
end;

procedure TdmAcdSummary.CalcData_Te_Daily;
begin
  with mdAcdTeDaily do
  begin
    DisableControls;
    First;

    try
      while not Eof do
      begin
        Edit;
        //�D���e�^�q��
        mdAcdTeDailyNotAcdCallout_Count.AsInteger :=
          mdAcdTeDailyCallout_Total.AsInteger - mdAcdTeDailyACD_ValidAns_Total.AsInteger;
        //---------------------------------------------------------------------
        CheckBrowseMode;
        Next;
      end;
    finally
      EnableControls;
    end;
  end;
end;

procedure TdmAcdSummary.mdPhoneOutSrcCalcFields(DataSet: TDataSet);
begin
  inherited;
  mdPhoneOutSrcRPHE011_REV.AsBoolean := (mdPhoneOutSrcRPHE011.AsString = 'N');
end;

function TdmAcdSummary.IsContracted(AFlag: string): Boolean;
begin
  AFlag := UpperCase(AFlag);
  Result := (AFlag = 'S') or (AFlag = 'N');
end;

procedure TdmAcdSummary.CheckDuplicateTK;
var
  aRecNo: Integer;
  aTK: string;
begin
  with mdAcdSrc do
  begin
    if not Active then Exit;
    DisableControls;
    aRecNo := RecNo;
    mdDupCheck.Close;
    mdDupCheck.Open;

    try
      First;
      while not Eof do
      begin
        aTK := Trim(mdAcdSrcTK.AsString);

        if (aTK <> '') then
        begin
          if mdDupCheck.Locate('TK', aTK, []) then
          begin
            mdDupCheck.Edit;
            mdDupCheckCount.AsInteger := mdDupCheckCount.AsInteger + 1;
          end
          else
          begin
            mdDupCheck.Append;
            mdDupCheckTK.AsString := aTK;
            mdDupCheckCount.AsInteger := 1;
          end;
          mdDupCheck.Post;
        end;
        Next;
      end;

      First;
      while not Eof do
      begin
        aTK := Trim(mdAcdSrcTK.AsString);

        if mdDupCheck.Locate('TK', aTK, []) and (mdDupCheckCount.AsInteger > 1) then
        begin
          Edit;
          mdAcdSrcDUP.AsBoolean := True;
          POst;
        end;
        Next;
      end;
    finally
      if (aRecNo <> -1) then
        RecNo := aRecNo;

      EnableControls;
      mdDupCheck.Close;
    end;
  end;
end;

function TdmAcdSummary.UseTimeFilter: Boolean;
begin
  Result := DaysBetween(FCalcInBegTime, FCalcInEndTime) > 0;
end;

function TdmAcdSummary.Get_IPH2004(IPHE001: Integer): string;
begin
  with qrGetIPH2004 do
  begin
    if Active then Close;
    ParamByName('IPHE001').AsInteger := IPHE001;
    Open;
    Result := qrGetIPH2004GUID.AsString;
    Close;
  end;
end;

procedure TdmAcdSummary.SaveAcdSiteDaily;
var
  aCmdUpd, aCmdNew: TUniSQL;
  aRecNo, aCC: Integer;

  procedure PrepareSQL;
  begin
    with aCmdNew.SQL do
    begin
      Clear;
    	Add('INSERT INTO WICSIPH4(');
      Add('IPH4001, IPH4002, IPH4003, IPH4004, IPH4005, IPH4006, IPH4007, IPH4008, IPH4009, IPH4010,');
      Add('IPH4011, IPH4012, IPH4013, IPH4099,');
      Add('IPH4205, IPH4209, IPH4210, IPH4211, IPH4213');
      Add(')');
    	Add('VALUES(');
      Add(':IPH4001, :IPH4002, :IPH4003, :IPH4004, :IPH4005, :IPH4006, :IPH4007, :IPH4008, :IPH4009, :IPH4010,');
      Add(':IPH4011, :IPH4012, :IPH4013, :IPH4099,');
      Add(':IPH4205, :IPH4209, :IPH4210, :IPH4211, :IPH4213');
      Add(')');
    end;

    with aCmdUpd.SQL do
    begin
      Clear;
    	Add('UPDATE WICSIPH4 SET');
    	Add('IPH4003 = :IPH4003, IPH4004 = :IPH4004, IPH4005 = :IPH4005, IPH4006 = :IPH4006,');
      Add('IPH4007 = :IPH4007, IPH4008 = :IPH4008, IPH4009 = :IPH4009, IPH4010 = :IPH4010,');
      Add('IPH4011 = :IPH4011, IPH4012 = :IPH4012, IPH4013 = :IPH4013, IPH4099 = :IPH4099,');
      Add('IPH4205 = :IPH4205, IPH4209 = :IPH4209, IPH4210 = :IPH4210, IPH4211 = :IPH4211,');
      Add('IPH4213 = :IPH4213');      
    	Add('WHERE (IPH4001 = :IPH4001) AND (IPH4002 = :IPH4002)');
    end;
  end;

  function PostData(ANewData: Boolean): Integer;
  var
    aCmd: TUniSQL;
  begin
    Result := 0;

    if ANewData then
      aCmd := aCmdNew
    else
      aCmd := aCmdUpd;

    with aCmd do
    begin
      //��~�B
      ParamByName('IPH4001').Value := dmReport.GetDescOfSite(dmReport.GetSiteName(mdAcdSiteDailySiteId.AsString));
      //�ӹq��
      ParamByName('IPH4002').Value := mdAcdSiteDailyPhoneDate.AsDateTime;
      //�Ⱦ��H�O
      ParamByName('IPH4003').Value := mdAcdSiteDailyDays.AsFloat;
      //��~�B��ť�v
      ParamByName('IPH4004').Value := mdAcdSiteDailyACD_Score.AsFloat;
      //�X���O�ɲv
      ParamByName('IPH4005').Value := mdAcdSiteDailyTimeOut_Rate_C.AsFloat;
      //ACD���e�`��
      ParamByName('IPH4006').Value := mdAcdSiteDailyACD_Total.AsInteger;
      //ACD���Ĭ��e��
      ParamByName('IPH4007').Value := mdAcdSiteDailyACD_ValidIn_Total.AsInteger;
      //ACD�B�z��
      ParamByName('IPH4008').Value := mdAcdSiteDailyACD_ValidAns_Total.AsInteger;
      //�X���^�q��
      ParamByName('IPH4009').Value := mdAcdSiteDailyPhoneOut_Count_C.AsInteger;
      //�̫��s�ɶ�
      ParamByName('IPH4099').Value := Now;
      //�X���O�ɳq�� Added by Joe 2017/08/01 09:47:14
      ParamByName('IPH4010').Value := mdAcdSiteDailyTimeOut_Count_C.AsInteger;
      //�X���ӹq�`�� Added by Joe 2017/08/01 09:47:14
      ParamByName('IPH4011').Value := mdAcdSiteDailyPhone_Count_C.AsInteger;
      //ACD������ť�� Added by Joe Lee 2017/10/31 14:15:23
      ParamByName('IPH4012').Value := mdAcdSiteDailyACD_Ans_Total.AsInteger;
      //���^�q�� Added by Joe Lee 2017/10/31 15:00:27
      ParamByName('IPH4013').Value := mdAcdSiteDailyNoAns_Count_C.AsInteger;
      //------------------------------------------------------------------------
      //�D�X���O�ɲv Added by Joe Lee 2017/10/31 16:25:13
      ParamByName('IPH4205').Value := mdAcdSiteDailyTimeOut_Rate_NC.AsFloat;
      //�D�X���^�q�`�� Added by Joe Lee 2017/10/31 16:25:13
      ParamByName('IPH4209').Value := mdAcdSiteDailyPhoneOut_Count_NC.AsInteger;
      //�D�X���O�ɳq�� Added by Joe Lee 2017/10/31 16:25:13
      ParamByName('IPH4210').Value := mdAcdSiteDailyTimeOut_Count_NC.AsInteger;
      //�D�X���ӹq�`�� Added by Joe Lee 2017/10/31 16:25:13
      ParamByName('IPH4211').Value := mdAcdSiteDailyPhone_Count_NC.AsInteger;
      //�D�X�����^�q�� Added by Joe Lee 2017/10/31 16:25:13
      ParamByName('IPH4213').Value := mdAcdSiteDailyNoAns_Count_NC.AsInteger;

      try
        Execute;
        Result := RowsAffected;
      except
        (*
        on E: Exception do
          MessageBox(Handle, PChar(E.Message), PChar(Application.Title), MB_OK + MB_ICONWARNING);
        *)
      end;
    end;
  end;
begin
	Log('�x�s[��~�B]ACD��T');
  aCmdUpd := dmReport.GetCmd_WintonTcrm;
  aCmdNew := dmReport.GetCmd_WintonTcrm;
  PrepareSQL;
  aCC := 0;
  aRecNo := JcGetRecNo(mdAcdSiteDaily);

  try
    with mdAcdSiteDaily do
    begin
      if not Active Then Exit;
      DisableControls;
      dmReport.StartTrans_WintonTcrm;
      First;

      try
        while not Eof do
        begin
          if PostData(True) < 1 then
            PostData(False);

          Next;
          Inc(aCC);
          dmReport.BatchCommit_WintonTcrm;

          if ((aCC mod 10) = 0) then
            Application.ProcessMessages;
        end;
        dmReport.BatchCommit_WintonTcrm;
      except
        dmReport.Rollback_WintonTcrm;
      end;
    end;
  finally
    JcSetRecNo(mdAcdSiteDaily, aRecNo);
    mdAcdSiteDaily.EnableControls;

    aCmdUpd.Free;
    aCmdNew.Free;
  end;
end;

procedure TdmAcdSummary.SaveAcdTeDaily;
var
  aCmdUpd, aCmdNew: TUniSQL;
  aRecNo, aCC: Integer;

  procedure PrepareSQL;
  begin
    with aCmdNew.SQL do
    begin
      Clear;
    	Add('INSERT INTO WICSIPH3(');
      Add('IPH3001, IPH3002, IPH3003, IPH3006, IPH3007,');
      Add('IPH3008, IPH3009, IPH3010, IPH3011, IPH3012,');
      Add('IPH3013, IPH3017, IPH3099');
      Add(')');
    	Add('VALUES(');
      Add(':IPH3001, :IPH3002, :IPH3003, :IPH3006, :IPH3007,');
      Add(':IPH3008, :IPH3009, :IPH3010, :IPH3011, :IPH3012,');
      Add(':IPH3013, :IPH3017, :IPH3099');
      Add(')');
    end;

    with aCmdUpd.SQL do
    begin
			Clear;
    	Add('UPDATE WICSIPH3 SET');
      Add('IPH3003 = :IPH3003, IPH3006 = :IPH3006, IPH3007 = :IPH3007, IPH3008 = :IPH3008, IPH3009 = :IPH3009,');
      Add('IPH3010 = :IPH3010, IPH3011 = :IPH3011, IPH3012 = :IPH3012, IPH3013 = :IPH3013, IPH3017 = :IPH3017,');
      Add('IPH3099 = :IPH3099');
    	Add('WHERE (IPH3001 = :IPH3001) AND (IPH3002 = :IPH3002)');
    end;
	end;

  function PostData(ANewData: Boolean): Integer;
  var
    aCmd: TUniSQL;
    aIPH3017: Extended;

    // Added by Joe Lee 2017/11/15 16:09:14
    // ����[TE�i��]�������Y����ڱ������e���Ⱦ��H��
    procedure Calc_IPH3017;
    begin
      //aIPH3017 := mdAcdTeDailyDays.AsFloat * 2;
      aIPH3017 := 0;

      if mdAcdTeDailyDuty_AM.AsString = STR_ACD_ON_DUTY then
        aIPH3017 := aIPH3017 + 1;

      if mdAcdTeDailyDuty_PM.AsString = STR_ACD_ON_DUTY then
        aIPH3017 := aIPH3017 + 1;
    end;

    // Added by Joe Lee 2017/11/15 16:09:23
    // �̾ڱ������e�H�ѡA�p��X�ӭ����O�_�F��
    function Calc_IPH3013: Boolean;
    begin
      // �p�G[ACD���e�Ⱦ��Ѽ�]��0�A�@�߳]���F��
      if (aIPH3017 = 0) then
        Result := True
      else
        Result := mdAcdTeDailyACD_ValidAns_Total.AsInteger >= (mdAcdTeDailyACD_DailyReqCount.AsInteger * aIPH3017 * 0.5);
    end;
  begin
    Result := 0;
    Calc_IPH3017; // Added by Joe Lee 2017/11/15 15:59:33

    if ANewData then
      aCmd := aCmdNew
    else
      aCmd := aCmdUpd;

    with aCmd do
    begin
      //�V�m�v
      ParamByName('IPH3001').Value := mdAcdTeDailyEmpId.AsString;
      //�ӹq��
      ParamByName('IPH3002').Value := mdAcdTeDailyPhoneDate.AsDateTime;
      //ACD�q��
      ParamByName('IPH3003').Value := mdAcdTeDailyACD_ValidAns_Total.AsInteger;
      //ACD��зǳq��
      ParamByName('IPH3006').Value := mdAcdTeDailyACD_DailyReqCount.AsInteger;
      //ACD�Ⱦ��Ѽ�
      ParamByName('IPH3007').Value := mdAcdTeDailyDays.AsFloat * 2;
      //ACD�B�z��
      ParamByName('IPH3008').Value := mdAcdTeDailyACD_ValidAns_Total.AsInteger;
      //�`�^�q��
      ParamByName('IPH3009').Value := mdAcdTeDailyCallout_Total.AsInteger;
      //�u�@���O(�W��)
      ParamByName('IPH3011').Value := mdAcdTeDailyDuty_AM.AsString;
      //�u�@���O(�U��)
      ParamByName('IPH3012').Value := mdAcdTeDailyDuty_PM.AsString;
      //�̫��s�ɶ�
      ParamByName('IPH3099').Value := Now;
      //�D���e�^�q�� Added by Joe Lee 2017/10/31 16:46:05
      ParamByName('IPH3010').Value := mdAcdTeDailyNotAcdCallout_Count.AsInteger;
      //ACD���e�Ⱦ��Ѽ� Added by Joe Lee 2017/11/15 15:59:53
      ParamByName('IPH3017').Value := aIPH3017;
      ParamByName('IPH3013').Value := Calc_IPH3013;

      try
        Execute;
        Result := RowsAffected;
      except
        on E: Exception do
        begin
//          aText := Format('SaveAcdTeDaily(), Err = %s', [E.Message]);
//          CodeSite.SendError(AnsiToUtf8(aText));
        end;
      end;
    end;
  end;
  // �P�_�O�_�����Ī�ACD�έp��� Added by Joe 2019/01/02 16:29:39
  function IsValidAcdData: Boolean;
  begin
    // �p�G[ACD�q��]�P[ACD�B�z��]�Ҭ�0�A�����L�Ī��έp��ơA���ݭn�x�s
    if(mdAcdTeDailyACD_ValidAns_Total.AsInteger = 0) and (mdAcdTeDailyACD_ValidAns_Total.AsInteger = 0) then
      Result := False
    else
      Result := True;
  end;
begin
	Log('�x�s[�V�m�v]ACD��T');
  aCmdUpd := dmReport.GetCmd_WintonTcrm;
  aCmdNew := dmReport.GetCmd_WintonTcrm;
  PrepareSQL;
  aCC := 0;
  aRecNo := JcGetRecNo(mdAcdTeDaily);

  try
    with mdAcdTeDaily do
    begin
      if not Active Then Exit;
      DisableControls;
      dmReport.StartTrans_WintonTcrm;
      First;

      try
        while not Eof do
        begin
          if IsValidAcdData then
          begin
            if PostData(True) < 1 then
              PostData(False);

            Inc(aCC);
          end;
          Next;
          dmReport.BatchCommit_WintonTcrm;

          if ((aCC mod 10) = 0) then
            Application.ProcessMessages;
        end;
        dmReport.BatchCommit_WintonTcrm;
      except
        dmReport.Rollback_WintonTcrm;
      end;
    end;
  finally
    JcSetRecNo(mdAcdTeDaily, aRecNo);
    mdAcdTeDaily.EnableControls;

    aCmdUpd.Free;
    aCmdNew.Free;
  end;
end;

procedure TdmAcdSummary.SaveAcdSwDaily;
var
  aCmdUpd, aCmdNew: TUniSQL;
  aRecNo, aCC: Integer;

  procedure PrepareSQL;
  begin
    with aCmdNew.SQL do
    begin
      Clear;
    	Add('INSERT INTO WICSIPH5(IPH5001, IPH5002, IPH5003, IPH5004, IPH5005, IPH5006, IPH5007, IPH5008, IPH5099)');
    	Add('VALUES(:IPH5001, :IPH5002, :IPH5003, :IPH5004, :IPH5005, :IPH5006, :IPH5007, :IPH5008, :IPH5099)');
    end;

    with aCmdUpd.SQL do
    begin
      Clear;
      Add('UPDATE WICSIPH5 SET');
      Add('IPH5004 = :IPH5004, IPH5005 = :IPH5005, IPH5006 = :IPH5006, IPH5007 = :IPH5007,');
      Add('IPH5008 = :IPH5008, IPH5099 = :IPH5099');
      Add('WHERE (IPH5001 = :IPH5001) AND (IPH5002 = :IPH5002) AND (IPH5003 = :IPH5003)');
    end;
  end;

  function PostData(ANewData: Boolean): Integer;
  var
    aCmd: TUniSQL;
  begin
    Result := 0;

    if ANewData then
      aCmd := aCmdNew
    else
      aCmd := aCmdUpd;

    with aCmd do
    begin
      //��~�B
      ParamByName('IPH5001').Value := mdAcdSwDailySiteName.AsString;
      //�ӹq��
      ParamByName('IPH5002').Value := mdAcdSwDailyPhoneDate.AsDateTime;
      //�t�ΧO
      ParamByName('IPH5003').Value := mdAcdSwDailySw.AsString;
      //�t�α�ť�v
      ParamByName('IPH5004').Value := mdAcdSwDailyACD_Score.AsFloat;
      //�X���O�ɲv
      ParamByName('IPH5005').Value := mdAcdSwDailyTimeOut_Rate_C.AsFloat;
      //ACD���e�`��
      ParamByName('IPH5006').Value := mdAcdSwDailyACD_Total.AsInteger;
      //ACD���Ĭ��e��
      ParamByName('IPH5007').Value := mdAcdSwDailyACD_ValidIn_Total.AsInteger;
      //ACD�B�z��
      ParamByName('IPH5008').Value := mdAcdSwDailyACD_ValidAns_Total.AsInteger;
      //�̫��s�ɶ�
      ParamByName('IPH5099').Value := Now;

      try
        Execute;
        Result := RowsAffected;
      except
        (*
        on E: Exception do
          MessageBox(Handle, PChar(E.Message), PChar(Application.Title), MB_OK + MB_ICONWARNING);
        *)
      end;
    end;
  end;
begin
	Log('�x�s[�t�ΧO]ACD��T');
  aCmdUpd := dmReport.GetCmd_WintonTcrm;
  aCmdNew := dmReport.GetCmd_WintonTcrm;
  PrepareSQL;
  aRecNo := JcGetRecNo(mdAcdSwDaily);
  aCC := 0;

  try
    with mdAcdSwDaily do
    begin
      if not Active Then Exit;
      DisableControls;
      dmReport.StartTrans_WintonTcrm;
      First;

      try
        while not Eof do
        begin
          if PostData(True) < 1 then
            PostData(False);

          Next;
          Inc(aCC);
          dmReport.BatchCommit_WintonTcrm;

          if ((aCC mod 10) = 0) then
            Application.ProcessMessages;
        end;
        dmReport.BatchCommit_WintonTcrm;
      except
        dmReport.Rollback_WintonTcrm;
      end;
    end;
  finally
    JcSetRecNo(mdAcdSwDaily, aRecNo);
    mdAcdSwDaily.EnableControls;

    aCmdUpd.Free;
    aCmdNew.Free;
  end;
end;

procedure TdmAcdSummary.PrepareData_ACD(ASiteId: string);
var
	aSiteName: string;
begin
  LogLine;
  aSiteName := dmReport.GetBranchOfDept(ASiteId);
  Log(Format('��zACD��T, Site = [%s] %s', [ASiteId, aSiteName]));
  InitReportConn(ASiteId);
  mdAcdSrc.Close;
  mdPhoneOutSrc.Close;
  //���o�Ⱦ����T Added by Joe 2017/11/10 15:51:04
  PrepreData_WICSIPHH(FCalcInBegTime, FCalcInEndTime);
  //���o�þ�z�ӹq���
  PrepareData_PhoneIn(ASiteId, FCalcInBegTime, FCalcInEndTime);
  Application.ProcessMessages;
  //���o�þ�z�^�q���
  PrepareData_PhoneOut(ASiteId, FCalcOutBegTime, FCalcOutEndTime);
  Application.ProcessMessages;
  Log('�M�z�^�q���');
  CleanData_PhoneOut(mdPhoneOutSrc);
  Application.ProcessMessages;
end;

class procedure TdmAcdSummary.Exec_CalcData(ACalcInBegTime, ACalcInEndTime, ACalcOutBegTime, ACalcOutEndTime: TDateTime;
  AForceReCalc: Boolean);
begin
	if not Assigned(dmAcdSummary) then
  	Application.CreateForm(TdmAcdSummary, dmAcdSummary);

	with dmAcdSummary do
  begin
    FForceCalcData := AForceReCalc;
		CalcReportData(ACalcInBegTime, ACalcInEndTime, ACalcOutBegTime, ACalcOutEndTime);
    Free;
  end;
end;

procedure TdmAcdSummary.InitData;
  procedure InitDatSet(ADataSet: TDataSet);
  begin
    with ADataSet do
    begin
      DisableControls;
      if Active then Close;
      Open;
    end;
  end;
begin
	mdAcdSrc.EmptyTable;
  mdPhoneOutSrc.EmptyTable;

  InitDatSet(mdAcdSrc);
  InitDatSet(mdPhoneOutSrc);
  InitDatSet(mdAcdTeDaily);
  InitDatSet(mdAcdSiteDaily);
  InitDatSet(mdAcdSwDaily);
end;

procedure TdmAcdSummary.CalcAcdData(ASiteId: string);
begin
  InitReportConn(ASiteId);
  // �ӤH -----------------------------------------------------------------
  Log(Format('�p��[�ӤH�ӹq]���(%s)', [ASiteId]));
  CalcData_PhoneIn(ASiteId);
  Application.ProcessMessages;
  Log(Format('�p��[�ӤH�^�q]���', [ASiteId]));
  CalcData_PhoneOut(ASiteId);
  Application.ProcessMessages;
  Log(Format('�p��[�ӤH�洫��]���', [ASiteId]));
  FillData_TeOnDutyCount(ASiteId, FCalcInBegTime, FCalcInEndTime);
  Application.ProcessMessages;
  Log(Format('�έp[�ӤH��ť�v]���', [ASiteId]));
  CalcData_Te_Daily;
  Application.ProcessMessages;
  // Added by Joe 2017/11/13 14:39:33
  Log(Format('�p��[�ӤH�u�@���O]���', [ASiteId]));
  FillData_TeDutyItem(ASiteId, FCalcInBegTime, FCalcInEndTime);
  Application.ProcessMessages;
  // ��~�B ---------------------------------------------------------------
  Log(Format('�p��[��~�B�C��ACD]���', [ASiteId]));
  CalcData_Site_Daily(ASiteId);
  Application.ProcessMessages;
  Log(Format('�p��[��~�B�洫��]���(�����ӹq)', [ASiteId]));
  ACD_GetSiteCount(dmReport.GetSiteName(ASiteId), FCalcInBegTime, FCalcInEndTime, False);
  Application.ProcessMessages;
  ACD_CalcData_Site(ASiteId, False);
  Application.ProcessMessages;
  Log(Format('�p��[��~�B�洫��]���(�������ӹq)', [ASiteId]));
  ACD_GetSiteCount(dmReport.GetSiteName(ASiteId), FCalcInBegTime, FCalcInEndTime, True);
  Application.ProcessMessages;
  ACD_CalcData_Site(ASiteId, True);
  Application.ProcessMessages;
  //�t�ΧO
  Log(Format('�p��[�t�ΧOACD]���', [ASiteId]));
  ACD_GetGroupCount(dmReport.GetSiteName(ASiteId), FCalcInBegTime, FCalcInEndTime, False);
  ACD_CalcData_Group(ASiteId);
  Application.ProcessMessages;
end;

procedure TdmAcdSummary.Log(AMsg: string);
begin
	fmMain.Log(AMsg);
end;

procedure TdmAcdSummary.LogLine(Ch: Char);
begin
	fmMain.LogLine(Ch);
end;

function TdmAcdSummary.CopyReportFromTemplate(ADate: TDateTime): string;
const
  XLS_FILE_TITLE = '�G�~�רӹq�����';
var
  aPath, aSrcFile, aDstFile: string;
  aDays: Word;
begin
  Result := '';
  aPath := IncludeTrailingPathDelimiter(dmReport.GetTemplatePath);
  aDays := DaysInMonth(ADate);
  aSrcFile := Format('%s%s_%d.xlsx', [aPath, XLS_FILE_TITLE, aDays]);
  aPath := IncludeTrailingPathDelimiter(dmReport.GetReportPath);
  aDstFile := Format('%s%s_%s.xlsx', [aPath, XLS_FILE_TITLE, FormatDateTime('yyyymmdd', ADate)]);

  if FileExists(aSrcFile) then
  begin
    Log(Format('�ƻsXLS����d�� %s -> %s', [aSrcFile, aDstFile]));

  	if CopyFile(PChar(aSrcFile), PChar(aDstFile), False) then
   		Result := aDstFile
    else
			Log(Format('!! �L�k�ƻsXLS����d�� %s', [aDstFile]));
  end
  else
    Log(Format('!! �䤣��XLS����d�� %s', [aSrcFile]));
end;

class procedure TdmAcdSummary.Exec_PrintReport(ADate: TDateTime);
begin
	if not Assigned(dmAcdSummary) then
  	Application.CreateForm(TdmAcdSummary, dmAcdSummary);

	with dmAcdSummary do
  begin
    PrintReport(ADate);

    if not fmMain.MailMode then
      ShellExecute(0, 'open', PChar(FXlsFileName), nil, nil, SW_SHOWMAXIMIZED);

  	Free;
  end;
end;

function TdmAcdSummary.PrepareData_SwACD(ASw: string; AYear, AMonth: Integer): TUniQuery;
begin
  Result := dmReport.GetQuery_WintonTcrm;

  with Result do
  begin
    SQL.Add('SELECT IPH5003, IPH5002, SUM(IPH5007) AS ACD_COUNT');
    SQL.Add('FROM WICSIPH5 WITH(NOLOCK)');
    SQL.Add('WHERE');
    SQL.Add('(DATEPART(YEAR, IPH5002) = :YEAR1 OR DATEPART(YEAR, IPH5002) = :YEAR2)');
    SQL.Add('AND (DATEPART(MONTH, IPH5002) = :MONTH)');
    SQL.Add('AND (IPH5003 = :IPH5003)');
    SQL.Add('GROUP BY IPH5003, IPH5002');

    ParamByName('YEAR1').Value := AYear;
    ParamByName('YEAR2').Value := AYear - 1;
    ParamByName('MONTH').Value := AMonth;
    ParamByName('IPH5003').Value := ASw;

    Open;
  end;
end;

procedure TdmAcdSummary.WriteDataToXls(ASw: string; AYear, AMonth, ADay: Word);
var
  aSheet: TXLSWorksheet;
  aXlsName: TXLSName;
  aSwCode, aName, aText: string;
  aData: TUniQuery;
  i, j, k, aEndDay, aRow: Word;
  aPrecedingTotal, aPrecedingTotal_Prev: Integer;

  procedure FillXlsData(AYear, AMonth, ADay, ACol: Integer);
  var
    i: Integer;
    aDate: TDateTime;
  begin
    for i := 1 to ADay do
    begin
      aDate := EncodeDate(AYear, AMonth, i);

      if aData.Locate('IPH5002', aDate, []) then
        aSheet.AsInteger[ACol, aXlsName.Row1+i-1] := aData.FieldByname('ACD_COUNT').AsInteger
      else
      	aSheet.AsInteger[ACol, aXlsName.Row1+i-1] := 0;
    end;
  end;
begin
  with XLSReadWriteII51 do
  begin
    aSheet := SheetByName(ASw);

    if (aSheet = nil) then
    begin
      Log(Format('!! �䤣��u�@��[%s], �L�k�g�J�u�@����', [ASw]));
      Exit;
    end;

    aName := 'DATA_' + ASw;
    aXlsName := XLSReadWriteII51.Names.Find(aName);
  end;

  if not Assigned(aXlsName) then
  begin
    Log(Format('!! �䤣��XLS�W��[%s], �L�k�g�J�u�@����', [aName]));
    Abort;
  end;

  try
    if (ASw = PID_WSTP) then
      aSwCode := 'WSTP2000'
    else if (ASw = PID_WBEC) then
    	aSwCode := 'WBEC2000'
    else if (ASw = PID_NTF) then
    	aSwCode := 'WSTF2000'
    else if (ASw = PID_MERP) then
    	aSwCode := 'MERP'
    else if (ASw = PID_NHR) then
    	aSwCode := 'WHRS'
    else
      aSwCode := ASw;

  	aData := PrepareData_SwACD(aSwCode, AYear, AMonth);
    //��J����(���~)���
    FillXlsData(AYear, AMonth, ADay, aXlsName.Col1);
		// �p�G�������ƾ�,���O���s����ݭn����
    if aData.Locate('IPH5002', EncodeDate(AYear, AMonth, ADay), []) then
    begin
      if (aData.FieldByname('ACD_COUNT').AsInteger > 0) then
        Inc(FNewRptCount);
		end;
    //��J�e��(�h�~)���
    aEndDay := DaysInAMonth(AYear, AMonth);
    FillXlsData(AYear-1, AMonth, aEndDay, aXlsName.Col1+1);
    //��J�~��
    aText := Format('%d', [AYear]);
    aSheet.AsString[0, 0] := StringReplace(aSheet.AsString[0, 0], '[yyyy]', aText, [rfReplaceAll, rfIgnoreCase]);
    //��J�~��
    aText := Format('%d/%d', [AMonth, ADay]);
    aSheet.AsString[6, 0] := StringReplace(aSheet.AsString[6, 0], '[mm/dd]', aText, [rfReplaceAll, rfIgnoreCase]);
    // ��J��~�׫e���֭p Added by Joe 2018/05/21 11:48:21
    GetPrecedingAcdTotal(aSwCode, aPrecedingTotal, aPrecedingTotal_Prev);
    aSheet.AsInteger[1, 2] := aPrecedingTotal;
    aSheet.AsInteger[2, 2] := aPrecedingTotal_Prev;
    //��J�����t������
    //aRow := ADay + 2;
    aRow := ADay + 3;
    aSheet.AsFormula[11, 0] := 'D' + IntToStr(aRow);
    aSheet.AsString[1, 1] := Format('%d�~', [AYear]);
    aSheet.AsString[2, 1] := Format('%d�~', [AYear-1]);
    // �аO�g���ζg��
		j := DaysInAMonth(AYear, AMonth);

    for i := 0 to j-1 do
    begin
      // ��J������
      aSheet.AsFloat[0, i+3] := EncodeDate(AYear, AMonth, i+1);
      // ���~
      k := DayOfWeek(EncodeDate(AYear, AMonth, i+1));

      if (k = 7) then
        aSheet.Cell[1, i+3].FillPatternForeColorRGB := $00BFFFBF
      else if (k = 1) then
        aSheet.Cell[1, i+3].FillPatternForeColorRGB := $00FFB5FF;
      // �h�~
      k := DayOfWeek(EncodeDate(AYear-1, AMonth, i+1));

      if (k = 7) then
        aSheet.Cell[2, i+3].FillPatternForeColorRGB := $00BFFFBF
      else if (k = 1) then
        aSheet.Cell[2, i+3].FillPatternForeColorRGB := $00FFB5FF;
    end;
		//�аO�X�ثe���p��I���
 		for i := 3 to 3 do
    begin
      if aSheet.Cell[i, aRow-1] <> nil then
      begin
        aSheet.Cell[i, aRow-1].FillPatternForeColorRGB := $00FFFFB9;
      end;
    end;
		// �N���֭p�W��ƼаO������
    with XLSReadWriteII51 do
    begin
      CmdFormat.BeginEdit(aSheet);
      CmdFormat.Font.Bold := True;
      CmdFormat.Font.Size := 16;
      CmdFormat.Apply(3, aRow-1);
    end;
  finally
    dmReport.CloseAndFree(aData);
  end;
end;

procedure TdmAcdSummary.UpdateWorkSheet_Summary(AYear, ADay: Integer);
const
  WORK_SHEET_NAME = '�έp��';
var
  aSheet: TXLSWorksheet;
  aText: string;
  aAreaRef1, aAreaRef2, aRcRef1: string;
begin
  with XLSReadWriteII51 do
  begin
    aSheet := SheetByName(WORK_SHEET_NAME);

    if (aSheet = nil) then
    begin
      Log(Format('!! �䤣��u�@��[%s], �L�k��s�έp����', [WORK_SHEET_NAME]));
      Exit;
    end;
    aText := Format('%d', [AYear]);
    aSheet.AsString[0, 0] := StringReplace(aSheet.AsString[0, 0], '[yyyy]', aText, [rfReplaceAll, rfIgnoreCase]);
    aSheet.AsString[1, 1] := StringReplace(aSheet.AsString[1, 1], '[yyyy]', aText, [rfReplaceAll, rfIgnoreCase]);
    aSheet.AsString[3, 1] := StringReplace(aSheet.AsString[3, 1], '[yyyy]', aText, [rfReplaceAll, rfIgnoreCase]);
    aSheet.AsString[5, 1] := StringReplace(aSheet.AsString[5, 1], '[yyyy]', aText, [rfReplaceAll, rfIgnoreCase]);
    aText := Format('%d', [AYear-1]);
    aSheet.AsString[1, 1] := StringReplace(aSheet.AsString[1, 1], '[yyyy-prev]', aText, [rfReplaceAll, rfIgnoreCase]);
    aSheet.AsString[3, 1] := StringReplace(aSheet.AsString[3, 1], '[yyyy-prev]', aText, [rfReplaceAll, rfIgnoreCase]);
    aSheet.AsString[5, 1] := StringReplace(aSheet.AsString[5, 1], '[yyyy-prev]', aText, [rfReplaceAll, rfIgnoreCase]);
    // Added by Joe 2018/05/18 15:41:36
    //��J�����W���Ҥ���
    // =(SUM(WSTP!B4:B23)/SUM(WSTP!C4:C23))-1
    aAreaRef1 := AreaToRefStr(1, 3, 1, ADay+2);
    aAreaRef2 := AreaToRefStr(2, 3, 2, ADay+2);
    aSheet.AsFormula[2, 2] := Format('(SUM(WSTP!%s)/SUM(WSTP!%s))-1', [aAreaRef1, aAreaRef2]);
    aSheet.AsFormula[2, 3] := Format('(SUM(WBEC!%s)/SUM(WBEC!%s))-1', [aAreaRef1, aAreaRef2]);
    aSheet.AsFormula[2, 4] := Format('(SUM(NTF!%s)/SUM(NTF!%s))-1', [aAreaRef1, aAreaRef2]);
    aSheet.AsFormula[2, 5] := Format('(SUM(MERP!%s)/SUM(MERP!%s))-1', [aAreaRef1, aAreaRef2]);
    aSheet.AsFormula[2, 6] := Format('(SUM(NHR!%s)/SUM(NHR!%s))-1', [aAreaRef1, aAreaRef2]);
    aSheet.AsFormula[2, 7] :=
      Format('(SUM(WSTP!%0:s,WBEC!%0:s,NTF!%0:s,MERP!%0:s,NHR!%0:s)/SUM(WSTP!%1:s,WBEC!%1:s,NTF!%1:s,MERP!%1:s,NHR!%1:s))-1',
      [aAreaRef1, aAreaRef2]);
    //��J�֭p�W���Ҥ���
    aRcRef1 :=  ColRowToRefStr(2, ADay+2);
    aAreaRef2 := AreaToRefStr(2, 2, 2, ADay+2);
    aSheet.AsFormula[6, 2] := Format('F3/SUM(WSTP!%s)', [aAreaRef2]);
    aSheet.AsFormula[6, 3] := Format('F4/SUM(WBEC!%s)', [aAreaRef2]);
    aSheet.AsFormula[6, 4] := Format('F5/SUM(NTF!%s)', [aAreaRef2]);
    aSheet.AsFormula[6, 5] := Format('F6/SUM(MERP!%s)', [aAreaRef2]);
    aSheet.AsFormula[6, 6] := Format('F7/SUM(NHR!%s)', [aAreaRef2]);
    aSheet.AsFormula[6, 7] :=
      Format('F8/(SUM(WSTP!%0:s)+SUM(WBEC!%0:s)+SUM(NTF!%0:s)+SUM(MERP!%0:s)+SUM(NHR!%0:s))', [aAreaRef2]);
  end;
end;

procedure TdmAcdSummary.PrintReport(ADate: TDateTime);
var
  aYear, aMonth, aDay: Word;
begin
  DecodeDate(ADate, aYear, aMonth, aDay);
  FXlsFileName := CopyReportFromTemplate(ADate);
  if (FXlsFileName = '') then Exit;

  XLSReadWriteII51.Filename := FXlsFileName;
  XLSReadWriteII51.Read;
  PrepareData_PrecedingAcdTotal;  // Added by Joe 2018/05/21 11:13:04
  WriteDataToXls(PID_WSTP, aYear, aMonth, aDay);
  WriteDataToXls(PID_WBEC, aYear, aMonth, aDay);
  WriteDataToXls(PID_NTF, aYear, aMonth, aDay);
  WriteDataToXls(PID_MERP, aYear, aMonth, aDay);
  WriteDataToXls(PID_NHR, aYear, aMonth, aDay);
  UpdateWorkSheet_Summary(aYear, aDay);
  
  XLSReadWriteII51.Write;
end;

class procedure TdmAcdSummary.Exec(ADate: TDateTime);
var
  aCalcInBegTime, aCalcInEndTime, aCalcOutBegTime, aCalcOutEndTime: TDateTime;
begin
	if not Assigned(dmAcdSummary) then
  	Application.CreateForm(TdmAcdSummary, dmAcdSummary);

	with dmAcdSummary do
  begin
    FNewRptCount := 0;
    aCalcInBegTime := DateOf(ADate);
    aCalcInEndTime := EndOfTheDay(ADate);
    aCalcOutBegTime:= aCalcInBegTime;
    aCalcOutEndTime:= aCalcInEndTime;
		CalcReportData(aCalcInBegTime, aCalcInEndTime, aCalcOutBegTime, aCalcOutEndTime);
    PrintReport(ADate);
    Log('�Ұ�EXCEL�Ӱ���t�s�s��');
    fmMain.CallExcelToSaveAs(FXlsFileName);

    if (FNewRptCount = 0) then
      Log('�����S�����,�����ͳ���')
    else if (FNewRptCount > 0) then
    begin
      if fmMain.MailMode then
      begin
        Log(Format('�H�e����-%s', [FXlsFileName]));
        SendMail;
      end
      else
        ShellExecute(0, 'open', PChar(FXlsFileName), nil, nil, SW_SHOWMAXIMIZED)
    end;
    LogLine('=');
  	Free;
  end;
end;

procedure TdmAcdSummary.CalcReportData(ACalcInBegTime, ACalcInEndTime, ACalcOutBegTime, ACalcOutEndTime: TDateTime);
const
  DATETIME_FMT = 'yyyy/mm/dd hh:nn';
var
  aText: string;
begin
  FCalcInBegTime := ACalcInBegTime;
  FCalcInEndTime := ACalcInEndTime;
  FCalcOutBegTime:= ACalcOutBegTime;
  FCalcOutEndTime:= ACalcOutEndTime;

  FInCalcData := True;
  InitExecute;
  BeginExecute;

  Log('�}�l�έp���x�sACD��ť�v���');
  aText := Format('�ӹq�ɶ�:%s~%s', [FormatDateTime(DATETIME_FMT, ACalcInBegTime), FormatDateTime(DATETIME_FMT, ACalcInEndTime)]);
  Log(aText);
  aText := Format('�^�q�ɶ�:%s~%s', [FormatDateTime(DATETIME_FMT, ACalcOutBegTime), FormatDateTime(DATETIME_FMT, ACalcOutEndTime)]);
  Log(aText);

  try
    PrepareData_ACD(SITE_ID_Taipei);
    CalcAcdData(SITE_ID_Taipei);
    //--------------------------------------------------------------------------
    PrepareData_ACD(SITE_ID_Taoyuan);
    CalcAcdData(SITE_ID_Taoyuan);
    //--------------------------------------------------------------------------
    PrepareData_ACD(SITE_ID_Taichung);
    CalcAcdData(SITE_ID_Taichung);
    //--------------------------------------------------------------------------
    PrepareData_ACD(SITE_ID_Tainan);
    CalcAcdData(SITE_ID_Tainan);
    //--------------------------------------------------------------------------
    //�p��t�ΧO�έp���
    CalcData_Sw_Daily;
    // Added by Joe 2015/12/02 17:22:04
    CheckDuplicateTK;
    //�g�J�έp���v���
    SaveAcdSiteDaily;
    SaveAcdTeDaily;
    SaveAcdSwDaily;
    LogLine;
  finally
    FInCalcData := False;
    EndExecute;
    connReport.Close;
  end;
end;

function TdmAcdSummary.MakeAdminNotifyMessage: TIdMessage;
begin
  // nothing to do now
  Result := nil;
end;

procedure TdmAcdSummary.MakeCCList(AEmailAddrList: TIdEmailAddressList);
begin
  with AEmailAddrList do
  begin
    if fmMain.DebugMode then
    begin
      Add.Address := 'joe0107@gmail.com';
      Add.Address := 'wintonjoelee@gmail.com';
    end
    else
    begin
      Add.Address := 'Tony@winton.com.tw';
      Add.Address := 'trista62@winton.com.tw';
      Add.Address := 'sky@winton.com.tw';
      Add.Address := 'jamesjuan@winton.com.tw';
      Add.Address := 'ray@winton.com.tw';
      Add.Address := 'joelee@winton.com.tw';
    end;
  end;
end;

function TdmAcdSummary.MakeNotifyMessage: TIdMessage;
var
  aDayOfWeek: string;
begin
  Result := TIdMessage.Create(Self);
  aDayOfWeek := GetChineseNumStr(DayOfWeek(FCalcInBegTime) - 1);
  if (aDayOfWeek = '�s') then aDayOfWeek := '��'; 

  with Result do
  begin
    //��J�����
    if fmMain.DebugMode then
    begin
      Recipients.Add.Address := 'joelee@winton.com.tw';
      //Recipients.Add.Address := 'f07@winton.com.tw';
    end
    else
    begin
      Recipients.Add.Address := 'orderchen@winton.com.tw';
      Recipients.Add.Address := 'ericl@winton.com.tw';
    end;
    //��J�ƥ�
    MakeCCList(CCList);
    //��J�l����Y��T
    Subject := Format('��~�רӹq�q���_%s(%s)', [FormatDateTime('yyyymmdd', FCalcInBegTime), aDayOfWeek]);
    //�H��H�a�}
    From.Address := 'rdrepl@winton.com.tw';
    ContentType := 'multipart/mixed';
    //��J�l�󤺮e
    Body.Text := '';
  end;

  if FileExists(FXlsFileName) then
  	TIdAttachmentFile.Create(Result.MessageParts, FXlsFileName);
end;

procedure TdmAcdSummary.SendMail;
var
  aMsg: TIdMessage;
begin
  if fmMain.MailMode then
  begin
    aMsg := MakeNotifyMessage;
    dmReport.SendNotofyMail_SSL(aMsg);
    Log('�w�z�L�l��ǰe����');
  end;
end;

function TdmAcdSummary.Is_OffDuty_TE: Boolean;
var
  aTime: TDateTime;
begin
  Result := False;
  //�p�G�OTE�פJ���d���A�ӥB�w�W�L�Ⱦ��ɶ��A���C�J�p��
  if (mdAcdSrcIPHE019.AsString <> CALLIN_KIND_TE) then Exit;
  aTime := TimeOf(mdAcdSrcIPHE004.AsDateTime);
  //�u�n�W�L�]���K�I�A�@�ߤ��C�J�p��
  Result := (aTime > EncodeTime(20, 0, 0, 0));
  if Result then Exit;
  //�p�G���O�]���ȯZ��A�W�L17:45���ӹq�d���@�ߤ��C�J�p��
  if not IsNightShift(DateOf(mdAcdSrcIPHE004.AsDateTime)) then
    Result := (aTime > EncodeTime(17, 45, 0, 0));
end;

procedure TdmAcdSummary.PrepreData_WICSIPHH(ABegTime, AEndTime: TDateTime);
begin
  with qrWICSIPHH do
  begin
    ParamByName('IPHH001B').AsDateTime := ABegTime;
    ParamByName('IPHH001E').AsDateTime := AEndTime;
    if Active then Refresh else Open;
  end;
end;

function TdmAcdSummary.IsNationalHoliday(ADate: TDateTime): Boolean;
var
  A: Variant;
begin
  Result := False;

  with qrWICSIPHH do
  begin
    if not Active then Exit;
    A := VarArrayOf([ADate, '��w����']);

    if Locate('IPHH001;IPHH002', A, []) then
      Result := True;
  end;
end;

function TdmAcdSummary.IsNightShift(ADate: TDateTime): Boolean;
var
  A: Variant;
begin
  Result := False;

  with qrWICSIPHH do
  begin
    if not Active then Exit;
    A := VarArrayOf([ADate, '�]���ȯZ']);

    if Locate('IPHH001;IPHH002', A, []) then
      Result := True;
  end;
end;

procedure TdmAcdSummary.FillData_TeDutyItem(ASiteId: string; ABegTime, AEndTime: TDateTime);
const
  SEARCH_FIELDS = 'CHEM001;CHEM004;CHEM002';
var
  aData: TUniQuery;
  A: Variant;
  aDeptId: string;
begin
  aData := GetTeDutyItem(ASiteID, ABegTime, AEndTime);

  try
    with mdAcdTeDaily do
    begin
      First;

      while not Eof do
      begin
        aDeptId := Copy(mdAcdTeDailyDeptId.AsString, 1, 2);
        //�վ�s�˻P�����������N��
        if (aDeptId = '06') then
          aDeptId := '05'
        else if (aDeptId = '09') then
          aDeptId := '08';
        //�u���s�ثe��~�B���H����ơA�_�h�i���л\���e�w���T�p�⪺��L��~�B�H�����Ⱦ�����
        if (aDeptId = ASiteID) then
        begin
          Edit;
          // �W�Ȫ��u�@����
          A := VarArrayOf([mdAcdTeDailyPhoneDate.AsDateTime, mdAcdTeDailyEmpId.AsString, '10']);

          if aData.Locate(SEARCH_FIELDS, A, []) then
            mdAcdTeDailyDuty_AM.AsString := aData.FieldByName('STM2003').AsString;
          // �U�Ȫ��u�@����
          A := VarArrayOf([mdAcdTeDailyPhoneDate.AsDateTime, mdAcdTeDailyEmpId.AsString, '20']);

          if aData.Locate(SEARCH_FIELDS, A, []) then
            mdAcdTeDailyDuty_PM.AsString := aData.FieldByName('STM2003').AsString;

          if Modified then
            Post
          else
            Cancel;
        end;
        Next;
      end;
    end;
  finally
    dmReport.CloseAndFree(aData);
  end;
end;

function TdmAcdSummary.GetTeDutyItem(ASiteId: string; ABegTime, AEndTime: TDateTime): TUniQuery;
begin
  Result := dmReport.GetQuery(connReport);

  with Result do
  begin
    SQL.Clear;
    SQL.Add('SELECT CHEM001, CHEM002, CHEM004, SALE002, STM2003');
    SQL.Add('FROM WICSCHEM A WITH(NOLOCK)');
    SQL.Add('LEFT JOIN WICSSALE S WITH(NOLOCK) ON SALE001 = CHEM004');
    SQL.Add('LEFT JOIN WICSSTM2 T WITH(NOLOCK) ON (T.STM2001 = CHEM005 AND T.STM2002 = CHEM006)');
    SQL.Add('LEFT JOIN WICSDEPT D WITH(NOLOCK) ON D.DEPT001 = S.SALE003');
    SQL.Add('LEFT JOIN WICSCLAS C WITH(NOLOCK) ON (D.DEPT005 = C.CLAS002 AND C.CLAS001 = ''B0'')');
    AddWhere('(CHEM001 >= :CHEM001B AND CHEM001 <= :CHEM001E)');
    AddWhere('(C.CLAS004 = ''�V�m��'')');

    ParamByName('CHEM001B').Value := DateOf(ABegTime);
    ParamByName('CHEM001E').Value := DateOf(AEndTime);
    Open;
  end;
end;

end.
