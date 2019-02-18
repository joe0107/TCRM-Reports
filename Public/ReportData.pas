unit ReportData;

interface

uses
  SysUtils, Classes, UniProvider, SQLServerUniProvider, DB, DBAccess, Forms, IniFiles, cxStyles, cxCustomData, cxData,
  cxDataStorage, cxEdit, cxDBData, cxClasses, Uni, MemDS, Variants, SQLiteUniProvider, IdEMailAddress, IdMessage,
  IdBaseComponent, IdComponent, IdSMTP, IdTCPConnection, IdTCPClient, IdMessageClient, IdExplicitTLSClientServerBase,
  IdSMTPBase, IdIOHandler, IdIOHandlerSocket, IdSSLOpenSSL, IdIOHandlerStack, IdSSL, IdAttachmentFile, Types, DateUtils,
  TcrmConstants;

type
  TdmReport = class(TDataModule)
    SQLServerUniProvider: TSQLServerUniProvider;
    connTeleContact: TUniConnection;
    UniConnTCRM: TUniConnection;
    UniConnWinton: TUniConnection;
    connTcrmPublic: TUniConnection;
    qrCust: TUniQuery;
    qrCustCUT1001: TStringField;
    qrCustCUT1002: TStringField;
    dsCust: TDataSource;
    qrCustFLAG_SW: TStringField;
    qrCustFLAG_HRS: TStringField;
    qrCustFLAG_HW: TStringField;
    qrTitle: TUniQuery;
    dsTitle: TDataSource;
    qrTitleCLAS002: TStringField;
    qrTitleCLAS004: TStringField;
    qrDbVersion: TUniQuery;
    dsDbVersion: TDataSource;
    qrDbVersionGUID: TStringField;
    qrDbVersionOPTN000: TStringField;
    qrDbVersionOPTN001: TStringField;
    qrDbVersionOPTN002: TStringField;
    qrDbVersionOPTN003: TStringField;
    qrDbVersionOPTN004: TStringField;
    qrEmp: TUniQuery;
    dsEmp: TDataSource;
    qrEmpSALE001: TStringField;
    qrEmpSALE002: TStringField;
    dsDept: TDataSource;
    qrDept: TUniQuery;
    qrDeptDEPT001: TStringField;
    qrDeptDEPT002: TStringField;
    qrClass_A1: TUniQuery;
    dsClass_A1: TDataSource;
    qrClass_A1GUID: TStringField;
    qrClass_A1CLAS002: TStringField;
    qrClass_A1CLAS004: TStringField;
    qrClass_J0: TUniQuery;
    StringField1: TStringField;
    StringField2: TStringField;
    StringField3: TStringField;
    dsClass_J0: TDataSource;
    qrClass_J1: TUniQuery;
    dsClass_J1: TDataSource;
    qrClass_J1GUID: TStringField;
    qrClass_J1CLAS002: TStringField;
    qrClass_J1CLAS003: TStringField;
    qrClass_J1CLAS004: TStringField;
    qrClass_J1CLAS005: TMemoField;
    UniConnLvUpd: TUniConnection;
    qrWICSCUT7: TUniQuery;
    qrEmpSALE_DESC: TStringField;
    SQLiteUniProvider: TSQLiteUniProvider;
    UniConnReport: TUniConnection;
    qrClass_10: TUniQuery;
    qrClass_10GUID: TStringField;
    qrClass_10CLAS002: TStringField;
    qrClass_10CLAS004: TStringField;
    dsClass_10: TDataSource;
    UniConnWcrm: TUniConnection;
    qrTcrmConfig: TUniQuery;
    qrTcrmConfigSiteID: TStringField;
    qrTcrmConfigSite: TStringField;
    qrTcrmConfigBranch: TStringField;
    qrTcrmConfigServer: TStringField;
    qrTcrmConfigDatabase: TStringField;
    qrTcrmConfigTE_Admin_Email: TStringField;
    qrTcrmConfigSite_Admin_Email: TStringField;
    qrTcrmConfigTE_Leader_Email: TStringField;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    FAdminEmail: string;
    function  GetIniFileName: string;
    procedure LoadConfig;
    procedure SaveConfig;
    procedure Init_LocalDB;
    procedure SetAdminEmail(const Value: string);
  public
    // Added by C45 at 2010/11/26 上午 11:17:15
    // TeleContact ReportDB Server address
    INI_RPTDB_Taipei    :string;
    INI_RPTDB_Taichung  :string;
    INI_RPTDB_Taoyuan   :string;
    INI_RPTDB_Tainan    :string;
  public
    procedure SetConn_TeleContact(ASiteName: string);
    function  SetConn_Tcrm(AServer: string): Boolean;
    procedure SetConn_TcrmPublic(AServer: string; ADB: string = 'WCRM'); overload;
    procedure SetConn_LvUpd(AServer: string; AConnected: Boolean = True);
    procedure SetConn_Wcrm(AServer: string; ADB: string = 'Winton_WCRM'); // Added by Joe 2016/03/15 17:35:18
    function  GetConnection(AConn: TUniConnection = nil): TUniConnection;
    function  GetBranchHostIp(ABranch: string): string;
    function  GetSiteIp(ASite: string): string;
  public
    procedure InitDatSet(ADataSet: TDataSet);
    procedure CloseAndFree(var ADataSet: TCustomUniDataSet); overload;
    procedure CloseAndFree(var ADataSet: TUniQuery); overload;
    //-------------------------------------------------------------------------
    function	GetCmd(AConn: TUniConnection = nil): TUniSQL;
    function	GetCmd_Tcrm: TUniSQL;
    function	GetCmd_LvUpd: TUniSQL;
    function	GetCmd_Local: TUniSQL;
    function	GetCmd_WintonTcrm: TUniSQL;
    //-------------------------------------------------------------------------
    function	GetQuery(AConn: TUniConnection = nil): TUniQuery;
    function	GetQuery_Tcrm: TUniQuery;
    function	GetQuery_TcrmPublic: TUniQuery;
    function	GetQuery_LvUpd: TUniQuery;
    function	GetQuery_WintonTcrm: TUniQuery;
    function	GetQuery_TeleContact: TUniQuery;
  public
    procedure InitLookup(ADataSet: TDataSet);
    procedure InitLookup_Cust;
    procedure InitLookup_Title;
    procedure InitLookup_DbVersion(ADbType: string = '');
    procedure InitLookup_Class_A1;
    procedure InitLookup_Class_J0;
    procedure InitLookup_Class_J1;
    procedure InitLookup_Class_10;
    procedure InitLookup_Dept;
    procedure InitLookup_Emp;
    procedure InitLookup_WICSCUT7;
  public
    function  ClassJ0_004_002(ACLAS004: string): string;  //WSTP -> 01
    function  Locate_WICSCUT7(const KeyFields: String; const KeyValues: Variant; Options: TLocateOptions): Boolean;
    function  Lookup_WICSCUT7(const KeyFields: String; const KeyValues: Variant; const ResultFields: String): Variant;
    function  Lookup_Emp(const KeyFields: String; const KeyValues: Variant; const ResultFields: String): Variant;
  public
    procedure Init_IdMessage(AMsg: TIdMessage);
    function  GetEmail_TE_Admin(const ASite: string): string;
    function  GetEmail_TE_Leader(const ASite: string): string;  // Added by Joe 2017/09/12 17:28:08
    function  GetEmail_Site_Admin(const ASite: string): string;
    function  GetAllEmail_TE_Admin: TStringList;
    function  GetAllEmail_TE_Leader: TStringList; // Added by Joe 2017/09/12 17:30:15
    function  GetAllEmail_Site_Admin: TStringList;
  public
    procedure StartTrans(AConn: TUniConnection = nil);
    function  InTrans(AConn: TUniConnection = nil): Boolean;
    procedure Commit(AConn: TUniConnection = nil);
    procedure BatchCommit(var ACount: Integer; const ACommitMax: Integer = 200; AConn: TUniConnection = nil); overload;
    procedure BatchCommit(AConn: TUniConnection = nil); overload;
    procedure Rollback(AConn: TUniConnection = nil);
    procedure CloseConn(AConn: TUniConnection = nil);
    //-------------------------------------------------------------------------
    procedure StartTrans_Tcrm;
    function  InTrans_Tcrm: Boolean;
    procedure Commit_Tcrm;
    procedure BatchCommit_Tcrm(var ACount: Integer; const ACommitMax: Integer = 200); overload;
    procedure BatchCommit_Tcrm; overload;
    procedure Rollback_Tcrm;
    procedure CloseConn_Tcrm;
    // Added by C45 2017/05/17 11:34:26
    procedure StartTrans_WintonTcrm;
    function  InTrans_WintonTcrm: Boolean;
    procedure Commit_WintonTcrm;
    procedure BatchCommit_WintonTcrm(var ACount: Integer; const ACommitMax: Integer = 200); overload;
    procedure BatchCommit_WintonTcrm; overload;
    procedure Rollback_WintonTcrm;
    procedure CloseConn_WintonTcrm;
    //-------------------------------------------------------------------------
    //將每周的第一天校正為 N Added by Joe 2016/07/02 09:52:53
    procedure SetDateFirst(ACmd: TUniSQL; N: Integer);
    procedure SetDateFirst_Tcrm(N: Integer);
    //轉換簡短日期描述字串 11/15(日)
    function  GetShortDateDesc(ADate: TDateTime): string;
    //-------------------------------------------------------------------------
    procedure SetUniConn(AConn: TUniConnection; AServer, ADB, AUser, APwd: string);
    procedure SetUniConn_TCRM(AConn: TUniConnection; AServer: string); overload;
    function  SetUniConn_TCRM(ASite: string): Boolean; overload;
    function  SetUniConn_TCRM(ASiteNdx: TWtnSiteNdx): Boolean; overload;
    function 	GetBranchOfDept(ADeptID: string; ABySite: Boolean = True): string;
    function 	Get_Dept_Name(ADeptId: string): string;
    function  GetSiteName(ASiteId: string): string; overload;
    function  GetSiteName(ASiteNdx: TWtnSiteNdx): string; overload;
    function  GetDescOfSite(ASite: string): string;
    function  GetDateDispTextDOW(ADate: TDateTime): string; // Added by Joe 2015/07/24 17:17:51
    //計算回電逾時資訊 2015.07.30
    function  Calc_CalloutDelay(AInTime, AOutTime: TDateTime; var ADelayMin: Integer): Boolean;
    //取得歸屬系統的相關資訊
    function  GetSwName_10(ACode: string): string;  		//20 -> WSTP2000
    //取得報表輸出目錄
    function  GetReportSockFolder: string;
    //計算並檢查執行報表的年月 Added by Joe 2017/09/18 11:52:18
    procedure Calc_YM(var AYear, AMonth: Word);
    //計算指定營業處的訓練師人數
    function  GetTeCount(ASite: string): Integer;
    //統計指定月份的工作(值機)天數 Added by Joe 2017/10/24 10:16:07
    function  GetOnDutyDays(AYear, AMonth: Word): Word; overload;
    //取得指定日的工作(值機)天數 Added by Joe 2017/11/08 16:56:09
    function  GetOnDutyDays(ADate: TDateTime): Word; overload;
  public
    procedure SendNotofyMail_SSL(AMsg: TIdMessage);
  public
    property 	AdminEmail: string read FAdminEmail write SetAdminEmail;
    function  GetTemplatePath: string;
    function  GetReportPath: string;
  end;

const
  ERR_dmMSSQL_NOT_CREATE  = 'dmMSSQL not created yet';
  STR_BEGIN_TO_EXE        = '開始執行';
  STR_BUILD_XLS_RPT       = '產生XLS報表';
  STR_NO_BUILD_RPT        = '不產生報表';
  STR_SND_RPT_BY_MAIL     = '已透過郵件傳送報';
  STR_TE_ADVANCED         = 'TE進階';
  STR_ACD_ON_DUTY         = 'ACD值機';

var
  dmReport: TdmReport;

implementation

uses JcNumUtils, Main;

{$R *.dfm}

{ TdmReport }

procedure TdmReport.SetConn_TeleContact(ASiteName: string);
var
  aServer: string;
begin
  if (ASiteName = SITE_NAME_Taipei) or (ASiteName = SITE_NAME_Taipei_TC)
   or (ASiteName = BRANCH_NAME_Taipei) or (ASiteName = BRANCH_NAME_Taipei_TC) then
    aServer := INI_RPTDB_Taipei
  else if (ASiteName = SITE_NAME_Taichung) or (ASiteName = SITE_NAME_Taichung_TC)
   or (ASiteName = BRANCH_NAME_Taichung) or (ASiteName = BRANCH_NAME_Taichung_TC) then
    aServer := INI_RPTDB_Taichung
  else if (ASiteName = SITE_NAME_Taoyuan) or (ASiteName = SITE_NAME_Taoyuan_TC)
   or (ASiteName = BRANCH_NAME_Taoyuan) or (ASiteName = BRANCH_NAME_Taoyuan_TC) then
    aServer := INI_RPTDB_Taoyuan
  else if (ASiteName = SITE_NAME_Tainan) or  (ASiteName = SITE_NAME_Tainan_TC)
   or (ASiteName = BRANCH_NAME_Tainan) or  (ASiteName = BRANCH_NAME_Tainan_TC) then
    aServer := INI_RPTDB_Tainan
  else
    raise Exception.Create('SetConn_Teletact() error, unknown site name');

  with connTeleContact do
  begin
    if Connected then
      Disconnect;

    Server := aServer;
  end;
end;

procedure TdmReport.LoadConfig;
var
  aIniFile: TIniFile;
begin
  aIniFile := TIniFile.Create(GetIniFileName);

  try
    with aIniFile do
    begin
      // Added by Joe 2017/04/21 08:56:24
      FAdminEmail := ReadString(INI_SEC_ADMIN, INI_NAME_EMAIL, 'joelee@winton.com.tw');
      // Added by C45 at 2010/11/26 上午 11:08:57
      INI_RPTDB_Taipei    := ReadString(INI_SEC_TELECONTACT, SITE_NAME_Taipei, '10.1.2.2');
      INI_RPTDB_Taichung  := ReadString(INI_SEC_TELECONTACT, SITE_NAME_Taichung, '10.5.1.1');
      INI_RPTDB_Taoyuan   := ReadString(INI_SEC_TELECONTACT, SITE_NAME_Taoyuan, '10.3.1.2');
      INI_RPTDB_Tainan    := ReadString(INI_SEC_TELECONTACT, SITE_NAME_Tainan, '10.6.1.200');
    end;
  finally
    aIniFile.Free;
  end;
end;

procedure TdmReport.SaveConfig;
var
  aIniFile: TIniFile;
begin
  aIniFile := TIniFile.Create(GetIniFileName);

  try
    with aIniFile do
    begin
      // Added by Joe 2017/04/21 08:59:42
      WriteString(INI_SEC_ADMIN, INI_NAME_EMAIL, FAdminEmail);
      // Added by C45 at 2010/11/26 上午 11:17:15
      WriteString(INI_SEC_TELECONTACT, SITE_NAME_Taipei, INI_RPTDB_Taipei);
      WriteString(INI_SEC_TELECONTACT, SITE_NAME_Taichung, INI_RPTDB_Taichung);
      WriteString(INI_SEC_TELECONTACT, SITE_NAME_Taoyuan, INI_RPTDB_Taoyuan);
      WriteString(INI_SEC_TELECONTACT, SITE_NAME_Tainan, INI_RPTDB_Tainan);
    end;
  finally
    aIniFile.Free;
  end;
end;

procedure TdmReport.DataModuleCreate(Sender: TObject);
begin
  LoadConfig;
  Init_LocalDB;
end;

procedure TdmReport.DataModuleDestroy(Sender: TObject);
begin
  SaveConfig;
end;

function TdmReport.SetConn_Tcrm(AServer: string): Boolean;
begin
  try
    with UniConnTCRM do
    begin
      if (Server <> AServer) then
      begin
        if Connected then Disconnect;
        Server := AServer;
      end;
      Connect;
    end;
    // Added by C45 2013/12/16 下午 03:50:52
    with UniConnWinton do
    begin
      Server := SERVER_IP_Winton;
      //Connect;
    end;
  except

  end;
  Result := UniConnTCRM.Connected;
end;

function TdmReport.GetCmd(AConn: TUniConnection): TUniSQL;
begin
	Result := TUniSQL.Create(Self);
	Result.Connection := GetConnection(AConn);
end;

function TdmReport.GetQuery(AConn: TUniConnection): TUniQuery;
begin
  Result := TUniQuery.Create(Self);
  Result.Connection := GetConnection(AConn);
end;

function TdmReport.GetQuery_Tcrm: TUniQuery;
begin
	Result := GetQuery(UniConnTCRM);
end;

function TdmReport.GetQuery_TcrmPublic: TUniQuery;
begin
	Result := GetQuery(connTcrmPublic);
end;

procedure TdmReport.SetConn_TcrmPublic(AServer, ADB: string);
begin
	with connTcrmPublic do
  begin
    if Connected then Close;
    Server := AServer;
    Database := ADB;
  end;
end;

procedure TdmReport.InitLookup_Cust;
begin
  InitLookup(qrCust);
end;

procedure TdmReport.InitLookup_Title;
begin
  InitLookup(qrTitle);
end;

procedure TdmReport.InitLookup_DbVersion(ADbType: string);
begin
  InitLookup(qrDbVersion);

  with qrDbVersion do
  begin
    Filtered := False;

    if (ADbType <> '') then
    begin
      Filter := Format('OPTN002 = %s', [QuotedStr(ADbType)]);
      Filtered := True;
    end;
  end;
end;

procedure TdmReport.BatchCommit(AConn: TUniConnection);
begin
  with GetConnection(AConn) do
  begin
    if Connected and InTransaction then
      Commit;
  end;
end;

procedure TdmReport.BatchCommit(var ACount: Integer; const ACommitMax: Integer; AConn: TUniConnection);
begin
  with GetConnection(AConn) do
  begin
    if Connected and InTransaction then
    begin
      Inc(ACount);

      if (ACount >= ACommitMax) then
      begin
        Commit;
        ACount := 0;
      end;
      
      Exit;
    end;
    //初始化批次更新狀態
    ACount := 0;

    if not Connected then
      Connect;
      
    StartTransaction;
  end;
end;

procedure TdmReport.CloseConn(AConn: TUniConnection);
begin
  with GetConnection(AConn) do
  begin
    if Connected then
    	Close;
  end;
end;

procedure TdmReport.Commit(AConn: TUniConnection);
begin
  with GetConnection(AConn) do
  begin
    if Connected and InTransaction then
      Commit;
  end;
end;

function TdmReport.InTrans(AConn: TUniConnection): Boolean;
begin
  with GetConnection(AConn) do
    Result := InTransaction;
end;

procedure TdmReport.Rollback(AConn: TUniConnection);
begin
  with GetConnection(AConn) do
  begin
    if Connected and InTransaction then
      Rollback;
  end;
end;

procedure TdmReport.StartTrans(AConn: TUniConnection);
begin
  with GetConnection(AConn) do
  begin
    if not Connected then
      Connect;

    if not InTransaction then
      StartTransaction;
  end;
end;

function TdmReport.GetConnection(AConn: TUniConnection): TUniConnection;
begin
  if (AConn = nil) then
		Result := UniConnTCRM
  else
    Result := AConn;
end;

procedure TdmReport.InitLookup(ADataSet: TDataSet);
begin
  with ADataSet do
  begin
    if not Active then
      Open
    else
      Refresh;
  end;
end;

procedure TdmReport.InitLookup_Class_A1;
begin
  InitLookup(qrClass_A1);
end;

procedure TdmReport.InitLookup_Class_J0;
begin
  InitLookup(qrClass_J0);
end;

procedure TdmReport.InitLookup_Class_J1;
begin
  InitLookup(qrClass_J1);
end;

procedure TdmReport.InitLookup_Dept;
begin
  InitLookup(qrDept);
end;

procedure TdmReport.InitLookup_Emp;
begin
  InitLookup(qrEmp);
end;

procedure TdmReport.CloseAndFree(var ADataSet: TCustomUniDataSet);
begin
  with ADataSet do
  begin
    if ADataSet = nil then Exit;
    if Active then Close;
    Free;
    ADataSet := nil;
  end;
end;

procedure TdmReport.CloseAndFree(var ADataSet: TUniQuery);
begin
  CloseAndFree(TCustomUniDataSet(ADataSet));
end;

function TdmReport.GetCmd_Tcrm: TUniSQL;
begin
	Result := GetCmd(UniConnTCRM);
end;

procedure TdmReport.StartTrans_Tcrm;
begin
  StartTrans(UniConnTCRM);
end;

function TdmReport.InTrans_Tcrm: Boolean;
begin
  Result := InTrans(UniConnTCRM);
end;

procedure TdmReport.Commit_Tcrm;
begin
  Commit(UniConnTCRM);
end;

procedure TdmReport.BatchCommit_Tcrm(var ACount: Integer; const ACommitMax: Integer);
begin
  BatchCommit(ACount, ACommitMax, UniConnTCRM);
end;

procedure TdmReport.BatchCommit_Tcrm;
begin
  BatchCommit(UniConnTCRM);
end;

procedure TdmReport.Rollback_Tcrm;
begin
  Rollback(UniConnTCRM);
end;

procedure TdmReport.CloseConn_Tcrm;
begin
  CloseConn(UniConnTCRM);
end;

procedure TdmReport.SetConn_LvUpd(AServer: string; AConnected: Boolean);
begin
  with UniConnLvUpd do
  begin
    if Connected then Disconnect;
    Server := AServer;
    if AConnected then Connect;
  end;
end;

function TdmReport.ClassJ0_004_002(ACLAS004: string): string;
begin
  InitLookup_Class_J0;
  ACLAS004 := UpperCase(ACLAS004);

  if Pos('HR', ACLAS004) > 0 then
    ACLAS004 := PID_HRS
  else if Pos(PID_NTF, ACLAS004) > 0 then
    ACLAS004 := PID_WSTF
  else if (ACLAS004 = PID_MERP) then
    ACLAS004 := PID_WMIS; 

  with qrClass_J0 do
  begin
    Result := VarToStr(Lookup('CLAS004', ACLAS004, 'CLAS002'));
  end;
end;

function TdmReport.Locate_WICSCUT7(const KeyFields: String; const KeyValues: Variant; Options: TLocateOptions): Boolean;
begin
  InitLookup_WICSCUT7;
  Result := qrWICSCUT7.Locate(KeyFields, KeyValues, Options);
end;

procedure TdmReport.InitLookup_WICSCUT7;
begin
  InitLookup(qrWICSCUT7);
end;

function TdmReport.Lookup_WICSCUT7(const KeyFields: String; const KeyValues: Variant; const ResultFields: String): Variant;
begin
  InitLookup_WICSCUT7;
  Result := qrWICSCUT7.Lookup(KeyFields, KeyValues, ResultFields);
end;

function TdmReport.GetCmd_LvUpd: TUniSQL;
begin
	Result := GetCmd(UniConnLvUpd);
end;

procedure TdmReport.Init_IdMessage(AMsg: TIdMessage);
begin
  with AMsg do
  begin
    //設定郵件屬性
    ContentType := 'multipart/mixed';
    CharSet :=  'UTF-8';
    //寄件人地址
    From.Address := 'rdrepl@winton.com.tw';
  end;
end;

procedure TdmReport.Init_LocalDB;
begin
  with UniConnReport do
  begin
    if Connected then Close;
    Database := ExtractFilePath(Application.ExeName) + 'TcrmReport.db';
  end;
end;

function TdmReport.GetCmd_Local: TUniSQL;
begin
	Result := GetCmd(UniConnReport);
end;

procedure TdmReport.InitLookup_Class_10;
begin
  InitLookup(qrClass_10);
end;

function TdmReport.GetQuery_LvUpd: TUniQuery;
begin
	Result := GetQuery(UniConnLvUpd);
end;

procedure TdmReport.SetConn_Wcrm(AServer, ADB: string);
begin
  with UniConnWcrm do
  begin
    if Connected then Disconnect;
    Server := AServer;
    Database := ADB;
  end;
end;

function TdmReport.GetQuery_WintonTcrm: TUniQuery;
begin
	Result := GetQuery(UniConnWinton);
end;

function TdmReport.GetCmd_WintonTcrm: TUniSQL;
begin
	Result := GetCmd(UniConnWinton);
end;

procedure TdmReport.SetDateFirst(ACmd: TUniSQL; N: Integer);
begin
  if (ACmd = nil) or (N < 1) or (N > 7) then
    Exit;

  with ACmd do
  begin
    SQL.Clear;
    SQL.Add(Format('SET DATEFIRST %d', [N]));
    Execute;
  end;
end;

procedure TdmReport.SetDateFirst_Tcrm(N: Integer);
var
  aCmd: TUniSQL;
begin
  aCmd := GetCmd_Tcrm;

  try
    SetDateFirst(aCmd, N);
  finally
    aCmd.Free;
  end;
end;

function TdmReport.GetAllEmail_Site_Admin: TStringList;
var
  aMail: string;
begin
  Result := TStringList.Create;

	with qrTcrmConfig do
  begin
    if not Active then Open;
    First;

    while not Eof do
    begin
      aMail := Trim(qrTcrmConfigSite_Admin_Email.AsString);

      if (aMail <> '') and (Result.IndexOf(aMail) = -1) then
        Result.Add(aMail);

      Next;
    end;
  end;
end;

function TdmReport.GetAllEmail_TE_Admin: TStringList;
var
  aMail: string;
begin
  Result := TStringList.Create;

	with qrTcrmConfig do
  begin
    if not Active then Open;
    First;

    while not Eof do
    begin
      aMail := Trim(qrTcrmConfigTE_Admin_Email.AsString);

      if (aMail <> '') and (Result.IndexOf(aMail) = -1) then
        Result.Add(aMail);

      Next;
    end;
  end;
end;

function TdmReport.GetAllEmail_TE_Leader: TStringList;
var
  aMail: string;
begin
  Result := TStringList.Create;

	with qrTcrmConfig do
  begin
    if not Active then Open;
    First;

    while not Eof do
    begin
      aMail := Trim(qrTcrmConfigTE_Leader_Email.AsString);

      if (aMail <> '') and (Result.IndexOf(aMail) = -1) then
        Result.Add(aMail);

      Next;
    end;
  end;
end;

function TdmReport.GetBranchHostIp(ABranch: string): string;
begin
  with qrTcrmConfig do
  begin
    if not Active then Open;

    if Locate('Branch', ABranch, []) then
    	Result := FieldByName('Server').AsString
    else
    	raise Exception.CreateFmt('GetBranchHostIp() error, Branch = %s', [ABranch]);
	end;
end;

function TdmReport.GetEmail_Site_Admin(const ASite: string): string;
begin
	with qrTcrmConfig do
  begin
    if not Active then Open;
    Result := Trim(VarToStr(Lookup('Site', ASite, 'Site_Admin_Email')));
  end;
end;

function TdmReport.GetEmail_TE_Admin(const ASite: string): string;
begin
	with qrTcrmConfig do
  begin
    if not Active then Open;
    Result := Trim(VarToStr(Lookup('Site', ASite, 'TE_Admin_Email')));
  end;
end;

function TdmReport.GetEmail_TE_Leader(const ASite: string): string;
begin
	with qrTcrmConfig do
  begin
    if not Active then Open;
    Result := Trim(VarToStr(Lookup('Site', ASite, 'TE_Leader_Email')));
  end;
end;

procedure TdmReport.SendNotofyMail_SSL(AMsg: TIdMessage);
var
  aSMTP: TIdSMTP;
  aSSL: TIdSSLIOHandlerSocketOpenSSL;
begin
  {$IFNDEF NO_MAIL}
  aSSL := TIdSSLIOHandlerSocketOpenSSL.Create(Self);
  aSMTP := TIdSMTP.Create(Self);

  with aSMTP do
  begin
    //IOHandler := aSSL;
    Host := 'mail.winton.com.tw';
    Port := 587;
    Username := 'rdrepl';
    Password := 'Wint0n2k';
  end;

  try
    try
      aSMTP.Connect;
      aSMTP.Send(AMsg);
    finally
      aSMTP.Disconnect;
      FreeAndNil(aSMTP);
      FreeAndNil(aSSL);
    end;
  except
    on E: Exception do
      fmMain.Log(Format('SendNotofyMail_SSL() err = %s', [E.Message]));
  end;
  {$ENDIF}
end;

function TdmReport.GetIniFileName: string;
begin
  Result := ExtractFileDir(Application.ExeName) + '\TcrmReport.ini';
end;

function TdmReport.GetQuery_TeleContact: TUniQuery;
begin
	Result := GetQuery(connTeleContact);
end;

procedure TdmReport.SetAdminEmail(const Value: string);
begin
  FAdminEmail := Value;
end;

function TdmReport.Lookup_Emp(const KeyFields: String; const KeyValues: Variant; const ResultFields: String): Variant;
begin
  InitLookup_Emp;
  Result := qrEmp.Lookup(KeyFields, KeyValues, ResultFields);
end;

function TdmReport.GetBranchOfDept(ADeptID: string; ABySite: Boolean): string;
var
  aDeptHead: string;
begin
  aDeptHead := Copy(ADeptID, 1, 2);

  if ABySite then
  begin
    if (aDeptHead = BRANCH_ID_Taipei) then
      Result := SITE_DESC_Taipei_TC
    else if (aDeptHead = BRANCH_ID_Taoyuan) or (aDeptHead = BRANCH_ID_Hsinchu) then
      Result := SITE_DESC_Taoyuan_TC
    else if (aDeptHead = BRANCH_ID_Taichung) then
      Result := SITE_DESC_Taichung_TC
    else if (aDeptHead = BRANCH_ID_Tainan) or (aDeptHead = BRANCH_ID_Kaohsiung) then
      Result := SITE_DESC_Tainan_TC
    else if (aDeptHead = BRANCH_ID_Shanghai) or (aDeptHead = BRANCH_ID_Xiamen) then
      Result := SITE_DESC_China_TC
    else
      Result := '';
  end
  else
  begin
    if (aDeptHead = SITE_ID_Taipei) then
      Result := '1.台北'
    else if (aDeptHead = BRANCH_ID_Taoyuan) then
      Result := '2.桃園'
    else if (aDeptHead = BRANCH_ID_Hsinchu) then
      Result := '3.新竹'
    else if (aDeptHead = BRANCH_ID_Taichung) then
      Result := '4.中區'
    else if (aDeptHead = BRANCH_ID_Tainan) then
      Result := '5.台南'
    else if (aDeptHead = BRANCH_ID_Kaohsiung) then
      Result := '6.高雄'
    else
      Result := '';
  end;
end;

function TdmReport.Get_Dept_Name(ADeptId: string): string;
var
  aFound: Boolean;
begin
  Result := '';
  ADeptId := Trim(ADeptId);

  with qrDept do
  begin
    if (not Active) or (ADeptId = '') then Exit;

    if (FieldByName('DEPT001').AsString = ADeptId) then
      aFound := True
    else
      aFound := Locate('DEPT001', ADeptId, []);

    if aFound then
      Result := VarToStr(FieldValues['DEPT002']);
  end;
end;

procedure TdmReport.SetUniConn(AConn: TUniConnection; AServer, ADB, AUser, APwd: string);
begin
	with AConn do
  begin
    if Connected then Close;
    Server := AServer;
    Database := ADB;
    Username := AUser;
    Password := APwd;
  end;
end;

procedure TdmReport.SetUniConn_TCRM(AConn: TUniConnection; AServer: string);
begin
  fmMain.Log(Format('設定TCRM連線, Srv = %s', [AServer]));
  SetUniConn(AConn, AServer, 'WCRM', 'tcrm', 'hellotcrm');
end;

function TdmReport.GetShortDateDesc(ADate: TDateTime): string;
var
  aYear, aMonth, aDay: Word;
  aWeekDay: Integer;
begin
  DecodeDate(ADate, aYear, aMonth, aDay);
  aWeekDay := DayOfWeek(aDate)-1;
  Result := GetChineseDayNumStr(aWeekDay);
  Result := Format('%.2d/%.2d(%s)', [aMonth, aDay, Result]);
end;

function TdmReport.GetSiteIp(ASite: string): string;
begin
  if (Pos(SITE_NAME_Winton_TC, ASite) > 0) or (Pos(SITE_NAME_Winton, ASite) > 0) then
    Result := SERVER_IP_Winton  
  else if (ASite = SITE_ID_Taipei) or (Pos(SITE_NAME_Taipei_TC, ASite) > 0) or (Pos(SITE_NAME_Taipei, ASite) > 0) then
    Result := SERVER_IP_Taipei
  else if (ASite = SITE_ID_Taoyuan) or (Pos(SITE_NAME_Taoyuan_TC, ASite) > 0) or (Pos(SITE_NAME_Taoyuan, ASite) > 0) then
    Result := SERVER_IP_Taoyuan
  else if (ASite = SITE_ID_Taichung) or (Pos(SITE_NAME_Taichung_TC, ASite) > 0) or (Pos(SITE_NAME_Taichung, ASite) > 0) then
    Result := SERVER_IP_Taichung
  else if (ASite = SITE_ID_Tainan)  or (Pos(SITE_NAME_Tainan_TC, ASite) > 0) or (Pos(SITE_NAME_Tainan, ASite) > 0) then
    Result := SERVER_IP_Tainan
  else
  	Result := '';
end;

function TdmReport.GetDateDispTextDOW(ADate: TDateTime): string;
var
  aDayOfWeekName: string;
begin
  aDayOfWeekName := GetChineseNumStr(DayOfWeek(aDate)-1);
  if aDayOfWeekName = '零' then aDayOfWeekName := '日';
  Result := FormatDateTime('MM/DD', aDate) + Format('(%s)', [aDayOfWeekName]);
end;

function TdmReport.Calc_CalloutDelay(AInTime, AOutTime: TDateTime; var ADelayMin: Integer): Boolean;
var
  aValueRelationship: TValueRelationship;
  aMinRest1, aMinRest2, aMinRest3: Integer; //以分鐘為單位
  aMinIn, aMinOut, aDayOfWeek: Integer;     //以分鐘為單位
  aYear, aMonth, aDay, aHour, aMin, aSec, aMSec: Word;
  aNewInTime: TDateTime;
begin
  Result := False;
  if (AOutTime < EncodeDate(2000, 1, 1)) then Exit;
  //計算出各段休息時間
  aMinRest1 := (60 * 12);       //中午 12:00
  aMinRest2 := (60 * 13) + 30;  //下午 13:30
  aMinRest3 := (60 * 17) + 45;  //下午 17:45
  //將來電及回電時間轉為分鐘
  DecodeDateTime(AInTime, aYear, aMonth, aDay, aHour, aMin, aSec, aMSec);
  aMinIn := (60 * aHour) + aMin;
  //調整來電計算基準時間
  //如果是午休時間的來電，重設來電時間為下午上班時間
  //如果下班時間的來電，重設來電時間為明天上班時間
  if (aMinIn >= aMinRest1) and (aMinIn < aMinRest2) then
    aNewInTime := RecodeTime(AInTime, 13, 30, 0, 0)
  else if (aMinIn > aMinRest3) then
    aNewInTime := IncDay(RecodeTime(AInTime, 8, 30, 0, 0))
  else
    aNewInTime := AInTime;

  //檢查重設後的來電時間是否為假日，若是，繼續調整為下一工作日
  aDayOfWeek :=  DayOfWeek(aNewInTime);

  if aDayOfWeek = 1 then // Sunday
    IncDay(aNewInTime, 1)
  else if aDayOfWeek = 7 then // Saturday
    IncDay(aNewInTime, 2);
  //以調整後的來電日與回電日做比較
  aValueRelationship := CompareDate(aNewInTime, AOutTime);

  if (aValueRelationship = EqualsValue) then  //同一天來回電
  begin
    //計算回電延遲時間，單位為分鐘
    if (aNewInTime >= AOutTime) then
      ADelayMin := 0
    else
      ADelayMin := MinutesBetween(aNewInTime, AOutTime);
    //調整回電計算基準時間
    DecodeDateTime(AOutTime, aYear, aMonth, aDay, aHour, aMin, aSec, aMSec);
    aMinOut := (60 * aHour) + aMin;

    if (aMinIn < aMinRest1) then  //午休前來電
    begin
      if (aMinOut >= aMinRest2) then  //午休後回電，扣掉午休時間
        Dec(ADelayMin, 90)
      else if ((aMinOut > aMinRest1) and (aMinOut < aMinRest2)) then  //午休期間回電
        Dec(ADelayMin, aMinOut - 60*12)
    end;

    if ADelayMin < 0 then ADelayMin := 0;

    Result := True;
  end
  else if (aValueRelationship = GreaterThanValue) then
  begin
    ADelayMin := 0;
    Result := True;
  end
  else  //非當日回電
  begin
    ADelayMin := 999;
    Result := True;
  end;
end;

function TdmReport.GetSwName_10(ACode: string): string;
begin
  Result := '';
  if (ACode = '') then Exit;

  if qrClass_10.Locate('CLAS002', ACode, []) then
    Result := Copy(qrClass_10.FieldValues['CLAS004'], 1, 12);
end;

function TdmReport.GetDescOfSite(ASite: string): string;
begin
  if (ASite = SITE_NAME_Taipei) or (Pos(SITE_NAME_Taipei_TC, ASite) > 0) or (Pos(SITE_NAME_Taipei_TC2, ASite) > 0) then
    Result := SITE_DESC_Taipei_TC
  else if (ASite = SITE_NAME_Taichung) or (Pos(SITE_NAME_Taichung_TC, ASite) > 0) or (Pos(SITE_NAME_Taichung_TC2, ASite) > 0) then
    Result := SITE_DESC_Taichung_TC
  else if (ASite = SITE_NAME_Taoyuan) or (Pos(SITE_NAME_Taoyuan_TC, ASite) > 0) then
    Result := SITE_DESC_Taoyuan_TC
  else if (ASite = SITE_NAME_Tainan) or (Pos(SITE_NAME_Tainan_TC, ASite) > 0) then
    Result := SITE_DESC_Tainan_TC
  else if (ASite = SITE_NAME_China) or (Pos(SITE_NAME_China_TC, ASite) > 0) then
    Result := SITE_DESC_China_TC
  else
    Result := SITE_DESC_MISC_TC;
end;

function TdmReport.GetSiteName(ASiteId: string): string;
begin
  if (ASiteId = SITE_ID_Taipei) then
    Result := SITE_NAME_Taipei
  else if (ASiteId = SITE_ID_Taichung) then
    Result := SITE_NAME_Taichung
  else if (ASiteId = SITE_ID_Taoyuan) then
    Result := SITE_NAME_Taoyuan
  else if (ASiteId = SITE_ID_Tainan) then
    Result := SITE_NAME_Tainan
  else if (ASiteId = SITE_ID_Tainan) then
    Result := SITE_NAME_Tainan
  else if (ASiteId = SITE_ID_China) then
    Result := SITE_NAME_China
  else
    raise Exception.Create('GetSiteName() error, unknown site id');
end;

procedure TdmReport.BatchCommit_WintonTcrm;
begin
  BatchCommit(UniConnWinton);
end;

procedure TdmReport.BatchCommit_WintonTcrm(var ACount: Integer; const ACommitMax: Integer);
begin
  BatchCommit(ACount, ACommitMax, UniConnWinton);
end;

procedure TdmReport.CloseConn_WintonTcrm;
begin
  CloseConn(UniConnWinton);
end;

procedure TdmReport.Commit_WintonTcrm;
begin
  Commit(UniConnWinton);
end;

function TdmReport.InTrans_WintonTcrm: Boolean;
begin
  Result := InTrans(UniConnWinton);
end;

procedure TdmReport.Rollback_WintonTcrm;
begin
  Rollback(UniConnWinton);
end;

procedure TdmReport.StartTrans_WintonTcrm;
begin
  StartTrans(UniConnWinton);
end;

function TdmReport.GetReportPath: string;
begin
  Result := ExtractFilePath(Application.ExeName) + 'ReportStock';
end;

function TdmReport.GetTemplatePath: string;
begin
	Result := ExtractFilePath(Application.ExeName) + 'Template';
end;


procedure TdmReport.Calc_YM(var AYear, AMonth: Word);
var
  aY, aM, aDay: Word;
begin
  //如果 Y/M = 0,自動依據當前日期取前一個月為計算值
  if (AYear = 0) or (AMonth = 0) then
  begin
    DecodeDate(IncMonth(Now, -1), aY, aM, aDay);
    AYear := aY;
    AMonth := aM;
  end;
  //防呆
  if (AMonth < 1) then AMonth := 1;
  if (AMonth > 12) then AMonth := 12;
end;

function TdmReport.GetReportSockFolder: string;
begin
  Result := ExtractFilePath(Application.ExeName) + 'ReportStock';
  ForceDirectories(Result);
end;

function TdmReport.GetTeCount(ASite: string): Integer;
var
  aQr: TUniQuery;
  aDeptList: string;
begin
  aQr := GetQuery_WintonTcrm;

  if (ASite = SITE_NAME_Taipei_TC) or (ASite = SITE_NAME_Taipei) then
    aDeptList := Format('%s,%s,%s,%s', [QuotedStr(TE_DEPT_021), QuotedStr(TE_DEPT_022),
      QuotedStr(TE_DEPT_023), QuotedStr(TE_DEPT_026)])
  else if (ASite = SITE_NAME_Taoyuan_TC) or (ASite = SITE_NAME_Taoyuan) then
    aDeptList := Format('%s,%s', [QuotedStr(TE_DEPT_052), QuotedStr(TE_DEPT_062)])
  else if (ASite = SITE_NAME_Taichung_TC) or (ASite = SITE_NAME_Taichung) then
    aDeptList := Format('%s,%s', [QuotedStr(TE_DEPT_075), QuotedStr(TE_DEPT_076)])
  else if (ASite = SITE_NAME_Tainan_TC) or (ASite = SITE_NAME_Tainan) then
    aDeptList := Format('%s,%s', [QuotedStr(TE_DEPT_082), QuotedStr(TE_DEPT_092)]);

  try
    try
      with aQr do
      begin
        SQL.Clear;
        SQL.Add('SELECT COUNT(*) AS _RC_ FROM WICSSALE WITH(NOLOCK)');
        AddWhere(Format('(SALE003 IN (%s))', [aDeptList]));
        AddWhere('(SALE008 IS NULL)');
        Open;
        Result := FieldByName('_RC_').AsINteger;
      end;
    except
      Result := 0;
    end;
  finally
    CloseAndFree(aQr);
  end;
end;

function TdmReport.GetOnDutyDays(AYear, AMonth: Word): Word;
var
  aQr: TUniQuery;
begin
  aQr := GetQuery_Tcrm;

  try
    try
      with aQr do
      begin
        SQL.Add('SELECT COUNT(DISTINCT CHEM001) AS _COUNT_');
        SQL.Add('FROM WICSCHEM WITH(NOLOCK)');
        SQL.Add('LEFT JOIN WICSSTM2 WITH(NOLOCK) ON STM2001 = CHEM005 AND STM2002 = CHEM006');
        AddWhere('(STM2004 = ''Y'')');
        AddWhere('(CHEM001 >= :CHEM001B AND CHEM001 < :CHEM001E)');
        ParamByName('CHEM001B').Value := StartOfAMonth(AYear, AMonth);
        ParamByName('CHEM001E').Value := EndOfAMonth(AYear, AMonth);
        Open;
        Result := FieldByName('_COUNT_').AsInteger;
      end;
    except
      Result := 0;
    end;
  finally
    CloseAndFree(aQr);
  end;
end;

function TdmReport.SetUniConn_TCRM(ASite: string): Boolean;
var
  aHostIp: string;
begin
  aHostIp := GetSiteIp(ASite);
  Result := SetConn_Tcrm(aHostIp);
end;

function TdmReport.SetUniConn_TCRM(ASiteNdx: TWtnSiteNdx): Boolean;
begin
  if (ASiteNdx = wsnTaipei) then
    Result := SetUniConn_TCRM(SITE_NAME_Taipei_TC)
  else if (ASiteNdx = wsnTaoyuan) then
    Result := SetUniConn_TCRM(SITE_NAME_Taoyuan_TC)
  else if (ASiteNdx = wsnTaichung) then
    Result := SetUniConn_TCRM(SITE_NAME_Taichung_TC)
  else if (ASiteNdx = wsnTainan) then
    Result := SetUniConn_TCRM(SITE_NAME_Tainan_TC)
  else
    Result := False;
end;

function TdmReport.GetSiteName(ASiteNdx: TWtnSiteNdx): string;
begin
  if (ASiteNdx = wsnTaipei) then
    Result := SITE_NAME_Taipei
  else if (ASiteNdx = wsnTaoyuan) then
    Result := SITE_NAME_Taoyuan
  else if (ASiteNdx = wsnTaichung) then
    Result := SITE_NAME_Taichung
  else if (ASiteNdx = wsnTainan) then
    Result := SITE_NAME_Tainan
  else
    raise Exception.Create('GetSiteName() error, unknown site index');
end;

procedure TdmReport.InitDatSet(ADataSet: TDataSet);
begin
  with ADataSet do
  begin
    DisableControls;
    if Active then Close;
    Open;
  end;
end;

function TdmReport.GetOnDutyDays(ADate: TDateTime): Word;
var
  aQr: TUniQuery;
begin
  aQr := GetQuery_Tcrm;

  try
    try
      with aQr do
      begin
        SQL.Add('SELECT COUNT(DISTINCT CHEM001) AS _COUNT_');
        SQL.Add('FROM WICSCHEM WITH(NOLOCK)');
        SQL.Add('LEFT JOIN WICSSTM2 WITH(NOLOCK) ON STM2001 = CHEM005 AND STM2002 = CHEM006');
        AddWhere('(STM2004 = ''Y'')');
        AddWhere('(CHEM001 >= :CHEM001B AND CHEM001 < :CHEM001E)');
        ParamByName('CHEM001B').Value := StartOfTheDay(ADate);
        ParamByName('CHEM001E').Value := EndOfTheDay(ADate);
        Open;
        Result := FieldByName('_COUNT_').AsInteger;
      end;
    except
      Result := 0;
    end;
  finally
    CloseAndFree(aQr);
  end;

end;

end.
