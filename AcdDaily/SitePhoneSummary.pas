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
    //�������o��ƨӷ������γs�u
    procedure InitReportConn;
    procedure InitData;
    procedure Log(AMsg: string);
    procedure LogLine;
  private
    procedure GetData(ASite: string; ABegTime, AEndTime: TDateTime);
    procedure GetData_Summary(ASite: string; ABegTime, AEndTime: TDateTime);
    //���o���w������Ⱦ��`�H��
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
    // ���ͫ��w����� XLS ����
    class procedure Exec(ADate: TDateTime);
	end;

var
  dmSitePhoneSummary: TdmSitePhoneSummary;

implementation

uses
  TcrmConstants, JcDateTimeUtils, JcNumUtils, JcDevExpressUtils, JcDataSetUtils, ReportData, Main;

const
  RPT_FILE_TITLE = '��~�B�^�q�Ĳv�έp��';
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
  Log(Format('�}�l����[%s]', [RPT_FILE_TITLE]));
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
    // �q�夤��Ʈw���o�έp�ƾ�
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
    // �ˬd���L[ACD�B�z��]
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
    Log(Format('�LACD�B�z�Ƹ�ơA�����ͳ���[%s]', [ASite]));
    Exit;
  end;
  //---------------------------------------------------------------------------
  Log(Format('����XLS����[%s]', [ASite]));

  with XLSRW do
  begin
    // �[�J���w��~�B���u�@��
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
  WORK_SHEET_NAME = '�έp��';
var
  aSheet: TXLSWorksheet;
  aText: string;
begin
  with XLSRW do
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
    Log('����L�H�Ⱦ��A�����ͳ���');
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
  if (aDayOfWeek = '�s') then aDayOfWeek := '��'; 

  with Result do
  begin
    //��J�����
    MakeRecipients(Recipients);
    //��J�ƥ�
    MakeCCList(CCList);
    //��J�l����Y��T
    Subject := Format('%s_%s(%s)', [RPT_FILE_TITLE, FormatDateTime('yyyymmdd', FCalcEndTime), aDayOfWeek]);
    //�H��H�a�}
    From.Address := 'rdrepl@winton.com.tw';
    ContentType := 'multipart/mixed';
    //��J�l�󤺮e
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
    Log('�w�z�L�l��ǰe����');
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
    AsString[RPT_START_COL, RPT_START_ROW]      := '���';
    MergeCells(RPT_START_COL+1, RPT_START_ROW, RPT_START_COL+1, RPT_START_ROW+1);
    AsString[RPT_START_COL+1, RPT_START_ROW]    := '�Ⱦ�'+#10+'�H��';
    //--------------------------------------------------------------------------
    MergeCells(RPT_START_COL+2, RPT_START_ROW, RPT_START_COL+10, RPT_START_ROW);
    AsString[RPT_START_COL+2, RPT_START_ROW]    := '�X��';
    AsString[RPT_START_COL+2, RPT_START_ROW+1]  := '�ӹq��';
    AsString[RPT_START_COL+3, RPT_START_ROW+1]  := '�^�q��';
    AsString[RPT_START_COL+4, RPT_START_ROW+1]  := 'ACD'+#10+'���Ĭ��e';
    AsString[RPT_START_COL+5, RPT_START_ROW+1]  := 'ACD'+#10+'������';
    AsString[RPT_START_COL+6, RPT_START_ROW+1]  := 'ACD'+#10+'�B�z��';
    AsString[RPT_START_COL+7, RPT_START_ROW+1]  := 'ACD'+#10+'��ť�v';
    AsString[RPT_START_COL+8, RPT_START_ROW+1]  := '�O�ɼ�';
    AsString[RPT_START_COL+9, RPT_START_ROW+1]  := '�O�ɲv';
    AsString[RPT_START_COL+10, RPT_START_ROW+1] := '���^��';
    //--------------------------------------------------------------------------
    MergeCells(RPT_START_COL+11, RPT_START_ROW, RPT_START_COL+15, RPT_START_ROW);
    AsString[RPT_START_COL+11, RPT_START_ROW]   := '�D�X��';
    AsString[RPT_START_COL+11, RPT_START_ROW+1] := '�ӹq��';
    AsString[RPT_START_COL+12, RPT_START_ROW+1] := '�^�q��';
    AsString[RPT_START_COL+13, RPT_START_ROW+1] := '�O�ɼ�';
    AsString[RPT_START_COL+14, RPT_START_ROW+1] := '�O�ɲv';
    AsString[RPT_START_COL+15, RPT_START_ROW+1] := '���^��';
    //--------------------------------------------------------------------------
    MergeCells(RPT_START_COL+16, RPT_START_ROW, RPT_START_COL+18, RPT_START_ROW);
    AsString[RPT_START_COL+16, RPT_START_ROW]   := '����';
    AsString[RPT_START_COL+16, RPT_START_ROW+1] := AsString[RPT_START_COL+5, RPT_START_ROW+1];
    AsString[RPT_START_COL+17, RPT_START_ROW+1] := AsString[RPT_START_COL+6, RPT_START_ROW+1];
    AsString[RPT_START_COL+18, RPT_START_ROW+1] := '�^�q��';
  end;
  // �]�w��e
  with ASheet do
  begin
    // ���
    Columns[RPT_START_COL].CharWidth := 10;
    // �Ⱦ��H��
    Columns[RPT_START_COL+1].CharWidth := 8;
  end;
  //============================================================================
  aCmdFormat := AXLS.CmdFormat;
  ASheet.CalcDimensions;

  with aCmdFormat do
  begin
    //�e�X�~�ءA���D�m���A���D�C��
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
  Log(Format('���o��~�B�Ӧ^�q���[%s]', [ASite]));

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
      // ���L�g��(��)�����
      // ���L��鱵ť�Ƭ�0�����(��ѫD�W�Z��)
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
        // ���
        AsString[RPT_START_COL, aRow] := aText;
        // �Ⱦ��H��
        AsFloat[RPT_START_COL+1, aRow] := FieldByName('IPH4003').AsFloat;
        Cell[RPT_START_COL+1, aRow].NumberFormat := '###0.0';
        //--------------------------------------------------------------------------
        // �ӹq�q��
        AsInteger[RPT_START_COL+2, aRow] := FieldByName('IPH4011').AsInteger;
        Cell[RPT_START_COL+2, aRow].NumberFormat := INT_FMT;
        // �^�q�q��
        AsInteger[RPT_START_COL+3, aRow] := FieldByName('IPH4009').AsInteger;
        Cell[RPT_START_COL+3, aRow].NumberFormat := INT_FMT;
        // ACD���Ĭ��e
        AsInteger[RPT_START_COL+4, aRow] := FieldByName('IPH4007').AsInteger;
        Cell[RPT_START_COL+4, aRow].NumberFormat := INT_FMT;
        // ACD������ť�q��
        AsInteger[RPT_START_COL+5, aRow] := FieldByName('IPH4012').AsInteger;
        Cell[RPT_START_COL+5, aRow].NumberFormat := INT_FMT;
        // ACD�B�z�q��
        AsInteger[RPT_START_COL+6, aRow] := FieldByName('IPH4008').AsInteger;
        Cell[RPT_START_COL+6, aRow].NumberFormat := INT_FMT;
        // ACD��ť�v
        //AsFloat[RPT_START_COL+7, aRow] := FieldByName('IPH4004').AsFloat/100;
        AsFloat[RPT_START_COL+7, aRow] := JcDivide(FieldByName('IPH4008').AsInteger, FieldByName('IPH4007').AsInteger);
        Cell[RPT_START_COL+7, aRow].NumberFormat := '##0.0 %';
        // �O�ɳq��
        AsInteger[RPT_START_COL+8, aRow] := FieldByName('IPH4010').AsInteger;
        Cell[RPT_START_COL+8, aRow].NumberFormat := INT_FMT;
        // �O�ɲv
        //AsFloat[RPT_START_COL+9, aRow] := FieldByName('IPH4005').AsFloat/100;
        AsFloat[RPT_START_COL+9, aRow] := JcDivide(FieldByName('IPH4010').AsInteger, FieldByName('IPH4011').AsInteger);
        Cell[RPT_START_COL+9, aRow].NumberFormat := '##0.0 %';
        // ���^�q��
        AsInteger[RPT_START_COL+10, aRow] := FieldByName('IPH4013').AsInteger;
        Cell[RPT_START_COL+10, aRow].NumberFormat := INT_FMT;
        //--------------------------------------------------------------------------
        // �ӹq�q��(�D�X��)
        AsInteger[RPT_START_COL+11, aRow] := FieldByName('IPH4211').AsInteger;
        Cell[RPT_START_COL+11, aRow].NumberFormat := INT_FMT;
        // �^�q�q��(�D�X��)
        AsInteger[RPT_START_COL+12, aRow] := FieldByName('IPH4209').AsInteger;
        Cell[RPT_START_COL+12, aRow].NumberFormat := INT_FMT;
        // �O�ɳq��(�D�X��)
        AsInteger[RPT_START_COL+13, aRow] := FieldByName('IPH4210').AsInteger;
        Cell[RPT_START_COL+13, aRow].NumberFormat := INT_FMT;
        // �O�ɲv(�D�X��)
        //AsFloat[RPT_START_COL+14, aRow] := FieldByName('IPH4205').AsFloat/100;
        AsFloat[RPT_START_COL+14, aRow] := JcDivide(FieldByName('IPH4210').AsInteger, FieldByName('IPH4211').AsInteger);
        Cell[RPT_START_COL+14, aRow].NumberFormat := '##0.0 %';
        // ���^�q��(�D�X��)
        AsInteger[RPT_START_COL+15, aRow] := FieldByName('IPH4213').AsInteger;
        Cell[RPT_START_COL+15, aRow].NumberFormat := INT_FMT;
        //--------------------------------------------------------------------------
        // ACD������ť�q��(����)
        aCellRef1 := ColRowToRefStr(RPT_START_COL+5, aRow);
        aCellRef2 := ColRowToRefStr(RPT_START_COL+1, aRow);
        AsFormula[RPT_START_COL+16, aRow] := Format('%s/%s', [aCellRef1, aCellRef2]);
        Cell[RPT_START_COL+16, aRow].NumberFormat := '###.0';
        // ACD�B�z�q��(����)
        aCellRef1 := ColRowToRefStr(RPT_START_COL+6, aRow);
        aCellRef2 := ColRowToRefStr(RPT_START_COL+1, aRow);
        AsFormula[RPT_START_COL+17, aRow] := Format('%s/%s', [aCellRef1, aCellRef2]);
        Cell[RPT_START_COL+17, aRow].NumberFormat := '###.0';
        // �^�q�q��(����)
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
  // ���ͦX�p�C
  //============================================================================
  with ASheet do
  begin
    CalcDimensions;
    AsString[RPT_START_COL, aRow] := '�X�p';
    Cell[RPT_START_COL, aRow].FillPatternForeColor := TXc12IndexColor(43);
    // ����������~,���������ͦA�ӧO�վ�
    for i := RPT_START_COL+1 to LastCol do
    begin
      AsFormula[i, aRow] := Format('SUM(%s)', [AreaToRefStr(i, RPT_START_ROW+2, i, LastRow)]);
      Cell[i, aRow].NumberFormat := '#,0';
      Cell[i, aRow].FillPatternForeColor := TXc12IndexColor(43);
    end;
    // �X�p-ACD��ť�v
    aCellRef1 := ColRowToRefStr(RPT_START_COL+6, aRow);
    aCellRef2 := ColRowToRefStr(RPT_START_COL+4, aRow);
    AsFormula[RPT_START_COL+7, aRow] := Format('%s/%s', [aCellRef1, aCellRef2]);
    // �X�p-�O�ɲv
    aCellRef1 := ColRowToRefStr(RPT_START_COL+8, aRow);
    aCellRef2 := ColRowToRefStr(RPT_START_COL+2, aRow);
    AsFormula[RPT_START_COL+9, aRow] := Format('%s/%s', [aCellRef1, aCellRef2]);
    // �X�p-�O�ɲv(�D�X��)
    aCellRef1 := ColRowToRefStr(RPT_START_COL+13, aRow);
    aCellRef2 := ColRowToRefStr(RPT_START_COL+11, aRow);
    AsFormula[RPT_START_COL+14, aRow] := Format('%s/%s', [aCellRef1, aCellRef2]);
    // ACD������ť�q��(����)
    aCellRef1 := ColRowToRefStr(RPT_START_COL+5, aRow);
    aCellRef2 := ColRowToRefStr(RPT_START_COL+1, aRow);
    AsFormula[RPT_START_COL+16, aRow] := Format('%s/%s', [aCellRef1, aCellRef2]);
    Cell[RPT_START_COL+16, aRow].NumberFormat := '###.0';
    // ACD�B�z�q��(����)
    aCellRef1 := ColRowToRefStr(RPT_START_COL+6, aRow);
    aCellRef2 := ColRowToRefStr(RPT_START_COL+1, aRow);
    AsFormula[RPT_START_COL+17, aRow] := Format('%s/%s', [aCellRef1, aCellRef2]);
    Cell[RPT_START_COL+17, aRow].NumberFormat := '###.0';
    // �^�q�q��(����)
    aCellRef1 := ColRowToRefStr(RPT_START_COL+3, aRow);
    aCellRef2 := ColRowToRefStr(RPT_START_COL+12, aRow);
    aCellRef3 := ColRowToRefStr(RPT_START_COL+1, aRow);
    AsFormula[RPT_START_COL+18, aRow] := Format('(%s+%s)/%s', [aCellRef1, aCellRef2, aCellRef3]);
    Cell[RPT_START_COL+18, aRow].NumberFormat := '###.0';
    //--------------------------------------------------------------------------
    // �վ���ܮ榡-�Ⱦ��H��
    Cell[RPT_START_COL+1, aRow].NumberFormat := '#,#.0';
    // �վ���ܮ榡-ACD��ť�v
    Cell[RPT_START_COL+7, aRow].NumberFormat := '##0.0 %';
    // �վ���ܮ榡-�O�ɲv
    Cell[RPT_START_COL+9, aRow].NumberFormat := '##0.0 %';
    // �վ���ܮ榡-�O�ɲv(�D�X��)
    Cell[RPT_START_COL+14, aRow].NumberFormat := '##0.0 %';
  end;

  with AXLS do
  begin
    CmdFormat.BeginEdit(ASheet);
    CmdFormat.Font.Bold := True;
    CmdFormat.Apply(0, aRow, ASheet.LastCol, aRow);
  end;
  //============================================================================
  // �e�X�~��
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
    // �إ߹w�]���x�s��榡
    CmdFormat.BeginEdit(Nil);
    CmdFormat.Font.Name := '�L�n������';
    CmdFormat.Font.Size := 10;
    CmdFormat.Alignment.Vertical := cvaCenter;
    CmdFormat.Alignment.WrapText := True; //�۰ʧ��
    aDefFmt := CmdFormat.AddAsDefault('DefFormat');
    DefaultFormat := aDefFmt;
  end;
end;

procedure TdmSitePhoneSummary.GetData_Summary(ASite: string; ABegTime, AEndTime: TDateTime);
begin
  Log('���o[�夤]���Ӧ^�q���`���');
  
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
