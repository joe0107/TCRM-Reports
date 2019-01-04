program AcdMonthly;

uses
  Forms,
  Main in 'Main.pas' {fmMain},
  ReportData in '..\Public\ReportData.pas' {dmReport: TDataModule},
  TcrmConstants in '..\..\TCRM\Source\TcrmConstants.pas',
  AcdSvcFailedAnalysis in 'AcdSvcFailedAnalysis.pas' {dmAcdSvcFailedAnalysis},
  PhoneAnalysis in 'PhoneAnalysis.pas' {dmPhoneAnalysis: TDataModule};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TdmReport, dmReport);
  Application.CreateForm(TdmPhoneAnalysis, dmPhoneAnalysis);
  Application.CreateForm(TdmAcdSvcFailedAnalysis, dmAcdSvcFailedAnalysis);
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.
