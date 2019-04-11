program AcdDaily;

uses
  madExcept,
  madLinkDisAsm,
  madListHardware,
  madListProcesses,
  madListModules,
  Forms,
  Main in 'Main.pas' {fmMain},
  ReportData in '..\Public\ReportData.pas' {dmReport: TDataModule},
  TcrmConstants in '..\..\TCRM\Source\TcrmConstants.pas',
  AcdSummary in 'AcdSummary.pas' {dmAcdSummary},
  TePhoneSummary in 'TePhoneSummary.pas' {dmTePhoneSummary},
  SitePhoneSummary in 'SitePhoneSummary.pas' {dmSitePhoneSummary: TDataModule};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TdmReport, dmReport);
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.
