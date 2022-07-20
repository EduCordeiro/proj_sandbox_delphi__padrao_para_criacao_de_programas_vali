program proj_sandbox_delphi__padrao_para_criacao_de_programas_vali;

uses
  Forms,
  uMain in 'uMain.pas' {frmMain},
  ucore in 'ucore.pas',
  udatatypes_apps in '..\classes-delphi__fingerprint_cpd_desenvolvimento\udatatypes_apps.pas',
  ClassLog in '..\classes-delphi__fingerprint_cpd_desenvolvimento\ClassLog.pas',
  ClassDirectory in '..\classes-delphi__fingerprint_cpd_desenvolvimento\ClassDirectory.pas',
  ClassStrings in '..\classes-delphi__fingerprint_cpd_desenvolvimento\ClassStrings.pas',
  ClassTextFile in '..\classes-delphi__fingerprint_cpd_desenvolvimento\ClassTextFile.pas',
  ClassIni in '..\classes-delphi__fingerprint_cpd_desenvolvimento\ClassIni.pas',
  ClassMySqlBases in '..\classes-delphi__fingerprint_cpd_desenvolvimento\ClassMySqlBases.pas',
  ClassConexoes in '..\classes-delphi__fingerprint_cpd_desenvolvimento\ClassConexoes.pas',
  ClassConf in '..\classes-delphi__fingerprint_cpd_desenvolvimento\ClassConf.pas',
  ClassArquivoIni in '..\classes-delphi__fingerprint_cpd_desenvolvimento\ClassArquivoIni.pas',
  ClassFuncoesWin in '..\classes-delphi__fingerprint_cpd_desenvolvimento\ClassFuncoesWin.pas',
  ClassLayoutArquivo in 'ClassLayoutArquivo.pas',
  ClassParametrosDeEntrada in 'ClassParametrosDeEntrada.pas',
  ClassFuncoesBancarias in '..\classes-delphi__fingerprint_cpd_desenvolvimento\ClassFuncoesBancarias.pas',
  ClassPlanoDeTriagem in '..\classes-delphi__fingerprint_cpd_desenvolvimento\ClassPlanoDeTriagem.pas',
  ClassExpressaoRegular in '..\classes-delphi__fingerprint_cpd_desenvolvimento\ClassExpressaoRegular.pas',
  ClassStatusProcessamento in '..\classes-delphi__fingerprint_cpd_desenvolvimento\ClassStatusProcessamento.pas',
  ClassDateTime in '..\classes-delphi__fingerprint_cpd_desenvolvimento\ClassDateTime.pas',
  ClassSMTPDelphi in '..\classes-delphi__fingerprint_cpd_desenvolvimento\ClassSMTPDelphi.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'proj-sandbox-delphi__esqueleto_para_criacao_de_programas';
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
