unit ucore;

interface

uses
  Windows, Messages, Variants, Graphics, Controls, FileCtrl,
  Dialogs, StdCtrls,  Classes, SysUtils, Forms,
  DB, ZConnection, ZAbstractRODataset, ZAbstractDataset, ZDataset, ZSqlProcessor,
  ADODb, DBTables,
  udatatypes_apps,
  // Classes
  ClassParametrosDeEntrada,
  ClassArquivoIni, ClassStrings, ClassConexoes, ClassConf, ClassMySqlBases,
  ClassTextFile, ClassDirectory, ClassLog, ClassFuncoesWin, ClassLayoutArquivo,
  ClassFuncoesBancarias, ClassPlanoDeTriagem, ClassExpressaoRegular,
  ClassStatusProcessamento, ClassDateTime, ClassSMTPDelphi;

type

  TCore = class(TObject)
  private

    __queryMySQL_processamento__    : TZQuery;
    __queryMySQL_Insert_            : TZQuery;
    __queryMySQL_plano_de_triagem__ : TZQuery;

    // FUNÇÃO DE PROCESSAMENTO
      Procedure PROCESSAMENTO();

      procedure StoredProcedure_Dropar(Nome: string; logBD:boolean=false; idprograma:integer=0);

      function StoredProcedure_Criar(Nome : string; scriptSQL: TStringList): boolean;

      procedure StoredProcedure_Executar(Nome: string; ComParametro:boolean=false; logBD:boolean=false; idprograma:integer=0);

      function Compactar_Arquivo_7z(Arquivo, destino : String; mover_arquivo: Boolean=false): integer;
      function Extrair_Arquivo_7z(Arquivo, destino : String): integer;

      PROCEDURE COMPACTAR_ARQUIVO(ARQUIVO_ORIGEM, PATH_DESTINO: String; MOVER_ARQUIVO: Boolean=FALSE);
      PROCEDURE EXTRAIR_ARQUIVO(ARQUIVO_ORIGEM, PATH_DESTINO: String);

      procedure Atualiza_arquivo_conf_C(ArquivoConf, sINP, sOUT, sTMP, sLOG, sRGP: String);
      procedure execulta_app_c(app, arquivo_conf: string);

  public

    __ListaPlanoDeTriagem__       : TRecordPlanoTriagemCorreios;

    objParametrosDeEntrada   : TParametrosDeEntrada;
    objConexao               : TMysqlDatabase;
    objPlanoDeTriagem        : TPlanoDeTriagem;
    objString                : TFormataString;
    objLogar                 : TArquivoDelog;
    objDateTime              : TFormataDateTime;
    objArquivoIni            : TArquivoIni;
    objArquivoDeConexoes     : TArquivoDeConexoes;
    objArquivoDeConfiguracao : TArquivoConf;
    objDiretorio             : TDiretorio;
    objFuncoesWin            : TFuncoesWin;
    objLayoutArquivoCliente  : TLayoutCliente;
    objFuncoesBancarias      : TFuncoesBancarias;
    objExpressaoRegular      : TExpressaoRegular;
    objStatusProcessamento   : TStausProcessamento;
    objEmail                 : TSMTPDelphi;

    PROCEDURE COMPACTAR();
    PROCEDURE EXTRAIR();

    function GERA_LOTE_PEDIDO(): String;
    Procedure VALIDA_LOTE_PEDIDO();
    Procedure AtualizaDadosTabelaLOG();

    function PesquisarLote(LOTE_PEDIDO : STRING; status : Integer): Boolean;

    procedure ExcluirBase(NomeTabela: String);
    procedure ExcluirTabela(NomeTabela: String);
    function EnviarEmail(Assunto: string=''; Corpo: string=''): Boolean;
    constructor create();

    procedure ReverterArquivos();

    procedure getListaDeArquivosJaProcessados();

    function ArquivoExieteTabelaTrack(Arquivo: string): Boolean;
    procedure CriaMovimento();

  end;

implementation

uses uMain, Math;

constructor TCore.create();
var
  sMSG                       : string;
  sArquivosScriptSQL         : string;
  stlScripSQL                : TStringList;
begin

  try

    stlScripSQL                                              := TStringList.Create();

    objStatusProcessamento                                   := TStausProcessamento.create();
    objParametrosDeEntrada                                   := TParametrosDeEntrada.Create();

    objParametrosDeEntrada.STL_LISTA_ARQUIVOS_JA_PROCESSADOS := TStringList.Create();
    objParametrosDeEntrada.STL_LISTA_ARQUIVOS_REVERTER       := TStringList.Create();

    objLogar                                                 := TArquivoDelog.Create();
    if FileExists(objLogar.getArquivoDeLog()) then
      objFuncoesWin.DelFile(objLogar.getArquivoDeLog());

    objFuncoesWin                        := TFuncoesWin.create(objLogar);
    objString                            := TFormataString.Create(objLogar);
    objDateTime                          := TFormataDateTime.Create(objLogar);
    objLayoutArquivoCliente              := TLayoutCliente.Create();
    objFuncoesBancarias                  := TFuncoesBancarias.Create();
    objExpressaoRegular                  := TExpressaoRegular.Create();

    objArquivoIni                        := TArquivoIni.create(objLogar,
                                                               objString,
                                                               ExtractFilePath(Application.ExeName),
                                                               ExtractFileName(Application.ExeName));

    objArquivoDeConexoes                 := TArquivoDeConexoes.create(objLogar,
                                                                      objString,
                                                                      objArquivoIni.getPathConexoes());

    objArquivoDeConfiguracao             := TArquivoConf.create(objArquivoIni.getPathConfiguracoes(),
                                                                ExtractFileName(Application.ExeName));

    objParametrosDeEntrada.ID_PROCESSAMENTO := objArquivoDeConfiguracao.getIDProcessamento;

    objConexao                           := TMysqlDatabase.Create();

    if objArquivoIni.getPathConfiguracoes() <> '' then
    begin

      objParametrosDeEntrada.PATHENTRADA                                := objArquivoDeConfiguracao.getConfiguracao('path_default_arquivos_entrada');
      objParametrosDeEntrada.PATHSAIDA                                  := objArquivoDeConfiguracao.getConfiguracao('path_default_arquivos_saida');
      objParametrosDeEntrada.TABELA_PROCESSAMENTO                       := objArquivoDeConfiguracao.getConfiguracao('tabela_processamento');
      objParametrosDeEntrada.TABELA_LOTES_PEDIDOS                       := objArquivoDeConfiguracao.getConfiguracao('TABELA_LOTES_PEDIDOS');
      objParametrosDeEntrada.TABELA_PLANO_DE_TRIAGEM                    := objArquivoDeConfiguracao.getConfiguracao('tabela_plano_de_triagem');
      objParametrosDeEntrada.CARREGAR_PLANO_DE_TRIAGEM_MEMORIA          := objArquivoDeConfiguracao.getConfiguracao('CARREGAR_PLANO_DE_TRIAGEM_MEMORIA');
      objParametrosDeEntrada.LIMITE_DE_SELECT_POR_INTERACOES_NA_MEMORIA := objArquivoDeConfiguracao.getConfiguracao('numero_de_select_por_interacoes_na_memoria');
      objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO                     := objArquivoDeConfiguracao.getConfiguracao('FORMATACAO_LOTE_PEDIDO');
      objParametrosDeEntrada.lista_de_caracteres_invalidos              := objArquivoDeConfiguracao.getConfiguracao('lista_de_caracteres_invalidos');
      objParametrosDeEntrada.eHost                                      := objArquivoDeConfiguracao.getConfiguracao('eHost');
      objParametrosDeEntrada.eUser                                      := objArquivoDeConfiguracao.getConfiguracao('eUser');
      objParametrosDeEntrada.eFrom                                      := objArquivoDeConfiguracao.getConfiguracao('eFrom');
      objParametrosDeEntrada.eTo                                        := objArquivoDeConfiguracao.getConfiguracao('eTo');

      objParametrosDeEntrada.EXTENCAO_ARQUIVOS                          := objArquivoDeConfiguracao.getConfiguracao('EXTENCAO_ARQUIVOS');
      
      objParametrosDeEntrada.COPIAR_LOG_PARA_SAIDA                      := StrTobool(objArquivoDeConfiguracao.getConfiguracao('COPIAR_LOG_PARA_SAIDA'));

      objParametrosDeEntrada.CRIAR_CSV_TRACK                            := StrTobool(objArquivoDeConfiguracao.getConfiguracao('CRIAR_CSV_TRACK'));
      objParametrosDeEntrada.PATH_TRACK                                 := objArquivoDeConfiguracao.getConfiguracao('PATH_TRACK');

      objParametrosDeEntrada.TABELA_TRACK                               := objArquivoDeConfiguracao.getConfiguracao('TABELA_TRACK');
      objParametrosDeEntrada.TABELA_TRACK_LINE                          := objArquivoDeConfiguracao.getConfiguracao('TABELA_TRACK_LINE');
      objParametrosDeEntrada.TABELA_TRACK_LINE_HISTORY                  := objArquivoDeConfiguracao.getConfiguracao('TABELA_TRACK_LINE_HISTORY');

      objParametrosDeEntrada.APP_C_GERA_IDX_EXE                         := objArquivoDeConfiguracao.getConfiguracao('APP_C_GERA_IDX_EXE');
      objParametrosDeEntrada.APP_C_GERA_IDX_CFG                         := objArquivoDeConfiguracao.getConfiguracao('APP_C_GERA_IDX_CFG');

      objParametrosDeEntrada.app_7z_32bits                              := objArquivoDeConfiguracao.getConfiguracao('app_7z_32bits');
      objParametrosDeEntrada.app_7z_64bits                              := objArquivoDeConfiguracao.getConfiguracao('app_7z_64bits');
      objParametrosDeEntrada.ARQUITETURA_WINDOWS                        := objArquivoDeConfiguracao.getConfiguracao('ARQUITETURA_WINDOWS');

      objParametrosDeEntrada.LOGAR                                      := objArquivoDeConfiguracao.getConfiguracao('LOGAR');

      //================
      //  LOGA USUÁRIO
      //========================================================================================================================================================
      objParametrosDeEntrada.APP_LOGAR                                  := objArquivoDeConfiguracao.getConfiguracao('APP_LOGAR');
      objParametrosDeEntrada.TABELA_LOTES_PEDIDOS_LOGIN                 := objArquivoDeConfiguracao.getConfiguracao('TABELA_LOTES_PEDIDOS_LOGIN');
      //========================================================================================================================================================

      objParametrosDeEntrada.ENVIAR_EMAIL                               := objArquivoDeConfiguracao.getConfiguracao('ENVIAR_EMAIL');



      objLogar.Logar('[DEBUG] TfrmMain.FormCreate() - Versão do programa: ' + objFuncoesWin.GetVersaoDaAplicacao());

      objParametrosDeEntrada.PathArquivo_TMP := objArquivoIni.getPathArquivosTemporarios();

      // Criando a Conexao
      objConexao.ConectarAoBanco(objArquivoDeConexoes.getHostName,
                                 'mysql',
                                 objArquivoDeConexoes.getUser,
                                 objArquivoDeConexoes.getPassword,
                                 objArquivoDeConexoes.getProtocolo
                                 );

      sArquivosScriptSQL := ExtractFileName(Application.ExeName);
      sArquivosScriptSQL := StringReplace(sArquivosScriptSQL, '.exe', '.sql', [rfReplaceAll, rfIgnoreCase]);

      stlScripSQL.LoadFromFile(objArquivoIni.getPathScripSQL() + sArquivosScriptSQL);
      objConexao.ExecutaScript(stlScripSQL);

      // Criando Objeto de Plano de Triagem
      if StrToBool(objParametrosDeEntrada.CARREGAR_PLANO_DE_TRIAGEM_MEMORIA) then
        objPlanoDeTriagem := TPlanoDeTriagem.create(objConexao,
                                                    objLogar,
                                                    objString,
                                                    objParametrosDeEntrada.TABELA_PLANO_DE_TRIAGEM, fac);



      objParametrosDeEntrada.stlRelatorioQTDE           := TStringList.Create();

      // LISTA DE ARUQIVOS JA PROCESSADOS
      getListaDeArquivosJaProcessados();


      objParametrosDeEntrada.STL_LOG_TXT                := TStringList.Create(); 

      IF StrToBool(objParametrosDeEntrada.LOGAR) THEN
      BEGIN

          //================
          //  LOGA USUÁRIO
          //==========================================================================================================================================================
          objParametrosDeEntrada.APP_LOGAR_PARAMETRO_TAB_INDEX      := '2';
          objParametrosDeEntrada.APP_LOGAR_PARAMETRO_NOME_APLICACAO := StringReplace(ExtractFileName(Application.ExeName), '.EXE', '', [rfReplaceAll, rfIgnoreCase]);
          objParametrosDeEntrada.APP_LOGAR_PARAMETRO_ARQUIVO_LOGAR  := ExtractFilePath(Application.ExeName) +
                                                                       StringReplace(ExtractFileName(objParametrosDeEntrada.APP_LOGAR), '.EXE', '.TXT', [rfReplaceAll, rfIgnoreCase]);

          objParametrosDeEntrada.APP_LOGAR_PARAMETRO_ARQUIVO_LOGAR  := StringReplace(objParametrosDeEntrada.APP_LOGAR_PARAMETRO_ARQUIVO_LOGAR, '\', '/', [rfReplaceAll, rfIgnoreCase]);



          objParametrosDeEntrada.STL_ARQUIVO_USUARIO_LOGADO := TStringList.Create();
          objFuncoesWin.ExecutarPrograma(objParametrosDeEntrada.APP_LOGAR
                                 + ' ' + objParametrosDeEntrada.APP_LOGAR_PARAMETRO_TAB_INDEX
                                 + ' ' + objParametrosDeEntrada.APP_LOGAR_PARAMETRO_NOME_APLICACAO
                                 + ' ' + objParametrosDeEntrada.APP_LOGAR_PARAMETRO_ARQUIVO_LOGAR);

          objParametrosDeEntrada.STL_ARQUIVO_USUARIO_LOGADO.LoadFromFile(objParametrosDeEntrada.APP_LOGAR_PARAMETRO_ARQUIVO_LOGAR);

          //=====================
          //   CAMPOS DE LOGIN
          //=====================
          objParametrosDeEntrada.USUARIO_LOGADO_APP           := objString.getTermo(1, ';', objParametrosDeEntrada.STL_ARQUIVO_USUARIO_LOGADO.Strings[0]);
          objParametrosDeEntrada.APP_LOGAR_CHAVE_APP          := objString.getTermo(2, ';', objParametrosDeEntrada.STL_ARQUIVO_USUARIO_LOGADO.Strings[0]);
          objParametrosDeEntrada.APP_LOGAR_LOTE               := objString.getTermo(3, ';', objParametrosDeEntrada.STL_ARQUIVO_USUARIO_LOGADO.Strings[0]);
          objParametrosDeEntrada.APP_LOGAR_USUARIO_LOGADO_WIN := objString.getTermo(4, ';', objParametrosDeEntrada.STL_ARQUIVO_USUARIO_LOGADO.Strings[0]);
          objParametrosDeEntrada.APP_LOGAR_IP                 := objString.getTermo(5, ';', objParametrosDeEntrada.STL_ARQUIVO_USUARIO_LOGADO.Strings[0]);
          objParametrosDeEntrada.APP_LOGAR_ID                 := objString.getTermo(6, ';', objParametrosDeEntrada.STL_ARQUIVO_USUARIO_LOGADO.Strings[0]);

          IF (Trim(objParametrosDeEntrada.USUARIO_LOGADO_APP) ='')
          or (Trim(objParametrosDeEntrada.APP_LOGAR_CHAVE_APP) ='')
          or (Trim(objParametrosDeEntrada.APP_LOGAR_LOTE) ='')
          or (Trim(objParametrosDeEntrada.APP_LOGAR_USUARIO_LOGADO_WIN) ='')
          or (Trim(objParametrosDeEntrada.APP_LOGAR_IP) ='')
          or (Trim(objParametrosDeEntrada.APP_LOGAR_ID) ='')
          THEN
            objParametrosDeEntrada.USUARIO_LOGADO_APP := '-1';
      END;

      //=========================
      //    DADOS DE REDE APP
      //=========================
      objParametrosDeEntrada.HOSTNAME                     := objFuncoesWin.getNetHostName;
      objParametrosDeEntrada.IP                           := objFuncoesWin.GetIP;
      objParametrosDeEntrada.USUARIO_SO                   := objFuncoesWin.GetUsuarioLogado;

      //========================
      //  GERA LOTE PEDIDO
      //========================
      if NOT StrToBool(objParametrosDeEntrada.LOGAR) then
      BEGIN

        objParametrosDeEntrada.PEDIDO_LOTE                  := GERA_LOTE_PEDIDO();

        objParametrosDeEntrada.USUARIO_LOGADO_APP           := objParametrosDeEntrada.USUARIO_SO;
        objParametrosDeEntrada.APP_LOGAR_CHAVE_APP          := objParametrosDeEntrada.ID_PROCESSAMENTO;
        objParametrosDeEntrada.APP_LOGAR_LOTE               := objParametrosDeEntrada.PEDIDO_LOTE;
        objParametrosDeEntrada.APP_LOGAR_USUARIO_LOGADO_WIN := objParametrosDeEntrada.USUARIO_SO;
        objParametrosDeEntrada.APP_LOGAR_IP                 := objParametrosDeEntrada.IP;
        objParametrosDeEntrada.APP_LOGAR_ID                 := objParametrosDeEntrada.ID_PROCESSAMENTO;

      END
      ELSE
      IF objParametrosDeEntrada.USUARIO_LOGADO_APP <> '-1' THEN
        objParametrosDeEntrada.PEDIDO_LOTE                := GERA_LOTE_PEDIDO();
      //==========================================================================================================================================================

    end;

  except
    on E:Exception do
    begin

      sMSG := '[ERRO] Não foi possível inicializar as configurações aq do programa. '+#13#10#13#10
            + ' EXCEÇÃO: '+E.Message+#13#10#13#10
            + ' O programa será encerrado agora.';

      showmessage(sMSG);

      objLogar.Logar(sMSG);

      Application.Terminate;
    end;
  end;

end;

function TCore.GERA_LOTE_PEDIDO(): String;
var
  sComando : string;
  sData    : string;
begin

  //==================
  //  CRIA NOVO LOTE
  //==================
  sData := FormatDateTime('YYYY-MM-DD hh:mm:ss', Now());

  sComando := ' insert into ' + objParametrosDeEntrada.TABELA_LOTES_PEDIDOS + '(VALIDO, DATA_CRIACAO, CHAVE, USUARIO_WIN, USUARIO_APP, IP, ID, LOTE_LOGIN, HOSTNAME)'
            + ' Value('
                      + '"'   + 'N'
                      + '","' + sData
                      + '","' + objParametrosDeEntrada.APP_LOGAR_CHAVE_APP
                      + '","' + objParametrosDeEntrada.APP_LOGAR_USUARIO_LOGADO_WIN
                      + '","' + objParametrosDeEntrada.USUARIO_LOGADO_APP
                      + '","' + objParametrosDeEntrada.APP_LOGAR_IP
                      + '","' + objParametrosDeEntrada.ID_PROCESSAMENTO
                      + '","' + objParametrosDeEntrada.APP_LOGAR_LOTE
                      + '","' + objParametrosDeEntrada.HOSTNAME
                      + '")';
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);

  //========================
  //  RETORNA LOTE CRIADO
  //========================
  sComando := ' SELECT LOTE_PEDIDO FROM  ' + objParametrosDeEntrada.TABELA_LOTES_PEDIDOS
            + ' WHERE '
                      + '     VALIDO        = "' + 'N'                                                 + '"'
                      + ' AND DATA_CRIACAO  = "' + sData                                               + '"'
                      + ' AND CHAVE         = "' + objParametrosDeEntrada.APP_LOGAR_CHAVE_APP          + '"'
                      + ' AND USUARIO_WIN   = "' + objParametrosDeEntrada.APP_LOGAR_USUARIO_LOGADO_WIN + '"'
                      + ' AND USUARIO_APP   = "' + objParametrosDeEntrada.USUARIO_LOGADO_APP           + '"'
                      + ' AND HOSTNAME      = "' + objParametrosDeEntrada.HOSTNAME                     + '"'
                      + ' AND LOTE_LOGIN    = "' + objParametrosDeEntrada.APP_LOGAR_LOTE               + '"'
                      + ' AND IP            = "' + objParametrosDeEntrada.APP_LOGAR_IP                 + '"'
                      + ' AND ID            = "' + objParametrosDeEntrada.ID_PROCESSAMENTO             + '"';
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

  Result := FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, __queryMySQL_processamento__.FieldByName('LOTE_PEDIDO').AsInteger);

end;

PROCEDURE TCore.VALIDA_LOTE_PEDIDO();
VAR
  sComando                : string;
BEGIN

  //========================
  //  RETORNA LOTE CRIADO
  //========================
  sComando := ' UPDATE  ' + objParametrosDeEntrada.TABELA_LOTES_PEDIDOS
            + ' set VALIDO         = "' + objParametrosDeEntrada.STATUS_PROCESSAMENTO  + '"'
            + '    ,RELATORIO_QTD  = "' + objParametrosDeEntrada.stlRelatorioQTDE.Text + '"'
            + '    ,LOTE_LOGIN     = "' + objParametrosDeEntrada.APP_LOGAR_LOTE    + '"'
            + ' WHERE '
            + '     LOTE_PEDIDO   = "' + objParametrosDeEntrada.PEDIDO_LOTE                   + '"'
            + ' AND VALIDO        = "' + 'N'                                                  + '"'
            + ' AND CHAVE         = "' + objParametrosDeEntrada.APP_LOGAR_CHAVE_APP           + '"'
            + ' AND USUARIO_WIN   = "' + objParametrosDeEntrada.APP_LOGAR_USUARIO_LOGADO_WIN  + '"'
            + ' AND HOSTNAME      = "' + objParametrosDeEntrada.HOSTNAME                      + '"'
            + ' AND IP            = "' + objParametrosDeEntrada.APP_LOGAR_IP                  + '"'
            + ' AND ID            = "' + objParametrosDeEntrada.ID_PROCESSAMENTO              + '"';
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);

end;

Procedure TCore.AtualizaDadosTabelaLOG();
var
  sComando                  : String;
Begin
  //=========================================================================
  //  GRAVA LOG NA TABELA DE LOGIN - SOMENTE SE O PARÂMETRO LOGAR FOR TRUE
  //=========================================================================
  if StrToBool(objParametrosDeEntrada.LOGAR) then
  begin
    objParametrosDeEntrada.STL_LOG_TXT.Text := StringReplace(objParametrosDeEntrada.STL_LOG_TXT.Text, '\', '\\', [rfReplaceAll, rfIgnoreCase]);

    sComando := ' update ' + objParametrosDeEntrada.TABELA_LOTES_PEDIDOS_LOGIN
              + ' SET '
              + '      LOG_APP          = "' + objParametrosDeEntrada.STL_LOG_TXT.Text                           + '"'
              + '     ,VALIDO           = "' + objParametrosDeEntrada.STATUS_PROCESSAMENTO                       + '"'
              + '     ,QTD_PROCESSADA   = "' + IntToStr(objParametrosDeEntrada.TOTAL_PROCESSADOS_LOG)            + '"'
              + '     ,QTD_INVALIDOS    = "' + IntToStr(objParametrosDeEntrada.TOTAL_PROCESSADOS_INVALIDOS_LOG)  + '"'
              + '     ,LOTE_APP         = "' + objParametrosDeEntrada.PEDIDO_LOTE                                + '"'
              + '     ,RELATORIO_QTD    = "' + objParametrosDeEntrada.stlRelatorioQTDE.Text                      + '"'
              + ' WHERE CHAVE       = "' + objParametrosDeEntrada.APP_LOGAR_CHAVE_APP          + '"'
              + '   AND LOTE_PEDIDO = "' + objParametrosDeEntrada.APP_LOGAR_LOTE               + '"'
              + '   AND USUARIO_WIN = "' + objParametrosDeEntrada.APP_LOGAR_USUARIO_LOGADO_WIN + '"'
              + '   AND USUARIO_APP = "' + objParametrosDeEntrada.USUARIO_LOGADO_APP           + '"'
              + '   AND HOSTNAME    = "' + objParametrosDeEntrada.HOSTNAME                     + '"'
              + '   AND IP          = "' + objParametrosDeEntrada.APP_LOGAR_IP                 + '"'
              + '   AND ID          = "' + objParametrosDeEntrada.APP_LOGAR_ID                 + '"';
    objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);
  end;

end;


Procedure TCore.PROCESSAMENTO();
Var


Arq_Arquivo_Entada   : TextFile;
Arq_Arquivo_Saida    : TextFile;

sArquivoEntrada      : string;
sArquivoSaida        : string;
sLinha               : string;
sValues              : string;
sComando             : string;
sCampos              : string;
sOperadora           : string;
sContrato            : string;
sCep                 : string;

iContArquivos        : Integer;
iTotalDeArquivos     : Integer;

// Variáveis de controle do select
iTotalDeRegistrosDaTabela : Integer;
iLimit : Integer;
iTotalDeInteracoesDeSelects : Integer;
iResto : Integer;
iRegInicial : Integer;
iQtdeRegistros : Integer;
iContInteracoesDeSelects : Integer;


begin

  //*********************************************************************************************
  //                         Alimentando nome dos campos da tabela de Cliente
  //*********************************************************************************************
  sComando := 'describe ' + objParametrosDeEntrada.tabela_processamento;
  objStatusProcessamento := objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

  while not __queryMySQL_processamento__.Eof do
  Begin
    sCampos := sCampos + __queryMySQL_processamento__.FieldByName('Field').AsString;
    __queryMySQL_processamento__.Next;
    if not __queryMySQL_processamento__.Eof then
      sCampos := sCampos + ',';
  end;

  iTotalDeArquivos := objParametrosDeEntrada.ListaDeArquivosDeEntrada.Count;

  for iContArquivos := 0 to iTotalDeArquivos - 1 do
  begin

    sComando := 'delete from ' + objParametrosDeEntrada.tabela_processamento;
    objStatusProcessamento := objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);

    sArquivoEntrada := objParametrosDeEntrada.ListaDeArquivosDeEntrada.Strings[iContArquivos];

    AssignFile(Arq_Arquivo_Entada, objString.AjustaPath(objParametrosDeEntrada.PathEntrada) + sArquivoEntrada);
    reset(Arq_Arquivo_Entada);

    while not eof(Arq_Arquivo_Entada) do
    Begin

      readln(Arq_Arquivo_Entada, sLinha);

      sLinha := objString.StringReplaceList(sLinha, objParametrosDeEntrada.lista_de_caracteres_invalidos);

      sOperadora := Copy(sLinha, 16, 3);
      sContrato  := Copy(sLinha, 23, 9);
      sCep       := Copy(sLinha, 393, 8);

      sValues := '"' + sOperadora + '",'
               + '"' + sContrato + '",'
               + '"' + sCep + '",'
               + '"' + sLinha + '"';

      sComando := 'Insert into ' + objParametrosDeEntrada.tabela_processamento + ' (' + sCampos + ') values(' + sValues + ')';
      objStatusProcessamento := objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);

    end;

    CloseFile(Arq_Arquivo_Entada);

    sComando := 'SELECT count(contrato) as qtde FROM ' + objParametrosDeEntrada.tabela_processamento;
    objStatusProcessamento := objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

    iTotalDeRegistrosDaTabela := __queryMySQL_processamento__.FieldByName('qtde').AsInteger;

    iLimit := StrToInt(objParametrosDeEntrada.LIMITE_DE_SELECT_POR_INTERACOES_NA_MEMORIA);
    iResto := iTotalDeRegistrosDaTabela mod iLimit;

    if iResto <> 0 then
      iTotalDeInteracoesDeSelects := iTotalDeRegistrosDaTabela div iLimit + 1
    else
      iTotalDeInteracoesDeSelects := iTotalDeRegistrosDaTabela div iLimit;

    iQtdeRegistros := 0;

    sArquivoSaida   := StringReplace(sArquivoEntrada, '.txt', '_SAIDA.TXT', [rfReplaceAll, rfIgnoreCase]);

    AssignFile(Arq_Arquivo_Saida, objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA) + sArquivoSaida);
    Rewrite(Arq_Arquivo_Saida);

    for iContInteracoesDeSelects := 0 to iTotalDeInteracoesDeSelects -1 do
    begin
      iRegInicial    := iQtdeRegistros;
      iQtdeRegistros := iQtdeRegistros + iLimit;

      sComando := 'SELECT * FROM ' + objParametrosDeEntrada.tabela_processamento + ' limit ' + IntToStr(iRegInicial) + ',' + IntToStr(iLimit);
      objStatusProcessamento := objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

      while not __queryMySQL_processamento__.Eof do
      begin

        sLinha := __queryMySQL_processamento__.FieldByName('LINHA').AsString;

        sCep   := Copy(sLinha, 393, 8);

        writeln(Arq_Arquivo_Saida, sLinha);

        __queryMySQL_processamento__.Next;

      end;

    end;

    CloseFile(Arq_Arquivo_Saida);

  end;

end;

procedure TCore.ExcluirBase(NomeTabela: String);
var
  sComando : String;
  sBase    : string;
begin

  sBase := objString.getTermo(1, '.', NomeTabela);

  sComando := 'drop database ' + sBase;
  objStatusProcessamento := objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);
end;

procedure TCore.ExcluirTabela(NomeTabela: String);
var
  sComando : String;
  sTabela  : String;
begin

  sTabela := objString.getTermo(2, '.', NomeTabela);

  sComando := 'drop table ' + sTabela;
  objStatusProcessamento := objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);
end;

procedure TCore.StoredProcedure_Dropar(Nome: string; logBD:boolean=false; idprograma:integer=0);
var
  sSQL: string;
  sMensagem: string;
begin
  try
    sSQL := 'DROP PROCEDURE if exists ' + Nome;
    objConexao.Executar_SQL(__queryMySQL_processamento__, sSQL, 1);
  except
    on E:Exception do
    begin
      sMensagem := '  StoredProcedure_Dropar(' + Nome + ') - Excecao:' + E.Message + ' . SQL: ' + sSQL;
      objLogar.Logar(sMensagem);
    end;
  end;

end;

function TCore.StoredProcedure_Criar(Nome : string; scriptSQL: TStringList): boolean;
var
  bExecutou    : boolean;
  sMensagem    : string;
begin


  bExecutou := objConexao.Executar_SQL(__queryMySQL_processamento__, scriptSQL.Text, 1).status;

  if not bExecutou then
  begin
    sMensagem := '  StoredProcedure_Criar(' + Nome + ') - Não foi possível carregar a stored procedure para execução.';
    objLogar.Logar(sMensagem);
  end;

  result := bExecutou;
end;

procedure TCore.StoredProcedure_Executar(Nome: string; ComParametro:boolean=false; logBD:boolean=false; idprograma:integer=0);
var

  sSQL        : string;
  sMensagem   : string;
begin

  try
    (*
    if not Assigned(con) then
    begin
      con := TZConnection.Create(Application);
      con.HostName  := objConexao.getHostName;
      con.Database  := sNomeBase;
      con.User      := objConexao.getUser;
      con.Protocol  := objConexao.getProtocolo;
      con.Password  := objConexao.getPassword;
      con.Properties.Add('CLIENT_MULTI_STATEMENTS=1');
      con.Connected := True;
    end;

    if not Assigned(QP) then
      QP := TZQuery.Create(Application);

    QP.Connection := con;
    QP.SQL.Clear;
    *)

    sSQL := 'CALL '+ Nome;
    if not ComParametro then
      sSQL := sSQL + '()';

    objConexao.Executar_SQL(__queryMySQL_processamento__, sSQL, 1);

  except
    on E:Exception do
    begin
      sMensagem := '[ERRO] StoredProcedure_Executar('+Nome+') - Excecao:'+E.Message+' . SQL: '+sSQL;
      objLogar.Logar(sMensagem);
      ShowMessage(sMensagem);
    end;
  end;

//  objConexao.Executar_SQL(__queryMySQL_processamento__, sSQL, 1)

end;

function TCore.EnviarEmail(Assunto: string=''; Corpo: string=''): Boolean;
var
  sHost    : string;
  suser    : string;
  sFrom    : string;
  sTo      : string;
  sAssunto : string;
  sCorpo   : string;
  sAnexo   : string;
  sAplicacao: string;

begin

  sAplicacao := ExtractFileName(Application.ExeName);
  sAplicacao := StringReplace(sAplicacao, '.exe', '', [rfReplaceAll, rfIgnoreCase]);

  sHost    := objParametrosDeEntrada.eHost;
  suser    := objParametrosDeEntrada.eUser;
  sFrom    := objParametrosDeEntrada.eFrom;
  sTo      := objParametrosDeEntrada.eTo;
  sAssunto := 'Processamento - ' + sAplicacao + ' - ' + objFuncoesWin.GetVersaoDaAplicacao() + ' [PROCESSAMENTO: ' + objParametrosDeEntrada.PEDIDO_LOTE + ']';
  sAssunto := sAssunto + ' ' + Assunto;
  sCorpo   := Corpo;

  sAnexo := objLogar.getArquivoDeLog();

  //sAnexo := StringReplace(anexo, '"', '', [rfReplaceAll, rfIgnoreCase]);
  //sAnexo := StringReplace(anexo, '''', '', [rfReplaceAll, rfIgnoreCase]);

  try

    objEmail := TSMTPDelphi.create(sHost, suser);

    if objEmail.ConectarAoServidorSMTP() then
    begin
      if objEmail.AnexarArquivo(sAnexo) then
      begin

          if not (objEmail.EnviarEmail(sFrom, sTo, sAssunto, sCorpo)) then
            ShowMessage('ERRO AO ENVIAR O E-MAIL')
          else
          if not objEmail.DesconectarDoServidorSMTP() then
            ShowMessage('ERRO AO DESCONECTAR DO SERVIDOR');
      end
      else
        ShowMessage('ERRO AO ANEXAR O ARQUIVO');
    end
    else
      ShowMessage('ERRO AO CONECTAR AO SERVIDOR');

  except
    ShowMessage('NÃO FOI POSSIVEL ENVIAR O E-MAIL.');
  end;
end;



function Tcore.PesquisarLote(LOTE_PEDIDO : STRING; status : Integer): Boolean;
var
  sComando : string;
  iPedido  : Integer;
  sStauts  : string;
begin

  case status of
    0: sStauts := 'S';
    1: sStauts := 'N';
  end;

  objParametrosDeEntrada.PEDIDO_LOTE_TMP := LOTE_PEDIDO;

  sComando := ' SELECT RELATORIO_QTD FROM  ' + objParametrosDeEntrada.TABELA_LOTES_PEDIDOS
            + ' WHERE LOTE_PEDIDO = ' + LOTE_PEDIDO + ' AND VALIDO = "' + sStauts + '"';
  objStatusProcessamento := objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

  objParametrosDeEntrada.stlRelatorioQTDE.Text := __queryMySQL_processamento__.FieldByName('RELATORIO_QTD').AsString;

  if __queryMySQL_processamento__.RecordCount > 0 then
    Result := True
  else
    Result := False;

end;

PROCEDURE TCORE.COMPACTAR();
Var
  sArquivo         : String;
  sPathEntrada     : String;
  sPathSaida       : String;

  iContArquivos    : Integer;
  iTotalDeArquivos : Integer;
BEGIN

  sPathEntrada := objString.AjustaPath(objParametrosDeEntrada.PATHENTRADA);
  sPathSaida   := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA);
  ForceDirectories(sPathSaida);

  iTotalDeArquivos := objParametrosDeEntrada.ListaDeArquivosDeEntrada.Count;

  for iContArquivos := 0 to iTotalDeArquivos - 1 do
  begin

    sArquivo := objParametrosDeEntrada.LISTADEARQUIVOSDEENTRADA.Strings[iContArquivos];
    COMPACTAR_ARQUIVO(sPathEntrada + sArquivo, sPathSaida, True);

  end;

end;

PROCEDURE TCORE.EXTRAIR();
Var
  sArquivo         : String;
  sPathEntrada     : String;
  sPathSaida       : String;

  iContArquivos    : Integer;
  iTotalDeArquivos : Integer;
BEGIN

  sPathEntrada := objString.AjustaPath(objParametrosDeEntrada.PATHENTRADA);
  sPathSaida   := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA);
  ForceDirectories(sPathSaida);

  iTotalDeArquivos := objParametrosDeEntrada.ListaDeArquivosDeEntrada.Count;

  for iContArquivos := 0 to iTotalDeArquivos - 1 do
  begin

    sArquivo := objParametrosDeEntrada.LISTADEARQUIVOSDEENTRADA.Strings[iContArquivos];
    EXTRAIR_ARQUIVO(sPathEntrada + sArquivo, sPathSaida);

  end;

end;


PROCEDURE TCORE.COMPACTAR_ARQUIVO(ARQUIVO_ORIGEM, PATH_DESTINO: String; MOVER_ARQUIVO: Boolean = FALSE);
begin

  Compactar_Arquivo_7z(ARQUIVO_ORIGEM, PATH_DESTINO, MOVER_ARQUIVO);

end;

PROCEDURE TCORE.EXTRAIR_ARQUIVO(ARQUIVO_ORIGEM, PATH_DESTINO: String);
begin

  Extrair_Arquivo_7z(ARQUIVO_ORIGEM, PATH_DESTINO);

end;

function TCORE.Compactar_Arquivo_7z(Arquivo, destino : String; mover_arquivo: Boolean=false): integer;
Var
  sComando                  : String;
  sArquivoDestino           : String;
  sParametros               : String;
  __AplicativoCompactacao__ : String;

  iRetorno                  : Integer;
Begin

    sArquivoDestino := ExtractFileName(Arquivo) + '.7Z';

    destino := objString.AjustaPath(destino);

    sParametros := ' a ';

    IF StrToInt(objParametrosDeEntrada.ARQUITETURA_WINDOWS) = 32 THEN
      __AplicativoCompactacao__ := objParametrosDeEntrada.app_7z_32bits;

    IF StrToInt(objParametrosDeEntrada.ARQUITETURA_WINDOWS) = 64 THEN
      __AplicativoCompactacao__ := objParametrosDeEntrada.app_7z_64bits;

    sComando := __AplicativoCompactacao__ + sParametros + ' "' + destino + sArquivoDestino + '" "' + Arquivo + '"';

    if mover_arquivo then
      sComando := sComando + ' -sdel';

    iRetorno := objFuncoesWin.WinExecAndWait32(sComando);

    Result   := iRetorno;

End;

function TCORE.Extrair_Arquivo_7z(Arquivo, destino : String): integer;
Var
  sComando                  : String;
  sParametros               : String;
  __AplicativoCompactacao__ : String;

  iRetorno                  : Integer;
Begin

    destino := objString.AjustaPath(destino);

    sParametros := ' e ';

    IF StrToInt(objParametrosDeEntrada.ARQUITETURA_WINDOWS) = 32 THEN
      __AplicativoCompactacao__ := objParametrosDeEntrada.app_7z_32bits;

    IF StrToInt(objParametrosDeEntrada.ARQUITETURA_WINDOWS) = 64 THEN
      __AplicativoCompactacao__ := objParametrosDeEntrada.app_7z_64bits;

    sComando := __AplicativoCompactacao__ + sParametros + ' ' + Arquivo +  ' -y -o"' + destino + '"';

    iRetorno := objFuncoesWin.WinExecAndWait32(sComando);

    Result   := iRetorno;

End;

procedure TCore.CriaMovimento();
var
  sPathEntrada                      : string;
  sPathMovimentoArquivos            : string;
  sPathMovimentoBackupZip           : string;
  sPathMovimentoCIF                 : string;
  sPathMovimentoRelatorio           : string;
  sPathComplemento                  : string;
  sPathMovimentoTRACK               : string;
  sPathMovimentoTMP                 : string;
  sArquivoZIP                       : string;
  sArquivoPDF                       : string;
  sArquivoTXT                       : string;
  sArquivoJRN                       : string;
  sArquivoAFP                       : string;
  sArquivoREL                       : string;
  sComando                          : string;
  sLinha                            : string;

  iContArquivos                     : Integer;
  iContArquivoZip                   : Integer;
  iTotalFolhas                      : Integer;
  iTotalPaginas                     : Integer;
  iTotalObjestos                    : Integer;


  stlFiltroArquivo                  : TStringList;
  stlRelatorio                      : TStringList;
  stlTrack                          : TStringList;

begin

  objParametrosDeEntrada.TIMESTAMP := now();

  //=======================================================================================================================================================================================
  //  LIMPANDO A TABELA DE PROCESSAMENTO
  //=======================================================================================================================================================================================
  sComando := 'DELETE FROM ' + objParametrosDeEntrada.TABELA_PROCESSAMENTO;
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);
  //=======================================================================================================================================================================================

  if objParametrosDeEntrada.TESTE then
    sPathComplemento := '_TESTE';

  stlFiltroArquivo                 := TStringList.create();
  stlRelatorio                     := TStringList.create();
  stlTrack                         := TStringList.create();

  //=======================================================================================================================================================================================
  //  DEFINE ESTRUTURA MOVIMENTO
  //=======================================================================================================================================================================================
  sPathEntrada                     := objString.AjustaPath(objParametrosDeEntrada.PATHENTRADA);
  sPathMovimentoArquivos           := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA) + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) + sPathComplemento + PathDelim + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, StrToInt(objParametrosDeEntrada.PEDIDO_LOTE)) + PathDelim + 'ARQUIVOS'   + PathDelim;
  sPathmovimentoBackupZip          := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA) + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) + sPathComplemento + PathDelim + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, StrToInt(objParametrosDeEntrada.PEDIDO_LOTE)) + PathDelim + 'BACKUP_ZIP' + PathDelim;
  sPathmovimentoCIF                := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA) + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) + sPathComplemento + PathDelim + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, StrToInt(objParametrosDeEntrada.PEDIDO_LOTE)) + PathDelim + 'CIF'        + PathDelim;
  sPathMovimentoRelatorio          := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA) + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) + sPathComplemento + PathDelim + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, StrToInt(objParametrosDeEntrada.PEDIDO_LOTE)) + PathDelim + 'RELATORIO'  + PathDelim;
  sPathMovimentoTRACK              := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA) + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) + sPathComplemento + PathDelim + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, StrToInt(objParametrosDeEntrada.PEDIDO_LOTE)) + PathDelim + 'TRACK'      + PathDelim;
  sPathMovimentoTMP                := objString.AjustaPath(objParametrosDeEntrada.PATHSAIDA) + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) + sPathComplemento + PathDelim + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, StrToInt(objParametrosDeEntrada.PEDIDO_LOTE)) + PathDelim + 'TMP'      + PathDelim;
  //=======================================================================================================================================================================================

  //===================================================================================================================================================================
  // CRIA PASTAS
  //===================================================================================================================================================================
  ForceDirectories(sPathMovimentoArquivos);
  ForceDirectories(sPathmovimentoBackupZip);
  ForceDirectories(sPathmovimentoCIF);
  ForceDirectories(sPathMovimentoRelatorio);
  ForceDirectories(sPathMovimentoTRACK);
  ForceDirectories(sPathMovimentoTMP);
  //===================================================================================================================================================================

  //===================================================================================================================================================================
  // EXTRAI E MOVE OS ARQUIVOS
  //===================================================================================================================================================================
  for iContArquivoZip := 0 to objParametrosDeEntrada.LISTADEARQUIVOSDEENTRADA.Count - 1 do
  begin

    objFuncoesWin.DeletarArquivosPorFiltro(sPathMovimentoTMP , '*.*');

    sArquivoZIP  := objParametrosDeEntrada.LISTADEARQUIVOSDEENTRADA.Strings[iContArquivoZip];

    EXTRAIR_ARQUIVO(sPathEntrada + sArquivoZIP, sPathMovimentoTMP);

    objFuncoesWin.CopiarArquivo(sPathEntrada + sArquivoZIP, sPathmovimentoBackupZip + sArquivoZIP);
    objFuncoesWin.CopiarArquivo(sPathEntrada + sArquivoZIP, sPathMovimentoTMP       + sArquivoZIP);
    DeleteFile(sPathMovimentoTMP + sArquivoZIP);

    //===================================================================================================================================================================
    // MOVENDO OS ARQUIVOS
    //===================================================================================================================================================================

      //===================================================================================================================================================================
      // PEGA LISTA DE ARQUIVOS PDF (LISTA DE POSTAGEM) E MOVE PARA A PASTA DE POSTAGEM
      //===================================================================================================================================================================
      stlFiltroArquivo.Clear;
      objFuncoesWin.ObterListaDeArquivosDeUmDiretorio(sPathMovimentoTMP, stlFiltroArquivo, '*.PDF');
      for iContArquivos := 0 to stlFiltroArquivo.Count - 1 do
      begin
        sArquivoPDF := stlFiltroArquivo.Strings[iContArquivos];
        if objFuncoesWin.CopiarArquivo(sPathMovimentoTMP + sArquivoPDF, sPathmovimentoCIF + sArquivoPDF) then
         DeleteFile(sPathMovimentoTMP + sArquivoPDF);
      end;
      //===================================================================================================================================================================

      //===================================================================================================================================================================
      // PEGA LISTA DE ARQUIVOS TXT (CIF) E MOVE PAA A PASTA DE POSTAGEM
      //===================================================================================================================================================================
      stlFiltroArquivo.Clear;
      objFuncoesWin.ObterListaDeArquivosDeUmDiretorio(sPathMovimentoTMP, stlFiltroArquivo, '*.TXT');
      for iContArquivos := 0 to stlFiltroArquivo.Count - 1 do
      begin
        sArquivoTXT := stlFiltroArquivo.Strings[iContArquivos];
        if objFuncoesWin.CopiarArquivo(sPathMovimentoTMP + sArquivoTXT, sPathmovimentoCIF + sArquivoTXT) then
         DeleteFile(sPathMovimentoTMP + sArquivoTXT);
      end;
      //===================================================================================================================================================================

      //===================================================================================================================================================================
      // PEGA LISTA DE ARQUIVOS TXT (AFP) E MOVE PAA A PASTA DE ARQUIVOS
      //===================================================================================================================================================================
      stlFiltroArquivo.Clear;
      objFuncoesWin.ObterListaDeArquivosDeUmDiretorio(sPathMovimentoTMP, stlFiltroArquivo, '*.AFP');
      for iContArquivos := 0 to stlFiltroArquivo.Count - 1 do
      begin
        sArquivoAFP := stlFiltroArquivo.Strings[iContArquivos];
        if objFuncoesWin.CopiarArquivo(sPathMovimentoTMP + sArquivoAFP, sPathMovimentoArquivos + sArquivoAFP) then
         DeleteFile(sPathMovimentoTMP + sArquivoAFP);
      end;
      //===================================================================================================================================================================

      //===================================================================================================================================================================
      // PEGA LISTA DE ARQUIVOS TXT (JRN) E MOVE PAA A PASTA DE ARQUIVOS
      //===================================================================================================================================================================
      stlFiltroArquivo.Clear;
      objFuncoesWin.ObterListaDeArquivosDeUmDiretorio(sPathMovimentoTMP, stlFiltroArquivo, '*.JRN');
      for iContArquivos := 0 to stlFiltroArquivo.Count - 1 do
      begin
        sArquivoJRN := stlFiltroArquivo.Strings[iContArquivos];
        if objFuncoesWin.CopiarArquivo(sPathMovimentoTMP + sArquivoJRN, sPathMovimentoArquivos + sArquivoJRN) then
         DeleteFile(sPathMovimentoTMP + sArquivoJRN);
      end;
      //===================================================================================================================================================================

    //===================================================================================================================================================================

    //===================================================================================================================================================================
    // CARREGA ARQUIVO JRN PARA BANCO PARA GERAR RELATÓRIOS
    //===================================================================================================================================================================
    stlFiltroArquivo.Clear;
    objFuncoesWin.ObterListaDeArquivosDeUmDiretorio(sPathMovimentoArquivos, stlFiltroArquivo, '*.JRN');
    for iContArquivos := 0 to stlFiltroArquivo.Count - 1 do
    begin

      sArquivoJRN := stlFiltroArquivo.Strings[iContArquivos];
      sArquivoAFP := StringReplace(sArquivoJRN, '.JRN', '.AFP', [rfReplaceAll, rfIgnoreCase]);

      sComando := ' LOAD DATA LOCAL INFILE "' + StringReplace(sPathMovimentoArquivos, '\', '\\', [rfReplaceAll, rfIgnoreCase]) + sArquivoJRN + '" '
               + '  INTO TABLE ' + objParametrosDeEntrada.TABELA_PROCESSAMENTO
               + '    CHARACTER SET latin1 '
               + '  FIELDS '
               + '    TERMINATED BY "|" '
               + '  LINES '
               + '    TERMINATED BY "\r\n" '
               + '   SET LOTE          = MID(CIF, 11, 5) '
               + '      ,DATA_POSTAGEM = MID(CIF, 29, 6) '
               + '      ,ARQUIVO_AFP   = "' + sArquivoAFP + '"'
               + '      ,ARQUIVO_ZIP   = "' + sArquivoZIP + '"'
               + '      ,MOVIMENTO     = "' + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) + '"'
               ;
      objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);
    end;
    //===================================================================================================================================================================

    //===================================================================================================
    // CRIANDO ARQUIVO IDX PARA OS ARQUIVOS AFP
    //===============================================================================================================================================
    Atualiza_arquivo_conf_C(objParametrosDeEntrada.APP_C_GERA_IDX_CFG, sPathMovimentoArquivos, sPathMovimentoArquivos, '', '', '');
    execulta_app_c(objParametrosDeEntrada.APP_C_GERA_IDX_EXE, objParametrosDeEntrada.APP_C_GERA_IDX_CFG);
    //===============================================================================================================================================

  END;
  //===================================================================================================================================================================


  //===================================================================================================
  // CABEÇALHO DO CSV TRACK
  //==================================================================================================================================================================
  stlTrack.Clear;
  sLinha      :=  'OF_FORMULARIO'
               + ';OF_ENVELOPE'
               + ';OF_ENCATE'
               + ';MOVIMENTO'
               + ';FILLER'
               + ';ARQUIVO'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';TIMESTAMP'
               + ';ACABAMENTO'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';LOTE_PROCESAMENTO'
               + ';QUANTIDADE_DE_OBJETOS_POR_OF'
               + ';QUANTIDADE_DE_FOLHAS_POR_OF'
               + ';QUANTIDADE_DE_PAGINAS_POR_OF'
               + ';FILLER'
               + ';FILLER'
               + ';CARTAO_POSTAGEM'
               + ';DATA_LOTE_QTD_POSTAGEM'
               + ';TOTAL_LOCAL'
               + ';TOTAL_ESTADUAL'
               + ';TOTAL_NACIONAL'
               + ';TOTAL'
               + ';PORTE[GR]'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';FILLER'
               + ';PAPEL'
               ;
    stlTrack.Add(sLinha);
  //==================================================================================================================================================================

  //===================================================================================================
  // CRIANDO RELATÓRIO DE QUANTIDADES
  //==================================================================================================================================================================
  stlRelatorio.Clear;
  sComando := 'SELECT ARQUIVO_ZIP, ARQUIVO_AFP, MOVIMENTO, DATA_POSTAGEM, LOTE, COUNT(CIF) AS QTD, SUM(PAGINAS) AS PAGINAS, SUM(FOLHAS) AS FOLHAS FROM ' + objParametrosDeEntrada.TABELA_PROCESSAMENTO
            + ' WHERE CIF <> "" '
            + ' GROUP BY ARQUIVO_AFP, DATA_POSTAGEM';
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

  sLinha := stringOfChar('-', 122)
  + #13 + #10 + 'RELATÓRIO DE QUANTIDADES - RECEBIMENTO' + sPathComplemento
  + #13 + #10 + stringOfChar('-', 122)
  + #13 + #10 + 'MOVIMENTO  ARQUIVO                                      DATA DE POSTAGEM LOTE DE POSTAGEM QUANTIDADE PAGINAS    FOLHAS'
  + #13 + #10 + '---------- -------------------------------------------- ---------------- ---------------- ---------- ---------- ----------';
  stlRelatorio.Add(sLinha);

  iTotalObjestos  := 0;
  iTotalFolhas    := 0;
  iTotalPaginas   := 0;

  while not __queryMySQL_processamento__.Eof do
  begin

    sLinha := objString.AjustaStr(__queryMySQL_processamento__.FieldByName('MOVIMENTO').AsString, 10)
      + ' ' + objString.AjustaStr(__queryMySQL_processamento__.FieldByName('ARQUIVO_AFP').AsString, 44)
      + ' ' + objString.AjustaStr(__queryMySQL_processamento__.FieldByName('DATA_POSTAGEM').AsString, 16, 1)
      + ' ' + objString.AjustaStr(__queryMySQL_processamento__.FieldByName('LOTE').AsString, 16, 1)
      + ' ' + FormatFloat('0000000000', __queryMySQL_processamento__.FieldByName('QTD').AsInteger)
      + ' ' + FormatFloat('0000000000',    __queryMySQL_processamento__.FieldByName('PAGINAS').AsInteger)
      + ' ' + FormatFloat('0000000000',    __queryMySQL_processamento__.FieldByName('FOLHAS').AsInteger)
      ;
    stlRelatorio.Add(sLinha);

    //=================================================================================================================================================================
    //  INSERE NA TABELA TRACK E CRIA CSV TRACK PRÉVIAS
    //=================================================================================================================================================================
    if not objParametrosDeEntrada.TESTE then
    begin
      sComando := 'INSERT INTO  ' + objParametrosDeEntrada.TABELA_TRACK
                + ' (ARQUIVO_ZIP, ARQUIVO_AFP, LOTE, TIMESTAMP, LINHAS, OBJETOS, FOLHAS, PAGINAS, STATUS_ARQUIVO, MOVIMENTO) '
                + ' VALUES("'
                +         __queryMySQL_processamento__.FieldByName('ARQUIVO_ZIP').AsString
                + '","' + __queryMySQL_processamento__.FieldByName('ARQUIVO_AFP').AsString
                + '","' + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, StrToInt(objParametrosDeEntrada.PEDIDO_LOTE))
                + '","' + FormatDateTime('YYYY-MM-DD hh:mm:ss', objParametrosDeEntrada.TIMESTAMP)
                + '","' + __queryMySQL_processamento__.FieldByName('QTD').AsString
                + '","' + __queryMySQL_processamento__.FieldByName('QTD').AsString
                + '","' + __queryMySQL_processamento__.FieldByName('FOLHAS').AsString
                + '","' + __queryMySQL_processamento__.FieldByName('PAGINAS').AsString
                + '","' + '0'
                + '","' + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO)
                + '")'
                ;
      objConexao.Executar_SQL(__queryMySQL_Insert_, sComando, 1);

      sLinha  := ''  // 'OF_FORMULARIO'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO)     // ';MOVIMENTO'
               + ';' // ';FILLER'
               + ';' + __queryMySQL_processamento__.FieldByName('ARQUIVO_AFP').AsString // ';ARQUIVO'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' + FormatDateTime('YYYY-MM-DD hh:mm:ss', objParametrosDeEntrada.TIMESTAMP)// ';TIMESTAMP'
               + ';' // ';ACABAMENTO'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' + FormatFloat(objParametrosDeEntrada.FORMATACAO_LOTE_PEDIDO, StrToInt(objParametrosDeEntrada.PEDIDO_LOTE))// ';LOTE_PROCESAMENTO'
               + ';' + __queryMySQL_processamento__.FieldByName('QTD').AsString     // ';QUANTIDADE_DE_OBJETOS_POR_OF'
               + ';' + __queryMySQL_processamento__.FieldByName('FOLHAS').AsString  // ';QUANTIDADE_DE_FOLHAS_POR_OF'
               + ';' + __queryMySQL_processamento__.FieldByName('PAGINAS').AsString // ';QUANTIDADE_DE_PAGINAS_POR_OF'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' // ';CARTAO_POSTAGEM'
               + ';' // ';DATA_LOTE_QTD_POSTAGEM'
               + ';' // ';TOTAL_LOCAL'
               + ';' // ';TOTAL_ESTADUAL'
               + ';' // ';TOTAL_NACIONAL'
               + ';' // ';TOTAL'
               + ';' // ';PORTE[GR]'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' // ';FILLER'
               + ';' // ';PAPEL'
               ;
      stlTrack.Add(sLinha);
      if objParametrosDeEntrada.CRIAR_CSV_TRACK then
        stlTrack.SaveToFile(objString.AjustaPath(objParametrosDeEntrada.PATH_TRACK) + StringReplace(__queryMySQL_processamento__.FieldByName('ARQUIVO_AFP').AsString, '.AFP', '.CSV', [rfReplaceAll, rfIgnoreCase]));

    end;

    if objParametrosDeEntrada.CRIAR_CSV_TRACK then
    BEGIN
      stlTrack.SaveToFile(sPathMovimentoTRACK                                     + StringReplace(__queryMySQL_processamento__.FieldByName('ARQUIVO_AFP').AsString, '.AFP', '.CSV', [rfReplaceAll, rfIgnoreCase]));
      objLogar.Logar(#13 + #10 + stlTrack.Text + #13 + #10);
    end;
    //=================================================================================================================================================================

    iTotalObjestos  := iTotalObjestos  + __queryMySQL_processamento__.FieldByName('QTD').AsInteger;
    iTotalFolhas    := iTotalFolhas    + __queryMySQL_processamento__.FieldByName('FOLHAS').AsInteger;
    iTotalPaginas   := iTotalPaginas   + __queryMySQL_processamento__.FieldByName('PAGINAS').AsInteger;
    
    __queryMySQL_processamento__.Next;
  end;

  sLinha := '---------- -------------------------------------------- ---------------- ---------------- ---------- ---------- ----------'
  + #13 + #10 + 'TOTAIS' + stringOfChar(' ', 84) + FormatFloat('0000000000', iTotalObjestos) + ' ' + FormatFloat('0000000000', iTotalPaginas) + ' ' + FormatFloat('0000000000', iTotalFolhas);
  stlRelatorio.Add(sLinha);

  sArquivoREL := sPathMovimentoRelatorio + 'RELATORIO_DE_QUANTIDADES_' + FormatDateTime('YYYYMMDD', objParametrosDeEntrada.MOVIMENTO) +'.REL';

  stlRelatorio.SaveToFile(sArquivoREL);
  objLogar.Logar(#13 + #10 + stlRelatorio.Text + #13 + #10);

  objFuncoesWin.ExecutarArquivoComProgramaDefault(sArquivoREL);

  objFuncoesWin.DelTree(sPathMovimentoTMP);
  //==================================================================================================================================================================


  objLogar.Logar('');


end;

procedure TCore.Atualiza_arquivo_conf_C(ArquivoConf, sINP, sOUT, sTMP, sLOG, sRGP: String);
var
  txtEntrada       : TextFile;
  sLinha           : string;
  sParametro       : string;
  stlArquivoConfC  : TStringList;
  sPathSaidaAFP    : string;
begin


  stlArquivoConfC := TStringList.Create();

  AssignFile(txtEntrada, ArquivoConf);
  Reset(txtEntrada);

  while not Eof(txtEntrada) do
  begin

    Readln(txtEntrada, sLinha);

    sParametro := AnsiUpperCase(Trim(objString.getTermo(1, '=', sLinha)));

    if sParametro = 'INP' then
      stlArquivoConfC.Add(sParametro + '=' + sINP);

    if sParametro = 'OUT' then
      stlArquivoConfC.Add(sParametro + '=' + sOUT);

    if sParametro = 'TMP' then
      stlArquivoConfC.Add(sParametro + '=' + sTMP);

    if sParametro = 'LOG' then
      stlArquivoConfC.Add(sParametro + '=' + sLOG);

    if sParametro = 'RGP' then
      stlArquivoConfC.Add(sParametro + '=' + sRGP);

  end;

  CloseFile(txtEntrada);

  stlArquivoConfC.SaveToFile(ArquivoConf);

end;

procedure TCore.execulta_app_c(app, arquivo_conf: string);
begin
  objFuncoesWin.ExecutarPrograma(app + ' "' + arquivo_conf + '"');
end;

function TCore.ArquivoExieteTabelaTrack(Arquivo: string): Boolean;
var
  sComando: string;
begin

  sComando := 'SELECT ARQUIVO_ZIP FROM ' + objParametrosDeEntrada.TABELA_TRACK
            + ' WHERE ARQUIVO_ZIP = "' + Arquivo + '" ';
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

  if __queryMySQL_processamento__.RecordCount > 0 then
   Result := True
  else
    Result := False;

end;

procedure TCore.getListaDeArquivosJaProcessados();
var
  sComando                   : string;
  sLinha                     : string;
begin

  sComando := 'SELECT * FROM ' + objParametrosDeEntrada.TABELA_TRACK
            + ' WHERE STATUS_ARQUIVO = "0" '
            + ' ORDER BY MOVIMENTO DESC';
  objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 2);

  objParametrosDeEntrada.STL_LISTA_ARQUIVOS_JA_PROCESSADOS.Clear;

  WHILE NOT __queryMySQL_processamento__.Eof do
  BEGIN

    sLinha := __queryMySQL_processamento__.FieldByName('MOVIMENTO').AsString
    + ' - ' + __queryMySQL_processamento__.FieldByName('ARQUIVO_ZIP').AsString
    + ' - ' + __queryMySQL_processamento__.FieldByName('ARQUIVO_AFP').AsString;

    objParametrosDeEntrada.STL_LISTA_ARQUIVOS_JA_PROCESSADOS.Add(sLinha);

    __queryMySQL_processamento__.Next;
  end;

end;

procedure TCore.ReverterArquivos();
var
  iContArquivos                       : Integer;
  sArquivoReverter                    : string;
  sComando                            : string;

begin

  for iContArquivos := 0 to objParametrosDeEntrada.STL_LISTA_ARQUIVOS_REVERTER.Count - 1 do
  begin

    sArquivoReverter := objParametrosDeEntrada.STL_LISTA_ARQUIVOS_REVERTER.Strings[iContArquivos];

    sComando := 'DELETE FROM ' + objParametrosDeEntrada.TABELA_TRACK
              + ' WHERE ARQUIVO_AFP = "' + sArquivoReverter + '" ';
    objConexao.Executar_SQL(__queryMySQL_processamento__, sComando, 1);

  end;

end;

end.
