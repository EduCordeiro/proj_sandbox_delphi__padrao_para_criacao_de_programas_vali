Unit ClassLayoutArquivo;

interface

  uses Classes, Dialogs, SysUtils, Forms, Controls, Graphics,
  StdCtrls, ComCtrls;

  type
    TLayoutCliente= Class
      // Propriedades da Classe ClassLayoutArquivo
      sCodigoOperadora              : String;
      sContratoNET                  : String;
      sContratoEBT                  : String;
      sNomeCliente                  : String;
      sEnderecoCliente              : String;
      sBairroCliente                : String;
      sCidadeCliente                : String;
      sCepCliente                   : String;
      sUFCliente                    : String;
      sNomeOperadora                : String;
      sEnderecoOperadora            : String;
      sBairroOperadora              : String;
      sCidadeOperadora              : String;
      sCEPOperadora                 : String;
      sUFOperadora                  : String;
      sBaseDaOperadora              : String;
      sMensagemStatus               : String;
      sforma_envio                  : String;
      sDestinoPostagem              : String;
      sTimeStampProcessamento       : String;
      sHASH_CODE                    : String;
      sBase                         : String;
      sNomeDoArquivo                : String;
      StatusInvalido                : String;
      sNumeroPedido                 : String;
      sMensagem                     : String;
      sEmailDestino                 : String;
      sEmailOrigem                  : String;
      sCidadeContrato               : String;
      sIDCarta                      : String;
      sLinhaTipo03                  : String;
      stlLinhaTipo04                : TStringList;
      sLinhaTipo05                  : String;

    end;

implementation


End.
