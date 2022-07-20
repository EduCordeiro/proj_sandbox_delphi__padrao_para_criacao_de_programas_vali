object frmMain: TfrmMain
  Left = 306
  Top = 120
  Width = 751
  Height = 594
  Caption = 'ESQUELO'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object lblIdProcessamento: TLabel
    Left = 172
    Top = 488
    Width = 138
    Height = 16
    Caption = 'ID do Processamento:'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object lblIdProcessamentoValor: TLabel
    Left = 382
    Top = 488
    Width = 203
    Height = 17
    Alignment = taCenter
    AutoSize = False
    Caption = 'lblIdProcessamentoValor'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object lblNumeroDoLotePedido: TLabel
    Left = 172
    Top = 506
    Width = 89
    Height = 16
    Caption = 'Lote / Pedido:'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object lblNumeroDoLotePedidoValor: TLabel
    Left = 382
    Top = 507
    Width = 203
    Height = 19
    Alignment = taCenter
    AutoSize = False
    Caption = 'lblIdProcessamentoValor'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object lblUsuarioLogado: TLabel
    Left = 172
    Top = 524
    Width = 104
    Height = 16
    Caption = 'Usu'#225'rio Logado:'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object lblUsuarioLogadoValor: TLabel
    Left = 382
    Top = 525
    Width = 203
    Height = 16
    Alignment = taCenter
    AutoSize = False
    Caption = 'lblUsuarioLogadoValor'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object lblLoteLogin: TLabel
    Left = 172
    Top = 544
    Width = 77
    Height = 16
    Caption = 'Lote Login:'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object lblloteLoginValor: TLabel
    Left = 382
    Top = 544
    Width = 203
    Height = 16
    Alignment = taCenter
    AutoSize = False
    Caption = 'lblloteLoginValor'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object btnSobre: TBitBtn
    Left = 8
    Top = 488
    Width = 130
    Height = 30
    Caption = '&Sobre'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    OnClick = btnSobreClick
    Glyph.Data = {
      36030000424D3603000000000000360000002800000010000000100000000100
      18000000000000030000130B0000130B00000000000000000000FFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF9BA5A4000000FF
      FFFFFFFFFFFFFFFF000000BABFBFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFF00000039BDB8000000C1C6C600000037D4CD000000FFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000003FF8F03E
      F3EC0000003FF7EF36D6CE000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFF0000003DF0E83FF8F03FF8F03FF8F031B2AD000000FFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFDBDFDF0000003FF8F03F
      F8F03FF8F03FF8F03CECE5000000F8F8F8FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFF0000003FF8F03FF8F03FF8F03FF8F03FF8F03FF8F03DF0E90000
      00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF8F8F893B7B60000000000003D
      F1EA3FF8F03BD3CC0000000000009CB4B3FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFF6F6F60000003EF6EE000000FFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF9B
      B4B3000000DDDFDFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF}
  end
  object btnSair: TBitBtn
    Left = 604
    Top = 488
    Width = 130
    Height = 30
    Caption = 'Sai&r'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
    OnClick = btnSairClick
    Glyph.Data = {
      36030000424D3603000000000000360000002800000010000000100000000100
      18000000000000030000130B0000130B00000000000000000000FFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00000000000000
      0000000000000000000000000000000000000000000000FFFFFFFFFFFFFFFFFF
      FFFFFFF9F9F9FEFEFE000000B5B7B759605C000000354A3B425C4948644F4B68
      534D6A54000000FFFFFFFEFFFEFEFFFEFEFFFE000000000000000000DEE0E0C7
      C9C99194934F54510000002F41343D5442435D49000000FFFFFFFEFFFEFEFFFE
      FEFFFE0000002C3FAA000000E0E2E2D1D4D4BEC1C1B2B4B48B8D8D000000293A
      2D354A3B000000FFFFFF0000000000000000000000000000882C3FAA000000D1
      D4D4BEC1C1B7B9B99B9C9C00000026352A314335000000FFFFFF0000002B3EC8
      0B10A10E14A30000880000882C3FAA000000BEC1C1B7B9B99B9C9C0000002737
      2B2C3D31000000FFFFFF0000004562E40000990000990000990000990000882C
      3FAA0000006A6C6C9B9C9C00000024312726352A000000FFFFFF0000004562E4
      0A16B00A16B00A16B00A16B0101BB12939990000007779795A5B5B0000001B25
      1D1D2820000000FFFFFF0000004562E44562E44562E44562E41E42DF364DB100
      0000BBBDBDA2A4A49B9C9C000000141C16161F19000000FFFFFF000000000000
      0000000000005B75E73E53B3000000CDD0D0BEC1C1B7B9B99B9C9C0000000C11
      0E0D120F000000FFFFFFFFFFFFFFFFFFFFFFFF0000004562E4000000DCDEDED0
      D3D3BEC1C1B7B9B99B9C9C000000060806070907000000FFFFFFFFFFFFFFFFFF
      FFFFFF000000000000000000E0E2E2D0D3D3BDC0C0B5B7B7989A9A0000000000
      00000000000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000E0E2E2D0
      D3D3BDC0C0ABADAD919292000000000000000000000000FFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFF0000000000000000000000000000000000000000000000
      00000000000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF}
  end
  object pgcMain: TPageControl
    Left = 8
    Top = 8
    Width = 726
    Height = 473
    ActivePage = tbsReverter
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
    OnChange = pgcMainChange
    object tbsEntrada: TTabSheet
      Caption = 'Entrada'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ImageIndex = 1
      ParentFont = False
      object lblCaminhoArquivosEntrada: TLabel
        Left = 5
        Top = 50
        Width = 302
        Height = 14
        Caption = 'Indique o caminho onde est'#227'o os arquivos de entrada. '
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
        WordWrap = True
      end
      object lblInfos: TLabel
        Left = 7
        Top = 416
        Width = 704
        Height = 14
        Alignment = taRightJustify
        AutoSize = False
        Color = clBlack
        Font.Charset = ANSI_CHARSET
        Font.Color = clWhite
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
      end
      object btnSelecionarTodos: TBitBtn
        Left = 457
        Top = 67
        Width = 125
        Height = 24
        Caption = '&Selecionar Todos'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -9
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        OnClick = btnSelecionarTodosClick
        Glyph.Data = {
          36030000424D3603000000000000360000002800000010000000100000000100
          18000000000000030000130B0000130B00000000000000000000000000000000
          000000000000000000000000C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0FFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFF000000EBEBEBDCDCDCDDDDDDDEDEDE000000FAFAFAF6
          F6F6F7F7F7F7F7F7C0C0C0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000FFFFFF
          ECECECEEEEEEDBDBDB000000FFFFFFFAFAFAFBFBFBF6F6F6C0C0C0FFFFFF6C4E
          31FFFFFFFFFFFFFFFFFF000000FFFFFFEAEAEAEBEBEBD9D9D9000000FFFFFFFA
          FAFA6C4E31F6F6F66C4E31FFFFFF6C4E316C4E31FFFFFFFFFFFF000000FFFFFF
          FFFFFFFFFFFFEBEBEB000000FFFFFFFFFFFFFFFFFFFAFAFAC0C0C0FFFFFF6C4E
          31FFFFFFFFFFFFFFFFFF000000000000000000000000000000000000C0C0C0C0
          C0C0C0C0C0C0C0C0C0C0C0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000EBEBEB
          DCDCDCDDDDDDDEDEDE000000FAFAFAF6F6F6F7F7F7F7F7F7C0C0C0FFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFF000000FFFFFFECECECEEEEEEDBDBDB000000FFFFFFFA
          FAFAFBFBFBF6F6F6C0C0C0FFFFFF6C4E31FFFFFFFFFFFFFFFFFF000000FFFFFF
          EAEAEAEBEBEBD9D9D9000000FFFFFFFAFAFA6C4E31F6F6F66C4E31FFFFFF6C4E
          316C4E31FFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFFEBEBEB000000FFFFFFFF
          FFFFFFFFFFFAFAFAC0C0C0FFFFFF6C4E31FFFFFFFFFFFFFFFFFF000000000000
          000000000000000000000000C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0FFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFF000000EBEBEBDCDCDCDDDDDDDEDEDE000000FAFAFAF6
          F6F6F7F7F7F7F7F7C0C0C0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000FFFFFF
          ECECECEEEEEEDBDBDB000000FFFFFFFAFAFAFBFBFBF6F6F6C0C0C0FFFFFF765A
          40FFFFFFFFFFFFFFFFFF000000FFFFFFEAEAEAEBEBEBD9D9D9000000FFFFFFFA
          FAFA6C4E31F6F6F66C4E31FFFFFF6C4E316C4E31FFFFFFFFFFFF000000FFFFFF
          FFFFFFFFFFFFEBEBEB000000FFFFFFFFFFFFFFFFFFFAFAFAC0C0C0FFFFFF6C4E
          31FFFFFFFFFFFFFFFFFF000000000000000000000000000000000000C0C0C0C0
          C0C0C0C0C0C0C0C0C0C0C0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF}
      end
      object btnLimparSelecao: TBitBtn
        Left = 586
        Top = 67
        Width = 125
        Height = 24
        Caption = '&Limpar sele'#231#227'o'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -9
        Font.Name = 'Arial'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnClick = btnLimparSelecaoClick
        Glyph.Data = {
          36030000424D3603000000000000360000002800000010000000100000000100
          18000000000000030000130B0000130B00000000000000000000FAFAFAFAFAFA
          FAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFA
          FAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFA
          FAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFA
          FAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFA
          FAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFA
          FAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAD4D4D49F9F9F
          9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F
          9F9F9F9F9F9F9FD4D4D49F9F9FF2F2F2FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFF000000B1B1B1000000FFFFFF9F9F9F9F9F9FEFEFEF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFCFCFC0000
          00FCFCFCFCFCFC9F9F9F9F9F9FEBEBEBFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFF8F8F8000000F8F8F8F8F8F89F9F9F9F9F9FE9E9E9
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF6F6F60000
          00F6F6F6F6F6F69F9F9F9F9F9FE6E6E6FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFF3F3F3000000F3F3F3F3F3F39F9F9F9F9F9FE2E2E2
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFEFEF0000
          00EFEFEFEFEFEF9F9F9F9F9F9FD4D4D4E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0
          E0E0E0E0E0E0E0E0E0E0E0000000BBBBBB000000E0E0E09F9F9FD4D4D49F9F9F
          9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F9F
          9F9F9F9F9F9F9FD4D4D4FAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFA
          FAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFA
          FAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFA
          FAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFA
          FAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFAFA}
      end
      object cltArquivos: TCheckListBox
        Left = 4
        Top = 96
        Width = 705
        Height = 317
        OnClickCheck = cltArquivosClickCheck
        ItemHeight = 13
        TabOrder = 2
      end
      object edtPathEntrada: TJvDirectoryEdit
        Left = 4
        Top = 70
        Width = 449
        Height = 21
        DialogKind = dkWin32
        TabOrder = 3
        Text = 'edtPathEntrada'
        OnChange = edtPathEntradaChange
      end
      object chkTeste: TCheckBox
        Left = 13
        Top = 15
        Width = 312
        Height = 17
        Caption = 'RODAR COMO TESTE'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 4
        OnClick = chkTesteClick
      end
    end
    object tbsSaida: TTabSheet
      Caption = 'Sa'#237'da'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ImageIndex = 2
      ParentFont = False
      object lblCaminhoArquivosSaida: TLabel
        Left = 9
        Top = 211
        Width = 228
        Height = 14
        Caption = 'Indique o caminho dos arquivos de sa'#237'da.'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object edtPathSaida: TJvDirectoryEdit
        Left = 8
        Top = 232
        Width = 689
        Height = 21
        DialogKind = dkWin32
        TabOrder = 0
        Text = 'edtPathSaida'
      end
    end
    object tbsExecutar: TTabSheet
      Caption = 'Executar'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ImageIndex = 3
      ParentFont = False
      object lblMovimento: TLabel
        Left = 20
        Top = 28
        Width = 87
        Height = 20
        Caption = 'Movimento'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object btnExecutar: TBitBtn
        Left = 576
        Top = 399
        Width = 130
        Height = 30
        Caption = 'E&xecutar'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 0
        OnClick = btnExecutarClick
        Glyph.Data = {
          36030000424D3603000000000000360000002800000010000000100000000100
          18000000000000030000130B0000130B00000000000000000000FFFFFFDCDCDC
          8282828282828282828282828282828282828282828282828282828282828282
          82828282828282DCDCDCFFFFFF828282EDEDEDFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEDEDED828282FFFFFF828282
          FFFFFFB1B1B10000000000000000000000000000000000000000000000000000
          009D9D9DFFFFFF828282FFFFFF828282FFFFFF000000DFDFDFCFCFCFD1D1D1D4
          D4D4DADADADCDCDCDFDFDFDFDFDFCFCFCF000000FFFFFF828282FFFFFF828282
          FFFFFF000000FFFFFFD9D9D9DFDFDFE4E4E4EFEFEFF4F4F4F9F9F9FFFFFFDFDF
          DF000000FFFFFF828282FFFFFF828282FFFFFF000000FFFFFFD4D4D4D9D9D951
          5151E9E9E9EFEFEFF4F4F4F9F9F9DFDFDF000000FFFFFF828282FFFFFF828282
          FFFFFF000000FFFFFFCFCFCFD4D4D4000000535353E9E9E9EFEFEFF4F4F4DCDC
          DC000000FFFFFF828282FFFFFF828282FFFFFF000000FFFFFFC9C9C9CFCFCF00
          0000000000535353E9E9E9EFEFEFDADADA000000FFFFFF828282FFFFFF828282
          FFFFFF000000FFFFFFC4C4C4C9C9C90000004F4F4FDFDFDFE4E4E4E9E9E9D7D7
          D7000000FFFFFF828282FFFFFF828282FFFFFF000000FFFFFFBFBFBFC4C4C449
          4949D4D4D4D9D9D9DFDFDFE4E4E4D4D4D4000000FFFFFF828282FFFFFF828282
          FFFFFF000000FFFFFFBFBFBFBFBFBFC4C4C4CFCFCFD4D4D4D9D9D9DFDFDFD1D1
          D1000000FFFFFF828282FFFFFF828282FFFFFF000000FFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000FFFFFF828282FFFFFF828282
          FFFFFFC8C8C80000000000000000000000000000000000000000000000000000
          009D9D9DFFFFFF828282FFFFFF828282EDEDEDFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEDEDED828282FFFFFFDCDCDC
          8282828282828282828282828282828282828282828282828282828282828282
          82828282828282DCDCDCFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF}
      end
      object dtpMovimento: TDateTimePicker
        Left = 19
        Top = 47
        Width = 106
        Height = 21
        Date = 43915.520581469910000000
        Time = 43915.520581469910000000
        TabOrder = 1
      end
    end
    object tbsRelatorios: TTabSheet
      Caption = 'Relatorios'
      ImageIndex = 3
      object lblLote: TLabel
        Left = 16
        Top = 24
        Width = 106
        Height = 13
        Caption = 'INFORME O LOTE'
      end
      object lblSalvarRelatorio: TLabel
        Left = 17
        Top = 402
        Width = 116
        Height = 13
        Caption = 'Salvar Relatorio em:'
      end
      object edtLote: TEdit
        Left = 16
        Top = 40
        Width = 121
        Height = 21
        TabOrder = 0
        Text = 'edtLote'
      end
      object btnPesquisar: TButton
        Left = 141
        Top = 39
        Width = 97
        Height = 25
        Caption = 'PESQUISAR'
        TabOrder = 1
        OnClick = btnPesquisarClick
      end
      object mmoRelatorio: TMemo
        Left = 16
        Top = 72
        Width = 681
        Height = 321
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -9
        Font.Name = 'Courier New'
        Font.Style = [fsBold]
        Lines.Strings = (
          'mmoRelatorio')
        ParentFont = False
        ReadOnly = True
        ScrollBars = ssVertical
        TabOrder = 2
      end
      object rgTipoLote: TRadioGroup
        Left = 448
        Top = 24
        Width = 241
        Height = 41
        Caption = 'Status do Lote'
        Columns = 2
        ItemIndex = 0
        Items.Strings = (
          'V'#225'lido'
          'Inv'#225'lido')
        TabOrder = 3
      end
      object edtSalvarRelatorio: TJvDirectoryEdit
        Left = 137
        Top = 399
        Width = 472
        Height = 21
        DialogKind = dkWin32
        TabOrder = 4
        Text = 'edtSalvarRelatorio'
      end
      object btnSalvarRelatorio: TButton
        Left = 619
        Top = 397
        Width = 75
        Height = 25
        Caption = 'Salvar'
        TabOrder = 5
        OnClick = btnSalvarRelatorioClick
      end
    end
    object tbsReverter: TTabSheet
      Caption = 'Rever Processo'
      ImageIndex = 4
      object lblReverterTitulo: TLabel
        Left = 22
        Top = 91
        Width = 117
        Height = 13
        Caption = 'Movimento   Arquivo'
      end
      object lblArquivoPesquisa: TLabel
        Left = 9
        Top = 9
        Width = 243
        Height = 13
        Caption = 'FILTRO PARA MARCA'#199#195'O DE ARQUIVOS'
      end
      object lblReverterArquivosMarcados: TLabel
        Left = 8
        Top = 412
        Width = 167
        Height = 13
        Caption = 'lblReverterArquivosMarcados'
      end
      object chklstArquivosProcessados: TCheckListBox
        Left = 5
        Top = 106
        Width = 705
        Height = 296
        ItemHeight = 13
        TabOrder = 0
      end
      object edtArquivoMarcar: TEdit
        Left = 5
        Top = 27
        Width = 376
        Height = 21
        TabOrder = 1
        Text = 'edtArquivoMarcar'
      end
      object btnMarcarArquivo: TButton
        Left = 383
        Top = 25
        Width = 128
        Height = 25
        Caption = 'Marcar Arquivo'
        TabOrder = 2
        OnClick = btnMarcarArquivoClick
      end
      object btnReverterArquivo: TButton
        Left = 529
        Top = 405
        Width = 180
        Height = 25
        Caption = 'Reverter arquvos marcados'
        TabOrder = 3
        OnClick = btnReverterArquivoClick
      end
      object btnMarcarTodosReverter: TButton
        Left = 426
        Top = 81
        Width = 140
        Height = 25
        Caption = 'MARCAR TODOS'
        TabOrder = 4
        OnClick = btnMarcarTodosReverterClick
      end
      object btnDesmarcarTodosReverter: TButton
        Left = 569
        Top = 81
        Width = 140
        Height = 25
        Caption = 'DESMARCAR TODOS'
        TabOrder = 5
        OnClick = btnDesmarcarTodosReverterClick
      end
    end
  end
  object tmrVerificacoesRegulares: TTimer
    Interval = 10
    OnTimer = tmrVerificacoesRegularesTimer
    Left = 694
    Top = 5
  end
end
