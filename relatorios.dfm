object Form1: TForm1
  Left = 66
  Top = 0
  Width = 1240
  Height = 728
  Caption = 'Relatorio'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  WindowState = wsMaximized
  PixelsPerInch = 96
  TextHeight = 13
  object Button2: TButton
    Left = 624
    Top = 640
    Width = 75
    Height = 25
    Caption = 'Button2'
    TabOrder = 0
  end
  object TabControl1: TTabControl
    Left = 24
    Top = 6
    Width = 1265
    Height = 610
    TabOrder = 1
    object TLabel
      Left = 136
      Top = 64
      Width = 3
      Height = 13
    end
    object Label2: TLabel
      Left = 136
      Top = 16
      Width = 75
      Height = 13
      Caption = 'DATA INICIAL: '
    end
    object Label1: TLabel
      Left = 440
      Top = 16
      Width = 68
      Height = 13
      Caption = 'DATA FINAL: '
    end
    object date_in: TDateTimePicker
      Left = 136
      Top = 40
      Width = 186
      Height = 21
      Date = 42150.690871585650000000
      Time = 42150.690871585650000000
      TabOrder = 0
    end
    object DBGrid1: TDBGrid
      Left = 23
      Top = 70
      Width = 1218
      Height = 515
      DataSource = DataSource1
      TabOrder = 1
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
    end
    object date_fin: TDateTimePicker
      Left = 440
      Top = 40
      Width = 186
      Height = 21
      Date = 42150.690897719910000000
      Time = 42150.690897719910000000
      TabOrder = 2
    end
    object Btn1: TButton
      Left = 848
      Top = 17
      Width = 89
      Height = 33
      Caption = 'Gerar Relat'#243'rio'
      TabOrder = 3
      OnClick = Btn1Click
    end
  end
  object DataSource1: TDataSource
    DataSet = ADOQuery1
    Left = 40
    Top = 8
  end
  object ADOConnection1: TADOConnection
    Connected = True
    ConnectionString = 
      'Provider=MSDAORA.1;Password=ninguemsabe;User ID=cnp;Data Source=' +
      'bravos;Persist Security Info=True'
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 72
    Top = 8
  end
  object ADOQuery1: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        'SELECT DISTINCT os.empresa, os.revenda,PF.CPF,fat_cliente.NOME,L' +
        'OGRADOURO_ENTREGA,COMPLEMENTO_ENTREGA,BAIRRO_ENTREGA,MUNICIPIO_E' +
        'NTREGA,UF_ENTREGA, CEP_ENTREGA,'
      'E_MAIL_CASA,DTA_NASCIMENTO,fat_cliente.ddd_TELEFONE,'
      'fat_cliente.TELEFONE,FAX,DDD_CELULAR,CELULAR,'
      
        'os.nro_os,os.dta_emissao,os.dta_encerramento,os.situacao_os,cons' +
        'ultor.nome AS NOME_CONSULTOR,mecanico.nome AS NOME_MECANICO,os.c' +
        'ategoria_os,oso.descricao,'
      
        'mo.des_modelo,ofi_ficha_seguimento.ano_fabricacao,ofi_ficha_segu' +
        'imento.ano_MODELO,ofi_atendimento.chassi,'
      
        'ofi_ficha_seguimento.placa,ofi_ficha_seguimento.dta_venda, os.ki' +
        'lometragem, dep.nome '
      ''
      
        'from OFI_ORDEM_SERVICO OS, ofi_servico_os oso, OFI_PARAMETRO PAR' +
        ', CAC_CONTATO CONTATO, OFI_ATENDIMENTO ATENDE, '
      
        '     FAT_CLIENTE , FAT_PESSOA_FISICA PF, ofi_atendimento, ofi_fi' +
        'cha_seguimento, vei_modelo mo,'
      
        '     fat_vendedor consultor, ofi_mecanico mecanico, ger_departam' +
        'ento dep'
      '     '
      'where ATENDE.EMPRESA = OS.EMPRESA '
      '  and ATENDE.REVENDA = OS.REVENDA '
      '  and PAR.EMPRESA = OS.EMPRESA '
      '  and PAR.REVENDA = OS.REVENDA '
      '  and ATENDE.CONTATO = OS.CONTATO '
      '  and CONTATO.EMPRESA = OS.EMPRESA '
      '  and CONTATO.REVENDA = OS.REVENDA '
      '  and CONTATO.CONTATO = OS.CONTATO '
      '  and fat_cliente.CLIENTE = CONTATO.CLIENTE '
      '  and dep.empresa = atende.empresa'
      '  and dep.revenda = atende.revenda'
      '  and dep.departamento = atende.departamento'
      '  and mo.modelo = ofi_ficha_seguimento.modelo'
      '  and OS.EMPRESA IN (1,4) '
      '  and OS.REVENDA IN (1,3,4) '
      '  and OS.SITUACAO_OS = 9 '
      '  and os.categoria_os in (1,2,3,4,6,7,8,9,11,22)'
      '  and FAT_CLIENTE.CLIENTE = PF.CLIENTE'
      
        '  and dta_encerramento between TO_DATE('#39'01042015'#39', '#39'DDMMYYYY'#39') a' +
        'nd TO_DATE('#39'30042015'#39', '#39'DDMMYYYY'#39')'
      '  and os.empresa = oso.empresa '
      '  and os.revenda = oso.revenda '
      '  and os.nro_os = oso.nro_os '
      '  and os.empresa = ofi_atendimento.empresa'
      '  and os.revenda = ofi_atendimento.revenda '
      '  and os.contato = ofi_atendimento.contato'
      '  and ofi_atendimento.chassi = ofi_ficha_seguimento.chassi'
      '  and ofi_atendimento.empresa = consultor.empresa(+)'
      '  and ofi_atendimento.revenda = consultor.revenda(+)'
      '  and ofi_atendimento.vendedor = consultor.vendedor(+)'
      '  and oso.empresa = mecanico.empresa(+)'
      '  and oso.revenda = mecanico.revenda(+)'
      '  and oso.mecanico = mecanico.mecanico(+)'
      '  order by os.nro_os')
    Left = 104
    Top = 8
  end
end
