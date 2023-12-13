Attribute VB_Name = "modDatabaseGeral"
Option Explicit

Const adOpenFowardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Public CONN As adoce.Connection

Public strTransfereDados As String

Public Sub CriarBase()
    '
    Dim rs
    '
    Dim intTabelas As Integer
    Dim intTamanho As Integer
    Dim intPassagem As Integer
    Dim intComprimento As Integer
    Dim intContador As Integer
    Dim intProgress As Integer
    '
    Screen.MousePointer = 0
    '
    AddStatus "Gerando Banco de Dados..."
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    '
    rs.Open "CREATE DATABASE '" & strPath & "\base.cdb" & "'"
    '
    AddStatus "Base de dados criada"
    '
    connOpen
    '
    intTamanho = 3400
    intPassagem = 1
    intComprimento = 0
    intContador = 0
    intTabelas = 26
    '
    ExecSQL "CREATE TABLE brand(codigo VARCHAR(3), descricao VARCHAR(20), status VARCHAR(1))"
    AddStatus "Criando Tabela de Brand"
    '
    intPassagem = 4
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    intProgress = 1
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ' If MsgBox("IntProgress = " & CInt(intProgress), vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
    '
    '=============================================================================================
    '
    ExecSQL "CREATE TABLE cidades(codigo VARCHAR(5), nome VARCHAR(30), uf VARCHAR(2))"
    AddStatus "Criando Tabela de Cidades"
    '
    intPassagem = 8
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE clientes(codigo_vendedor VARCHAR(5), codigo_cliente VARCHAR(5), " _
            & "nome_fantasia VARCHAR(20), razao_social VARCHAR(40), endereco_entrega VARCHAR(40) " _
            & ", cep_entrega VARCHAR(8), bairro_entrega VARCHAR(20), cidade_entrega VARCHAR(5), " _
            & "endereco_cobranca VARCHAR(40), cep_cobranca VARCHAR(8), bairro_cobranca VARCHAR(20), " _
            & "cidade_cobranca VARCHAR(5), telefone VARCHAR(12), fax VARCHAR(12), email VARCHAR(30), " _
            & "www VARCHAR(30), data_fundacao VARCHAR(8), predio_proprio VARCHAR(1), " _
            & "referencia_bancaria_1 VARCHAR(20), referencia_bancaria_2 VARCHAR(20), " _
            & "referencia_comercial_1 VARCHAR(20), referencia_comercial_2 VARCHAR(20), " _
            & "contato VARCHAR(25), data_ultima_compra VARCHAR(8), valor_ultima_compra VARCHAR(10), " _
            & "cnpjmf VARCHAR(14), incricao_estadual VARCHAR(20), cpf VARCHAR(14), rg VARCHAR(20), " _
            & "desconto_maximo VARCHAR(5), limite_credito VARCHAR(10), bloqueado VARCHAR(1), " _
            & "condicao_pagamento_padrao VARCHAR(2), forma_pagamento_padrao VARCHAR(2), " _
            & "periodicidade_visita VARCHAR(2), status VARCHAR(1), ramo_atividade VARCHAR(3));"
            '
    AddStatus "Criando Tabela de Clientes"
    '
    intPassagem = 12
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    '============================== CRIA TABELA DE CONTADORES ==============================
    '
    ExecSQL "CREATE TABLE contadores(contadores VARCHAR(3), contador_filtro VARCHAR(3));"
    ExecSQL "INSERT INTO contadores (contadores, contador_filtro) VALUES ('000','000');"
    '
    AddStatus "Criando Tabela de Condições de Pagamentos"
    '
    '=======================================================================================
    '
    ExecSQL "CREATE TABLE condicao_pagamento(codigo_condicao VARCHAR(2), codigo_tabela_preco VARCHAR(1), descricao VARCHAR(20), dia_pgto_1p VARCHAR(2), dia_pgto_2p VARCHAR(2), dia_pgto_3p VARCHAR(2), dia_pgto_4p VARCHAR(2), dia_pgto_5p VARCHAR(2), pedido_minimo VARCHAR(10), status VARCHAR(1));"
    AddStatus "Criando Tabela de Condições de Pagamentos"
    '
    intPassagem = 16
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    '=================================================================================
    '
    '==================================================================================
    '
    ExecSQL "CREATE TABLE destinatarios(codigo_destinatario VARCHAR(3), nome VARCHAR(20),  status VARCHAR(1));"
    AddStatus "Criando Tabela de Destinatarios"
    '
    intPassagem = 18
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    '=================================================================================
    '
    ExecSQL "CREATE TABLE descricao_tabela(codigo_tabela VARCHAR(1), descricao VARCHAR(10));"
    AddStatus "Criando Tabela de Descrições"
    '
    intPassagem = 20
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE desconto_canal(sub_brand VARCHAR(3), cliente VARCHAR(5), desconto VARCHAR(6));"
    AddStatus "Criando Tabela de Descontos de Canal"
    '
    intPassagem = 24
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE estoque(codigo_cliente VARCHAR(5), codigo_produto VARCHAR(6), data VARCHAR(8), dias VARCHAR(3), media_diaria VARCHAR(12), media VARCHAR(8), estoque VARCHAR(8), sugestao VARCHAR(8), pedido VARCHAR(8));"
    AddStatus "Criando Tabela de Estoque"
    '
    ExecSQL "CREATE TABLE historico(codigo_cliente VARCHAR(5), codigo_produto VARCHAR(6), data VARCHAR(8), dias VARCHAR(3), media_diaria VARCHAR(12), media VARCHAR(8), estoque VARCHAR(8), sugestao VARCHAR(8), pedido VARCHAR(8));"
    AddStatus "Criando Tabela de Históricos"
    '
    intPassagem = 28
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE fabricante(codigo VARCHAR(3), descricao VARCHAR(50), status VARCHAR(1));"
    AddStatus "Criando Tabela de Fabricantes"
    '
    intPassagem = 32
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE forma_pagamento(codigo VARCHAR(2), descricao VARCHAR(15), status VARCHAR(1));"
    AddStatus "Criando Tabela de Formas de pagamentos"
    '
    intPassagem = 36
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE itens_pedido(numero_pedido VARCHAR(6), codigo_produto VARCHAR(6), qtd_pedida VARCHAR(8), qtd_faturada VARCHAR(8), valor_unitario VARCHAR(10), desconto VARCHAR(6));"
    '
    AddStatus "Criando Tabela de Itens de Pedidos"
    '
    intPassagem = 40
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE justificativa_nao_venda(codigo VARCHAR(3), tipo VARCHAR(1), descricao VARCHAR(50));"
    AddStatus "Criando Tabela de Justificativa de não Venda"
    '
    intPassagem = 44
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE mensagens(tipo_mensagem VARCHAR(1), codigo_vendedor_origem VARCHAR(5), codigo_vendedor_destino VARCHAR(5), assunto VARCHAR(20), data VARCHAR(8), hora VARCHAR(6), status VARCHAR(1) , mensagem text);"
    AddStatus "Criando Tabela de Mensagens"
    '
    intPassagem = 48
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE objetivo_venda(codigo_vendedor VARCHAR(5), codigo_produto VARCHAR(6), cota_qtde VARCHAR(10), realizado VARCHAR(10));"
    AddStatus "Criando Tabela de Objetivos de Vendas"
    '
    intPassagem = 52
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ' ExecSQL "CREATE TABLE observacoes_clientes(cliente VARCHAR(7), observacao TEXT);" Euclides
    '
    ExecSQL "CREATE TABLE observacoes_clientes(cliente VARCHAR(5), observacao TEXT);"
    AddStatus "Criando Tabela de Observações de Clientes"
    '
    intPassagem = 56
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE pedido(codigo_vendedor VARCHAR(5), codigo_cliente VARCHAR(5), numero_pedido_interno VARCHAR(6), numero_pedido_externo VARCHAR(6), pedido_cliente VARCHAR(10), data_emissao VARCHAR(8), hora_emissao VARCHAR(6), data_entrega VARCHAR(8), acrescimo_valor VARCHAR(10), desconto_valor VARCHAR(10), forma_pgto VARCHAR(2), condicao_pgto VARCHAR(2), tipo_movimento VARCHAR(1), status VARCHAR(1), observacao TEXT);"
    AddStatus "Criando Tabela de Pedidos"
    '
    intPassagem = 60
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE produtos(fabricante VARCHAR(3), brand VARCHAR(3), sub_brand VARCHAR(3), codigo_produto VARCHAR(6), descricao VARCHAR(50), unidade VARCHAR(3), qtd_disponivel VARCHAR(4), aliquota_icms VARCHAR(5), aliquota_ipi VARCHAR(5), substituicao_tributaria VARCHAR(1), quantidade_embalagem VARCHAR(4), desconto_acrescimo VARCHAR(6), peso VARCHAR(9), volume VARCHAR(9), desconto_maximo VARCHAR(6), embalagem VARCHAR(10), empresa VARCHAR(2), filial VARCHAR(2), preco1 VARCHAR(10), preco2 VARCHAR(10), preco3 VARCHAR(10), preco4 VARCHAR(10), preco5 VARCHAR(10), media_diaria VARCHAR(12), estoque VARCHAR(10), sugestao VARCHAR(10), pedido VARCHAR(10), filtro VARCHAR(3) ,status VARCHAR(1));"
    AddStatus "Criando Tabela de Produtos"
    '
    intPassagem = 64
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE promocoes(codigo_produto VARCHAR(6), qtd_inicial VARCHAR(8), qtd_final VARCHAR(8), preco_promocional VARCHAR(10));"
    AddStatus "Criando Tabela de Promoções"
    '
    intPassagem = 68
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE ramo_atividade(codigo VARCHAR(3), descricao VARCHAR(50), status VARCHAR(1));"
    AddStatus "Criando Tabela de Ramos de Atividade"
    '
    intPassagem = 72
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE relatorios(descricao VARCHAR(25), relatorio text, data VARCHAR(8), hora VARCHAR(6));"
    AddStatus "Criando Tabela de Relatórios"
    '
    intPassagem = 76
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE roteiro_percorrido(codigo_vendedor_atual VARCHAR(5), data_base VARCHAR(8), codigo_cliente VARCHAR(5), data_visita VARCHAR(8), hora_visita VARCHAR(6), motivo_nao_visita VARCHAR(3), visita_extra VARCHAR(1), automatico VARCHAR(1), dia_visita VARCHAR(1),  status VARCHAR(1));"
    AddStatus "Criando Tabela de Roteiro Percorrido"
    '
    intPassagem = 80
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE sistematica_visita(codigo_vendedor VARCHAR(5), codigo_cliente VARCHAR(5), dia VARCHAR(1), numero_visita VARCHAR(3));"
    AddStatus "Criando Tabela de Sistemática de Visita"
    '
    intPassagem = 84
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE sub_brand(codigo VARCHAR(3), descricao VARCHAR(50), observacao VARCHAR(40), status VARCHAR(1));"
    AddStatus "Criando Tabela de Sub Brand"
    '
    intPassagem = 88
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE tabela_precos(tabela_precos VARCHAR(1), produto VARCHAR(6), preco VARCHAR(10));" ' , desconto VARCHAR(6)
    AddStatus "Criando Tabela de Preços"
    '
    intPassagem = 92
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE tipo_movimento(codigo VARCHAR(1), descricao VARCHAR(15));"
    AddStatus "Criando Tabela de Tipos de Movimentos"
    '
    intPassagem = 96
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE titulos_aberto(codigo_cliente VARCHAR(5), numero_documento VARCHAR(20), data_vencimento VARCHAR(8), vencimento_data VARCHAR(8), valor VARCHAR(10));"
    AddStatus "Criando Tabela de Titulos em Aberto"
    '
    intPassagem = 98
    intComprimento = (Int(intTamanho / 100) * intPassagem)
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    ExecSQL "CREATE TABLE vendedor(codigo_vendedor VARCHAR(5), senha VARCHAR(6), nome VARCHAR(50), aceita_pedido_bloq VARCHAR(1), contra_senha VARCHAR(1), codigo_proximo_cliente VARCHAR(5), extra1 VARCHAR(1), extra2 VARCHAR(1), numero_proximo_pedido VARCHAR(8), tempo_maximo VARCHAR(3), habilitar_desconto_item VARCHAR(1), habilitar_desconto_pedido VARCHAR(1), habilitar_edicao_preco VARCHAR(1), habilitar_acrescimo VARCHAR(1), habilitar_cobranca_titulo VARCHAR(1), empresa VARCHAR(2), filial VARCHAR(2), data_cortetitulos VARCHAR(8), pedido_minimo VARCHAR(10), mensagem TEXT, status VARCHAR(1));"
    '
    AddStatus "Criando Tabela de Vendedor"
    '
    intPassagem = 100
    intComprimento = intTamanho
    '
    Progresso intProgress, intComprimento, intTamanho, intPassagem
    '
    Screen.MousePointer = 11
    '
    Set rs = Nothing
    '
    connClose
    '
    On Error GoTo 0
    '
End Sub

Public Function ExecSQL(paramSQL As String) As Boolean
    On Error Resume Next
    CONN.Execute (paramSQL)
    If CONN.Errors.Count > 0 Then
        ExecSQL = False
    Else
        ExecSQL = True
    End If
End Function

Public Function connOpen() As Boolean
    On Error Resume Next
    connOpen = True
    If CONN Is Nothing Then
        Set CONN = CreateObject("ADOCE.Connection.3.0")
        CONN.Open strPath & "\base.cdb"
        If CONN.Errors.Count > 0 Then connOpen = False
    End If
    On Error GoTo 0
End Function

Public Sub connClose()
    On Error Resume Next
    CONN.Close
    Set CONN = Nothing
    On Error GoTo 0
End Sub

Public Sub EncheTitulosPendentes(ByVal strNomeFantasia As String)
    '
    Dim rs
    '
    Dim strCodigo   As String
    Dim strFormataNumero   As String
    Dim strControle As String
    '
    Dim mpos As Integer
    '
    Dim valVencidas As Double
    Dim valVencer   As Double
    Dim valTotal    As Double
    Dim strDataHoje As String
    Dim strDataCorte As String
    '
    Screen.MousePointer = 11
    '
    frmContas.GridCtrl.Visible = False
    '
    valVencidas = 0
    valVencer = 0
    valTotal = 0
    '
    frmContas.Label2.Caption = FormatCurrency(CStr(valVencidas), 2, vbTrue, vbTrue, vbTrue)
    frmContas.Label5.Caption = FormatCurrency(CStr(valVencer), 2, vbTrue, vbTrue, vbTrue)
    frmContas.lblTotal.Caption = FormatCurrency(CStr(valTotal), 2, vbTrue, vbTrue, vbTrue)
    '
    frmContas.GridCtrl.Rows = 1
    frmContas.GridCtrl.Clear
    frmContas.GridCtrl.Row = 0
    frmContas.GridCtrl.Col = 0
    frmContas.GridCtrl.CellBackColor = &HC0C0C0
    frmContas.GridCtrl.CellFontBold = True
    frmContas.GridCtrl.Row = 0
    frmContas.GridCtrl.Col = 1
    frmContas.GridCtrl.CellBackColor = &HC0C0C0
    frmContas.GridCtrl.CellFontBold = True
    frmContas.GridCtrl.Row = 0
    frmContas.GridCtrl.Col = 2
    frmContas.GridCtrl.CellBackColor = &HC0C0C0
    '
    frmContas.GridCtrl.CellFontBold = True
    '
    frmContas.GridCtrl.ColWidth(0) = 1660 ' 1600
    frmContas.GridCtrl.ColWidth(1) = 900 ' 2000
    frmContas.GridCtrl.ColWidth(2) = 800 ' 1500
    '
    frmContas.GridCtrl.TextMatrix(0, 0) = "Documento"
    frmContas.GridCtrl.TextMatrix(0, 1) = "Venci/to"
    frmContas.GridCtrl.TextMatrix(0, 2) = "Valor"
    '
    connOpen
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    '
    rs.Open "SELECT * FROM vendedor WHERE codigo_vendedor='" & Trim(usrCodigoVendedor) & "';", CONN, adOpenDynamic, adLockOptimistic
    '
    strDataCorte = ""
    '
    If rs.RecordCount = 1 Then
       '
       strDataHoje = Mid(rs("data_cortetitulos"), 5, 4) & Mid(rs("data_cortetitulos"), 3, 2) & Mid(rs("data_cortetitulos"), 1, 2)
       '
    Else
       '
       strDataHoje = Mid(RetornaDataString(Now), 5, 4) & Mid(RetornaDataString(Now), 3, 2) & Mid(RetornaDataString(Now), 1, 2)
       '
       strDataCorte = strDataHoje
       '
    End If
    '
    If strDataHoje = "00000000" Then
       '
       strDataHoje = Mid(RetornaDataString(Now), 5, 4) & Mid(RetornaDataString(Now), 3, 2) & Mid(RetornaDataString(Now), 1, 2)
       '
       strDataCorte = strDataHoje
       '
    End If
    '
    ' MsgBox "Data de Corte=" & strDataHoje, vbOKOnly + vbInformation, App.Title
    '
    frmContas.Label8.Caption = Trim(Mid(strDataHoje, 7, 2)) & "/" & Trim(Mid(strDataHoje, 5, 2)) & "/" & Trim(Mid(strDataHoje, 1, 4))
    '
    If rs.State = 1 Then rs.Close
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    '
    rs.Open "SELECT * FROM clientes WHERE nome_fantasia='" & Trim(strNomeFantasia) & "';", CONN, adOpenDynamic, adLockReadOnly
    '
    If rs.RecordCount <= 0 Then
       '
       frmContas.GridCtrl.Visible = True
       '
       Screen.MousePointer = 0
       '
       MsgBox "Cliente inexistente.", vbOKOnly + vbInformation, App.Title
       '
       frmContas.lblTotal.Caption = "R$0,00"
       '
       Exit Sub
       '
    End If
    '
    strCodigo = rs("codigo_cliente")
    '
    If rs.State = 1 Then rs.Close
    '
    rs.Open "SELECT * FROM titulos_aberto WHERE codigo_cliente='" & Trim(strCodigo) & "' order by vencimento_data ASC;", CONN, adOpenDynamic, adLockReadOnly
    '
    If rs.RecordCount <= 0 Then
       '
       frmContas.GridCtrl.Visible = True
       '
       Screen.MousePointer = 0
       '
       MsgBox "Não há títulos pendentes para este cliente.", vbOKOnly + vbInformation, App.Title
       '
       frmContas.lblTotal.Caption = "R$0,00"
       '
       Exit Sub
       '
    End If
    '
    Do Until rs.EOF
       '
       frmContas.GridCtrl.Rows = frmContas.GridCtrl.Rows + 1
       '
       frmContas.GridCtrl.TextMatrix(frmContas.GridCtrl.Rows - 1, 0) = rs("numero_documento")
       frmContas.GridCtrl.TextMatrix(frmContas.GridCtrl.Rows - 1, 1) = Trim(Mid(rs("data_vencimento"), 1, 2)) & "/" & Trim(Mid(rs("data_vencimento"), 3, 2)) & "/" & Trim(Mid(rs("data_vencimento"), 5, 4))
       '
       ' strFormataNumero = FormatCurrency(rs("valor"), 2, vbTrue, vbTrue, vbTrue)
       '
       '=========================================================================
       '
       strFormataNumero = rs("valor")
       '
       If InStr(1, Trim(strFormataNumero), ",", vbTextCompare) = 0 Then
          '
          strFormataNumero = Trim(strFormataNumero) & ",00"
          '
       Else
          '
          strFormataNumero = strFormataNumero & "00"
          '
          strFormataNumero = Mid(Trim(strFormataNumero), 1, InStr(1, Trim(strFormataNumero), ",", vbTextCompare) + 2)
          '
       End If
       '
       While Mid(Trim(strFormataNumero), 1, 1) = "0"
          '
          strFormataNumero = Mid(Trim(strFormataNumero), 2, Len(Trim(strFormataNumero)) - 1)
          '
       Wend
       '
       If Mid(Trim(strFormataNumero), 1, 1) = "," Then strFormataNumero = "0" & strFormataNumero
       '
       frmContas.GridCtrl.TextMatrix(frmContas.GridCtrl.Rows - 1, 2) = strFormataNumero
       '
       '=========================================================================
       '
       frmContas.GridCtrl.Col = 1
       frmContas.GridCtrl.CellAlignment = flexAlignRightCenter
       '
       frmContas.GridCtrl.Col = 0
       '
       If Mid(rs("data_vencimento"), 5, 4) & Mid(rs("data_vencimento"), 3, 2) & Mid(rs("data_vencimento"), 1, 2) < strDataHoje Then
          '
          valVencidas = valVencidas + CDbl(rs("valor"))
          '
       Else
          '
          valVencer = valVencer + CDbl(rs("valor"))
          '
       End If
       '
       rs.MoveNext
       '
    Loop
    '
    valTotal = valVencer + valVencidas
    '
    frmContas.Label2.Caption = FormatCurrency(CStr(valVencidas), 2, vbTrue, vbTrue, vbTrue)
    frmContas.Label5.Caption = FormatCurrency(CStr(valVencer), 2, vbTrue, vbTrue, vbTrue)
    frmContas.lblTotal.Caption = FormatCurrency(CStr(valTotal), 2, vbTrue, vbTrue, vbTrue)
    '
    If rs.State = 1 Then rs.Close
    '
    connClose
    '
    Set rs = Nothing
    '
    Screen.MousePointer = 0
    '
    frmContas.GridCtrl.Visible = True
    '
    If strDataCorte <> "" Then
       '
       ' MsgBox "Data de Corte inválida para os Títulos, Assume a data atual:" & Now, vbOKOnly + vbInformation, App.Title
       '
    End If
    '
End Sub

Public Sub EncheComboVendedores()
  '
  Dim rs
  '
  Screen.MousePointer = 11
  '
  connOpen
  '
  Set rs = CreateObject("ADOCE.Recordset.3.0")
  '
  rs.Open "SELECT * from Destinatarios order by nome;", CONN, adOpenForwardOnly, adLockReadOnly
  '
  frmMensagens.cboNDestinatario.Clear
  '
  frmMensagens.cboEDestinatario.Clear
  '
  frmMensagens.cboRemetente.Clear
  '
  Do Until rs.EOF
     '
     frmMensagens.cboNDestinatario.AddItem rs("nome")
     frmMensagens.cboEDestinatario.AddItem rs("nome")
     frmMensagens.cboRemetente.AddItem rs("nome")
     '
     rs.MoveNext
     '
  Loop
  '
  If rs.State = 1 Then rs.Close
  '
  connClose
  '
  Set rs = Nothing
  '
  Screen.MousePointer = 0
  '
End Sub

Public Function SalvaNovaMensagem() As Boolean
   '
   Dim strCodigoVendedor As String
   Dim rs
   '
   Screen.MousePointer = 11
   '
   SalvaNovaMensagem = False
   '
   Set rs = CreateObject("ADOCE.Recordset.3.0")
   '
   connOpen
   '
   rs.Open "SELECT * FROM destinatarios WHERE nome='" & Trim(frmMensagens.cboNDestinatario.List(frmMensagens.cboNDestinatario.ListIndex)) & "';", CONN, adOpenDynamic, adLockOptimistic
   '
   If rs.RecordCount > 0 Then
      '
      strCodigoVendedor = rs("codigo_destinatario")
      '
      ' MsgBox "Encontrou Destinatario=" & Trim(rs("nome")) & " = " & rs("codigo_destinatario"), vbOKOnly + vbInformation, App.Title
      '
   Else
      '
      MsgBox "Nao Encontrou Destinatario:(" & Trim(frmMensagens.cboNDestinatario.List(frmMensagens.cboNDestinatario.ListIndex)) & ")", vbOKOnly + vbInformation, App.Title
      '
   End If
   '
   strCodigoVendedor = "00" & rs("codigo_destinatario")
   '
   If rs.State = 1 Then rs.Close
   '
   rs.Open "SELECT * FROM mensagens WHERE codigo_vendedor_destino='" & Trim(strCodigoVendedor) & "' AND assunto ='" & frmMensagens.txtNAssunto.Text & "' AND tipo_mensagem='3';", CONN, adOpenDynamic, adLockOptimistic
   '
   If rs.RecordCount <= 0 Then rs.AddNew
   '
   ' Tipo de Mensagem: Nova
   '
   rs("tipo_mensagem") = "3"
   '
   rs("codigo_vendedor_origem") = Trim(usrCodigoVendedor)
   rs("codigo_vendedor_destino") = Trim(strCodigoVendedor)
   rs("assunto") = Trim(frmMensagens.txtNAssunto.Text)
   rs("data") = RetornaDataString(Now)
   rs("hora") = RetornaHoraString(CStr(Hour(Now)), CStr(Minute(Now)), CStr(Second(Now)))
   '
   rs("status") = "D"
   '
   rs("mensagem") = Trim(frmMensagens.txtNMensagem.Text)
   '
   rs.Update
   '
   If Err.Number <> 0 Then
      SalvaNovaMensagem = False
   Else
      SalvaNovaMensagem = True
   End If
   '
   If rs.State = 1 Then rs.Close
   '
   connClose
   '
   Set rs = Nothing
   '
   Screen.MousePointer = 0
   '
End Function

Public Sub EncheMensagem(ByVal strNomeDestinatario As String, ByVal valTipoMensagem As Integer, ByVal strComboNovaAssunto As String)
    '
    '1 - Recebida
    '
    '2 - Enviada
    '
    '3 - Nova
    '
    ' ExecSQL "CREATE TABLE mensagens(tipo_mensagem VARCHAR(1), codigo_vendedor_origem VARCHAR(5), codigo_vendedor_destino VARCHAR(5), assunto VARCHAR(20), data VARCHAR(8), hora VARCHAR(6), mensagem text);"
    '
    Dim strCodigoVendedor As String
    '
    Dim rs
    '
    Screen.MousePointer = 11
    '
    connOpen
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    '
    rs.Open "SELECT * FROM destinatarios WHERE nome='" & Trim(strNomeDestinatario) & "';", CONN, adOpenDynamic, adLockOptimistic
    '
    If rs.RecordCount > 0 Then
       '
       strCodigoVendedor = "00" & rs("codigo_destinatario")
       '
       ' MsgBox "Encontrou Destinatario:" & strCodigoVendedor & " (" & Trim(strNomeDestinatario) & ")", vbOKOnly + vbInformation, App.Title
       '
    Else
       '
       MsgBox "Nao Encontrou Destinatario:(" & Trim(strNomeDestinatario) & ")", vbOKOnly + vbInformation, App.Title
       '
       If rs.State = 1 Then rs.Close
       '
       connClose
       '
       Set rs = Nothing
       '
       Exit Sub
       '
    End If
    '
    If rs.State = 1 Then rs.Close
    '
    Select Case valTipoMensagem
    Case 1
      '
      frmMensagens.cboRMensagem.Text = vbNullString
      '
      If Len(Trim(strComboNovaAssunto)) > 0 Then
         '
         rs.Open "SELECT * FROM mensagens WHERE codigo_vendedor_origem='" & Trim(strCodigoVendedor) & "' AND assunto ='" & strComboNovaAssunto & "' AND tipo_mensagem='1';", CONN, adOpenDynamic, adLockOptimistic ' AND codigo_vendedor_destino='" & usrCodigoVendedor & "'
         '
         frmMensagens.cboRMensagem.Text = rs("mensagem")
         '
      Else
         '
         ' MsgBox "Procura mensagens:" & strCodigoVendedor & " (" & Trim(strNomeDestinatario) & ")", vbOKOnly + vbInformation, App.Title
         '
         rs.Open "SELECT * FROM mensagens WHERE codigo_vendedor_origem='" & Trim(strCodigoVendedor) & "' AND tipo_mensagem='1';", CONN, adOpenDynamic, adLockOptimistic '  AND codigo_vendedor_destino='" & usrCodigoVendedor & "'
         '
         frmMensagens.cboRAssunto.Clear
         '
         If rs.RecordCount > 0 Then
            '
            Do Until rs.EOF
               '
               ' MsgBox "Encontrou Remetente=" & strNomeDestinatario & " = " & rs("codigo_vendedor_origem"), vbOKOnly + vbInformation, App.Title
               '
               frmMensagens.cboRAssunto.AddItem rs("assunto")
               '
               rs.MoveNext
               '
            Loop
            '
         Else
            '
            MsgBox "Não foram encontradas mensagens de:(" & strNomeDestinatario & ")", vbOKOnly + vbInformation, App.Title
            '
         End If
         '
      End If
      '
    Case 2
      '
      frmMensagens.txtEMensagem.Text = vbNullString
      '
      If Len(strComboNovaAssunto) > 0 Then
         '
         rs.Open "SELECT * FROM mensagens WHERE codigo_vendedor_destino='" & Trim(strCodigoVendedor) & "' AND assunto ='" & strComboNovaAssunto & "' AND tipo_mensagem='2';", CONN, adOpenDynamic, adLockOptimistic
         '
          If rs.RecordCount = 1 Then
            '
            frmMensagens.txtEMensagem.Text = rs("mensagem")
            '
         End If
         '
      Else
         '
         rs.Open "SELECT * FROM mensagens WHERE codigo_vendedor_destino='" & Trim(strCodigoVendedor) & "' AND tipo_mensagem='2';", CONN, adOpenDynamic, adLockOptimistic
         '
         frmMensagens.cboEAssunto.Clear
         '
         If rs.RecordCount > 0 Then
            '
            Do Until rs.EOF
               '
               frmMensagens.cboEAssunto.AddItem rs("assunto")
               '
               rs.MoveNext
               '
            Loop
            '
         Else
            '
            MsgBox "Não foram encontradas mensagens para:(" & strNomeDestinatario & ")", vbOKOnly + vbInformation, App.Title
            '
         End If
         '
      End If
      '
    Case 3
      '
      frmMensagens.txtNAssunto.Text = vbNullString
      frmMensagens.txtNMensagem.Text = vbNullString
      '
      If Len(strComboNovaAssunto) > 0 Then
         '
         If Trim(strComboNovaAssunto) = "(Nova)" Then
            '
            frmMensagens.cboNAssunto.Visible = True ' False
            '
            frmMensagens.txtNAssunto.ZOrder vbBringToFront ' vbSendToBack
            frmMensagens.txtNAssunto.SetFocus
            '
         Else
            '
            rs.Open "SELECT * FROM mensagens WHERE codigo_vendedor_destino='" & Trim(strCodigoVendedor) & "' AND assunto ='" & strComboNovaAssunto & "' AND tipo_mensagem='3';", CONN, adOpenDynamic, adLockOptimistic
            '
            frmMensagens.txtNAssunto.Text = strComboNovaAssunto
            frmMensagens.txtNMensagem.Text = rs("mensagem")
            '
         End If
         '
      Else
         '
         rs.Open "SELECT * FROM mensagens WHERE codigo_vendedor_destino='" & Trim(strCodigoVendedor) & "' AND tipo_mensagem='3';", CONN, adOpenDynamic, adLockOptimistic
         '
         frmMensagens.cboNAssunto.Clear
         '
         frmMensagens.cboNAssunto.AddItem "(Nova)"
         '
         If rs.RecordCount > 0 Then
            '
            Do Until rs.EOF
               '
               frmMensagens.cboNAssunto.AddItem rs("assunto")
               '
               rs.MoveNext
               '
            Loop
            '
         Else
            '
            MsgBox "Não foram encontradas mensagens para:(" & strNomeDestinatario & ")", vbOKOnly + vbInformation, App.Title
            '
         End If
         '
      End If
      '
    End Select
    '
    If rs.State = 1 Then rs.Close
    '
    connClose
    '
    Set rs = Nothing
    '
Screen.MousePointer = 0
'
End Sub

Public Sub EncheSistematica(ByVal strDiaVisita As String, ByVal strDataSemana As String)
   '
   Dim rsSistematica, rsClientes, rsJustificativa, rsRoteiro
   Dim strNumeroJustificativa As String
   '
   Screen.MousePointer = 11
   '
   strDiaVisita = Left(Trim(strDiaVisita), 1)
   '
   frmSistematica.GridCtrl.Visible = False
   '
   frmSistematica.GridCtrl.Rows = 1
   frmSistematica.GridCtrl.Clear
   frmSistematica.GridCtrl.Row = 0
   frmSistematica.GridCtrl.Col = 0
   frmSistematica.GridCtrl.CellBackColor = &HC0C0C0
   frmSistematica.GridCtrl.CellFontBold = True
   frmSistematica.GridCtrl.Row = 0
   frmSistematica.GridCtrl.Col = 1
   frmSistematica.GridCtrl.CellBackColor = &HC0C0C0
   frmSistematica.GridCtrl.CellFontBold = True
   frmSistematica.GridCtrl.Row = 0
   frmSistematica.GridCtrl.Col = 2
   frmSistematica.GridCtrl.CellBackColor = &HC0C0C0
   frmSistematica.GridCtrl.CellFontBold = True
   frmSistematica.GridCtrl.ColWidth(0) = 2500
   frmSistematica.GridCtrl.ColWidth(1) = 1800
   frmSistematica.GridCtrl.ColWidth(2) = 2200
   frmSistematica.GridCtrl.TextMatrix(0, 0) = "Cliente"
   frmSistematica.GridCtrl.TextMatrix(0, 1) = "Telefone"
   frmSistematica.GridCtrl.TextMatrix(0, 2) = "Justificativa"
   '
   connOpen
   '
   Set rsSistematica = CreateObject("ADOCE.Recordset.3.0")
   Set rsClientes = CreateObject("ADOCE.Recordset.3.0")
   Set rsJustificativa = CreateObject("ADOCE.Recordset.3.0")
   Set rsRoteiro = CreateObject("ADOCE.Recordset.3.0")
   '
   ' Acha registros em SISTEMATICA DE VISITA que corresponda ao dia selecionado.
   '
   rsSistematica.Open "SELECT * FROM sistematica_visita " _
           & "WHERE(dia = '" & strDiaVisita & "')" _
           & "ORDER BY sistematica_visita.numero_visita;", CONN, adOpenDynamic, adLockReadOnly
           '
   '
   ' Se não encontrou nenhum registro.
   '
   If rsSistematica.RecordCount <= 0 Then
       '
       frmSistematica.GridCtrl.Row = 0
       frmSistematica.GridCtrl.Col = 0
       '
       frmSistematica.GridCtrl.Visible = True
       '
       Screen.MousePointer = 0
       '
       MsgBox "Não há visitas para este dia.", vbOKOnly + vbInformation, App.Title
       '
       rsSistematica.Close
       '
       connClose
       '
       Set rsSistematica = Nothing
       Set rsClientes = Nothing
       Set rsJustificativa = Nothing
       '
       Exit Sub
       '
   End If
   '
   ' Se encontrou vai varrer a tebela de SISTEMATICA
   '
   Do Until rsSistematica.EOF
      '
      ' Procura se existe o CLIENTE
      '
      If rsClientes.State = 1 Then rsClientes.Close
      '
      rsClientes.Open "SELECT * FROM clientes WHERE codigo_cliente='" & rsSistematica("codigo_cliente") & "';", CONN, adOpenForwardOnly, adLockPessimistic
      '
      If rsClientes.RecordCount <= 0 Then
         '
         ' Se não existe mostra como não cadastrado.
         '
         frmSistematica.GridCtrl.Rows = frmSistematica.GridCtrl.Rows + 1
         '
         frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Rows - 1, 0) = "Cliente não cadastrado"
         frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Rows - 1, 1) = "-"
         frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Rows - 1, 2) = "-"
         '
      Else
         '
         ' Se existe mostra CODIGO, NOME, TELEFONE E MOTIVO.
         '
         frmSistematica.GridCtrl.Rows = frmSistematica.GridCtrl.Rows + 1
         '
         frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Rows - 1, 0) = rsClientes("codigo_cliente") & " - " & rsClientes("nome_fantasia")
         frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Rows - 1, 1) = rsClientes("telefone")
         '
         ' Procura em ROTEIRO PERCORRIDO se existe registro correspondente para o cliente
         ' e mesma data base.
         '
         If rsRoteiro.State = 1 Then rsRoteiro.Close
         '
         rsRoteiro.Open "SELECT * FROM roteiro_percorrido WHERE codigo_cliente='" & _
         rsSistematica("codigo_cliente") & "' AND data_base='" & strDataSemana & "';", CONN, adOpenForwardOnly, adLockPessimistic
         '
         ' Se não existe mostra como cliente a ser visitado - Vermelho Claro.
         '
         If rsRoteiro.RecordCount <= 0 Then
            '
            frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Rows - 1, 2) = "-"
            '
            For I = 0 To 2
                '
                frmSistematica.GridCtrl.Row = frmSistematica.GridCtrl.Rows - 1
                frmSistematica.GridCtrl.Col = I
                '
                ' &HC0FFC0 - Verde Claro
                ' &HC0C0FF - Vermelho Claro
                '
                frmSistematica.GridCtrl.CellBackColor = VermelhoClaro ' VermelhoClaro = &HC0C0FF
                '
            Next
            '
         Else
            '
            ' Se existe mostra CLIENTE como VISITADO - Amarelo Claro.
            '
            strNumeroJustificativa = rsRoteiro("motivo_nao_visita")
            '
            ' Procura pela descrição do MOTIVO e mostra.
            '
            If rsJustificativa.State = 1 Then rsJustificativa.Close
            '
            rsJustificativa.Open "SELECT * FROM justificativa_nao_venda WHERE codigo='" & strNumeroJustificativa & "';", CONN, adOpenForwardOnly, adLockPessimistic
            '
            If rsJustificativa.RecordCount = 1 Then
               '
               frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Rows - 1, 2) = rsJustificativa("descricao")
               '
            Else
               '
               frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Rows - 1, 2) = "Não encontrada"
               '
            End If
            '
            For I = 0 To 2
                '
                frmSistematica.GridCtrl.Row = frmSistematica.GridCtrl.Rows - 1
                frmSistematica.GridCtrl.Col = I
                '
                ' &HC0FFC0 - Verde Claro
                '
                frmSistematica.GridCtrl.CellBackColor = AmareloClaro
                '
            Next
            '
            If rsJustificativa.State = 1 Then rsJustificativa.Close
            '
         End If
         '
         rsRoteiro.Close
         '
      End If
      '
      rsClientes.Close
      '
      rsSistematica.MoveNext
      '
   Loop
   '
   ' Procura em ROTEIRO PERCORRIDO se houve VISITA EXTRA
   '
   If rsRoteiro.State = 1 Then rsRoteiro.Close
   '
   rsRoteiro.Open "SELECT * FROM roteiro_percorrido " & _
   " WHERE visita_extra='S' AND dia_visita = '" & strDiaVisita & "';" _
   , CONN, adOpenForwardOnly, adLockPessimistic
   '
   If rsRoteiro.RecordCount > 0 Then
      '
      Do Until rsRoteiro.EOF
         '
         ' Procura se existe o CLIENTE
         '
         If rsClientes.State = 1 Then rsClientes.Close
         '
         rsClientes.Open "SELECT * FROM clientes WHERE codigo_cliente='" & _
         rsRoteiro("codigo_cliente") & "';", CONN, adOpenForwardOnly, adLockPessimistic
         '
         ' Se não existe mostra como cliente a ser visitado - Vermelho Claro.
         '
         If rsClientes.RecordCount <= 0 Then
            '
            ' Se não existe mostra como não cadastrado.
            '
            frmSistematica.GridCtrl.Rows = frmSistematica.GridCtrl.Rows + 1
            '
            frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Rows - 1, 0) = "Cliente não cadastrado"
            frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Rows - 1, 1) = "-"
            frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Rows - 1, 2) = "-"
            '
            For I = 0 To 2
                '
                frmSistematica.GridCtrl.Row = frmSistematica.GridCtrl.Rows - 1
                frmSistematica.GridCtrl.Col = I
                '
                ' &HC0FFC0 - Verde Claro
                '
                frmSistematica.GridCtrl.CellBackColor = VermelhoClaro
                '
            Next
            '
         Else
            '
            ' Se existe mostra CODIGO, NOME, TELEFONE E MOTIVO.
            '
            strNumeroJustificativa = rsRoteiro("motivo_nao_visita")
            '
            frmSistematica.GridCtrl.Rows = frmSistematica.GridCtrl.Rows + 1
            '
            frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Rows - 1, 0) = rsClientes("codigo_cliente") & " - " & rsClientes("nome_fantasia")
            frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Rows - 1, 1) = rsClientes("telefone")
            '
            ' Se existe mostra CLIENTE como VISITADO - Amarelo Claro.
            '
            ' Procura pela descrição do MOTIVO e mostra.
            '
            If rsJustificativa.State = 1 Then rsJustificativa.Close
            '
            rsJustificativa.Open "SELECT * FROM justificativa_nao_venda WHERE codigo='" & strNumeroJustificativa & "';", CONN, adOpenForwardOnly, adLockPessimistic
            '
            If rsJustificativa.RecordCount = 1 Then
               '
               frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Rows - 1, 2) = rsJustificativa("descricao")
               '
            Else
               '
               frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Rows - 1, 2) = "Não encontrada"
               '
            End If
            '
            For I = 0 To 2
                '
                frmSistematica.GridCtrl.Row = frmSistematica.GridCtrl.Rows - 1
                frmSistematica.GridCtrl.Col = I
                '
                ' &HC0FFC0 - Verde Claro
                '
                frmSistematica.GridCtrl.CellBackColor = VerdeClaro
                '
            Next
            '
            rsJustificativa.Close
            '
            rsClientes.Close
            '
         End If
         '
         rsRoteiro.MoveNext
         '
      Loop
      '
      rsRoteiro.Close
      '
   End If
   '
   rsSistematica.Close
   '
   connClose
   '
   frmSistematica.GridCtrl.Row = 0
   frmSistematica.GridCtrl.Col = 0
   '
   Set rsSistematica = Nothing
   Set rsClientes = Nothing
   Set rsJustificativa = Nothing
   '
   Screen.MousePointer = 0
   '
   frmSistematica.GridCtrl.Visible = True
   '
End Sub

Public Sub EncheComboRelatorios()
  '
  Dim rs
  '
  Screen.MousePointer = 11
  '
  connOpen
  '
  Set rs = CreateObject("ADOCE.Recordset.3.0")
  '
  rs.Open "SELECT * FROM relatorios;", CONN, adOpenForwardOnly, adLockReadOnly
  '
  frmRelatorios.cboDescricao.Clear
  '
  Do Until rs.EOF
     '
     frmRelatorios.cboDescricao.AddItem rs("descricao")
     '
     rs.MoveNext
     '
  Loop
  '
  If rs.State = 1 Then rs.Close
  '
  connClose
  '
  Set rs = Nothing
  '
  Screen.MousePointer = 0
  '
End Sub

Public Sub EncheFormularioRelatorios(ByVal strDescricao As String)
  '
  Dim rs
  Dim strRelatorioTodo As String
  Dim strRelatorioPArte As String
  '
  Screen.MousePointer = 11
  '
  connOpen
  '
  Set rs = CreateObject("ADOCE.Recordset.3.0")
  '
  rs.Open "SELECT * FROM relatorios WHERE descricao='" & Trim(strDescricao) & "';", CONN, adOpenForwardOnly, adLockReadOnly
  '
  frmRelatorios.txtData.Text = Trim(Mid(rs("data"), 1, 2)) & "/" & Trim(Mid(rs("data"), 3, 2)) & "/" & Trim(Mid(rs("data"), 5, 4))
  frmRelatorios.txtHora.Text = Trim(Mid(rs("hora"), 1, 2)) & ":" & Trim(Mid(rs("hora"), 3, 2)) & ":" & Trim(Mid(rs("hora"), 5, 2))
  '
  frmRelatorios.txtDetalhes.Text = rs("relatorio")
  '
  If rs.State = 1 Then rs.Close
  '
  connClose
  '
  Set rs = Nothing
  '
  Screen.MousePointer = 0
  '
End Sub

Function RetornaStringRelatorio(StringExtrato As String) As String
  '
  RetornaStringRelatorio = Right(StringExtrato, Len(StringExtrato) - cstLarguraRelatorio)
  '
End Function

Public Function RetornaCodigoCidade(strNomeCidade As String, valParamaetro As Integer) As String
  '
  ' 1 - Retorna o Código a partir da string da cidade
  ' 2 - Retorna a cidade a partir do código
  Screen.MousePointer = 11
  If Len(strNomeCidade) > 0 Then
     Dim rs
     connOpen
     Set rs = CreateObject("ADOCE.Recordset.3.0")
     If valParamaetro = 1 Then
        rs.Open "SELECT * FROM cidades WHERE nome='" & strNomeCidade & "';", CONN, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
           MsgBox "Não existem cidades com este Nome. Tente novamente.", vbOKOnly + vbCritical, App.Title
           If rs.State = 1 Then rs.Close
           connClose
           Set rs = Nothing
           Screen.MousePointer = 0
           Exit Function
        End If
        RetornaCodigoCidade = rs("codigo")
     Else
        rs.Open "SELECT * FROM cidades WHERE codigo='" & strNomeCidade & "';", CONN, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount <= 0 Then
           MsgBox "Não existem cidades com esse Código (" & strNomeCidade & "). Tente novamente.", vbOKOnly + vbCritical, App.Title
           If rs.State = 1 Then rs.Close
           connClose
           Set rs = Nothing
           Screen.MousePointer = 0
           Exit Function
        End If
        RetornaCodigoCidade = rs("nome")
     End If
     '
     If rs.State = 1 Then rs.Close
     '
     connClose
     Set rs = Nothing
  Else
     RetornaCodigoCidade = ""
  End If
     Screen.MousePointer = 0
End Function

Public Function RetornaAtividade(strNomeAtividade As String, valParamaetro As Integer) As String
   '
   ' 1 - Retorna o Código a partir da string da Atividade
   ' 2 - Retorna a Atividade a partir do código
   '
   Dim rs
   '
   Screen.MousePointer = 11
   '
   If Len(strNomeAtividade) > 0 Then
      '
       connOpen
       '
       Set rs = CreateObject("ADOCE.Recordset.3.0")
       '
       If valParamaetro = 1 Then
          '
          rs.Open "SELECT * FROM ramo_atividade WHERE descricao='" & strNomeAtividade & "';", CONN, adOpenForwardOnly, adLockReadOnly
          '
          If rs.RecordCount <= 0 Then
             '
             MsgBox "Não existe Ramos de Atividade com esta Descrição. Tente novamente.", vbOKOnly + vbCritical, App.Title
             '
             If rs.State = 1 Then rs.Close
             '
             connClose
             '
             Set rs = Nothing
             '
             Screen.MousePointer = 0
             '
             Exit Function
             '
          End If
          '
          RetornaAtividade = rs("codigo")
          '
       Else
          rs.Open "SELECT * FROM ramo_atividade WHERE codigo='" & strNomeAtividade & "';", CONN, adOpenForwardOnly, adLockReadOnly
          '
          If rs.RecordCount <= 0 Then
             '
             MsgBox "Não existe Ramo de Atividade com esse Código (" & strNomeAtividade & "). Tente novamente.", vbOKOnly + vbCritical, App.Title
             '
             If rs.State = 1 Then rs.Close
             '
             connClose
             '
             Set rs = Nothing
             '
             Screen.MousePointer = 0
             '
             Exit Function
             '
          End If
          '
          RetornaAtividade = rs("descricao")
          '
       End If
       '
       If rs.State = 1 Then rs.Close
       '
       connClose
       '
       Set rs = Nothing
       '
   Else
       '
       RetornaAtividade = ""
       '
   End If
   '
   Screen.MousePointer = 0
   '
End Function

Public Sub EncheComboCidade()
  '
  Screen.MousePointer = 11
  '
  Dim rs
  '
  Set rs = CreateObject("ADOCE.Recordset.3.0")
  '
  rs.Open "SELECT * FROM cidades ORDER BY nome;", CONN, adOpenForwardOnly, adLockReadOnly
  '
  frmClientes.cboCidadeCobranca.Clear
  frmClientes.cboCidadeEndereco.Clear
  '
  Do Until rs.EOF
     '
     frmClientes.cboCidadeCobranca.AddItem rs("nome")
     frmClientes.cboCidadeEndereco.AddItem rs("nome")
     '
     rs.MoveNext
     '
  Loop
  '
  If rs.State = 1 Then rs.Close
  '
  Set rs = Nothing
  '
  Screen.MousePointer = 0
  '
End Sub

Public Sub EncheComboAtividade()
    '
    Screen.MousePointer = 11
    '
    Dim rs
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    rs.Open "SELECT * FROM ramo_atividade ORDER BY descricao;", CONN, adOpenForwardOnly, adLockReadOnly
    '
    frmClientes.cboAtividade.Clear
    '
    Do Until rs.EOF
       frmClientes.cboAtividade.AddItem rs("descricao")
       rs.MoveNext
    Loop
    '
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    Screen.MousePointer = 0
    '
End Sub

Public Sub EncheFormularioRoteiroNE(ByVal strDataSerial As String)
  '
  Dim strNow As String
  Dim intAchouEspaco As Integer
  '
  'If Len(frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Row, 0)) <= 0 Then
  '   '
  '   MsgBox "Selecione uma visita para poder cadastrar o roteiro.", vbOKOnly + vbCritical, App.Title
  '   '
  '   Exit Sub
  '   '
  'End If
  '
  If frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Row, 0) = "Cliente" Then
     '
     Screen.MousePointer = 0
     '
     MsgBox "Selecione uma visita para poder cadastrar o roteiro.", vbOKOnly + vbCritical, App.Title
     '
     Exit Sub
     '
  End If
  '
  intAchouEspaco = InStr(1, frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Row, 0), " ", vbTextCompare)
  '
  intAchouEspaco = intAchouEspaco - 1
  '
  If intAchouEspaco > 7 Then
     '
     Screen.MousePointer = 0
     '
     MsgBox "Selecione uma visita para poder cadastrar o roteiro.", vbOKOnly + vbCritical, App.Title
     '
     Exit Sub
     '
  End If
  '
  ' Screen.MousePointer = 11 ' xixixi
  '
  frmRoteiro.Visible = True
  '
  strNow = RetornaDataString(Now)
  '
  frmRoteiro.cboCliente.Visible = False
  '
  frmRoteiro.txtVendedor.Text = usrCodigoVendedor
  '
  frmRoteiro.TxtCodigoCliente.Text = Mid(frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Row, 0), 1, 5)
  '
  frmRoteiro.txtCliente.Text = frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Row, 0)
  '
  frmRoteiro.txtData.Text = Trim(Mid(strDataSerial, 1, 2)) & "/" & Trim(Mid(strDataSerial, 3, 2)) & "/" & Trim(Mid(strDataSerial, 5, 4))
  frmRoteiro.txtDataPrevista.Text = Trim(Mid(strNow, 1, 2)) & "/" & Trim(Mid(strNow, 3, 2)) & "/" & Trim(Mid(strNow, 5, 4))
  '
  frmRoteiro.optNao.Value = True
  frmRoteiro.optSim.Value = False
  '
  Screen.MousePointer = 0
  '
End Sub

Public Sub EncheFormularioRoteiroEE(ByVal strDataSerial As String)
  '
  Dim strNow As String
  '
  Screen.MousePointer = 11
  '
  frmRoteiro.Visible = True
  '
  strNow = RetornaDataString(Now)
  '
  frmRoteiro.cboCliente.Visible = True
  '
  frmRoteiro.txtVendedor.Text = usrCodigoVendedor
  '
  connOpen
  '
  EncheComboClientes frmRoteiro.cboCliente, 2
  '
  connClose
  '
  frmRoteiro.txtData.Text = Trim(Mid(strDataSerial, 1, 2)) & "/" & Trim(Mid(strDataSerial, 3, 2)) & "/" & Trim(Mid(strDataSerial, 5, 4))
  frmRoteiro.txtDataPrevista.Text = Trim(Mid(strNow, 1, 2)) & "/" & Trim(Mid(strNow, 3, 2)) & "/" & Trim(Mid(strNow, 5, 4))
  '
  frmRoteiro.optNao.Value = False
  frmRoteiro.optSim.Value = True
  '
  Screen.MousePointer = 0
  '
End Sub

Public Function CadastraRoteiroPercorrido() As Boolean
  '
  Dim rs, rsJustificativa
  Dim mTipoJustificativa As String
  Dim mStatusJustificativa As String
  Dim strHoraVisita As String
  Dim mComando As String
  '
  CadastraRoteiroPercorrido = False
  '
  Screen.MousePointer = 11
  '
  If frmRoteiro.optNao.Value = True Then
     '
     mTipoJustificativa = "N"
     '
  Else
     '
     mTipoJustificativa = "S"
     '
  End If
  '
  If mRoteiroExtra <> 1 And mRoteiroExtra <> 2 Then
     '
     MsgBox "Houve erro na tentiva de Gravação.", vbOKOnly + vbCritical, App.Title
     '
     Screen.MousePointer = 0
     '
     Exit Function
     '
  End If
  '
  connOpen
  '
  Set rs = CreateObject("ADOCE.Recordset.3.0")
  '
  If mRoteiroExtra = 1 Then
     '
     ' strCodigoCliente = Left(frmSistematica.GridCtrl.TextMatrix(frmSistematica.GridCtrl.Row, 0), 1, 5)
     '
     mComando = "SELECT * FROM clientes WHERE codigo_cliente='" & frmRoteiro.TxtCodigoCliente.Text & "';"
     '
     ' MsgBox "Comando:(" & mComando & ")", vbOKOnly + vbCritical, App.Title
     '
     rs.Open mComando, CONN, adOpenForwardOnly, adLockReadOnly
     '
  End If
  '
  If mRoteiroExtra = 2 Then
     '
     strFantasia = Trim(frmRoteiro.cboCliente.Text)
     '
     mComando = "SELECT * FROM clientes WHERE nome_fantasia='" & Trim(strFantasia) & "';"
     '
     ' MsgBox "Comando:(" & mComando & ")", vbOKOnly + vbCritical, App.Title
     '
     rs.Open mComando, CONN, adOpenForwardOnly, adLockReadOnly
     '
  End If
  '
  If rs.RecordCount <= 0 Then
     '
     MsgBox "Selecione um Cliente.", vbOKOnly + vbCritical, App.Title
     '
     If rs.State = 1 Then rs.Close
     '
     Set rs = Nothing
     '
     connClose
     '
     Screen.MousePointer = 0
     '
     Exit Function
     '
  End If
  '
  Set rsJustificativa = CreateObject("ADOCE.Recordset.3.0")
  '
  rsJustificativa.Open "SELECT * FROM justificativa_nao_venda WHERE codigo='" & Left(frmRoteiro.cboMotivo.Text, 3) & "';", CONN, adOpenForwardOnly, adLockReadOnly
  '
  If rsJustificativa.RecordCount <= 0 Then
     '
     MsgBox "Selecione um Motivo.", vbOKOnly + vbCritical, App.Title
     '
     If rsJustificativa.State = 1 Then rs.Close
     '
     Set rsJustificativa = Nothing
     '
     connClose
     '
     Screen.MousePointer = 0
     '
     Exit Function
     '
  End If
  '
  If rsJustificativa.RecordCount > 0 Then
     '
     mStatusJustificativa = rsJustificativa("tipo")
     '
  Else
     '
     mStatusJustificativa = "*"
     '
  End If
  '
  If rsJustificativa.State = 1 Then rsJustificativa.Close
  '
  Set rsJustificativa = Nothing
  '
  Set rs = CreateObject("ADOCE.Recordset.3.0")
  '
  'rs.Open "SELECT * FROM roteiro_percorrido WHERE codigo_cliente='" & frmRoteiro.txtCliente.Text & _
  '"' AND data_base='" & Trim(Mid(frmRoteiro.txtData.Text, 1, 2)) & Trim(Mid(frmRoteiro.txtData.Text, 4, 2)) & Trim(Mid(frmRoteiro.txtData.Text, 7, 4)) & _
  '"' AND hora_visita='" & Trim(Mid(frmRoteiro.txtHoraVisita.Text, 1, 2)) & Trim(Mid(frmRoteiro.txtHoraVisita.Text, 4, 2)) & Trim(Mid(frmRoteiro.txtHoraVisita.Text, 7, 2)) & _
  '"';", CONN, adOpenDynamic, adLockOptimistic
  '
  rs.Open "SELECT * FROM roteiro_percorrido WHERE codigo_cliente='" & frmRoteiro.TxtCodigoCliente.Text & _
  "' AND data_base='" & Trim(Mid(frmRoteiro.txtData.Text, 1, 2)) & Trim(Mid(frmRoteiro.txtData.Text, 4, 2)) & Trim(Mid(frmRoteiro.txtData.Text, 7, 4)) & _
  "';", CONN, adOpenDynamic, adLockOptimistic
  '
  If rs.RecordCount <= 0 Then
     '
     rs.AddNew
     '
  End If
  '
  If Len(Trim(frmRoteiro.txtHoraVisita.Text)) = 8 Then
     '
     strHoraVisita = Trim(Mid(frmRoteiro.txtHoraVisita.Text, 1, 2)) & Trim(Mid(frmRoteiro.txtHoraVisita.Text, 4, 2)) & Trim(Mid(frmRoteiro.txtHoraVisita.Text, 7, 2))
     '
  Else
     '
     strHoraVisita = RetornaHoraString(CStr(Hour(Now)), CStr(Minute(Now)), CStr(Second(Now)))
     '
  End If
  '
  rs("codigo_vendedor_atual") = usrCodigoVendedor
  '
  rs("data_base") = Trim(Mid(frmRoteiro.txtData.Text, 1, 2)) & Trim(Mid(frmRoteiro.txtData.Text, 4, 2)) & Trim(Mid(frmRoteiro.txtData.Text, 7, 4))
  '
  rs("codigo_cliente") = frmRoteiro.TxtCodigoCliente.Text
  '
  rs("data_visita") = Trim(Mid(frmRoteiro.txtDataPrevista.Text, 1, 2)) & Trim(Mid(frmRoteiro.txtDataPrevista.Text, 4, 2)) & Trim(Mid(frmRoteiro.txtDataPrevista.Text, 7, 4))
  '
  rs("hora_visita") = strHoraVisita
  '
  ' MsgBox "Tentou gravar Roteiro.", vbOKOnly + vbCritical, App.Title
  '
  rs("motivo_nao_visita") = Left(frmRoteiro.cboMotivo.Text, 3)
  '
  rs("visita_extra") = mTipoJustificativa
  '
  rs("automatico") = "N"
  '
  rs("dia_visita") = Mid(frmSistematica.cboDiadaSemana.Text, 1, 1)
  '
  rs("status") = mStatusJustificativa
  '
  rs.Update
  '
  If rs.State = 1 Then rs.Close
  '
  Set rs = Nothing
  '
  connClose
  '
  Screen.MousePointer = 0
  '
  CadastraRoteiroPercorrido = True
  '
End Function

Public Function AlteraDiasVendedor(ByVal strDias As String) As Boolean
    '
    Dim strUserName As String
    '
    AlteraDiasVendedor = False
    Screen.MousePointer = 11
    strDias = Right(strDias, 3)
    connOpen
    Dim rs
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    '
    rs.Open "SELECT * FROM vendedor WHERE codigo_vendedor='" & Trim(frmValidarVendedor.txtCodigoVendedorAtual.Text) & "';", CONN, adOpenDynamic, adLockOptimistic
    '
    If rs.RecordCount > 0 Then
        rs("tempo_maximo") = strDias
        rs.Update
        AlteraDiasVendedor = True
    End If
    '
    If rs.State = 1 Then rs.Close
    '
    Set rs = Nothing
    frmLogin.File.Open strPath & "\last.txt", fsModeOutput, fsAccessWrite
    frmLogin.File.LinePrint RetornaDataString(Now)
    frmLogin.File.Close
    connClose
    Screen.MousePointer = 0
    '
End Function
'
