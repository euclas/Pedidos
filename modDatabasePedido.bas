Attribute VB_Name = "modDatabasePedido"
Option Explicit
Public Verificador As String
Public strNomeFantasia As String
'
'
'
Public Sub EncheCombosPedidos()
  '
  Dim rs
  '
  Dim strLine As String
  Dim nbrecords As Integer
  '
  Screen.MousePointer = 11
  '
  Set rs = CreateObject("ADOCE.Recordset.3.0")
  '
  rs.Open "SELECT * from forma_pagamento ORDER BY descricao;", CONN, adOpenForwardOnly, adLockReadOnly
  If rs.RecordCount > 0 Then
      frmPedido.cboFPagto.Clear
      Do Until rs.EOF
          frmPedido.cboFPagto.AddItem rs("descricao")
          rs.MoveNext
      Loop
  End If
  If rs.State = 1 Then rs.Close
  '
  rs.Open "SELECT * from condicao_pagamento ORDER BY descricao;", CONN, adOpenForwardOnly, adLockReadOnly
  If rs.RecordCount > 0 Then
      frmPedido.cboCPagto.Clear
      Do Until rs.EOF
          frmPedido.cboCPagto.AddItem rs("descricao")
          rs.MoveNext
      Loop
  End If
  If rs.State = 1 Then rs.Close
  '
  rs.Open "SELECT * from tipo_movimento ORDER BY descricao;", CONN, adOpenForwardOnly, adLockReadOnly
  If rs.RecordCount > 0 Then
      frmPedido.cboTmov.Clear
      Do Until rs.EOF
          frmPedido.cboTmov.AddItem rs("descricao")
          rs.MoveNext
      Loop
  End If
  If rs.State = 1 Then rs.Close
  '
  rs.Open "SELECT * from produtos ORDER BY descricao;", CONN, adOpenForwardOnly, adLockReadOnly
  '
  If rs.RecordCount > 0 Then
     frmPedido.cboProdutos.Clear
     Do Until rs.EOF
        frmPedido.cboProdutos.AddItem rs("descricao")
        rs.MoveNext
     Loop
  End If
  If rs.State = 1 Then rs.Close
  '
  Set rs = Nothing
  '
  Screen.MousePointer = 0
  '
End Sub

Public Sub EncheCombosEstoque()
  '
  Dim rs
  '
  Dim strLine As String
  Dim nbrecords As Integer
  '
  Screen.MousePointer = 11
  '
  Set rs = CreateObject("ADOCE.Recordset.3.0")
  '
  rs.Open "SELECT * from forma_pagamento ORDER BY descricao;", CONN, adOpenForwardOnly, adLockReadOnly
  If rs.RecordCount > 0 Then
      frmEstoque.cboFPagto.Clear
      Do Until rs.EOF
          frmEstoque.cboFPagto.AddItem rs("descricao")
          rs.MoveNext
      Loop
  End If
  If rs.State = 1 Then rs.Close
  '
  rs.Open "SELECT * from condicao_pagamento ORDER BY descricao;", CONN, adOpenForwardOnly, adLockReadOnly
  If rs.RecordCount > 0 Then
      frmEstoque.cboCPagto.Clear
      Do Until rs.EOF
          frmEstoque.cboCPagto.AddItem rs("descricao")
          rs.MoveNext
      Loop
  End If
  If rs.State = 1 Then rs.Close
  '
  rs.Open "SELECT * from tipo_movimento ORDER BY descricao;", CONN, adOpenForwardOnly, adLockReadOnly
  If rs.RecordCount > 0 Then
      frmEstoque.cboTMovto.Clear
      Do Until rs.EOF
          frmEstoque.cboTMovto.AddItem rs("descricao")
          rs.MoveNext
      Loop
  End If
  If rs.State = 1 Then rs.Close
  '
'  rs.Open "SELECT * from produtos ORDER BY descricao;", CONN, adOpenForwardOnly, adLockReadOnly
'  '
'  If rs.RecordCount > 0 Then
'     frmEstoque.cboProdutos.Clear
'     Do Until rs.EOF
'        frmEstoque.cboProdutos.AddItem rs("descricao")
'        rs.MoveNext
'     Loop
'  End If
'  If rs.State = 1 Then rs.Close
  '
  Set rs = Nothing
  '
  Screen.MousePointer = 0
  '
End Sub

Public Sub EncheHistoricoPedidos(ByVal strNomeFantasia As String)
  '
  Screen.MousePointer = 11
  '
  Dim rs, rsItem
  Dim strCodigo As String
  Dim valPedido As Double
  Dim valDesconto As Double
  Dim strFormataNumero As String
  '
  frmHistorico.GridCtrl.Rows = 1
  frmHistorico.GridCtrl.Clear
  frmHistorico.GridCtrl.Row = 0
  frmHistorico.GridCtrl.Col = 0
  frmHistorico.GridCtrl.CellBackColor = &HC0C0C0
  frmHistorico.GridCtrl.CellFontBold = True
  frmHistorico.GridCtrl.Row = 0
  frmHistorico.GridCtrl.Col = 1
  frmHistorico.GridCtrl.CellBackColor = &HC0C0C0
  frmHistorico.GridCtrl.CellFontBold = True
  frmHistorico.GridCtrl.Row = 0
  frmHistorico.GridCtrl.Col = 2
  frmHistorico.GridCtrl.CellBackColor = &HC0C0C0
  frmHistorico.GridCtrl.CellFontBold = True
  frmHistorico.GridCtrl.Row = 0
  frmHistorico.GridCtrl.Col = 3
  frmHistorico.GridCtrl.CellBackColor = &HC0C0C0
  frmHistorico.GridCtrl.CellFontBold = True
  '
  frmHistorico.GridCtrl.ColWidth(0) = 800 ' 1300
  frmHistorico.GridCtrl.ColWidth(1) = 900 ' 1800
  frmHistorico.GridCtrl.ColWidth(2) = 1000 ' 1800
  frmHistorico.GridCtrl.ColWidth(3) = 1500 ' 1800
  '
  frmHistorico.GridCtrl.TextMatrix(0, 0) = "Pedido"
  frmHistorico.GridCtrl.TextMatrix(0, 1) = "Data"
  frmHistorico.GridCtrl.TextMatrix(0, 2) = "Vl. Líquido"
  frmHistorico.GridCtrl.TextMatrix(0, 3) = "Status"
  '
  frmHistorico.GridCtrl.Col = 0
  '
  connOpen
  '
  Set rs = CreateObject("ADOCE.Recordset.3.0")
  '
  Set rsItem = CreateObject("ADOCE.Recordset.3.0")
  '
  rs.Open "SELECT * FROM clientes WHERE nome_fantasia='" & Trim(strNomeFantasia) & "';", CONN, adOpenDynamic, adLockReadOnly
  '
  If rs.RecordCount <= 0 Then
     '
     Screen.MousePointer = 0
     '
     MsgBox "Não foi possível consultar esse cliente neste momento.", vbOKOnly + vbInformation, App.Title
     '
     Exit Sub
     '
  End If
  '
  strCodigo = rs("codigo_cliente")
  '
  If rs.State = 1 Then rs.Close
  '
  ' codigo_vendedor VARCHAR(5),
  ' codigo_cliente VARCHAR(5),
  ' numero_pedido_interno VARCHAR(6),
  ' numero_pedido_externo VARCHAR(6),
  ' pedido_cliente VARCHAR(10),
  ' data_emissao VARCHAR(8),
  ' hora_emissao VARCHAR(6),
  ' data_entrega VARCHAR(8),
  ' acrescimo_valor VARCHAR(10),
  ' desconto_valor VARCHAR(10),
  ' forma_pgto VARCHAR(2),
  ' condicao_pgto VARCHAR(2),
  ' tipo_movimento VARCHAR(1),
  ' status VARCHAR(1),
  ' observacao TEXT);"
  '
  rs.Open "SELECT * FROM pedido WHERE codigo_cliente='" & Trim(strCodigo) & "' ORDER BY numero_pedido_interno DESC;", CONN, adOpenDynamic, adLockReadOnly
  '
  If rs.RecordCount <= 0 Then
     '
     Screen.MousePointer = 0
     '
     MsgBox "Não há histórico de pedidos para este cliente.", vbOKOnly + vbInformation, App.Title
     '
     Exit Sub
     '
  End If
  '
  Do Until rs.EOF
     '
     frmHistorico.GridCtrl.Rows = frmHistorico.GridCtrl.Rows + 1
     '
     If IsNumeric(rs("numero_pedido_externo")) Then
        '
        frmHistorico.GridCtrl.TextMatrix(frmHistorico.GridCtrl.Rows - 1, 0) = rs("numero_pedido_externo")
        '
     Else
        '
        frmHistorico.GridCtrl.TextMatrix(frmHistorico.GridCtrl.Rows - 1, 0) = rs("numero_pedido_interno")
        '
     End If
     '
     frmHistorico.GridCtrl.TextMatrix(frmHistorico.GridCtrl.Rows - 1, 1) = Mid(rs("data_emissao"), 1, 2) & "/" & Mid(rs("data_emissao"), 3, 2) & "/" & Mid(rs("data_emissao"), 5, 4)
     '
     valPedido = 0
     '
     If IsNumeric(rs("acrescimo_valor")) Then
        '
        valPedido = valPedido + CDbl(rs("acrescimo_valor"))
        '
     End If
     
     If IsNumeric(rs("desconto_valor")) Then
        '
        valPedido = valPedido - CDbl(rs("desconto_valor"))
        '
     End If
     '
     rsItem.Open "SELECT * FROM itens_pedido WHERE numero_pedido='" & Trim(frmHistorico.GridCtrl.TextMatrix(frmHistorico.GridCtrl.Rows - 1, 0)) & "';", CONN, adOpenForwardOnly, adLockReadOnly
     '
     If rsItem.RecordCount > 0 Then
        '
        Do Until rsItem.EOF
           '
           If IsNumeric(rsItem("desconto")) And Len(Trim(rsItem("desconto"))) > 0 Then
              '
              If Left(rsItem("desconto"), 1) = "-" Then
                 '
                 valDesconto = ((CDbl(rsItem("desconto")) - 200) / 100) * -1
                 '
              Else
                 '
                 valDesconto = CDbl(rsItem("desconto")) / 100
                 '
              End If
              '
           Else
              '
              valDesconto = 1
              '
           End If
           '
           ' MsgBox "ValDesconto=" & CStr(valDesconto), vbOKOnly + vbInformation, App.Title
           '
           If IsNumeric(rsItem("qtd_faturada")) And Len(Trim(rsItem("qtd_faturada"))) > 0 Then
              '
              If Trim(rsItem("qtd_faturada")) <> "-" Then
                 '
                 valPedido = valPedido + (CDbl(rsItem("qtd_faturada")) * CDbl(rsItem("valor_unitario"))) - ((CDbl(rsItem("qtd_faturada")) * CDbl(rsItem("valor_unitario"))) * valDesconto)
                 '
              End If
              '
           Else
              '
              If IsNumeric(rsItem("qtd_pedida")) And Len(Trim(rsItem("qtd_pedida"))) > 0 Then
                 '
                 valPedido = valPedido + (CDbl(rsItem("qtd_pedida")) * CDbl(rsItem("valor_unitario"))) - ((CDbl(rsItem("qtd_pedida")) * CDbl(rsItem("valor_unitario"))) * valDesconto)
                 '
              End If
              '
           End If
           '
           rsItem.MoveNext
           '
        Loop
        '
        strFormataNumero = FormatCurrency(CStr(valPedido), 2, vbTrue, vbTrue, vbTrue)
        strFormataNumero = Mid(strFormataNumero, 3, Len(strFormataNumero) - 2)
        '
        ' strFormataNumero = FormatCurrency(CDbl(txtQtdPedida.Text) * (CDbl(txtPrecoUnit.Text) * (1 - (CDbl(txtDescItem.Text) / 100))), 2, vbTrue, vbTrue, vbTrue)
        '
        '=========================================================================
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
        '===========================================================================================
        '
        frmHistorico.GridCtrl.TextMatrix(frmHistorico.GridCtrl.Rows - 1, 2) = strFormataNumero
        '
        frmHistorico.GridCtrl.Col = 2
        frmHistorico.GridCtrl.CellAlignment = flexAlignRightCenter
        frmHistorico.GridCtrl.Col = 0
        '
     Else
         '
         frmHistorico.GridCtrl.TextMatrix(frmHistorico.GridCtrl.Rows - 1, 2) = FormatCurrency("0", 2, vbTrue, vbTrue, vbTrue)
         '
     End If
     '
     rsItem.Close
     '
     Select Case LCase(rs("status"))
     Case "d"
         frmHistorico.GridCtrl.TextMatrix(frmHistorico.GridCtrl.Rows - 1, 3) = "Digitado"
     Case "t"
         frmHistorico.GridCtrl.TextMatrix(frmHistorico.GridCtrl.Rows - 1, 3) = "Transmitido"
     Case "f"
         frmHistorico.GridCtrl.TextMatrix(frmHistorico.GridCtrl.Rows - 1, 3) = "Faturado"
     Case "c"
         frmHistorico.GridCtrl.TextMatrix(frmHistorico.GridCtrl.Rows - 1, 3) = "Cancelado"
     End Select
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
  Set rsItem = Nothing
  '
  Screen.MousePointer = 0
  '
  frmHistorico.GridCtrl.Col = 0
  '
End Sub

Public Sub EnchePedidoVelho(ByVal strNumeroPedido As String)
    '
    Screen.MousePointer = 11
    '
    Dim rs, rsaux
    Dim valPedido As Double
    Dim valDesconto As Double
    Dim dblTotal As Double
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    Set rsaux = CreateObject("ADOCE.Recordset.3.0")
    '
    LimpaFormularioPEdido
    '
    connOpen
    '
    rs.Open "SELECT * FROM pedido WHERE numero_pedido_interno='" & Trim(strNumeroPedido) & "';", CONN, adOpenForwardOnly, adLockReadOnly
    '
    If rs.RecordCount <= 0 Then
       '
       If rs.State = 1 Then rs.Close
       '
       rs.Open "SELECT * FROM pedido WHERE numero_pedido_externo='" & Trim(strNumeroPedido) & "';", CONN, adOpenForwardOnly, adLockReadOnly
       '
       If rs.RecordCount <= 0 Then
          '
          MsgBox "Código de pedido inválido na base de dados.", vbOKOnly + vbCritical, App.Title
          '
          Exit Sub
          '
       End If
       '
    End If
    '
    frmPedido.txtObs.Text = rs("observacao")
    '
    Select Case Len(Trim(strNumeroPedido))
        Case 1
            frmPedido.txtNumeroPedido.Text = "00000" & Trim(strNumeroPedido)
        Case 2
            frmPedido.txtNumeroPedido.Text = "0000" & Trim(strNumeroPedido)
        Case 3
            frmPedido.txtNumeroPedido.Text = "000" & Trim(strNumeroPedido)
        Case 4
            frmPedido.txtNumeroPedido.Text = "00" & Trim(strNumeroPedido)
        Case 5
            frmPedido.txtNumeroPedido.Text = "0" & Trim(strNumeroPedido)
        Case 6
            frmPedido.txtNumeroPedido.Text = Trim(strNumeroPedido)
    End Select
    '
    ' frmPedido.txtNumeroPedido.Text = frmPedido.txtNumeroPedido.Text
    '
    frmPedido.txtEntrega.Text = "-"
    frmPedido.txtAcrescimo.Text = rs("acrescimo_valor")
    frmPedido.txtDesconto.Text = rs("desconto_valor")
    '
    ' frmPedido.txtEmissao.Text = Mid(rs("data_emissao"), 1, 2) & "/" & Mid(rs("data_emissao"), 3, 2) & "/" & Mid(rs("data_emissao"), 5, 4) Euclides
    '
    frmPedido.txtEntrega.Text = Mid(rs("data_entrega"), 1, 2) & "/" & Mid(rs("data_entrega"), 3, 2) & "/" & Mid(rs("data_entrega"), 5, 4)
    '
    'Nome do Cliente ----->
    '
    rsaux.Open "SELECT * FROM clientes WHERE codigo_cliente='" & rs("codigo_cliente") & "';", CONN, adOpenForwardOnly, adLockReadOnly
    '
    If rsaux.RecordCount > 0 Then frmPedido.txtCliente.Text = rsaux("nome_fantasia")
    '
    rsaux.Close
    '
    'Forma de Pagamento ----->
    '
    rsaux.Open "SELECT * FROM forma_pagamento WHERE codigo='" & rs("forma_pgto") & "';", CONN, adOpenForwardOnly, adLockReadOnly
    '
    If rsaux.RecordCount > 0 Then frmPedido.txtFPgto.Text = rsaux("descricao")
    '
    rsaux.Close
    '
    'Condição de Pagamento ----->
    '
    rsaux.Open "SELECT * FROM condicao_pagamento WHERE codigo_condicao='" & rs("condicao_pgto") & "';", CONN, adOpenForwardOnly, adLockReadOnly
    '
    If rsaux.RecordCount > 0 Then frmPedido.txtCPagto.Text = rsaux("descricao")
    '
    rsaux.Close
    '
    'Tipo de movimento ----->
    rsaux.Open "SELECT * FROM tipo_movimento WHERE codigo='" & rs("tipo_movimento") & "';", CONN, adOpenForwardOnly, adLockReadOnly
    '
    If rsaux.RecordCount > 0 Then frmPedido.txtTMvto.Text = rsaux("descricao")
    '
    rsaux.Close
    '--------->
    If rs.State = 1 Then rs.Close
    '
    rs.Open "SELECT * FROM itens_pedido WHERE numero_pedido='" & frmPedido.txtNumeroPedido.Text & "';", CONN, adOpenForwardOnly, adLockReadOnly
    '
    If rs.RecordCount > 0 Then
       '
        Do Until rs.EOF
            '
            frmPedido.GridCtrl.Rows = frmPedido.GridCtrl.Rows + 1
            '
            rsaux.Open "SELECT * FROM produtos WHERE codigo_produto='" & Trim(rs("codigo_produto")) & "';", CONN, adOpenForwardOnly, adLockReadOnly
            '
            If rsaux.RecordCount > 0 Then frmPedido.GridCtrl.TextMatrix(frmPedido.GridCtrl.Rows - 1, 0) = rsaux("descricao")
            '
            rsaux.Close
            '
            frmPedido.GridCtrl.TextMatrix(frmPedido.GridCtrl.Rows - 1, 1) = rs("qtd_pedida")
            frmPedido.GridCtrl.TextMatrix(frmPedido.GridCtrl.Rows - 1, 2) = rs("qtd_faturada")
            frmPedido.GridCtrl.TextMatrix(frmPedido.GridCtrl.Rows - 1, 3) = rs("valor_unitario")
            frmPedido.GridCtrl.TextMatrix(frmPedido.GridCtrl.Rows - 1, 4) = rs("desconto")
            '
            If IsNumeric(frmPedido.GridCtrl.TextMatrix(frmPedido.GridCtrl.Rows - 1, 2)) And _
               Len(Trim(frmPedido.GridCtrl.TextMatrix(frmPedido.GridCtrl.Rows - 1, 2))) > 0 Then
               '
               frmPedido.GridCtrl.TextMatrix(frmPedido.GridCtrl.Rows - 1, 5) = CStr(CDbl(frmPedido.GridCtrl.TextMatrix(frmPedido.GridCtrl.Rows - 1, 2)) * CDbl(frmPedido.GridCtrl.TextMatrix(frmPedido.GridCtrl.Rows - 1, 3)))
               '
            Else
               '
               frmPedido.GridCtrl.TextMatrix(frmPedido.GridCtrl.Rows - 1, 5) = CStr(CDbl(frmPedido.GridCtrl.TextMatrix(frmPedido.GridCtrl.Rows - 1, 1)) * CDbl(frmPedido.GridCtrl.TextMatrix(frmPedido.GridCtrl.Rows - 1, 3)))
               '
            End If
            '
            rs.MoveNext
            '
        Loop
        '
    End If
    '
    frmPedido.GridCtrl.Col = 0
    frmPedido.GridCtrl.Row = 0
    '
    If rs.State = 1 Then rs.Close
    '
    connClose
    '
    Call RecalculaGrid
    '
    ' kride
    '
    'dblTotal = 0
    'For I = 1 To (frmPedido.GridCtrl.Rows - 1)
    '    dblTotal = dblTotal + CDbl(frmPedido.GridCtrl.TextMatrix(I, 5))
    'Next
    ''
    'frmPedido.Text2.Text = FormatCurrency(CStr(dblTotal), 2, vbTrue, vbTrue, vbTrue)
    ''
    'If Len(frmPedido.txtAcrescimo.Text) > 0 And IsNumeric(frmPedido.txtAcrescimo.Text) Then dblTotal = (dblTotal + CDbl(frmPedido.txtAcrescimo.Text))
    'If Len(frmPedido.txtDesconto.Text) > 0 And IsNumeric(frmPedido.txtDesconto.Text) Then dblTotal = (dblTotal - CDbl(frmPedido.txtDesconto.Text))
    ''
    'frmPedido.txtLiquido.Text = FormatCurrency(CStr(dblTotal), 2, vbTrue, vbTrue, vbTrue)
    ''
    Set rs = Nothing
    ''
    Set rsaux = Nothing
    '
    frmPedido.Show
    '
    Screen.MousePointer = 0
    '
End Sub

Public Sub RecalculaGrid()
    Dim rs
    Dim strCodigoProduto As String
    Dim strCodigoCliente As String
    Dim strPrecoProduto As String
    Dim intTXTaPreencher As Integer
    Dim strCodigoTabela As String
    Dim strSubBrand As String
    Dim dblTotal As Double
    Dim mCondPagto As String
    Dim strFormataNumero As String
    '
    Dim valPreco1 As String
    Dim valPreco2 As String
    Dim valPreco3 As String
    Dim valPreco4 As String
    Dim valPreco5 As String
    '
    If Trim(frmPedido.cboPedidoCliente.Text) <> "" Then
       '
       strNomeFantasia = frmPedido.cboPedidoCliente.Text
       '
    Else
       '
       strNomeFantasia = frmPedido.txtCliente.Text
       '
    End If
    '
    connOpen
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    '
    rs.Open "SELECT * FROM clientes WHERE nome_fantasia='" & Trim(strNomeFantasia) & "';", CONN, adOpenForwardOnly, adLockReadOnly
    '
    If rs.RecordCount > 0 Then
       '
       strCodigoCliente = rs("codigo_cliente")
       '
    Else
       '
       MsgBox "Cliente inexistente:(" & strNomeFantasia & ")", vbOKOnly + vbCritical, App.Title
       '
       If rs.State = 1 Then rs.Close
       '
       Screen.MousePointer = 0
       '
       Exit Sub
       '
    End If
    '
    If rs.State = 1 Then rs.Close
    '
    For I = 1 To (frmPedido.GridCtrl.Rows - 1)
        '
        rs.Open "SELECT * FROM produtos WHERE descricao='" & Trim(frmPedido.GridCtrl.TextMatrix(I, 0)) & "';", CONN, adOpenForwardOnly, adLockReadOnly
        '
        If rs.RecordCount > 0 Then
           '
            strCodigoProduto = rs("codigo_produto")
            strSubBrand = rs("sub_brand")
            '
            valPreco1 = rs("preco1")
            valPreco2 = rs("preco2")
            valPreco3 = rs("preco3")
            valPreco4 = rs("preco4")
            valPreco5 = rs("preco5")
            '
            If rs.State = 1 Then rs.Close
            '
            'Condição de Pagamento----->
            '
            mCondPagto = ""
            '
            If Trim(frmPedido.cboCPagto.Text) <> "" Then
               '
               mCondPagto = frmPedido.cboCPagto.Text
               '
            Else
               '
               mCondPagto = frmPedido.txtCPagto.Text
               '
            End If
            '
            rs.Open "SELECT * FROM condicao_pagamento WHERE descricao='" & Trim(mCondPagto) & "';", CONN, adOpenForwardOnly, adLockReadOnly
            '
            If rs.RecordCount > 0 Then
               '
               strCodigoTabela = rs("codigo_tabela_preco")
               '
            Else
               '
               MsgBox "Selecione uma Condição de Pagamento (1).", vbOKOnly + vbCritical, App.Title
               '
               If rs.State = 1 Then rs.Close
               '
               Screen.MousePointer = 0
               '
               Exit Sub
               '
            End If
            '
            If rs.State = 1 Then rs.Close
            '
            'Tabela de Preços----------> euclides 24/09/2002
            '
''            ' rs.Open "SELECT * FROM tabela_precos WHERE produto='" & strCodigoProduto & "' AND tabela_precos='" & strCodigoTabela & "';", CONN, adOpenForwardOnly, adLockReadOnly
''            '
''            rs.Open "SELECT * FROM tabela_precos WHERE tabela_precos='" & strCodigoTabela & "' AND produto='" & strCodigoProduto & "';", CONN, adOpenForwardOnly, adLockReadOnly
''            '
''            If rs.RecordCount > 0 Then
''               '
''               strPrecoProduto = rs("preco")
''               '
''            Else
''               '
''               MsgBox "Não há preço cadastrado para este produto nesta condição de pagamento.", vbOKOnly + vbCritical, App.Title
''               '
''               If rs.State = 1 Then rs.Close
''               '
''               Screen.MousePointer = 0
''               '
''               Exit Sub
''               '
''            End If
''            '
''            If rs.State = 1 Then rs.Close
              '
              '
              ' strCodigoTabela
              '
              Select Case CInt(strCodigoTabela)
              Case 1
                   strPrecoProduto = valPreco1
              Case 2
                   strPrecoProduto = valPreco2
              Case 3
                   strPrecoProduto = valPreco3
              Case 4
                   strPrecoProduto = valPreco4
              Case 5
                   strPrecoProduto = valPreco5
              Case Else
                   strPrecoProduto = valPreco5
              End Select
              '
              '
              '
            'Desconto do Canal--------->
            '
            ' kkride
            '
            rs.Open "SELECT * FROM desconto_canal WHERE sub_brand='" & Trim(strSubBrand) & "' AND cliente='" & Trim(strCodigoCliente) & "';", CONN, adOpenForwardOnly, adLockReadOnly
            '
            If rs.RecordCount > 0 Then
               '
               ' frmPedido.GridCtrl.TextMatrix(I, 5) = FormatCurrency(CStr(CDbl(strPrecoProduto) * ((100 - CDbl(rs("desconto"))) / 100)), 2, vbTrue, vbTrue, vbTrue)
               '
               ' strFormataNumero = CStr(CDbl(strPrecoProduto) * ((100 - CDbl(rs("desconto"))) / 100))
               '
               strFormataNumero = CStr(CDbl(frmPedido.GridCtrl.TextMatrix(I, 1)) * CDbl(frmPedido.GridCtrl.TextMatrix(I, 3)) * ((100 - CDbl(rs("desconto"))) / 100))
               '
               frmPedido.lblDescontoCanal.Visible = True
               '
               frmPedido.txtDescItem.Enabled = False
               '
            Else
               '
               ' rs("qtd_pedida") = frmPedido.GridCtrl.TextMatrix(I, 1)
               ' rs("qtd_faturada") = frmPedido.GridCtrl.TextMatrix(I, 2)
               ' rs("valor_unitario") = frmPedido.GridCtrl.TextMatrix(I, 3)
               ' rs("desconto") = frmPedido.GridCtrl.TextMatrix(I, 4)
               '
               ' euclides - 24/09/2002
               '
               ' frmPedido.GridCtrl.TextMatrix(I, 5) = FormatCurrency(CStr(CDbl(frmPedido.GridCtrl.TextMatrix(I, 1)) * CDbl(frmPedido.GridCtrl.TextMatrix(I, 3)) * ((100 - CDbl(frmPedido.GridCtrl.TextMatrix(I, 4))) / 100)))
               '
               strFormataNumero = CStr(CDbl(frmPedido.GridCtrl.TextMatrix(I, 1)) * CDbl(frmPedido.GridCtrl.TextMatrix(I, 3)) * ((100 - CDbl(frmPedido.GridCtrl.TextMatrix(I, 4))) / 100))
               '
               If usrhabilitar_desconto_item = True Then frmPedido.txtDescItem.Enabled = True
               '
               ' frmPedido.GridCtrl.TextMatrix(I, 5) = FormatCurrency(strPrecoProduto, 2, vbTrue, vbTrue, vbTrue)
               '
               ' frmPedido.GridCtrl.TextMatrix(I, 5) = FormatCurrency(CStr(CDbl(strPrecoProduto) * ((100 - CDbl(rs("desconto"))) / 100)), 2, vbTrue, vbTrue, vbTrue)
               '
            End If
            '
            'If rs.RecordCount > 0 Then
            '   '
            '   frmPedido.txtPrecoOriginal.Text = FormatCurrency(CStr(CDbl(strPrecoProduto) * ((100 - CDbl(rs("desconto"))) / 100)), 2, vbTrue, vbTrue, vbTrue)
            '   '
            'Else
            '   '
            '   frmPedido.txtPrecoOriginal.Text = FormatCurrency(strPrecoProduto, 2, vbTrue, vbTrue, vbTrue)
            '   '
            'End If
            '
            ' strFormataNumero = FormatCurrency(CDbl(txtQtdPedida.Text) * (CDbl(txtPrecoUnit.Text) * (1 - (CDbl(txtDescItem.Text) / 100))), 2, vbTrue, vbTrue, vbTrue)
            '
'            strFormataNumero = frmPedido.GridCtrl.TextMatrix(I, 5)
            '
            '================================================================================
            '
            'strFormataNumero = Mid(Trim(strFormataNumero), 3, Len(Trim(strFormataNumero)) - 2)
            ''
            'If InStr(1, Trim(strFormataNumero), ",", vbTextCompare) <> 0 Then
            '   '
            '   strFormataNumero = Mid(Trim(strFormataNumero), 1, InStr(1, Trim(strFormataNumero), ",", vbTextCompare) + 2)
            '   '
            'Else
            '   '
            '   strFormataNumero = Trim(strFormataNumero) & ",00"
            '   '
            'End If
            ''
            'While Mid(Trim(strFormataNumero), 1, 1) = "0"
            '      '
            '      strFormataNumero = Mid(Trim(strFormataNumero), 2, Len(Trim(strFormataNumero)) - 1)
            '      '
            'Wend
            ''
            'If Mid(Trim(strFormataNumero), 1, 1) = "," Then strFormataNumero = "0" & strFormataNumero
            ''
            '' GridCtrl.TextMatrix(GridCtrl.Rows - 1, 5) = FormatCurrency(CDbl(txtQtdPedida.Text) * (CDbl(Text1.Text) * (1 - (CDbl(txtDescItem.Text) / 100))), 2, vbTrue, vbTrue, vbTrue)
            ''
            ''
            '' GridCtrl.TextMatrix(GridCtrl.Rows - 1, 5) = strFormataNumero
            ''
            ''
            ''=========================================================================
            ''
            'strFormataNumero = rs("valor")
            '
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
            'frmContas.GridCtrl.TextMatrix(frmContas.GridCtrl.Rows - 1, 2) = strFormataNumero
            '
            '=========================================================================
            '
            frmPedido.GridCtrl.TextMatrix(I, 5) = strFormataNumero
            '
            frmPedido.GridCtrl.Col = 5
            frmPedido.GridCtrl.CellAlignment = flexAlignRightCenter
            '
            If rs.State = 1 Then rs.Close
            '
        End If
        '
        If rs.State = 1 Then rs.Close
        '
    Next
    '
    connClose
    '
    Set rs = Nothing
    '
    Screen.MousePointer = 0
    '
    dblTotal = 0
    '
    For I = 1 To (frmPedido.GridCtrl.Rows - 1)
        '
        dblTotal = dblTotal + CDbl(frmPedido.GridCtrl.TextMatrix(I, 5))
        '
    Next
    '
    frmPedido.Text2.Text = FormatCurrency(CStr(dblTotal), 2, vbTrue, vbTrue, vbTrue)
    '
    If Len(frmPedido.txtAcrescimo.Text) > 0 And IsNumeric(frmPedido.txtAcrescimo.Text) Then dblTotal = (dblTotal + CDbl(frmPedido.txtAcrescimo.Text))
    If Len(frmPedido.txtDesconto.Text) > 0 And IsNumeric(frmPedido.txtDesconto.Text) Then dblTotal = (dblTotal - CDbl(frmPedido.txtDesconto.Text))
    '
    frmPedido.txtLiquido.Text = FormatCurrency(CStr(dblTotal), 2, vbTrue, vbTrue, vbTrue)
    '
End Sub

Public Function ValidaCliqueItem() As Boolean
  '
  ' kride
  '
  ValidaCliqueItem = False
  '
  If Len(Trim(frmPedido.cboProdutos.Text)) <= 0 Then
     '
     MsgBox "Selecione um produto.", vbOKOnly + vbCritical, App.Title
     Screen.MousePointer = 0
     '
     Exit Function
     '
  End If
  '
  If Not IsNumeric(frmPedido.txtQtdPedida.Text) Or CDbl(frmPedido.txtQtdPedida.Text) <= 0 Then
     '
     MsgBox "Quantidade é inválida.", vbOKOnly + vbCritical, App.Title
     '
     frmPedido.txtQtdPedida.Text = 0 ' vbNullString
     frmPedido.txtQtdPedida.SelStart = 1
     frmPedido.txtQtdPedida.SelLength = Len(frmPedido.txtQtdPedida.Text)
     '
     Screen.MousePointer = 0
     '
     Exit Function
  End If
  If frmPedido.optPromocional01.Value = True Then
     If CDbl(frmPedido.txtQtdPedida.Text) > CDbl(frmPedido.txtPF01.Text) Or CDbl(frmPedido.txtQtdPedida.Text) < CDbl(frmPedido.txtPI01.Text) Then
        MsgBox "Quantidade pedida deve ser compatível com a promoção.", vbOKOnly + vbCritical, App.Title
        frmPedido.txtQtdPedida.Text = 0 ' vbNullString
        Screen.MousePointer = 0
        Exit Function
     End If
  End If
  If frmPedido.optPromocional02.Value = True Then
     If CDbl(frmPedido.txtQtdPedida.Text) > CDbl(frmPedido.txtPF02.Text) Or CDbl(frmPedido.txtQtdPedida.Text) < CDbl(frmPedido.txtPI02.Text) Then
        MsgBox "Quantidade pedida deve ser compatível com a promoção.", vbOKOnly + vbCritical, App.Title
        frmPedido.txtQtdPedida.Text = 0 ' vbNullString
        Screen.MousePointer = 0
        Exit Function
     End If
  End If
  If frmPedido.optPromocional03.Value = True Then
     If CDbl(frmPedido.txtQtdPedida.Text) > CDbl(frmPedido.txtPF03.Text) Or CDbl(frmPedido.txtQtdPedida.Text) < CDbl(frmPedido.txtPI03.Text) Then
        MsgBox "Quantidade pedida deve ser compatível com a promoção.", vbOKOnly + vbCritical, App.Title
        frmPedido.txtQtdPedida.Text = 0 ' vbNullString
        Screen.MousePointer = 0
        Exit Function
     End If
  End If
  If frmPedido.optPromocional04.Value = True Then
     If CDbl(frmPedido.txtQtdPedida.Text) > CDbl(frmPedido.txtPF04.Text) Or CDbl(frmPedido.txtQtdPedida.Text) < CDbl(frmPedido.txtPI04.Text) Then
        MsgBox "Quantidade pedida deve ser compatível com a promoção.", vbOKOnly + vbCritical, App.Title
        frmPedido.txtQtdPedida.Text = 0 ' vbNullString
        Screen.MousePointer = 0
        Exit Function
     End If
  End If
  If frmPedido.optPromocional05.Value = True Then
     If CDbl(frmPedido.txtQtdPedida.Text) > CDbl(frmPedido.txtPF05.Text) Or CDbl(frmPedido.txtQtdPedida.Text) < CDbl(frmPedido.txtPI05.Text) Then
        MsgBox "Quantidade pedida deve ser compatível com a promoção.", vbOKOnly + vbCritical, App.Title
        frmPedido.txtQtdPedida.Text = 0 ' vbNullString
        Screen.MousePointer = 0
        Exit Function
     End If
  End If
  '
  If Trim(frmPedido.txtQtdPedida.Text) = "" Then frmPedido.txtQtdPedida.Text = 0
  If Trim(frmPedido.txtPrecoUnit.Text) = "" Then frmPedido.txtPrecoUnit.Text = 0
  If Trim(frmPedido.txtDescItem.Text) = "" Then frmPedido.txtDescItem.Text = 0
  If Trim(frmPedido.txtTotal.Text) = "" Then frmPedido.txtTotal.Text = 0
  '
  If Not IsNumeric(frmPedido.txtQtdPedida.Text) Or Not IsNumeric(frmPedido.txtPrecoUnit.Text) Or Not IsNumeric(frmPedido.txtDescItem.Text) Then
     MsgBox "Verifique os valores digitados em Quantidade pedida, Valor do Item e Desconto.", vbOKOnly + vbCritical, App.Title
     Screen.MousePointer = 0
     Exit Function
  End If
  '
  ValidaCliqueItem = True
  '
End Function

Public Function VerificaStatusPedido(ByVal strCodigoaPesquisar As String) As Integer
    '
    '1 - verifica se cliente está ativo ou não
    '
    '2 - verifica se produto está em promoção ou não
    '
    'Retornos
    '
    '0 - True
    '1 - False
    '2 - Não pode fazer pedido
    '
    Screen.MousePointer = 11
    '
    Dim rs, rsObservacao
    Dim strCodigoCliente As String
    Dim mCondPagto As String
    Dim mTotalTitulos As Double
    '
    '---------- Novas
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
    '
    Dim valPreco1 As String
    Dim valPreco2 As String
    Dim valPreco3 As String
    Dim valPreco4 As String
    Dim valPreco5 As String
    '
    valVencidas = 0
    valVencer = 0
    valTotal = 0
    '
    '-------- Novas
    '
    connOpen
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    Set rsObservacao = CreateObject("ADOCE.Recordset.3.0")
    '
    Select Case Verificador
    Case "1"
         '
         On Error Resume Next
         '
         ' Passa para ultimo cliente a seleção corrente
         '
         mUltimoCliente = strCodigoaPesquisar
         '
         VerificaStatusPedido = 0
         '
         rs.Open "SELECT * FROM clientes WHERE nome_fantasia='" & Trim(strCodigoaPesquisar) & "';", CONN, adOpenForwardOnly, adLockReadOnly
         '
         If rs.RecordCount > 0 Then
            '
            strCodigoCliente = rs("codigo_cliente")
            '
            '
            ' observacoes_clientes
            '
            ' cliente VARCHAR(5)
            ' observacao TEXT
            '
            If rsObservacao.State = 1 Then rsObservacao.Close
            '
            rsObservacao.Open "SELECT * FROM observacoes_clientes WHERE cliente='" & Trim(strCodigoCliente) & "';", CONN, adOpenForwardOnly, adLockReadOnly
            '
            If rsObservacao.RecordCount > 0 Then
               '
               frmPedido.txtDMPP.Text = rsObservacao("observacao")
               '
               fraPedido.txtDMPP.ZOrder vbBringToFront ' vbSendToBack
               '
            End If
            '
            If rsObservacao.State = 1 Then rsObservacao.Close
            '
            strCodigo = rs("codigo_cliente")
            '
            If UCase(rs("status")) = "A" Then
               '
               VerificaStatusPedido = 0
               '
            Else
               '
               If rs.State = 1 Then rs.Close
               '
               rs.Open "SELECT * FROM vendedor WHERE codigo_vendedor='" & Trim(usrCodigoVendedor) & "';", CONN, adOpenForwardOnly, adLockReadOnly
               '
               If rs("aceita_pedido_bloq") = "S" Then
                  '
                  VerificaStatusPedido = 0
                  '
               Else
                  '
                  If rs("contra_senha") = "S" Then
                     '
                     frmValidaCliente.txtCodigoClienteAtual.Text = strCodigoCliente
                     frmValidaCliente.txtCodigoVendedorAtual.Text = usrCodigoVendedor
                     frmValidaCliente.txtDataAtual.Text = Mid(RetornaDataString(Now), 1, 4) + Mid(RetornaDataString(Now), 7, 2)
                     frmValidaCliente.txtNumeroPedido.Text = frmPedido.txtNumeroPedido.Text
                     '
                     Primos
                     '
                     VerificaStatusPedido = 1
                     '
                  Else
                     '
                     MsgBox "Cliente bloqueado e não pode fazer pedidos.", vbOKOnly + vbCritical, App.Title
                     '
                     VerificaStatusPedido = 2
                     '
                  End If
                  '
               End If
               '
            End If
            '
            If UCase(rs("bloqueado")) = "N" Then
               '
               VerificaStatusPedido = 0
               '
            Else
               '
               If rs.State = 1 Then rs.Close
               '
               rs.Open "SELECT * FROM vendedor WHERE codigo_vendedor='" & Trim(usrCodigoVendedor) & "';", CONN, adOpenForwardOnly, adLockReadOnly
               '
               If rs("aceita_pedido_bloq") = "S" Then
                  '
                  VerificaStatusPedido = 0
                  '
               Else
                  '
                  If rs("contra_senha") = "S" Then
                     '
                     frmValidaCliente.txtCodigoClienteAtual.Text = strCodigoCliente
                     frmValidaCliente.txtCodigoVendedorAtual.Text = usrCodigoVendedor
                     frmValidaCliente.txtDataAtual.Text = Mid(RetornaDataString(Now), 1, 4) + Mid(RetornaDataString(Now), 7, 2)
                     frmValidaCliente.txtNumeroPedido.Text = frmPedido.txtNumeroPedido.Text
                     '
                     Primos
                     '
                     VerificaStatusPedido = 1
                     '
                  Else
                     '
                     MsgBox "Cliente bloqueado e não pode fazer pedidos.", vbOKOnly + vbCritical, App.Title
                     '
                     VerificaStatusPedido = 2
                     '
                  End If
                  '
               End If
               '
            End If
            '
            If CDbl(rs("limite_credito")) > 0 Then
               '
               VerificaStatusPedido = 0
               '
            Else
               '
               If rs.State = 1 Then rs.Close
               '
               rs.Open "SELECT * FROM vendedor WHERE codigo_vendedor='" & Trim(usrCodigoVendedor) & "';", CONN, adOpenForwardOnly, adLockReadOnly
               '
               If rs("aceita_pedido_bloq") = "S" Then
                  '
                  VerificaStatusPedido = 0
                  '
               Else
                  '
                  If rs("contra_senha") = "S" Then
                     '
                     frmValidaCliente.txtCodigoClienteAtual.Text = strCodigoCliente
                     frmValidaCliente.txtCodigoVendedorAtual.Text = usrCodigoVendedor
                     frmValidaCliente.txtDataAtual.Text = Mid(RetornaDataString(Now), 1, 4) + Mid(RetornaDataString(Now), 7, 2)
                     frmValidaCliente.txtNumeroPedido.Text = frmPedido.txtNumeroPedido.Text
                     '
                     Primos
                     '
                     VerificaStatusPedido = 1
                     '
                  Else
                     '
                     MsgBox "Cliente bloqueado e não pode fazer pedidos.", vbOKOnly + vbCritical, App.Title
                     '
                     VerificaStatusPedido = 2
                     '
                  End If
                  '
               End If
               '
            End If
            '
            If rs.State = 1 Then rs.Close
            '
            rs.Open "SELECT * FROM vendedor WHERE codigo_vendedor='" & Trim(usrCodigoVendedor) & "';", CONN, adOpenForwardOnly, adLockReadOnly
            '
            If rs.RecordCount = 1 Then
               '
               strDataHoje = Mid(rs("data_cortetitulos"), 5, 4) & Mid(rs("data_cortetitulos"), 3, 2) & Mid(rs("data_cortetitulos"), 1, 2)
               '
            Else
               '
               MsgBox "Data de Corte inválida para os Títulos, Assume a data atual:" & Now, vbOKOnly + vbInformation, App.Title
               '
               strDataHoje = Mid(RetornaDataString(Now), 5, 4) & Mid(RetornaDataString(Now), 3, 2) & Mid(RetornaDataString(Now), 1, 2)
               '
            End If
            '
            If strDataHoje = "00000000" Then
               '
               MsgBox "Data de Corte inválida para os Títulos, Assume a data atual:" & Now, vbOKOnly + vbInformation, App.Title
               '
               strDataHoje = Mid(RetornaDataString(Now), 5, 4) & Mid(RetornaDataString(Now), 3, 2) & Mid(RetornaDataString(Now), 1, 2)
               '
            End If
            '
            ' MsgBox "Data de Corte=" & strDataHoje, vbOKOnly + vbInformation, App.Title
            '
            If rs.State = 1 Then rs.Close
            '
            '------------ Verifica se há títulos em aberto para este cliente.
            '
            rs.Open "SELECT * FROM titulos_aberto WHERE codigo_cliente='" & Trim(strCodigo) & "' order by vencimento_data ASC;", CONN, adOpenDynamic, adLockReadOnly
            '
            If rs.RecordCount > 0 Then
               '
               Do Until rs.EOF
                  '
                  ' strFormataNumero = FormatCurrency(rs("valor"), 2, vbTrue, vbTrue, vbTrue)
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
               If rs.State = 1 Then rs.Close
               '
               If valVencidas > 0.01 Then
                  '
                  If rs.State = 1 Then rs.Close
                  '
                  rs.Open "SELECT * FROM vendedor WHERE codigo_vendedor='" & Trim(usrCodigoVendedor) & "';", CONN, adOpenForwardOnly, adLockReadOnly
                  '
                  If rs("aceita_pedido_bloq") = "S" Then
                     '
                     VerificaStatusPedido = 0
                     '
                  Else
                     '
                     If rs("contra_senha") = "S" Then
                        '
                        frmValidaCliente.txtCodigoClienteAtual.Text = strCodigoCliente
                        frmValidaCliente.txtCodigoVendedorAtual.Text = usrCodigoVendedor
                        frmValidaCliente.txtDataAtual.Text = Mid(RetornaDataString(Now), 1, 4) + Mid(RetornaDataString(Now), 7, 2)
                        frmValidaCliente.txtNumeroPedido.Text = frmPedido.txtNumeroPedido.Text
                        '
                        Primos
                        '
                        VerificaStatusPedido = 1
                        '
                     Else
                        '
                        MsgBox "Cliente possui títulos vencidos(Total:" & FormatCurrency(CStr(valVencidas), 2, vbTrue, vbTrue, vbTrue) & ")e não pode fazer pedido.", vbOKOnly + vbCritical, App.Title
                        '
                        VerificaStatusPedido = 2
                        '
                     End If
                     '
                  End If
                  '
               End If
               '
            End If
            '
            '------------ Entrando
            '
         Else
            '
            ' Cliente inexistente
            '
            MsgBox "Cliente inexistente:(" & Trim(strCodigoaPesquisar) & ")", vbOKOnly + vbCritical, App.Title
            '
            strCodigoCliente = ""
            '
            VerificaStatusPedido = 2
            '
         End If
         '
         On Error GoTo 0
         '
    Case "2"
            '
         Dim strCodigoProduto As String
         Dim strSubBrand As String
         Dim strPrecoProduto As String
         Dim intTXTaPreencher As Integer
         Dim strCodigoTabela As String
         '
         If Trim(frmPedido.cboPedidoCliente.Text) <> "" Then
            '
            strNomeFantasia = frmPedido.cboPedidoCliente.Text
            '
         Else
            '
            strNomeFantasia = frmPedido.txtCliente.Text
            '
         End If
         '
         ' Passa para ultimo cliente a seleção corrente
         '
         mUltimoCliente = strNomeFantasia
         '
         rs.Open "SELECT * FROM clientes WHERE nome_fantasia='" & Trim(strNomeFantasia) & "';", CONN, adOpenForwardOnly, adLockReadOnly
         '
         If rs.RecordCount > 0 Then
            '
            strCodigoCliente = rs("codigo_cliente")
            '
         Else
            '
            MsgBox "Cliente inexistente:(" & strNomeFantasia & ")", vbOKOnly + vbCritical, App.Title
            '
            If rs.State = 1 Then rs.Close
            '
            Screen.MousePointer = 0
            '
            Exit Function
            '
         End If
         '
         If rs.State = 1 Then rs.Close
         '
         rs.Open "SELECT * FROM produtos WHERE descricao='" & Trim(strCodigoaPesquisar) & "';", CONN, adOpenForwardOnly, adLockReadOnly
         '
         If rs.RecordCount > 0 Then
            '
            strCodigoProduto = rs("codigo_produto")
            '
            strSubBrand = rs("sub_brand")
            '
            valPreco1 = rs("preco1")
            valPreco2 = rs("preco2")
            valPreco3 = rs("preco3")
            valPreco4 = rs("preco4")
            valPreco5 = rs("preco5")
            '
            If rs.State = 1 Then rs.Close
            '
            ' Condição de Pagamento----->
            '
            mCondPagto = ""
            '
            If Trim(frmPedido.cboCPagto.Text) <> "" Then
               '
               mCondPagto = frmPedido.cboCPagto.Text
               '
            Else
               '
               mCondPagto = frmPedido.txtCPagto.Text
               '
            End If
            '
            rs.Open "SELECT * FROM condicao_pagamento WHERE descricao='" & Trim(mCondPagto) & "';", CONN, adOpenForwardOnly, adLockReadOnly
            '
            ' MsgBox "Cond. Pagto:(" & Trim(mCondPagto) & ")", vbOKOnly + vbCritical, App.Title
            '
            ' MsgBox "Encontrou:" & CInt(rs.RecordCount) & " Registros", vbOKOnly + vbCritical, App.Title
            '
            If rs.RecordCount > 0 Then
               '
               strCodigoTabela = rs("codigo_tabela_preco")
               '
            Else
               '
               MsgBox "Selecione uma Condição de Pagamento(2).", vbOKOnly + vbCritical, App.Title
               '
               If rs.State = 1 Then rs.Close
               '
               Screen.MousePointer = 0
               '
               Exit Function
               '
            End If
            '
            If rs.State = 1 Then rs.Close
              '
''            'Tabela de Preços----------> euclides - 24/09/2002
''            '
''            '
''            ' rs.Open "SELECT * FROM tabela_precos WHERE produto='" & strCodigoProduto & "' AND tabela_precos='" & strCodigoTabela & "';", CONN, adOpenForwardOnly, adLockReadOnly
''            '
''            rs.Open "SELECT * FROM tabela_precos WHERE tabela_precos='" & strCodigoTabela & "' AND produto='" & strCodigoProduto & "';", CONN, adOpenForwardOnly, adLockReadOnly
''            '
''            If rs.RecordCount > 0 Then
''               strPrecoProduto = rs("preco")
''            Else
''               MsgBox "Não há preço cadastrado para este produto nesta condição de pagamento.", vbOKOnly + vbCritical, App.Title
''               '
''               If rs.State = 1 Then rs.Close
''               '
''               Screen.MousePointer = 0
''               '
''               Exit Function
''               '
''            End If
''            '
              Select Case CInt(strCodigoTabela)
              Case 1
                   strPrecoProduto = valPreco1
              Case 2
                   strPrecoProduto = valPreco2
              Case 3
                   strPrecoProduto = valPreco3
              Case 4
                   strPrecoProduto = valPreco4
              Case 5
                   strPrecoProduto = valPreco5
              Case Else
                   strPrecoProduto = valPreco5
              End Select
              '
              '
              '
            If rs.State = 1 Then rs.Close
            '
            'Desconto do Canal--------->
            '
            ' MsgBox "SubBrand=" & Trim(strSubBrand) & " cliente='" & Trim(strCodigoCliente), vbOKOnly + vbCritical, App.Title
            '
            rs.Open "SELECT * FROM desconto_canal WHERE sub_brand='" & Trim(strSubBrand) & "' AND cliente='" & Trim(strCodigoCliente) & "';", CONN, adOpenForwardOnly, adLockReadOnly
            '
            If rs.RecordCount > 0 Then
               '
               frmPedido.txtPrecoOriginal.Text = FormatCurrency(CStr(CDbl(strPrecoProduto) * ((100 - CDbl(rs("desconto"))) / 100)), 2, vbTrue, vbTrue, vbTrue)
               '
               frmPedido.lblDescontoCanal.Visible = True
               '
               ' lblDescontoCanal.Caption = FormatCurrency(CStr(CDbl(rs("desconto")) / 100), 2, vbTrue, vbTrue, vbTrue)
               '
               frmPedido.txtDescItem.Enabled = False
               '
               frmPedido.txtDescItem.Text = Mid(FormatCurrency(CStr(CDbl(rs("desconto"))), 2, vbTrue, vbTrue, vbTrue), 3, Len(FormatCurrency(CStr(CDbl(rs("desconto"))), 2, vbTrue, vbTrue, vbTrue)) - 2)
               '
            Else
               '
               frmPedido.txtPrecoOriginal.Text = FormatCurrency(strPrecoProduto, 2, vbTrue, vbTrue, vbTrue)
               '
               If usrhabilitar_desconto_item = True Then frmPedido.txtDescItem.Enabled = True
               '
               frmPedido.txtDescItem.Text = "0"
               '
            End If
            '
            '
            '
            If rs.State = 1 Then rs.Close
            '
            '-------------------------->
            'Promoções----------------->
            '
            frmPedido.txtPF01.Text = vbNullString
            frmPedido.txtPF02.Text = vbNullString
            frmPedido.txtPF03.Text = vbNullString
            frmPedido.txtPF04.Text = vbNullString
            frmPedido.txtPF05.Text = vbNullString
            '
            frmPedido.txtPI01.Text = vbNullString
            frmPedido.txtPI02.Text = vbNullString
            frmPedido.txtPI03.Text = vbNullString
            frmPedido.txtPI04.Text = vbNullString
            frmPedido.txtPI05.Text = vbNullString
            '
            frmPedido.txtPP01.Text = vbNullString
            frmPedido.txtPP02.Text = vbNullString
            frmPedido.txtPP03.Text = vbNullString
            frmPedido.txtPP04.Text = vbNullString
            frmPedido.txtPP05.Text = vbNullString
            '
            rs.Open "SELECT * FROM promocoes WHERE codigo_produto='" & Trim(strCodigoProduto) & "' ORDER BY qtd_inicial;", CONN, adOpenDynamic, adLockOptimistic
            '
            If rs.RecordCount > 0 Then
               '
               frmPedido.lblPromocao.Visible = True
               '
               intTXTaPreencher = 1
               '
               VerificaStatusPedido = 0
               '
               Do Until rs.EOF
                  '
                  Select Case intTXTaPreencher
                  Case 1
                       frmPedido.txtPF01.Text = CStr(CDbl(rs("qtd_final")))
                       frmPedido.txtPI01.Text = CStr(CDbl(rs("qtd_inicial")))
                       frmPedido.txtPP01.Text = FormatCurrency(rs("preco_promocional"), 2, vbTrue, vbTrue, vbTrue)
                  Case 2
                       frmPedido.txtPF02.Text = CStr(CDbl(rs("qtd_final")))
                       frmPedido.txtPI02.Text = CStr(CDbl(rs("qtd_inicial")))
                       frmPedido.txtPP02.Text = FormatCurrency(rs("preco_promocional"), 2, vbTrue, vbTrue, vbTrue)
                  Case 3
                       frmPedido.txtPF03.Text = CStr(CDbl(rs("qtd_final")))
                       frmPedido.txtPI03.Text = CStr(CDbl(rs("qtd_inicial")))
                       frmPedido.txtPP03.Text = FormatCurrency(rs("preco_promocional"), 2, vbTrue, vbTrue, vbTrue)
                  Case 4
                       frmPedido.txtPF04.Text = CStr(CDbl(rs("qtd_final")))
                       frmPedido.txtPI04.Text = CStr(CDbl(rs("qtd_inicial")))
                       frmPedido.txtPP04.Text = FormatCurrency(rs("preco_promocional"), 2, vbTrue, vbTrue, vbTrue)
                  Case 5
                       frmPedido.txtPF05.Text = CStr(CDbl(rs("qtd_final")))
                       frmPedido.txtPI05.Text = CStr(CDbl(rs("qtd_inicial")))
                       frmPedido.txtPP05.Text = FormatCurrency(rs("preco_promocional"), 2, vbTrue, vbTrue, vbTrue)
                  Case Else
                       Exit Do
                  End Select
                  '
                  rs.MoveNext
                  '
                  intTXTaPreencher = intTXTaPreencher + 1
                  '
               Loop
               '
            Else
               '
               frmPedido.lblPromocao.Visible = False
               '
               VerificaStatusPedido = 1
               '
            End If
            '
            If rs.State = 1 Then rs.Close
            '
            'Sub_Brand ------>
            '
            rs.Open "SELECT * FROM sub_brand WHERE codigo='" & Trim(strSubBrand) & "';", CONN, adOpenForwardOnly, adLockReadOnly
            '
            If rs.RecordCount > 0 Then
               '
               frmPedido.lblSubBrand.Caption = rs("descricao")
               '
               frmPedido.txtMensagem.Text = "Mensagem:" & rs("observacao")
               '
            End If
            '
         Else
            '
            VerificaStatusPedido = 1
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
End Function

Public Sub LimpaFormularioPEdido()
    '
    frmPedido.txtAcrescimo.Text = vbNullString
    frmPedido.txtCliente.Text = vbNullString
    frmPedido.txtCPagto.Text = vbNullString
    frmPedido.txtDesconto.Text = vbNullString
    
    '
    frmPedido.txtEntrega.Text = vbNullString
    frmPedido.txtFPgto.Text = vbNullString
    '
    frmPedido.txtLiquido.Text = vbNullString
    '
    frmPedido.txtNomeProdutoPromo.Text = vbNullString
    frmPedido.txtNumeroPedido.Text = vbNullString
    frmPedido.txtObs.Text = vbNullString
    '
    frmPedido.txtPF01.Text = vbNullString
    frmPedido.txtPF02.Text = vbNullString
    frmPedido.txtPF03.Text = vbNullString
    frmPedido.txtPF04.Text = vbNullString
    frmPedido.txtPF05.Text = vbNullString
    '
    frmPedido.txtPI01.Text = vbNullString
    frmPedido.txtPI02.Text = vbNullString
    frmPedido.txtPI03.Text = vbNullString
    frmPedido.txtPI04.Text = vbNullString
    frmPedido.txtPI05.Text = vbNullString
    '
    frmPedido.txtPP01.Text = vbNullString
    frmPedido.txtPP02.Text = vbNullString
    frmPedido.txtPP03.Text = vbNullString
    frmPedido.txtPP04.Text = vbNullString
    frmPedido.txtPP05.Text = vbNullString
    '
    frmPedido.txtQtdPedida.Text = vbNullString
    frmPedido.txtPrecoUnit.Text = vbNullString
    frmPedido.txtPrecoOriginal.Text = vbNullString
    '
    frmPedido.txtDescItem.Text = vbNullString
    '
    frmPedido.txtTotal.Text = vbNullString
    '
    frmPedido.Text2.Text = vbNullString
    frmPedido.txtTMvto.Text = vbNullString
    '
    frmPedido.GridCtrl.Clear
    frmPedido.GridCtrl.Rows = 1
    frmPedido.GridCtrl.Row = 0
    frmPedido.GridCtrl.Col = 0
    frmPedido.GridCtrl.CellBackColor = &HC0C0C0
    frmPedido.GridCtrl.CellFontBold = True
    frmPedido.GridCtrl.Row = 0
    frmPedido.GridCtrl.Col = 1
    frmPedido.GridCtrl.CellBackColor = &HC0C0C0
    frmPedido.GridCtrl.CellFontBold = True
    frmPedido.GridCtrl.Row = 0
    frmPedido.GridCtrl.Col = 2
    frmPedido.GridCtrl.CellBackColor = &HC0C0C0
    frmPedido.GridCtrl.CellFontBold = True
    frmPedido.GridCtrl.Row = 0
    frmPedido.GridCtrl.Col = 3
    frmPedido.GridCtrl.CellBackColor = &HC0C0C0
    frmPedido.GridCtrl.CellFontBold = True
    frmPedido.GridCtrl.Row = 0
    frmPedido.GridCtrl.Col = 4
    frmPedido.GridCtrl.CellBackColor = &HC0C0C0
    frmPedido.GridCtrl.CellFontBold = True
    frmPedido.GridCtrl.Row = 0
    frmPedido.GridCtrl.Col = 5
    frmPedido.GridCtrl.CellBackColor = &HC0C0C0
    frmPedido.GridCtrl.CellFontBold = True
    '
    frmPedido.GridCtrl.ColWidth(0) = 2500
    frmPedido.GridCtrl.ColWidth(1) = 800 ' 1500
    frmPedido.GridCtrl.ColWidth(2) = 800 ' 1500
    frmPedido.GridCtrl.ColWidth(3) = 800 ' 1500
    frmPedido.GridCtrl.ColWidth(4) = 800 ' 1500
    frmPedido.GridCtrl.ColWidth(5) = 800 ' 1500
    '
    frmPedido.GridCtrl.TextMatrix(0, 0) = "Produto"
    frmPedido.GridCtrl.TextMatrix(0, 1) = "Q. Pedida"
    frmPedido.GridCtrl.TextMatrix(0, 2) = "Q. Faturada"
    frmPedido.GridCtrl.TextMatrix(0, 3) = "$ Unitário"
    frmPedido.GridCtrl.TextMatrix(0, 4) = "% Desc."
    frmPedido.GridCtrl.TextMatrix(0, 5) = "Total"
    '
    frmPedido.cboCPagto.Visible = False
    frmPedido.cboFPagto.Visible = False
    frmPedido.cboPedidoCliente.Visible = False
    frmPedido.cboTmov.Visible = False
    '
    frmPedido.lblPromocao.Visible = False
    '
    frmPedido.lblDescontoCanal.Visible = False
    '
    IntIncrDataPed = 0
    '
End Sub

Public Sub LimpaFormularioPedEstoque()
    '
    frmEstoque.txtCPagto.Visible = False
    frmEstoque.txtFPagto.Visible = False
    frmEstoque.txtTMovto.Visible = False
    '
    frmEstoque.cboCPagto.Visible = True
    frmEstoque.cboFPagto.Visible = True
    frmEstoque.cboTMovto.Visible = True
    '
    IntIncrDataPed = 0
    '
End Sub

Public Function IncluiPedido() As Boolean
  '
  ' Passa para ultimo cliente a seleção corrente
  '
  mUltimoCliente = frmPedido.cboPedidoCliente.Text
  '
  Dim rs, rsaux
  '
  Dim strCodigoCliente As String
  Dim strNumPedido As String
  Dim strCondPagto As String
  Dim strFormaPagto As String
  Dim strTipoMovto  As String
  Dim strCodProduto As String
  '
  Dim mHora As String
  '
  IncluiPedido = False
  '
  If frmPedido.GridCtrl.Rows - 1 = 0 Then
     '
     MsgBox "Não será possível cadastrar esse pedido pois é necessário selecionar algum item para realizar um pedido.", vbOKOnly + vbCritical, App.Title
     '
     Exit Function
     '
  End If
  '
  connOpen
  '
  Set rs = CreateObject("ADOCE.Recordset.3.0")
  '
  Set rsaux = CreateObject("ADOCE.Recordset.3.0")
  '
  rs.Open "SELECT * FROM pedido WHERE numero_pedido_interno='" & frmPedido.txtNumeroPedido.Text & "';", CONN, adOpenDynamic, adLockOptimistic
  '
  If rs.RecordCount = 1 Then
     '
     MsgBox "Pedido existente:(" & frmPedido.txtNumeroPedido.Text & ")", vbOKOnly + vbCritical, App.Title
     '
     Exit Function
     '
  End If
  '
  rsaux.Open "SELECT * FROM clientes WHERE nome_fantasia='" & Trim(frmPedido.cboPedidoCliente.Text) & "';", CONN, adOpenForwardOnly, adLockReadOnly
  '
  If rsaux.RecordCount = 1 Then
     '
     strCodigoCliente = rsaux("codigo_cliente")
     '
  Else
     '
     MsgBox "Cliente inexistente:(" & Trim(frmPedido.cboPedidoCliente.Text) & ")", vbOKOnly + vbCritical, App.Title
     '
     If rsaux.State = 1 Then rsaux.Close
     '
     Screen.MousePointer = 0
     '
     Exit Function
     '
  End If
  '
  rsaux.Close
  '
  rsaux.Open "SELECT * FROM forma_pagamento WHERE descricao='" & Trim(frmPedido.cboFPagto.Text) & "';", CONN, adOpenForwardOnly, adLockReadOnly
  '
  If rsaux.RecordCount = 1 Then
     '
     strFormaPagto = rsaux("codigo")
     '
  Else
     '
     MsgBox "Forma de Pagamento inexistente:(" & Trim(frmPedido.cboFPagto.Text) & ")", vbOKOnly + vbCritical, App.Title
     '
     If rsaux.State = 1 Then rsaux.Close
     '
     Screen.MousePointer = 0
     '
     Exit Function
     '
  End If
  '
  rsaux.Close
  '
  rsaux.Open "SELECT * FROM condicao_pagamento WHERE descricao='" & Trim(frmPedido.cboCPagto.Text) & "';", CONN, adOpenForwardOnly, adLockReadOnly
  '
  If rsaux.RecordCount = 1 Then
     '
     strCondPagto = rsaux("codigo_condicao")
     '
     If rsaux("pedido_minimo") <> "0000000000" Then
        '
        If CDbl(rsaux("pedido_minimo")) > CDbl(frmPedido.txtLiquido.Text) Then
           '
           MsgBox "Pedido não satisfaz o valor mínimo para esta Condicao de Pagamento:(" & Trim(FormatCurrency(CDbl(rsaux("pedido_minimo")), 2, vbTrue, vbTrue, vbTrue)) & ")", vbOKOnly + vbCritical, App.Title
           '
           Screen.MousePointer = 0
           '
           Exit Function
           '
        End If
        '
     End If
     '
  Else
     '
     MsgBox "Condição de Pagamento inexistente:(" & Trim(frmPedido.cboCPagto.Text) & ")", vbOKOnly + vbCritical, App.Title
     '
     If rsaux.State = 1 Then rsaux.Close
     '
     Screen.MousePointer = 0
     '
     Exit Function
     '
  End If
  '
  rsaux.Close
  '
  rsaux.Open "SELECT * FROM tipo_movimento WHERE descricao='" & Trim(frmPedido.cboTmov.Text) & "';", CONN, adOpenForwardOnly, adLockReadOnly
  '
  If rsaux.RecordCount = 1 Then
     '
     strTipoMovto = rsaux("codigo")
     '
  Else
     '
     MsgBox "Tipo de Movimento inexistente:(" & Trim(frmPedido.cboTmov.Text) & ")", vbOKOnly + vbCritical, App.Title
     '
     If rsaux.State = 1 Then rsaux.Close
     '
     Screen.MousePointer = 0
     '
     Exit Function
     '
  End If
  '
  rsaux.Close
  '
  strNumPedido = Trim(frmPedido.txtNumeroPedido.Text)
  '
  Select Case Len(strNumPedido)
  Case 1
       strNumPedido = "00000" & strNumPedido
  Case 2
       strNumPedido = "0000" & strNumPedido
  Case 3
       strNumPedido = "000" & strNumPedido
  Case 4
       strNumPedido = "00" & strNumPedido
  Case 5
       strNumPedido = "0" & strNumPedido
  Case 6
       strNumPedido = strNumPedido
  End Select
  '
  rs.AddNew
  '
  ' codigo_vendedor VARCHAR(5),
  ' codigo_cliente VARCHAR(5),
  ' numero_pedido_interno VARCHAR(6),
  ' numero_pedido_externo VARCHAR(6),
  ' pedido_cliente VARCHAR(10),
  ' data_emissao VARCHAR(8),
  ' hora_emissao VARCHAR(6),
  ' data_entrega VARCHAR(8),
  ' acrescimo_valor VARCHAR(10),
  ' desconto_valor VARCHAR(10),
  ' forma_pgto VARCHAR(2),
  ' condicao_pgto VARCHAR(2),
  ' tipo_movimento VARCHAR(1),
  ' status VARCHAR(1),
  ' observacao TEXT);"
  '
  '
  ' rs("codigo_vendedor") = usrCodigoVendedor
  '
  Select Case Len(Trim(usrCodigoVendedor))
  Case 1
       rs("codigo_vendedor") = "0000" & Trim(usrCodigoVendedor)
  Case 2
       rs("codigo_vendedor") = "000" & Trim(usrCodigoVendedor)
  Case 3
       rs("codigo_vendedor") = "00" & Trim(usrCodigoVendedor)
  Case 4
       rs("codigo_vendedor") = "0" & Trim(usrCodigoVendedor)
  Case 5
       rs("codigo_vendedor") = Trim(usrCodigoVendedor)
  Case Else
       rs("codigo_vendedor") = Left(usrCodigoVendedor, 5)
  End Select
  '
  rs("codigo_cliente") = strCodigoCliente
  '
  rs("numero_pedido_interno") = strNumPedido
  '
  rs("numero_pedido_externo") = "-"
  '
  rs("pedido_cliente") = "-"
  '
  rs("data_emissao") = RetornaDataString(Now) ' CDate(frmPedido.txtEmissao.Text)) ' Euclides
  '
  rs("hora_emissao") = RetornaHoraString(CStr(Hour(Now)), CStr(Minute(Now)), CStr(Second(Now)))
  '
  rs("data_entrega") = RetornaDataString(CDate(frmPedido.txtEntrega.Text))
  rs("acrescimo_valor") = frmPedido.txtAcrescimo.Text
  rs("desconto_valor") = frmPedido.txtDesconto.Text
  '
  rs("forma_pgto") = strFormaPagto
  '
  rs("condicao_pgto") = strCondPagto
  '
  rs("tipo_movimento") = strTipoMovto
  '
  rs("status") = "D"
  '
  rs("observacao") = frmPedido.txtObs.Text
  '
  rs.Update
  '
  If rs.State = 1 Then rs.Close
  '
  rs.Open "SELECT * FROM itens_pedido WHERE numero_pedido='" & frmPedido.txtNumeroPedido.Text & "';", CONN, adOpenDynamic, adLockOptimistic
  '
  For I = 1 To (frmPedido.GridCtrl.Rows - 1)
      '
      rsaux.Open "SELECT * FROM produtos WHERE descricao='" & Trim(frmPedido.GridCtrl.TextMatrix(I, 0)) & "';", CONN, adOpenForwardOnly, adLockReadOnly
      '
      If rsaux.RecordCount = 1 Then
         '
         strCodProduto = rsaux("codigo_produto")
         '
      Else
         '
         MsgBox "Produto inexistente:(" & Trim(frmPedido.GridCtrl.TextMatrix(I, 0)) & ")", vbOKOnly + vbCritical, App.Title
         '
         Exit Function
         '
      End If
      '
      rsaux.Close
      '
      rs.AddNew
      '
      ' numero_pedido VARCHAR(6),
      ' codigo_produto VARCHAR(6),
      ' qtd_pedida VARCHAR(8),
      ' qtd_faturada VARCHAR(8),
      ' valor_unitario VARCHAR(10),
      ' desconto VARCHAR(6));"
      '
      rs("numero_pedido") = strNumPedido
      rs("codigo_produto") = strCodProduto
      '
      rs("qtd_pedida") = frmPedido.GridCtrl.TextMatrix(I, 1)
      rs("qtd_faturada") = frmPedido.GridCtrl.TextMatrix(I, 2)
      rs("valor_unitario") = frmPedido.GridCtrl.TextMatrix(I, 3)
      rs("desconto") = frmPedido.GridCtrl.TextMatrix(I, 4)
      '
      rs.Update
      '
  Next
  '
  If rs.State = 1 Then rs.Close
  '
  rs.Open "SELECT * FROM roteiro_percorrido;", CONN, adOpenDynamic, adLockOptimistic
  '
  rs.AddNew
  '
  rs("codigo_vendedor_atual") = usrCodigoVendedor
  '
  rs("data_base") = Trim(Mid(frmPedido.txtEntrega.Text, 1, 2)) & Trim(Mid(frmPedido.txtEntrega.Text, 4, 2)) & Trim(Mid(frmPedido.txtEntrega.Text, 7, 4))
  '
  rs("codigo_cliente") = strCodigoCliente
  '
  rs("data_visita") = RetornaDataString(Now)
  '
  mHora = RetornaHoraString(CStr(Hour(Now)), CStr(Minute(Now)), CStr(Second(Now)))
  '
  rs("hora_visita") = mHora
  '
  rs("motivo_nao_visita") = "---"
  '
  rs("visita_extra") = "N"
  '
  rs("automatico") = "N"
  '
  rs("dia_visita") = Trim(CStr(Weekday(Now, vbSunday)))
  '
  rs("status") = "*"
  '
  rs.Update
  '
  If rs.State = 1 Then rs.Close
  '
  If rsaux.State = 1 Then rs.Close
  '
  connClose
  '
  Set rs = Nothing
  '
  Set rsaux = Nothing
  '
  Screen.MousePointer = 0
  '
  IncluiPedido = True
  '
End Function

Public Function EditarPedido() As Boolean
    '
    Dim mDataEntrega As String
    '
    Dim strCodigoCliente As String
    Dim strNumPedido As String
    Dim strCondPagto As String
    Dim strFormaPagto As String
    Dim strTipoMovto  As String
    Dim strCodProduto As String
    '
    ' Passa para ultimo cliente a seleção corrente
    '
    mUltimoCliente = frmPedido.cboPedidoCliente.Text
    '
    If frmPedido.GridCtrl.Rows - 1 = 0 Then
       '
       MsgBox "Não será possível alterar esse pedido pois é necessário selecionar algum item para realizar um pedido.", vbOKOnly + vbCritical, App.Title
       '
       IncluiPedido = False
       '
       Exit Function
       '
    End If
    '
    Dim rs, rsaux
    Dim strNumeroPedido As String
    '
    EditarPedido = False
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    Set rsaux = CreateObject("ADOCE.Recordset.3.0")
    '
    connOpen
    '
    ' pedido
    '
    ' codigo_vendedor VARCHAR(5),
    ' codigo_cliente VARCHAR(5),
    ' numero_pedido_interno VARCHAR(6),
    ' numero_pedido_externo VARCHAR(6),
    ' pedido_cliente VARCHAR(10),
    ' data_emissao VARCHAR(8),
    ' hora_emissao VARCHAR(6),
    ' data_entrega VARCHAR(8),
    ' acrescimo_valor VARCHAR(10),
    ' desconto_valor VARCHAR(10),
    ' forma_pgto VARCHAR(2),
    ' condicao_pgto VARCHAR(2),
    ' tipo_movimento VARCHAR(1),
    ' status VARCHAR(1),
    ' observacao TEXT);"
    '
    strNumPedido = Trim(frmPedido.txtNumeroPedido.Text)
    '
    Select Case Len(strNumPedido)
    Case 1
         strNumPedido = "00000" & strNumPedido
    Case 2
         strNumPedido = "0000" & strNumPedido
    Case 3
         strNumPedido = "000" & strNumPedido
    Case 4
         strNumPedido = "00" & strNumPedido
    Case 5
         strNumPedido = "0" & strNumPedido
    Case 6
         strNumPedido = strNumPedido
    End Select
    '
    rs.Open "SELECT * FROM pedido WHERE numero_pedido_interno='" & strNumPedido & "';", CONN, adOpenDynamic, adLockOptimistic
    '
    '=============================================================================
    '
    If rs.RecordCount > 0 Then
        '
        ' rs("codigo_vendedor") = usrCodigoVendedor
        '
        'Select Case Len(Trim(usrCodigoVendedor))
        'Case 1
        '     rs("codigo_vendedor") = "0000" & Trim(usrCodigoVendedor)
        'Case 2
        '     rs("codigo_vendedor") = "000" & Trim(usrCodigoVendedor)
        'Case 3
        '     rs("codigo_vendedor") = "00" & Trim(usrCodigoVendedor)
        'Case 4
        '     rs("codigo_vendedor") = "0" & Trim(usrCodigoVendedor)
        'Case 5
        '     rs("codigo_vendedor") = Trim(usrCodigoVendedor)
        'Case Else
        '     rs("codigo_vendedor") = Left(usrCodigoVendedor, 5)
        'End Select
        '
        rs("data_emissao") = RetornaDataString(Now)
        '
        If IsDate(frmPedido.txtEntrega.Text) Then
           '
           mDataEntrega = RetornaDataString(CDate(frmPedido.txtEntrega.Text))
           '
        Else
           '
           mDataEntrega = ""
           '
        End If
        '
        rs("data_entrega") = mDataEntrega
        '
        '====================================================================================
        '
        rsaux.Open "SELECT * FROM clientes WHERE nome_fantasia='" & Trim(frmPedido.cboPedidoCliente.Text) & "';", CONN, adOpenForwardOnly, adLockReadOnly
        '
        If rsaux.RecordCount = 1 Then
           '
           strCodigoCliente = rsaux("codigo_cliente")
           '
        Else
           '
           MsgBox "Cliente inexistente:(" & Trim(frmPedido.cboPedidoCliente.Text) & ")", vbOKOnly + vbCritical, App.Title
           '
           If rsaux.State = 1 Then rsaux.Close
           '
           Screen.MousePointer = 0
           '
           Exit Function
           '
        End If
        '
        rsaux.Close
        '
        rsaux.Open "SELECT * FROM forma_pagamento WHERE descricao='" & Trim(frmPedido.cboFPagto.Text) & "';", CONN, adOpenForwardOnly, adLockReadOnly
        '
        If rsaux.RecordCount = 1 Then
           '
           strFormaPagto = rsaux("codigo")
           '
        Else
           '
           MsgBox "Forma de Pagamento inexistente:(" & Trim(frmPedido.cboFPagto.Text) & ")", vbOKOnly + vbCritical, App.Title
           '
           If rsaux.State = 1 Then rsaux.Close
           '
           Screen.MousePointer = 0
           '
           Exit Function
           '
        End If
        '
        rsaux.Close
        '
        rsaux.Open "SELECT * FROM condicao_pagamento WHERE descricao='" & Trim(frmPedido.cboCPagto.Text) & "';", CONN, adOpenForwardOnly, adLockReadOnly
        '
        If rsaux.RecordCount = 1 Then
           '
           strCondPagto = rsaux("codigo_condicao")
           '
           If rsaux("pedido_minimo") <> "0000000000" Then
              '
              If CDbl(rsaux("pedido_minimo")) > CDbl(frmPedido.txtLiquido.Text) Then
                 '
                 MsgBox "Pedido não satisfaz o valor mínimo para esta Condicao de Pagamento:(" & Trim(FormatCurrency(CDbl(rsaux("pedido_minimo")), 2, vbTrue, vbTrue, vbTrue)) & ")", vbOKOnly + vbCritical, App.Title
                 '
                 Screen.MousePointer = 0
                 '
                 Exit Function
                 '
              End If
              '
           End If
           '
        Else
           '
           MsgBox "Condição de Pagamento inexistente:(" & Trim(frmPedido.cboCPagto.Text) & ")", vbOKOnly + vbCritical, App.Title
           '
           If rsaux.State = 1 Then rsaux.Close
           '
           Screen.MousePointer = 0
           '
           Exit Function
           '
        End If
        '
        rsaux.Close
        '
        rsaux.Open "SELECT * FROM tipo_movimento WHERE descricao='" & Trim(frmPedido.cboTmov.Text) & "';", CONN, adOpenForwardOnly, adLockReadOnly
        '
        If rsaux.RecordCount = 1 Then
           '
           strTipoMovto = rsaux("codigo")
           '
        Else
           '
           MsgBox "Tipo de Movimento inexistente:(" & Trim(frmPedido.cboTmov.Text) & ")", vbOKOnly + vbCritical, App.Title
           '
           If rsaux.State = 1 Then rsaux.Close
           '
           Screen.MousePointer = 0
           '
           Exit Function
           '
        End If
        '
        rsaux.Close
        '
        ' codigo_vendedor VARCHAR(5),
        ' codigo_cliente VARCHAR(5),
        ' numero_pedido_interno VARCHAR(6),
        ' numero_pedido_externo VARCHAR(6),
        ' pedido_cliente VARCHAR(10),
        ' data_emissao VARCHAR(8),
        ' hora_emissao VARCHAR(6),
        ' data_entrega VARCHAR(8),
        ' acrescimo_valor VARCHAR(10),
        ' desconto_valor VARCHAR(10),
        ' forma_pgto VARCHAR(2),
        ' condicao_pgto VARCHAR(2),
        ' tipo_movimento VARCHAR(1),
        ' status VARCHAR(1),
        ' observacao TEXT);"
        '
        rs("codigo_cliente") = strCodigoCliente
        '
        rs("numero_pedido_interno") = strNumPedido
        '
        'rs("numero_pedido_externo") = "-"
        '
        'rs("pedido_cliente") = "-"
        '
        rs("data_emissao") = RetornaDataString(Now)
        '
        rs("hora_emissao") = RetornaHoraString(CStr(Hour(Now)), CStr(Minute(Now)), CStr(Second(Now)))
        '
        rs("data_entrega") = RetornaDataString(CDate(frmPedido.txtEntrega.Text))
        '
        rs("acrescimo_valor") = frmPedido.txtAcrescimo.Text
        '
        rs("desconto_valor") = frmPedido.txtDesconto.Text
        '
        rs("forma_pgto") = strFormaPagto
        '
        rs("condicao_pgto") = strCondPagto
        '
        rs("tipo_movimento") = strTipoMovto
        '
        rs("status") = "D"
        '
        rs("observacao") = frmPedido.txtObs.Text
        '
        rs.Update
        '
    End If
    '
    If rs.State = 1 Then rs.Close
    '
    rs.Open "SELECT * FROM itens_pedido WHERE numero_pedido='" & strNumPedido & "';", CONN, adOpenDynamic, adLockOptimistic
    '
    On Error Resume Next
    '
    Do Until rs.EOF
       '
       rs.MoveFirst
       rs.Delete
       '
    Loop
    '
    On Error GoTo 0
    '
    'Select Case Len(frmPedido.txtNumeroPedido.Text)
    'Case 1
    '     rs("numero_pedido") = "00000" & frmPedido.txtNumeroPedido.Text
    'Case 2
    '     rs("numero_pedido") = "0000" & frmPedido.txtNumeroPedido.Text
    'Case 3
    '     rs("numero_pedido") = "000" & frmPedido.txtNumeroPedido.Text
    'Case 4
    '     rs("numero_pedido") = "00" & frmPedido.txtNumeroPedido.Text
    'Case 5
    '     rs("numero_pedido") = "0" & frmPedido.txtNumeroPedido.Text
    'Case 6
    '     rs("numero_pedido") = frmPedido.txtNumeroPedido.Text
    'End Select
    '
    For I = 1 To (frmPedido.GridCtrl.Rows - 1)
        '
        rsaux.Open "SELECT * FROM produtos WHERE descricao='" & Trim(frmPedido.GridCtrl.TextMatrix(I, 0)) & "';", CONN, adOpenForwardOnly, adLockReadOnly
        '
        If rsaux.RecordCount > 0 Then
           '
           strCodProduto = rsaux("codigo_produto")
           '
           rsaux.Close
           '
           rs.AddNew
           '
           rs("numero_pedido") = strNumPedido
           rs("codigo_produto") = strCodProduto
           '
           rs("qtd_pedida") = frmPedido.GridCtrl.TextMatrix(I, 1)
           rs("qtd_faturada") = frmPedido.GridCtrl.TextMatrix(I, 2)
           rs("valor_unitario") = frmPedido.GridCtrl.TextMatrix(I, 3)
           rs("desconto") = frmPedido.GridCtrl.TextMatrix(I, 4)
           '
           rs.Update
           '
        Else
           '
           MsgBox "Produto Inexistente:(" & Trim(frmPedido.GridCtrl.TextMatrix(I, 0)) & ")", vbOKOnly + vbCritical, App.Title
           '
           If rsaux.State = 1 Then rsaux.Close
           '
           Screen.MousePointer = 0
           '
           Exit Function
           '
        End If
        '
    Next
    '
    If rs.State = 1 Then rs.Close
    If rsaux.State = 1 Then rsaux.Close
    '
    connClose
    '
    Set rs = Nothing
    Set rsaux = Nothing
    '
    EditarPedido = True
    '
    Screen.MousePointer = 0
    '
End Function

Public Function ExcluiPedido() As Boolean
  '
  ' Passa para ultimo cliente a seleção corrente
  '
  mUltimoCliente = frmPedido.cboPedidoCliente.Text
  '
  Dim rs
  '
  Screen.MousePointer = 11
  '
  ExcluiPedido = False
  '
  connOpen
  '
  Set rs = CreateObject("ADOCE.Recordset.3.0")
  '
  rs.Open "SELECT * FROM pedido WHERE numero_pedido_interno='" & frmPedido.txtNumeroPedido.Text & "';", CONN, adOpenDynamic, adLockOptimistic
  '
  If rs.RecordCount > 0 Then
     '
     rs.Delete
     '
     If rs.State = 1 Then rs.Close
     '
     rs.Open "SELECT * FROM itens_pedido WHERE numero_pedido='" & frmPedido.txtNumeroPedido.Text & "';", CONN, adOpenDynamic, adLockOptimistic
     '
     If rs.RecordCount > 0 Then
        '
        On Error Resume Next
        '
        Do Until rs.EOF
           '
           rs.Delete
           rs.MoveFirst
           '
        Loop
        '
        On Error GoTo 0
        '
     End If
     '
     LimpaFormularioPEdido
     '
     ExcluiPedido = True
     '
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
