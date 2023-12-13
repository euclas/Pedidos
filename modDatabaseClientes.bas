Attribute VB_Name = "modDatabaseClientes"
Option Explicit

Public Function IncluiCliente() As Boolean
    '
    Dim rs
    '
    Dim strCodigo As String
    Dim strDataFundacao As String
    Dim strPredioProprio As String
    '
    Dim mCidadeEndereco As String
    Dim mCidadeCobranca As String
    Dim mAtividade As String
    '
    mCidadeEndereco = RetornaCodigoCidade(frmClientes.cboCidadeEndereco.Text, 1)
    mCidadeCobranca = RetornaCodigoCidade(frmClientes.cboCidadeCobranca.Text, 1)
    mAtividade = RetornaAtividade(Trim(frmClientes.cboAtividade.Text), 1)
    '
    Screen.MousePointer = 11
    '
    connOpen
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    '
    strCodigo = Trim(frmClientes.cboCodigo.Text)
    '
    Select Case Len(strCodigo)
    Case 1
         strCodigo = "0000" & Trim(strCodigo)
    Case 2
         strCodigo = "000" & Trim(strCodigo)
    Case 3
         strCodigo = "00" & Trim(strCodigo)
    Case 4
         strCodigo = "0" & Trim(strCodigo)
    Case 5
         strCodigo = Trim(strCodigo)
    Case Else
         strCodigo = Left(Trim(strCodigo), 5)
    End Select
    '
    rs.Open "SELECT * FROM clientes WHERE codigo_cliente='" & Trim(strCodigo) & "';", CONN, adOpenDynamic, adLockOptimistic
    '
    If rs.RecordCount < 1 Then
       '
       If rs.State = 1 Then rs.Close
       '
       rs.Open "SELECT * FROM clientes;", CONN, adOpenDynamic, adLockOptimistic
       '
       rs.AddNew
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
       Select Case Len(strCodigo)
       Case 1
            rs("codigo_cliente") = "0000" & Trim(strCodigo)
       Case 2
            rs("codigo_cliente") = "000" & Trim(strCodigo)
       Case 3
            rs("codigo_cliente") = "00" & Trim(strCodigo)
       Case 4
            rs("codigo_cliente") = "0" & Trim(strCodigo)
       Case 5
            rs("codigo_cliente") = Trim(strCodigo)
       Case Else
            rs("codigo_cliente") = Left(Trim(strCodigo), 5)
       End Select
       '
       rs("nome_fantasia") = frmClientes.txtFantasia.Text
       rs("razao_social") = frmClientes.txtRazaoSocial.Text
       rs("endereco_entrega") = frmClientes.txtEntrega.Text
       rs("cep_entrega") = frmClientes.txtCEP.Text
       rs("bairro_entrega") = frmClientes.txtBairro.Text
       '
       rs("cidade_entrega") = mCidadeEndereco
       '
       rs("endereco_cobranca") = frmClientes.txtCobranca.Text
       rs("cep_cobranca") = frmClientes.txtCobrancaCEP.Text
       rs("bairro_cobranca") = frmClientes.txtCobrancaBairro.Text
       '
       rs("cidade_cobranca") = mCidadeCobranca
       '
       rs("telefone") = frmClientes.txtTelefone.Text
       rs("fax") = frmClientes.txtFax.Text
       rs("email") = frmClientes.txtEmail.Text
       rs("www") = frmClientes.txtWWW.Text
       '
       If IsDate(Trim(frmClientes.txtDataFundacao.Text)) Then
          '
          strDataFundacao = RetornaDataString(CDate(frmClientes.txtDataFundacao.Text))
          '
       Else
          '
          MsgBox "(4) - Data de Fundação não é uma Data Válida:(" & frmClientes.txtDataFundacao.Text & ")", vbOKOnly + vbCritical, App.Title
          '
          strDataFundacao = ""
          '
       End If
       '
       rs("data_fundacao") = strDataFundacao
       '
       If frmClientes.optNao = True Then
          '
          strPredioProprio = "N"
          '
       Else
          '
          strPredioProprio = "S"
          '
       End If
       '
       rs("predio_proprio") = strPredioProprio
       '
       rs("referencia_bancaria_1") = frmClientes.txtRefBanc01.Text
       rs("referencia_bancaria_2") = frmClientes.txtRefBanc02.Text
       rs("referencia_comercial_1") = frmClientes.txtRefPess01.Text
       rs("referencia_comercial_2") = frmClientes.txtRefPess02.Text
       rs("contato") = frmClientes.txtContato.Text
       rs("cnpjmf") = frmClientes.txtCNPJMF.Text
       rs("incricao_estadual") = frmClientes.txtIEST.Text
       rs("cpf") = frmClientes.txtCPF.Text
       rs("rg") = frmClientes.txtRG.Text
       rs("ramo_atividade") = mAtividade
       '
       rs("status") = mStatusCliente
       '
       rs("predio_proprio") = "N"
       rs("data_ultima_compra") = ""
       rs("valor_ultima_compra") = ""
       rs("desconto_maximo") = ""
       rs("limite_credito") = ""
       rs("bloqueado") = "S"
       rs("condicao_pagamento_padrao") = ""
       rs("forma_pagamento_padrao") = ""
       rs("periodicidade_visita") = ""
       '
       rs.Update
       '
       ' MsgBox "Vendedor:" & rs("codigo_vendedor") & "  Codigo:(" & strCodigo & ")", vbOKOnly + vbCritical, App.Title
       '
       If rs.State = 1 Then rs.Close
       '
       rs.Open "SELECT * FROM observacoes_clientes WHERE cliente='" & Trim(strCodigo) & "';", CONN, adOpenDynamic, adLockOptimistic
       '
       If rs.RecordCount > 0 Then
          rs("observacao") = frmClientes.txtObs.Text
       Else
          '
          If rs.State = 1 Then rs.Close
          '
          rs.Open "SELECT * FROM observacoes_clientes;", CONN, adOpenDynamic, adLockOptimistic
          '
          rs.AddNew
          '
          Select Case Len(strCodigo)
          Case 1
               rs("cliente") = "0000" & Trim(strCodigo)
          Case 2
               rs("cliente") = "000" & Trim(strCodigo)
          Case 3
               rs("cliente") = "00" & Trim(strCodigo)
          Case 4
               rs("cliente") = "0" & Trim(strCodigo)
          Case 5
               rs("cliente") = Trim(strCodigo)
          Case Else
               rs("cliente") = Left(Trim(strCodigo), 5)
          End Select
          '
          rs("observacao") = frmClientes.txtObs.Text
          '
       End If
       '
       rs.Update
       '
    Else
       '
       MsgBox "Não foi possível incluir cliente.", vbOKOnly + vbCritical, App.Title
       '
    End If
    '
    If rs.State = 1 Then rs.Close
    '
    Set rs = Nothing
    '
    connClose
    '
    AcertaVendedor 1
    '
    Screen.MousePointer = 0
    '
    If Err.Number <> 0 Then
       '
       IncluiCliente = False
       '
    Else
       '
       IncluiCliente = True
       '
    End If
    '
End Function
'
Public Sub EncheComboClientes(cboCombo As ComboBox, valTipoCombo As Integer)
  '
  Dim rs
  Set rs = CreateObject("ADOCE.Recordset.3.0")
  '
  Select Case valTipoCombo
  Case 1 'Codigo do cliente
       rs.Open "SELECT * FROM clientes ORDER BY codigo_cliente;", CONN, adOpenFowardOnly, adLockReadOnly
       cboCombo.Clear
       If rs.RecordCount > 0 Then
          Do Until rs.EOF
             cboCombo.AddItem rs("codigo_cliente") & " " & rs("nome_fantasia")
             rs.MoveNext
             Loop
       End If
  Case 2 'Nome Fantasia
       rs.Open "SELECT * FROM clientes ORDER BY nome_fantasia;", CONN, adOpenFowardOnly, adLockReadOnly
       cboCombo.Clear
       If rs.RecordCount > 0 Then
          Do Until rs.EOF
             cboCombo.AddItem rs("nome_fantasia")
             rs.MoveNext
             Loop
       End If
  Case 3 'Razão Social
       rs.Open "SELECT * FROM clientes ORDER BY razao_social;", CONN, adOpenFowardOnly, adLockReadOnly
       cboCombo.Clear
       If rs.RecordCount > 0 Then
          Do Until rs.EOF
             cboCombo.AddItem rs("razao_social")
             rs.MoveNext
             Loop
       End If
  Case 4 'Todos - apenas para formulário de clientes
       '
       frmClientes.cboCodigo.Clear
       frmClientes.cboFantasia.Clear
       frmClientes.cboRSocial.Clear
       '
       If rs.RecordCount > 0 Then
          Do Until rs.EOF
             frmClientes.cboCodigo.AddItem rs("codigo_cliente")
             frmClientes.cboFantasia.AddItem rs("nome_fantasia")
             frmClientes.cboRSocial.AddItem rs("razao_social")
             rs.MoveNext
             Loop
       End If
       '
  End Select
  '
  If rs.State = 1 Then rs.Close
  '
  Set rs = Nothing
  '
End Sub

Public Sub EncheFormularioCliente(strParametro As String, valTipoParamentro As Integer)
   '
    Dim rs
    Dim strCodigo As String
    '
    '1 - Código do Cliente
    '2 - Nome Fantasia
    '3 - Razão Social
    '
    Screen.MousePointer = 11
    '
    If Len(Trim(strParametro)) < 1 Then
       '
       MsgBox "Não pode encontrar registro com Chave vazia:" & Trim(strParametro), vbOKOnly + vbCritical, App.Title
       '
       Exit Sub
       '
    End If
    '
    connOpen
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    '
    Select Case valTipoParamentro
        Case 1
            '
            rs.Open "SELECT * FROM clientes WHERE codigo_cliente='" & strParametro & "';", CONN, adOpenFowardOnly, adLockReadOnly
            '
        Case 2
            '
            rs.Open "SELECT * FROM clientes WHERE nome_fantasia='" & strParametro & "';", CONN, adOpenFowardOnly, adLockReadOnly
            '
        Case 3
            '
            rs.Open "SELECT * FROM clientes WHERE razao_social='" & strParametro & "';", CONN, adOpenFowardOnly, adLockReadOnly
            '
    End Select
    '
    ' MsgBox "Encontrar:" & Trim(strParametro) & " Encontrado:" & Trim(rs("nome_fantasia")), vbOKOnly + vbCritical, App.Title
    '
    If rs.BOF Or rs.EOF Then
       '
       MsgBox "Registro Inválido.", vbOKOnly + vbCritical, App.Title
       '
       If rs.State = 1 Then rs.Close
       '
       connClose
       '
       Set rs = Nothing
       '
       Screen.MousePointer = 0
       '
       Exit Sub
       '
    End If
    '
    If rs.RecordCount < 0 Then
       '
       MsgBox "Registro inexistente:" & Trim(strParametro), vbOKOnly + vbCritical, App.Title
       '
       If rs.State = 1 Then rs.Close
       '
       connClose
       '
       Set rs = Nothing
       '
       Screen.MousePointer = 0
       '
       Exit Sub
       '
    End If
    '
    If rs.RecordCount > 1 Then
       '
       MsgBox "Existe mais de um registro com esta Chave:" & Trim(strParametro), vbOKOnly + vbCritical, App.Title
       '
       If rs.State = 1 Then rs.Close
       '
       connClose
       '
       Set rs = Nothing
       '
       Screen.MousePointer = 0
       '
       Exit Sub
       '
    End If
    '
    ' MsgBox "Encontrado:" & Trim(rs("codigo_cliente")), vbOKOnly + vbCritical, App.Title
    '
    '============================================================================
    '
    frmClientes.cboCodigo.Clear
    '
    frmClientes.cboCodigo.AddItem strParametro
    '
    frmClientes.cboCodigo.ListIndex = 0
    '
    '============================================================================
    '
    strCodigo = rs("codigo_cliente")
    '
    frmClientes.cboCodigo.Text = strCodigo
    '
    frmClientes.cboFantasia.Text = rs("nome_fantasia")
    frmClientes.cboRSocial.Text = rs("razao_social")
    '
    frmClientes.txtEntrega.Text = rs("endereco_entrega")
    frmClientes.txtCEP.Text = rs("cep_entrega")
    frmClientes.txtBairro.Text = rs("bairro_entrega")
    frmClientes.txtCidade.Text = rs("cidade_entrega")
    '
    frmClientes.txtCobranca.Text = rs("endereco_cobranca")
    frmClientes.txtCobrancaCEP.Text = rs("cep_cobranca")
    frmClientes.txtCobrancaBairro.Text = rs("bairro_cobranca")
    frmClientes.txtCobrancaCidade.Text = rs("cidade_cobranca")
    frmClientes.txtTelefone.Text = rs("telefone")
    frmClientes.txtFax.Text = rs("fax")
    frmClientes.txtEmail.Text = rs("email")
    frmClientes.txtWWW.Text = rs("www")
    '
    frmClientes.txtDataFundacao.Text = Left(rs("data_fundacao"), 2) & "/" & Mid(rs("data_fundacao"), 3, 2) & "/" & Right(rs("data_fundacao"), 4)
    '
    If rs("predio_proprio") = "N" Then
       frmClientes.optNao = True
    Else
       frmClientes.optSim = True
    End If
    '
    frmClientes.txtRefBanc01.Text = rs("referencia_bancaria_1")
    frmClientes.txtRefBanc02.Text = rs("referencia_bancaria_2")
    frmClientes.txtRefPess01.Text = rs("referencia_comercial_1")
    frmClientes.txtRefPess02.Text = rs("referencia_comercial_2")
    frmClientes.txtContato.Text = rs("contato")
    frmClientes.txtCNPJMF.Text = rs("cnpjmf")
    frmClientes.txtIEST.Text = rs("incricao_estadual")
    frmClientes.txtCPF.Text = rs("cpf")
    frmClientes.txtRG.Text = rs("rg")
    frmClientes.txtAtividade.Text = rs("ramo_atividade")
    '
    ' Euclides - Coloca nos Combos a descrição da cidade.
    '
    If IsNumeric(frmClientes.txtCobrancaCidade.Text) Then
       frmClientes.txtCobrancaCidade.Text = RetornaCodigoCidade(frmClientes.txtCobrancaCidade.Text, 2)
       ' frmClientes.txtCidade.Text = RetornaCodigoCidade(frmClientes.txtCidade.Text, 2)
    Else
       frmClientes.txtCobrancaCidade.Text = RetornaCodigoCidade(frmClientes.txtCobrancaCidade.Text, 1)
       'frmClientes.txtCidade.Text = RetornaCodigoCidade(frmClientes.txtCidade.Text, 1)
    End If
    '
    If IsNumeric(frmClientes.txtCidade.Text) Then
       'frmClientes.txtCobrancaCidade.Text = RetornaCodigoCidade(frmClientes.txtCobrancaCidade.Text, 2)
       frmClientes.txtCidade.Text = RetornaCodigoCidade(frmClientes.txtCidade.Text, 2)
    Else
       'frmClientes.txtCobrancaCidade.Text = RetornaCodigoCidade(frmClientes.txtCobrancaCidade.Text, 1)
       frmClientes.txtCidade.Text = RetornaCodigoCidade(frmClientes.txtCidade.Text, 1)
    End If
    '
    frmClientes.txtAtividade.Text = RetornaAtividade(Trim(frmClientes.txtAtividade.Text), 2)
    frmClientes.cboAtividade.Text = frmClientes.txtAtividade.Text
    '
    If rs.State = 1 Then rs.Close
    '
    connClose
    '
    Set rs = Nothing
    '
    connOpen
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    '
    rs.Open "SELECT * FROM observacoes_clientes WHERE cliente='" & Trim(strCodigo) & "';", CONN, adOpenFowardOnly, adLockReadOnly
    '
    ' MsgBox "Encontrou " & CInt(rs.RecordCount) & " Registros: " & Trim(strCodigo), vbOKOnly + vbCritical, App.Title
    '
    If rs.RecordCount > 0 Then
       '
       frmClientes.txtObs.Text = rs("observacao")
       '
    Else
       '
       frmClientes.txtObs.Text = "-"
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
End Sub

Public Sub ExcluiCliente()
    '
    Dim rs
    Dim strCodigo As String
    '
    If Box("Deseja realmente excluir esse cliente ?", vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
    '
    If Len(Trim(frmClientes.cboCodigo.Text)) <= 0 Then
       '
       MsgBox "É necessário selecionar um cliente para executar essa operação.", vbOKOnly + vbCritical, App.Title
       Exit Sub
       '
    End If
    '
    Screen.MousePointer = 11
    '
    connOpen
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    '
    rs.Open "SELECT * FROM clientes WHERE codigo_cliente='" & Trim(frmClientes.cboCodigo.Text) & "';", CONN, adOpenDynamic, adLockOptimistic
    '
    strCodigo = rs("codigo_cliente")
    '
    rs.Delete
    '
    If rs.State = 1 Then rs.Close
    '
    rs.Open "SELECT * FROM observacoes_clientes WHERE cliente='" & Trim(strCodigo) & "';", CONN, adOpenDynamic, adLockOptimistic
    '
    If rs.RecordCount > 0 Then rs.Delete
    '
    If rs.State = 1 Then rs.Close
    '
    connClose
    '
    Set rs = Nothing
    '
    connOpen
    '
    EncheComboClientes frmClientes.cboCodigo, 1
    '
    EncheComboClientes frmClientes.cboFantasia, 2
    '
    EncheComboClientes frmClientes.cboRSocial, 3
    '
    connClose
    '
    LimpaControlesCliente opcExclusao
    '
    LimpaControlesCliente opcConsulta
    '
    If Err.Number <> 0 Then
       MsgBox "Ocorreu um erro ao tentar excluir.", vbOKOnly + vbCritical, App.Title
    Else
       MsgBox "Cliente excluído com sucesso !!!", vbOKOnly + vbInformation, App.Title
    End If
    '
    Screen.MousePointer = 0
    '
End Sub

Public Function EditarCliente(ByVal strCodigo As String) As Boolean
  '
  Dim rs
  '
  ' Dim strCodigo As String
  '
  Dim strDataFundacao As String
  Dim strPredioProprio As String
  '
  Dim mCidadeEndereco As String
  Dim mCidadeCobranca As String
  Dim mAtividade As String
  '
  mCidadeEndereco = RetornaCodigoCidade(frmClientes.cboCidadeEndereco.Text, 1)
  mCidadeCobranca = RetornaCodigoCidade(frmClientes.cboCidadeCobranca.Text, 1)
  mAtividade = RetornaAtividade(Trim(frmClientes.cboAtividade.Text), 1)
  '
  Screen.MousePointer = 11
  '
  connOpen
  '
  Set rs = CreateObject("ADOCE.Recordset.3.0")
  strCodigo = frmClientes.cboCodigo.Text
  '
  rs.Open "SELECT * FROM clientes WHERE codigo_cliente='" & Trim(strCodigo) & "';", CONN, adOpenDynamic, adLockOptimistic
  '
  If rs.RecordCount > 0 Then
     '
     If rs("status") = "D" Then mStatusCliente = "D"
     '
     rs("nome_fantasia") = frmClientes.txtFantasia.Text
     rs("razao_social") = frmClientes.txtRazaoSocial.Text
     rs("endereco_entrega") = frmClientes.txtEntrega.Text
     rs("cep_entrega") = frmClientes.txtCEP.Text
     rs("bairro_entrega") = frmClientes.txtBairro.Text
     '
     rs("cidade_entrega") = mCidadeEndereco
     '
     rs("endereco_cobranca") = frmClientes.txtCobranca.Text
     rs("cep_cobranca") = frmClientes.txtCobrancaCEP.Text
     rs("bairro_cobranca") = frmClientes.txtCobrancaBairro.Text
     '
     rs("cidade_cobranca") = mCidadeCobranca
     '
     rs("telefone") = frmClientes.txtTelefone.Text
     rs("fax") = frmClientes.txtFax.Text
     rs("email") = frmClientes.txtEmail.Text
     rs("www") = frmClientes.txtWWW.Text
     '
     If IsDate(Trim(frmClientes.txtDataFundacao.Text)) Then
        '
        strDataFundacao = RetornaDataString(CDate(frmClientes.txtDataFundacao.Text))
        '
     Else
        '
        MsgBox "(4) - Data de Fundação não é uma Data Válida:(" & frmClientes.txtDataFundacao.Text & ")", vbOKOnly + vbCritical, App.Title
        '
        strDataFundacao = ""
        '
     End If
     '
     rs("data_fundacao") = strDataFundacao
     '
     If frmClientes.optNao = True Then
        '
        strPredioProprio = "N"
        '
     Else
        '
        strPredioProprio = "S"
        '
     End If
     '
     rs("predio_proprio") = strPredioProprio
     '
     rs("referencia_bancaria_1") = frmClientes.txtRefBanc01.Text
     rs("referencia_bancaria_2") = frmClientes.txtRefBanc02.Text
     rs("referencia_comercial_1") = frmClientes.txtRefPess01.Text
     rs("referencia_comercial_2") = frmClientes.txtRefPess02.Text
     rs("contato") = frmClientes.txtContato.Text
     rs("cnpjmf") = frmClientes.txtCNPJMF.Text
     rs("incricao_estadual") = frmClientes.txtIEST.Text
     rs("cpf") = frmClientes.txtCPF.Text
     rs("rg") = frmClientes.txtRG.Text
     '
     rs("status") = mStatusCliente
     '
     rs("ramo_atividade") = mAtividade
     '
     rs.Update
     '
     If rs.State = 1 Then rs.Close
     '
     rs.Open "SELECT * FROM observacoes_clientes WHERE cliente='" & Trim(strCodigo) & "';", CONN, adOpenDynamic, adLockOptimistic
     '
     If rs.RecordCount > 0 Then
        rs("observacao") = frmClientes.txtObs.Text
     Else
        '
        If rs.State = 1 Then rs.Close
        '
        rs.Open "SELECT * FROM observacoes_clientes;", CONN, adOpenDynamic, adLockOptimistic
        '
        rs.AddNew
        '
        Select Case Len(strCodigo)
        Case 1
             rs("cliente") = "0000" & Trim(strCodigo)
        Case 2
             rs("cliente") = "000" & Trim(strCodigo)
        Case 3
             rs("cliente") = "00" & Trim(strCodigo)
        Case 4
             rs("cliente") = "0" & Trim(strCodigo)
        Case 5
             rs("cliente") = Trim(strCodigo)
        Case Else
             rs("cliente") = Left(Trim(strCodigo), 5)
        End Select
        '
        rs("observacao") = frmClientes.txtObs.Text
        '
     End If
     '
     rs.Update
     '
  Else
     '
     MsgBox "Não foi possível alterar este cliente no momento pois não foi localizado na base de dados. Existe algum erro em sua base de dados.", vbOKOnly + vbCritical, App.Title
     '
  End If
  '
  If rs.State = 1 Then rs.Close
  '
  Set rs = Nothing
  '
  connClose
  '
  If Err.Number <> 0 Then
     '
     EditarCliente = False
     '
  Else
     '
     EditarCliente = True
     '
  End If
  '
  Screen.MousePointer = 0
  '
End Function
'
'
'
Public Function ValidaDataCliente(ByVal strDate As String) As Boolean
   '
   Dim msai As Boolean
   '
   If (Trim(strDate) = "") Or (Trim(strDate) = "**/**/****") Or (Trim(strDate) = "//") Then
      '
      msai = True
      '
   Else
      '
      If msai = True And Len(Trim(strDate)) <> 10 Then
         msai = False
      Else
         If msai = True And Mid(strDate, 3, 1) <> "/" Then
            msai = False
         Else
            If msai = True And Mid(strDate, 6, 1) <> "/" Then
               msai = False
            Else
               If msai = True And (IsNumeric(Mid(strDate, 1, 2))) = False Then
                  msai = False
               Else
                  If msai = True And (IsNumeric(Mid(strDate, 4, 2))) = False Then
                     msai = False
                  Else
                     If msai = True And (IsNumeric(Mid(strDate, 7, 4))) = False Then
                        msai = False
                     Else
                        If msai = True And CInt(Mid(strDate, 1, 2)) > 31 Then
                           msai = False
                        Else
                           If msai = True And CInt(Mid(strDate, 4, 2)) > 12 Then
                              msai = False
                           Else
                              If msai = True And CInt(Mid(strDate, 7, 4)) > 2005 Then
                                 msai = False
                              Else
                                 If msai = True And (Not IsDate(strDate)) Then
                                    msai = False
                                 Else
                                    msai = True
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
      '
   End If
   '
   ValidaDataCliente = msai
   '
End Function
