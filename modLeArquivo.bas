Attribute VB_Name = "modLeArquivo"
Option Explicit
Dim rs
Dim strMensagem As String
Dim strRelatorio As String
'
Dim strObsCliente As String
Dim mCodigoCliente As String
Dim strLinha As String
Dim strObservacaoPed As String
Dim strRosto As String
'
Dim bolDescontoCanal As Boolean
Dim bolTitulosAberto As Boolean
Dim bolSistematica As Boolean
Dim bolPromocoes As Boolean

Public strNomeArquivo As String

Public Sub PegaLinha(ByVal FileName As String, ByVal intProgress As Integer)
    '
    Screen.MousePointer = 0
    '
    ' If MsgBox("IntProgress = " & CInt(IntProgress), vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
    '
    If intProgress = 1 Then
       '
       frmStart.SIPVisible = False
       '
       frmStart.Frame1.Visible = True
       '
       frmStart.PictureBox1.Refresh
       '
       frmStart.Refresh
       '
    End If
    '
    If intProgress = 2 Then
      '
      frmGerenciador.Frame1.Visible = True
      '
      frmGerenciador.PictureBox1.Refresh
       '
    End If
    '
    If intProgress = 1 Then
       '
       '
    End If
    '
    Dim strDia As String
    Dim strMes As String
    Dim strFiller As String
    '
    Dim DblPercentagem As Double
    '
    Dim bolMudaPasso As Boolean
    '
    Dim intTamanho As Integer
    Dim intRegistros As Integer
    Dim intContador As Integer
    Dim intLidos As Integer
    Dim intPasso As Integer
    Dim intComprimento As Integer
    Dim intPassagem As Integer
    Dim intAsomar As Integer
    '
    Dim intTipoArquivo As Integer
    Dim strTipoArquivo As String
    '
    Dim strNeutro As String
    '
    bolDescontoCanal = False
    bolTitulosAberto = False
    bolSistematica = False
    bolPromocoes = False
    '
    ' Pega ultima data de acesso e atualiza o arquivo onde ela fica acumulada.
    '
    frmLogin.File.Open strPath & "\last.txt", fsModeOutput, fsAccessWrite
    '
    strDia = Trim(Day(Now))
    '
    If Len(Trim(strDia)) = 1 Then strDia = "0" & strDia
    '
    strMes = Trim(Month(Now))
    '
    If Len(Trim(strMes)) = 1 Then strMes = "0" & strMes
    '
    frmLogin.File.LinePrint strDia & strMes & Trim(Year(Now))
    '
    frmLogin.File.Close
    '
    ' Abre arquivo texto para ler os dados.
    '
    frmLogin.File.Open FileName, fsModeInput, fsAccessRead, fsLockRead
    '
    connOpen
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    '
    intRegistros = 1
    '
    ' Conta número de linhas
    '
    Do
      '
      strLinha = frmLogin.File.LineInputString
      '
      If Left(strLinha, 2) = "99" Or intRegistros > 10001 Then Exit Do
      '
      intRegistros = intRegistros + 1
      '
    Loop
    '
    frmLogin.File.Close
    '
    Set rs = Nothing
    '
    ' If MsgBox("Número de Itens:" & CStr(intRegistros), vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
    '
    frmLogin.File.Open FileName, fsModeInput, fsAccessRead, fsLockRead
    '
    connOpen
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    '
    strMensagem = vbNullString
    strRelatorio = vbNullString
    '
    strFiller = ""
    intTamanho = 3400
    '
    If (intRegistros Mod 100) <> 0 Then
       '
       intPasso = CInt(intRegistros / 100) + 1
       '
       bolMudaPasso = True
       '
       intAsomar = CInt(((intRegistros Mod 100)) / intPasso) + 1
       '
    Else
       '
       intPasso = CInt(intRegistros / 100)
       '
       bolMudaPasso = False
       '
       intAsomar = 0
       '
    End If
    '
    ' If MsgBox("Passo:" & CStr(intAsomar), vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
    '
    intPassagem = 1
    intComprimento = 0
    intContador = 0
    intLidos = 0
    '
    strLinha = ""
    strObsCliente = ""
    '
    Do
        '
        If Trim(strLinha) = "" Then strLinha = frmLogin.File.LineInputString
        '
        intContador = intContador + 1
        '
        intLidos = intLidos + 1
        '
        If (intContador Mod intPasso) = 0 Then
           '
           intComprimento = intComprimento + (intTamanho / 100)
           '
           intPassagem = intPassagem + 1
           '
           Progresso intProgress, intComprimento, intTamanho, intPassagem
           '
        End If
        '
        If intContador > 10000 Then
           '
           Exit Do
           '
        End If
        '
        Select Case Left(strLinha, 2)
                '
            Case "99"
                '
                AddStatus "Manutenção concluída com sucesso"
                '
                Progresso intProgress, intTamanho, intTamanho, 100
                '
                Exit Do
                '
            Case "00"
                '
                strLinha = ""
                strFiller = "00"
                '
            Case "01"
                '
                '==========================================================================
                '
                '                        APAGA E RECRIA O BANCO DE DADOS
                '
                '==========================================================================
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Excluindo Banco de Dados"
                   '
                   connClose
                   '
                   Set rs = Nothing
                   '
                   If frmStart.FileSystem.Dir(strPath & "\base.cdb") <> "" Then
                      '
                      frmStart.FileSystem.Kill strPath & "\base.cdb"
                      '
                   End If
                   '
                   CriarBase
                   '
                   connOpen
                   '
                   Set rs = CreateObject("ADOCE.Recordset.3.0")
                   '
                End If
                '
                strLinha = ""
                strFiller = "01"
                '
            Case "08" ' 01
                '
                '==========================================================================
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Excluindo dados das tabelas"
                   '
                   If UCase(Mid(strLinha, 1, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 2, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 3, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 4, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 5, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 6, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 7, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 8, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 9, 1)) = "Z" Then strNeutro = ""
                   '
                   If UCase(Mid(strLinha, 10, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 11, 1)) = "Z" Then ExecSQL "DELETE * FROM Vendedor;"
                   If UCase(Mid(strLinha, 12, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 13, 1)) = "Z" Then ExecSQL "DELETE * FROM forma_pagamento;"
                   If UCase(Mid(strLinha, 14, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 15, 1)) = "Z" Then ExecSQL "DELETE * FROM condicao_pagamento;"
                   If UCase(Mid(strLinha, 16, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 17, 1)) = "Z" Then ExecSQL "DELETE * FROM ramo_atividade;"
                   If UCase(Mid(strLinha, 18, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 19, 1)) = "Z" Then ExecSQL "DELETE * FROM cidades;"
                   '
                   If UCase(Mid(strLinha, 20, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 21, 1)) = "Z" Then ExecSQL "DELETE * FROM estoque;"
                   If UCase(Mid(strLinha, 22, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 23, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 24, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 25, 1)) = "Z" Then ExecSQL "DELETE * FROM destinatarios;"
                   If UCase(Mid(strLinha, 26, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 27, 1)) = "Z" Then ExecSQL "DELETE * FROM clientes;"
                   If UCase(Mid(strLinha, 28, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 29, 1)) = "Z" Then ExecSQL "DELETE * FROM observacoes_clientes;"
                   '
                   If UCase(Mid(strLinha, 30, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 31, 1)) = "Z" Then ExecSQL "DELETE * FROM sistematica_visita;"
                   If UCase(Mid(strLinha, 32, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 33, 1)) = "Z" Then ExecSQL "DELETE * FROM fabricante;"
                   If UCase(Mid(strLinha, 34, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 35, 1)) = "Z" Then ExecSQL "DELETE * FROM brand;"
                   If UCase(Mid(strLinha, 36, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 37, 1)) = "Z" Then ExecSQL "DELETE * FROM sub_brand;"
                   If UCase(Mid(strLinha, 38, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 39, 1)) = "Z" Then ExecSQL "DELETE * FROM produtos;"
                   '
                   If UCase(Mid(strLinha, 40, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 41, 1)) = "Z" Then ExecSQL "DELETE * FROM objetivo_venda;"
                   If UCase(Mid(strLinha, 42, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 43, 1)) = "Z" Then ExecSQL "DELETE * FROM tipo_movimento;"
                   If UCase(Mid(strLinha, 44, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 45, 1)) = "Z" Then ExecSQL "DELETE * FROM promocoes;"
                   If UCase(Mid(strLinha, 46, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 47, 1)) = "Z" Then ExecSQL "DELETE * FROM descricao_tabela;"
                   If UCase(Mid(strLinha, 48, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 49, 1)) = "Z" Then ExecSQL "DELETE * FROM tabela_precos;"
                   '
                   If UCase(Mid(strLinha, 50, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 51, 1)) = "Z" Then ExecSQL "DELETE * FROM pedidos;"
                   If UCase(Mid(strLinha, 52, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 53, 1)) = "Z" Then ExecSQL "DELETE * FROM itens_pedido;"
                   If UCase(Mid(strLinha, 54, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 55, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 56, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 57, 1)) = "Z" Then ExecSQL "DELETE * FROM titulos_aberto;"
                   If UCase(Mid(strLinha, 58, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 59, 1)) = "Z" Then ExecSQL "DELETE * FROM justificativa_nao_venda;"
                   '
                   If UCase(Mid(strLinha, 60, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 61, 1)) = "Z" Then ExecSQL "DELETE * FROM desconto_canal;"
                   If UCase(Mid(strLinha, 62, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 63, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 64, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 65, 1)) = "Z" Then ExecSQL "DELETE * FROM mensagens;"
                   If UCase(Mid(strLinha, 66, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 67, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 68, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 69, 1)) = "Z" Then ExecSQL "DELETE * FROM relatorios;"
                   '
                   If UCase(Mid(strLinha, 70, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 71, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 72, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 73, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 74, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 75, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 76, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 77, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 78, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 79, 1)) = "Z" Then strNeutro = ""
                   '
                   If UCase(Mid(strLinha, 80, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 81, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 82, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 83, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 84, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 85, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 86, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 87, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 88, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 89, 1)) = "Z" Then strNeutro = ""
                   '
                   If UCase(Mid(strLinha, 90, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 91, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 92, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 93, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 94, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 95, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 96, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 97, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 98, 1)) = "Z" Then strNeutro = ""
                   If UCase(Mid(strLinha, 99, 1)) = "Z" Then strNeutro = ""
                   '
                End If
                '
                '==========================================================================
                '
                strLinha = ""
                strFiller = "08"
                '
            Case "11"
                '
                ' AddStatus Left(strLinha, 25) & strFiller
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo/Alterando Tabela Vendedores"
                   '
                   'If rs.State = 1 Then rs.Close
                   '
                   rs.Open "SELECT * FROM vendedor WHERE codigo_vendedor='" & Trim(Mid(strLinha, 3, 5)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   '
                   If rs.RecordCount <= 0 Then
                      '
                      rs.AddNew
                      '
                   End If
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                LeArquivoIncluiVendedor strLinha
                '
                strLinha = ""
                '
            Case "13"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Formas de Pagamento"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM forma_pagamento WHERE codigo='" & Trim(Mid(strLinha, 3, 2)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiFormaPgto strLinha
                '
                strLinha = ""
                '
            Case "15"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Condições de Pagamento"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM condicao_pagamento WHERE codigo_condicao='" & Trim(Mid(strLinha, 3, 2)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiCondicaoPgto strLinha
                '
                strLinha = ""
                '
            Case "17"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Ramos de Atividade"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM ramo_atividade WHERE codigo='" & Trim(Mid(strLinha, 3, 3)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiRamoAtiv strLinha
                '
                strLinha = ""
                '
            Case "19"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Cidades"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM cidades WHERE codigo='" & Trim(Mid(strLinha, 3, 5)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiCidades strLinha
                '
                strLinha = ""
                '
            Case "21"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Estoque"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM estoque where codigo_cliente=" & Mid(strLinha, 3, 5) & " and codigo_produto=" & Mid(strLinha, 8, 6) & ";", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiEstoque strLinha
                '
                strLinha = ""
                '
            Case "23"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Históricos"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM historico where codigo_cliente=" & Mid(strLinha, 3, 5) & " and codigo_produto=" & Mid(strLinha, 8, 6) & " and data=" & Mid(strLinha, 18, 4) & Mid(strLinha, 16, 2) & Mid(strLinha, 14, 2) & ";", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiHistorico strLinha
                '
                strLinha = ""
                '
            Case "25"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   rs.Close
                   '
                   AddStatus "Incluindo Tabela de Destinatarios"
                   '
                   rs.Open "SELECT * FROM Destinatarios WHERE codigo_destinatario='" & Trim(Mid(strLinha, 3, 3)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiDestinatario strLinha
                '
                strLinha = ""
                '
            Case "27"
                '
                strObsCliente = ""
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Clientes"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM clientes WHERE codigo_cliente='" & Trim(Mid(strLinha, 8, 5)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiCliente strLinha
                '
                strLinha = ""
                '
            Case "29"
                '
                ' Left(strLinha, 2)
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo na Tabela de Observações de Clientes"
                   '
                   mCodigoCliente = Trim(Mid(strLinha, 3, 5))
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM observacoes_clientes WHERE cliente='" & Trim(mCodigoCliente) & "';", CONN, adOpenDynamic, adLockOptimistic
                   '
                   If rs.RecordCount <= 0 Then
                      '
                      rs.AddNew
                      '
                      rs("cliente") = Trim(mCodigoCliente)
                      '
                   End If
                   '
                   strObsCliente = ""
                   '
                   While Left(strLinha, 2) = "29"
                      '
                      ' LeArquivoIncluiObsCliente strLinha
                      '
                      ' Sub LeArquivoIncluiObsCliente(ByVal strLinha As String)
                      '
                      strObsCliente = strObsCliente & Mid(strLinha, 10, cstLarguraObsCliente)
                      '
                      ' End Sub
                      '
                      intContador = intContador + 1
                      '
                      intLidos = intLidos + 1
                      '
                      If (intContador Mod intPasso) = 0 Then
                         '
                         intComprimento = intComprimento + (intTamanho / 100)
                         '
                         intPassagem = intPassagem + 1
                         '
                         Progresso intProgress, intComprimento, intTamanho, intPassagem
                         '
                      End If
                      '
                      If intContador > 10000 Then
                         '
                         Exit Do
                         '
                      End If
                      '
                      strLinha = frmLogin.File.LineInputString
                      '
                   Wend
                   '
                   If Trim(strObsCliente) <> "" Then
                      '
                      Do While InStr(1, strObsCliente, "@%", vbTextCompare) > 0
                         '
                         strObsCliente = Mid(strObsCliente, 1, InStr(1, strObsCliente, "@%", vbTextCompare) - 1) & Chr(13) & Chr(10) _
                         & Mid(strObsCliente, InStr(1, strObsCliente, "@%", vbTextCompare) + 2, Len(strObsCliente) - InStr(1, strObsCliente, "@%", vbTextCompare) + 1)
                         '
                         ' MsgBox "mlinhaObs(*): *" & mObservacao & "* len=" & Len(mObservacao), vbOKOnly + vbCritical, App.Title
                         '
                      Loop
                      '
                      rs("observacao") = strObsCliente
                      '
                      rs.Update
                      '
                   End If
                   '
                   mCodigoCliente = ""
                   '
                   strObsCliente = ""
                   '
                   strFiller = "29"
                   '
                End If
                '
            Case "31"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Sistematica de Visita"
                   '
                   'If bolSistematica = False Then
                   '   '
                   '   ExecSQL "DROP TABLE sistematica_visita;"
                   '   ExecSQL "CREATE TABLE sistematica_visita(codigo_vendedor VARCHAR(5), codigo_cliente VARCHAR(5), dia VARCHAR(1), numero_visita VARCHAR(3)  );"
                   '   '
                   '   AddStatus "Sistemática OK"
                   '   '
                   bolSistematica = True
                   '   '
                   'End If
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM sistematica_visita;", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiSistematica strLinha
                '
                strLinha = ""
                '
            Case "33"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Fabricantes"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM fabricante WHERE codigo='" & Trim(Mid(strLinha, 3, 3)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiFabricante strLinha
                '
                strLinha = ""
                '
            Case "35"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Brand"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM brand WHERE codigo='" & Trim(Mid(strLinha, 3, 3)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiBrand strLinha
                '
                strLinha = ""
                '
            Case "37"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Sub Brand"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM sub_brand WHERE codigo='" & Trim(Mid(strLinha, 3, 3)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluisubbrand strLinha                   '
                '
                strLinha = ""
                '
            Case "39"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Produtos"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM produtos WHERE codigo_produto='" & Trim(Mid(strLinha, 12, 6)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiProduto strLinha
                '
                strLinha = ""
                '
            Case "41"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Objetivos de Venda"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM objetivo_venda WHERE codigo_vendedor='" & Trim(Mid(strLinha, 3, 5)) & "' AND codigo_produto='" & Trim(Mid(strLinha, 8, 6)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiObjetivo strLinha
                '
                strLinha = ""
                '
            Case "43"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Tipos de Movimento"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM tipo_movimento WHERE codigo='" & Trim(Mid(strLinha, 3, 1)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiTipoMovimento strLinha
                '
                strLinha = ""
                '
            Case "45"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Promoções"
                   '
                   'If bolPromocoes = False Then
                   '   '
                   '   ExecSQL "DROP TABLE promocoes;"
                   '   ExecSQL "CREATE TABLE promocoes(codigo_produto VARCHAR(6), qtd_inicial VARCHAR(8), qtd_final VARCHAR(8), preco_promocional VARCHAR(10)  );"
                   '   '
                   '   AddStatus "Promoções OK"
                   '   '
                   bolPromocoes = True
                   '   '
                   'End If
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM promocoes WHERE codigo_produto='" & Trim(Mid(strLinha, 3, 6)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiPromocao strLinha
                '
                strLinha = ""
                '
            Case "47"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Descrições de Tabelas de Preços"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM descricao_tabela WHERE codigo_tabela='" & Trim(Mid(strLinha, 3, 1)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiDescricaoTabela strLinha
                '
                strLinha = ""
                '
            Case "49"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Precos"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM tabela_precos WHERE tabela_precos='" & Trim(Mid(strLinha, 3, 2)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiTabelaPrecos strLinha
                '
                strLinha = ""
                '
            Case "51"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Atualizando Tabelas de Pedidos"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM pedido WHERE numero_pedido_interno='" & Trim(Mid(strLinha, 13, 6)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
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
                rs.AddNew
                '
                LeArquivoIncluiPedido strLinha
                '
                strLinha = ""
                '
            Case "53"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Atualizando Tabela de Itens de Pedidos"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM itens_pedido WHERE numero_pedido='" & Trim(Mid(strLinha, 3, 6)) & "' AND codigo_produto='" & Trim(Mid(strLinha, 9, 6)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiItenPedido strLinha
                '
                strLinha = ""
                '
            Case "55"
                '
                '================== Tabela de Observações de Pedidos
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Atualizando Tabela de Observações de Pedidos"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM pedido WHERE numero_pedido_interno='" & Trim(Mid(strLinha, 8, 6)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   '
                End If
                '
                strObservacaoPed = Trim(Mid(strLinha, 16, 80))
                '
                Do While InStr(1, strObservacaoPed, "@%", vbTextCompare) > 0
                   '
                   strObservacaoPed = Mid(strObservacaoPed, 1, InStr(1, strObservacaoPed, "@%", vbTextCompare) - 1) & Chr(13) & Chr(10) _
                   & Mid(strObservacaoPed, InStr(1, strObservacaoPed, "@%", vbTextCompare) + 2, Len(strObservacaoPed) - InStr(1, strObservacaoPed, "@%", vbTextCompare) + 1)
                   '
                   ' MsgBox "mlinhaObs(*): *" & mObservacao & "* len=" & Len(mObservacao), vbOKOnly + vbCritical, App.Title
                   '
                Loop
                '
                strFiller = "33"
                '
                If rs.RecordCount <= 0 Then
                   '
                   rs.AddNew
                   '
                End If
                '
                LeArquivoIncluiOBSPedido strLinha
                '
                strLinha = ""
                '
            Case "57"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Títulos em aberto"
                   '
                   'If bolTitulosAberto = False Then
                   '   '
                   '   ExecSQL "DROP TABLE titulos_aberto;"
                   '   ExecSQL "CREATE TABLE titulos_aberto(codigo_cliente VARCHAR(5), numero_documento VARCHAR(20), data_vencimento VARCHAR(10), valor VARCHAR(10)  );"
                   '   '
                   '   AddStatus "Titulos aberto OK"
                   '   '
                   bolTitulosAberto = True
                   '   '
                   'End If
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM titulos_aberto;", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiTituloAberto strLinha
                '
                strLinha = ""
                '
            Case "59"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Justificativas de Visita"
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM justificativa_nao_venda WHERE codigo='" & Trim(Mid(strLinha, 3, 3)) & "';", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiNaoVenda strLinha
                '
                strLinha = ""
                '
            Case "61"
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Incluindo Tabela de Descontos de Canal"
                   '
                   'If bolDescontoCanal = False Then
                   '   '
                   '   ExecSQL "DROP TABLE desconto_canal;"
                   '   ExecSQL "CREATE TABLE desconto_canal(sub_brand VARCHAR(3), cliente VARCHAR(5), desconto VARCHAR(6)  );"
                   '   '
                   '   AddStatus "Desconto Canal OK"
                   '   '
                   bolDescontoCanal = True
                   '
                   'End If
                   '
                   rs.Close
                   '
                   rs.Open "SELECT * FROM desconto_canal;", CONN, adOpenDynamic, adLockOptimistic
                   '
                   strFiller = Left(strLinha, 2)
                   '
                End If
                '
                rs.AddNew
                '
                LeArquivoIncluiDescontoCanal strLinha
                '
                strLinha = ""
                '
            Case "63"
                '
                AddStatus "Agrupando textos da Tabela de Mensagens"
                '
                strFiller = ""
                '
                LeArquivoIncluiCorpoMensagem strLinha
                '
                strLinha = ""
                '
            Case "65"
                '
                AddStatus "Incluindo Tabela de Mensagems"
                '
                strFiller = ""
                '
                LeArquivoIncluiMensagem strLinha
                '
                strLinha = ""
                '
            Case "66"
                '
                '================== Mensagem de Rosto
                '
                If Left(strLinha, 2) <> strFiller Then
                   '
                   AddStatus "Agrupando textos da Mensagem de Rosto"
                   '
                   strRosto = ""
                   '
                   While Left(strLinha, 2) = "66"
                      '
                      ' LeArquivoIncluiObsCliente strLinha
                      '
                      ' Sub LeArquivoIncluiObsCliente(ByVal strLinha As String)
                      '
                      strRosto = strRosto & Mid(strLinha, 3, cstLarguraRosto)
                      '
                      ' End Sub
                      '
                      intContador = intContador + 1
                      '
                      intLidos = intLidos + 1
                      '
                      If (intContador Mod intPasso) = 0 Then
                         '
                         intComprimento = intComprimento + (intTamanho / 100)
                         '
                         intPassagem = intPassagem + 1
                         '
                         Progresso intProgress, intComprimento, intTamanho, intPassagem
                         '
                      End If
                      '
                      If intContador > 10000 Then
                         '
                         Exit Do
                         '
                      End If
                      '
                      strLinha = frmLogin.File.LineInputString
                      '
                   Wend
                   '
                   If Trim(strRosto) <> "" Then
                      '
                      Do While InStr(1, strRosto, "@%", vbTextCompare) > 0
                         '
                         strRosto = Mid(strRosto, 1, InStr(1, strRosto, "@%", vbTextCompare) - 1) & Chr(13) & Chr(10) _
                         & Mid(strRosto, InStr(1, strRosto, "@%", vbTextCompare) + 2, Len(strRosto) - InStr(1, strRosto, "@%", vbTextCompare) + 1)
                         '
                      Loop
                      '
                      ' MsgBox "Rosto:" & strRosto, vbOKOnly + vbCritical, App.Title
                      '
                      rs.Close
                      '
                      rs.Open "SELECT * FROM vendedor WHERE codigo_vendedor='" & Trim(usrCodigoVendedor) & "';", CONN, adOpenDynamic, adLockOptimistic
                      '
                      If rs.RecordCount > 0 Then
                         '
                         rs("mensagem") = strRosto
                         '
                         rs.Update
                         '
                      End If
                      '
                   End If
                   '
                   strRosto = ""
                   '
                   strFiller = "66"
                   '
                End If
                '
            Case "67"
                '
                AddStatus "Agrupando Tabela de Textos de Relatórios"
                '
                strFiller = ""
                '
                LeArquivoIncluiCorpoRelatorio strLinha
                '
                strLinha = ""
                '
            Case "69"
                '
                AddStatus "Incluindo Tabela de Relatórios"
                '
                strFiller = ""
                '
                LeArquivoIncluiRelatorio strLinha
                '
                strLinha = ""
                '
        End Select
        '
    Loop
    '
    '
    ' If MsgBox("Número de Itens:" & CStr(intLidos), vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
    '
    rs.Close
    '
    Set rs = Nothing
    '
    ' vbTextCompare , vbBinaryCompare , vbDatabaseCompare
    '
    frmLogin.File.Close
    '
    If InStr(1, strNomeArquivo, "I0000000", vbTextCompare) <> 0 Then
       '
       For intTipoArquivo = 1 To 999
           '
           strTipoArquivo = Right("000" + Trim(intTipoArquivo), 3) & ".000"
           '
           ' If MsgBox("Arquivo:" & strPath & "\I0000" & strTipoArquivo & "*", vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
           '
           If frmLogin.FileSystem.Dir(strPath & "\I0000" & strTipoArquivo) = "" Then
              '
              ' Arquivo a receber o nome NÃO existe
              '
              ' If MsgBox("Nome do Arquivo:" & strNomeArquivo & "=> " & strPath & "\I0000" & strTipoArquivo, vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
              '
              frmLogin.FileSystem.MoveFile strNomeArquivo, strPath & "\I0000" & strTipoArquivo
              '
              Exit For
              '
           End If
           '
       Next
       '
    Else
       '
       If frmLogin.FileSystem.Dir(strPath & "\-R" & Right(RetornaDataString(Now), 2) & Mid(RetornaDataString(Now), 3, 2) & Left(RetornaDataString(Now), 2) & Right(strNomeArquivo, 5)) <> "" Then
          '
          ' Arquivo a receber o nome já existe
          '
          frmLogin.FileSystem.Kill strPath & "\-R" & Right(RetornaDataString(Now), 2) & Mid(RetornaDataString(Now), 3, 2) & Left(RetornaDataString(Now), 2) & Right(strNomeArquivo, 5)
          '
       End If
       '
       ' frmLogin.File.Close
       '
       ' If MsgBox("Nome do Arquivo:" & strNomeArquivo & "=>" & strPath & "\-R" & Right(RetornaDataString(Now), 2) & Mid(RetornaDataString(Now), 3, 2) & Left(RetornaDataString(Now), 2) & Right(strNomeArquivo, 5), vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
       '
       frmLogin.FileSystem.MoveFile strNomeArquivo, strPath & "\-R" & Right(RetornaDataString(Now), 2) & Mid(RetornaDataString(Now), 3, 2) & Left(RetornaDataString(Now), 2) & Right(strNomeArquivo, 5)
       '
    End If
    '
    connClose
    '
    Screen.MousePointer = 11
    '
End Sub
'
'Filler 01 - Leitura - Vendedor
'
Sub LeArquivoIncluiVendedor(ByVal strLinha As String)
    '
    Dim varDouble As Double
    '
    ' rs.Open "SELECT * FROM vendedor WHERE codigo_vendedor='" & Trim(Mid(strLinha, 3, 5)) & "';", CONN, adOpenDynamic, adLockOptimistic
    '
    usrCodigoVendedor = Trim(Mid(strLinha, 3, 5))
    '
    rs("codigo_vendedor") = usrCodigoVendedor
    '
    rs("senha") = Trim(Mid(strLinha, 8, 6))
    rs("nome") = Trim(LCase(Mid(strLinha, 14, 30)))
    rs("aceita_pedido_bloq") = Trim(Mid(strLinha, 44, 1))
    rs("contra_senha") = Trim(Mid(strLinha, 45, 1))
    rs("codigo_proximo_cliente") = Trim(Mid(strLinha, 46, 5))
    rs("extra1") = Trim(Mid(strLinha, 53, 1))
    rs("extra2") = Trim(Mid(strLinha, 54, 1))
    rs("numero_proximo_pedido") = Trim(Mid(strLinha, 55, 6))
    rs("tempo_maximo") = Trim(Mid(strLinha, 59, 3))
    rs("habilitar_desconto_item") = Trim(Mid(strLinha, 62, 1))
    rs("habilitar_desconto_pedido") = Trim(Mid(strLinha, 63, 1))
    rs("habilitar_edicao_preco") = Trim(Mid(strLinha, 64, 1))
    '
    rs("habilitar_acrescimo") = Trim(Mid(strLinha, 65, 1))
    '
    rs("habilitar_cobranca_titulo") = Trim(Mid(strLinha, 66, 1))
    rs("empresa") = Trim(Mid(strLinha, 67, 2))
    rs("filial") = Trim(Mid(strLinha, 69, 2))
    rs("data_cortetitulos") = Trim(Mid(strLinha, 71, 8))
    '
    varDouble = 0
    '
    If IsNumeric(Trim(Mid(strLinha, 79, 10))) Then
        varDouble = CDbl(Trim(Mid(strLinha, 79, 10))) / 100
    End If
    '
    rs("pedido_minimo") = CStr(varDouble)
    '
    '============ Pega mensagem de outro Filler
    '
    rs("mensagem") = ""
    rs("status") = LCase(Trim(Mid(strLinha, 89, 1)))
    '
    rs.Update
    '
End Sub
'
'Filler 21 - Leitura - Estoque
'
Sub LeArquivoIncluiEstoque(ByVal strLinha As String)
    '
    Dim varDouble As Double
    '
    ' estoque
    '
    ' codigo_cliente VARCHAR(5)
    ' codigo_produto VARCHAR(6)
    ' data VARCHAR(8)
    ' dias VARCHAR(3)
    ' media_diaria VARCHAR(12)
    ' media VARCHAR(8)
    ' estoque VARCHAR(8)
    ' sugestao VARCHAR(8)
    ' pedido VARCHAR(8)
    '
    rs("codigo_cliente") = Trim(Mid(strLinha, 3, 5))
    rs("codigo_produto") = Trim(Mid(strLinha, 8, 6))
    rs("data") = Trim(Mid(strLinha, 18, 4)) & Trim(Mid(strLinha, 16, 2)) & Trim(Mid(strLinha, 14, 2))
    rs("dias") = ""
    rs("media_diaria") = Trim(Mid(strLinha, 22, 12))
    rs("media") = ""
    rs("estoque") = Trim(Mid(strLinha, 34, 8))
    rs("sugestao") = Trim(Mid(strLinha, 42, 8))
    rs("pedido") = Trim(Mid(strLinha, 50, 8))
    '
    rs.Update
    '
End Sub
'
'Filler 23 - Leitura - Histórico
'
Sub LeArquivoIncluiHistorico(ByVal strLinha As String)
    '
    Dim varDouble As Double
    '
    ' Historico
    '
    ' codigo_cliente VARCHAR(5)
    ' codigo_produto VARCHAR(6)
    ' data VARCHAR(8)
    ' dias VARCHAR(3)
    ' media_diaria VARCHAR(12)
    ' media VARCHAR(8)
    ' estoque VARCHAR(8)
    ' sugestao VARCHAR(8)
    ' pedido VARCHAR(8)
    '
    rs("codigo_cliente") = Trim(Mid(strLinha, 3, 5))
    rs("codigo_produto") = Trim(Mid(strLinha, 8, 6))
    rs("data") = Trim(Mid(strLinha, 18, 4)) & Trim(Mid(strLinha, 16, 2)) & Trim(Mid(strLinha, 14, 2))
    rs("dias") = ""
    rs("media_diaria") = "000000000000" ' Trim(Mid(strLinha, 26, 12))
    rs("media") = ""
    rs("estoque") = "00000000" ' Trim(Mid(strLinha, 39, 8))
    rs("sugestao") = "00000000" ' Trim(Mid(strLinha, 48, 8))
    rs("pedido") = Trim(Mid(strLinha, 22, 8))
    '
    rs.Update
    '
End Sub
'
'Filler 03 - Leitura - Forma de Pagamento
'
Sub LeArquivoIncluiFormaPgto(ByVal strLinha As String)
    '
    rs("codigo") = Trim(Mid(strLinha, 3, 2))
    rs("descricao") = Trim(Mid(strLinha, 5, 15))
    rs("status") = LCase(Trim(Mid(strLinha, 20, 1)))
    '
    rs.Update
    '
End Sub
'
'Filler 05 - Leitura - Condição de Pagamento
'
Sub LeArquivoIncluiCondicaoPgto(ByVal strLinha As String)
    '
    Dim varDouble As Double
    '
    rs("codigo_condicao") = Trim(Mid(strLinha, 3, 2))
    rs("codigo_tabela_preco") = Trim(Mid(strLinha, 5, 1))
    rs("descricao") = Trim(Mid(strLinha, 6, 20))
    rs("dia_pgto_1p") = Trim(Mid(strLinha, 26, 2))
    rs("dia_pgto_2p") = Trim(Mid(strLinha, 28, 2))
    rs("dia_pgto_3p") = Trim(Mid(strLinha, 30, 2))
    rs("dia_pgto_4p") = Trim(Mid(strLinha, 32, 2))
    rs("dia_pgto_5p") = Trim(Mid(strLinha, 34, 2))
    '
    varDouble = 0
    '
    If IsNumeric(Trim(Mid(strLinha, 36, 10))) Then
       '
       varDouble = CDbl(Trim(Mid(strLinha, 36, 10))) / 100
       '
    End If
    '
    rs("pedido_minimo") = CStr(varDouble)
    '
    rs("status") = LCase(Trim(Mid(strLinha, 46, 1)))
    '
    rs.Update
    '
End Sub
'
'Filler 07 - Leitura - Ramo de Atividade
'
Sub LeArquivoIncluiRamoAtiv(ByVal strLinha As String)
    '
    rs("codigo") = Trim(Mid(strLinha, 3, 3))
    rs("descricao") = Trim(Mid(strLinha, 6, 30))
    rs("status") = LCase(Trim(Mid(strLinha, 36, 1)))
    '
    rs.Update
    '
End Sub
'
'Filler 08 - Leitura - Cidades
'
Sub LeArquivoIncluiCidades(ByVal strLinha As String)
    '
    rs("codigo") = Trim(Mid(strLinha, 3, 5))
    rs("nome") = Trim(Mid(strLinha, 8, 30))
    rs("uf") = LCase(Trim(Mid(strLinha, 38, 2)))
    '
    rs.Update
    '
End Sub
'
'Filler 09 - Leitura - Clientes
'
Sub LeArquivoIncluiCliente(ByVal strLinha As String)
    '
    ' rs("codigo_vendedor") = Trim(Mid(strLinha, 3, 5))
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
    rs("codigo_cliente") = Trim(Mid(strLinha, 8, 5))
    rs("nome_fantasia") = Trim(Mid(strLinha, 13, 20))
    rs("razao_social") = Trim(Mid(strLinha, 33, 40))
    rs("endereco_entrega") = Trim(Mid(strLinha, 73, 40))
    rs("cep_entrega") = Trim(Mid(strLinha, 113, 8))
    rs("bairro_entrega") = Trim(Mid(strLinha, 121, 20))
    rs("cidade_entrega") = Trim(Mid(strLinha, 141, 5))
    rs("endereco_cobranca") = Trim(Mid(strLinha, 146, 40))
    rs("cep_cobranca") = Trim(Mid(strLinha, 186, 8))
    rs("bairro_cobranca") = Trim(Mid(strLinha, 194, 20))
    rs("cidade_cobranca") = Trim(Mid(strLinha, 214, 5))
    rs("telefone") = Trim(Mid(strLinha, 219, 12))
    rs("fax") = Trim(Mid(strLinha, 231, 12))
    rs("email") = Trim(Mid(strLinha, 243, 30))
    rs("www") = Trim(Mid(strLinha, 273, 30))
    rs("data_fundacao") = Trim(Mid(strLinha, 303, 8))
    rs("predio_proprio") = Trim(Mid(strLinha, 311, 1))
    rs("referencia_bancaria_1") = Trim(Mid(strLinha, 312, 20))
    rs("referencia_bancaria_2") = Trim(Mid(strLinha, 332, 20))
    rs("referencia_comercial_1") = Trim(Mid(strLinha, 352, 20))
    rs("referencia_comercial_2") = Trim(Mid(strLinha, 372, 20))
    rs("contato") = Trim(Mid(strLinha, 392, 25))
    rs("data_ultima_compra") = Trim(Mid(strLinha, 417, 8))
    rs("valor_ultima_compra") = Trim(Mid(strLinha, 425, 10))
    rs("cnpjmf") = Trim(Mid(strLinha, 435, 14))
    rs("incricao_estadual") = Trim(Mid(strLinha, 449, 20))
    rs("cpf") = Trim(Mid(strLinha, 469, 14))
    rs("rg") = Trim(Mid(strLinha, 483, 20))
    rs("desconto_maximo") = Trim(Mid(strLinha, 503, 4))
    rs("limite_credito") = Trim(Mid(strLinha, 507, 10))
    rs("bloqueado") = Trim(Mid(strLinha, 517, 1))
    rs("condicao_pagamento_padrao") = Trim(Mid(strLinha, 518, 2))
    rs("forma_pagamento_padrao") = Trim(Mid(strLinha, 520, 2))
    rs("periodicidade_visita") = Trim(Mid(strLinha, 522, 2))
    rs("ramo_atividade") = Trim(Mid(strLinha, 524, 3))
    rs("status") = LCase(Trim(Mid(strLinha, 527, 1)))
    '
    rs.Update
    '
End Sub
'
'Filler 10 - Leitura - Destinatarios
'
Sub LeArquivoIncluiDestinatario(ByVal strLinha As String)
    '
    rs("codigo_destinatario") = Trim(Mid(strLinha, 3, 3))
    rs("nome") = Trim(Mid(strLinha, 6, 20))
    rs("status") = LCase(Trim(Mid(strLinha, 26, 1)))
    '
    rs.Update
    '
End Sub
'
'Filler 11 - Leitura - Sistematica de Visita
'
Sub LeArquivoIncluiSistematica(ByVal strLinha As String)
    '
    ' rs("codigo_vendedor") = Trim(Mid(strLinha, 3, 5))
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
    rs("codigo_cliente") = Trim(Mid(strLinha, 8, 5))
    rs("dia") = Trim(Mid(strLinha, 13, 1))
    rs("numero_visita") = Trim(Mid(strLinha, 14, 3))
    '
    rs.Update
    '
End Sub
'
'Filler 12 - Leitura - Fabricante
'
Sub LeArquivoIncluiFabricante(ByVal strLinha As String)
    '
    rs("codigo") = Trim(Mid(strLinha, 3, 3))
    rs("descricao") = Trim(Mid(strLinha, 6, 20))
    rs("status") = LCase(Trim(Mid(strLinha, 26, 1)))
    '
    rs.Update
    '
End Sub
'
'Filler 14 - Leitura - Brand
'
Sub LeArquivoIncluiBrand(ByVal strLinha As String)
    '
    rs("codigo") = Trim(Mid(strLinha, 3, 3))
    rs("descricao") = Trim(Mid(strLinha, 6, 20))
    rs("status") = LCase(Trim(Mid(strLinha, 26, 1)))
    '
    rs.Update
    '
End Sub
'
'Filler 16 - Leitura - Sub-Brand
'
Sub LeArquivoIncluisubbrand(ByVal strLinha As String)
    '
    rs("codigo") = Trim(Mid(strLinha, 3, 3))
    rs("descricao") = Trim(Mid(strLinha, 6, 20))
    rs("observacao") = Trim(Mid(strLinha, 26, 40))
    rs("status") = LCase(Trim(Mid(strLinha, 46, 1)))
    rs.Update
    '
End Sub
'
'Filler 18 - Leitura - Produtos
'
Sub LeArquivoIncluiProduto(ByVal strLinha As String)
    Dim varDouble As Double
    '
    rs("fabricante") = Trim(Mid(strLinha, 3, 3))
    rs("brand") = Trim(Mid(strLinha, 6, 3))
    rs("sub_brand") = Trim(Mid(strLinha, 9, 3))
    rs("codigo_produto") = Trim(Mid(strLinha, 12, 6))
    rs("descricao") = Trim(Mid(strLinha, 18, 40))
    rs("unidade") = Trim(Mid(strLinha, 58, 3))
    rs("qtd_disponivel") = Trim(Mid(strLinha, 61, 4))
    '
    If IsNumeric(Trim(Mid(strLinha, 65, 4))) Then
        varDouble = CDbl(Trim(Mid(strLinha, 65, 4))) / 100
    Else
        varDouble = 0
    End If
    '
    rs("aliquota_icms") = CStr(varDouble)
    '
    If IsNumeric(Trim(Mid(strLinha, 69, 4))) Then
        varDouble = CDbl(Trim(Mid(strLinha, 69, 4))) / 100
    Else
        varDouble = 0
    End If
    '
    rs("aliquota_ipi") = CStr(varDouble)
    '
    rs("substituicao_tributaria") = Trim(Mid(strLinha, 73, 1))
    rs("quantidade_embalagem") = Trim(Mid(strLinha, 74, 4))
    '
    If Left(Trim(Mid(strLinha, 78, 5)), 1) = "-" Then
       If IsNumeric(Trim(Mid(strLinha, 79, 4))) Then
          varDouble = (CDbl(Trim(Mid(strLinha, 79, 4))) / 100) * -1
       Else
          varDouble = 0
       End If
    Else
        If IsNumeric(Trim(Mid(strLinha, 78, 5))) Then
            varDouble = CDbl(Trim(Mid(strLinha, 78, 5))) / 100
        Else
            varDouble = 0
        End If
    End If
    '
    rs("desconto_acrescimo") = CStr(varDouble)
    '
    If IsNumeric(Trim(Mid(strLinha, 83, 9))) Then
        varDouble = CDbl(Trim(Mid(strLinha, 83, 9))) / 1000
    Else
        varDouble = 0
    End If
    '
    rs("peso") = CStr(varDouble)
    '
    rs("volume") = Trim(Mid(strLinha, 92, 9))
    '
    If IsNumeric(Trim(Mid(strLinha, 101, 4))) Then
        varDouble = CDbl(Trim(Mid(strLinha, 101, 4))) / 100
    Else
        varDouble = 0
    End If
    '
    rs("desconto_maximo") = CStr(varDouble)
    '
    rs("embalagem") = Trim(Mid(strLinha, 105, 10))
    rs("empresa") = Trim(Mid(strLinha, 115, 2))
    rs("filial") = Trim(Mid(strLinha, 117, 2))
    '
    '=================================================================================================
    '
    If IsNumeric(Trim(Mid(strLinha, 119, 10))) Then
        varDouble = CDbl(Trim(Mid(strLinha, 119, 10))) / 100
    Else
        varDouble = 0
    End If
    '
    rs("preco1") = CStr(varDouble)
    '
    If IsNumeric(Trim(Mid(strLinha, 129, 10))) Then
        varDouble = CDbl(Trim(Mid(strLinha, 129, 10))) / 100
    Else
        varDouble = 0
    End If
    '
    rs("preco2") = CStr(varDouble)
    '
    If IsNumeric(Trim(Mid(strLinha, 139, 10))) Then
        varDouble = CDbl(Trim(Mid(strLinha, 139, 10))) / 100
    Else
        varDouble = 0
    End If
    '
    rs("preco3") = CStr(varDouble)
    '
    If IsNumeric(Trim(Mid(strLinha, 149, 10))) Then
        varDouble = CDbl(Trim(Mid(strLinha, 149, 10))) / 100
    Else
        varDouble = 0
    End If
    '
    rs("preco4") = CStr(varDouble)
    '
    If IsNumeric(Trim(Mid(strLinha, 159, 10))) Then
        varDouble = CDbl(Trim(Mid(strLinha, 159, 10))) / 100
    Else
        varDouble = 0
    End If
    '
    rs("preco5") = CStr(varDouble)
    '
    '=================================================================================
    '
    rs("media_diaria") = "000000000000"
    rs("estoque") = "0000000000"
    rs("sugestao") = "0000000000"
    rs("pedido") = "0000000000"
    '
    rs("filtro") = "000"
    '
    rs("status") = LCase(Trim(Mid(strLinha, 169, 1)))
    '
    rs.Update
    '
End Sub
'
'Filler 22 - Leitura - Objetivos de Vendas
'
Sub LeArquivoIncluiObjetivo(ByVal strLinha As String)
    '
    ' rs("codigo_vendedor") = Trim(Mid(strLinha, 3, 5))
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
    rs("codigo_produto") = Trim(Mid(strLinha, 8, 6))
    rs("cota_qtde") = Trim(Mid(strLinha, 14, 10))
    rs("realizado") = Trim(Mid(strLinha, 24, 10))
    '
    rs.Update
    '
End Sub
'
'Filler 23 - Leitura - Tipo de Movimentos
'
Sub LeArquivoIncluiTipoMovimento(ByVal strLinha As String)
    '
    rs("codigo") = Trim(Mid(strLinha, 3, 1))
    rs("descricao") = Trim(Mid(strLinha, 4, 15))
    '
    rs.Update
    '
End Sub
'
'Filler 25 - Leitura - Promoções
'
Sub LeArquivoIncluiPromocao(ByVal strLinha As String)
    Dim varDouble As Double
    '
    rs("codigo_produto") = Trim(Mid(strLinha, 3, 6))
    rs("qtd_inicial") = Trim(Mid(strLinha, 9, 8))
    rs("qtd_final") = Trim(Mid(strLinha, 17, 8))
    '
    If IsNumeric(Trim(Mid(strLinha, 25, 10))) Then
        varDouble = CDbl(Trim(Mid(strLinha, 25, 10))) / 100
    Else
        varDouble = 0
    End If
    '
    rs("preco_promocional") = CStr(varDouble)
    '
    rs.Update
    '
End Sub
'
'Filler 26 - Leitura - Descrições das Tabelas de Preços
'
Sub LeArquivoIncluiDescricaoTabela(ByVal strLinha As String)
    Dim varDouble As Double
    '
    rs("codigo_tabela") = Trim(Mid(strLinha, 3, 1))
    rs("descricao") = Trim(Mid(strLinha, 4, 10))
    '
    rs.Update
    '
End Sub
'
'Filler 27 - Leitura - Tabelas de Preços
'
Sub LeArquivoIncluiTabelaPrecos(ByVal strLinha As String)
    Dim varDouble As Double
    '
    rs("tabela_precos") = Trim(Mid(strLinha, 3, 1))
    '
    'rs("descricao") = Trim(Mid(strlinha, 4, 6))
    '
    rs("produto") = Trim(Mid(strLinha, 4, 6))
    '
    If IsNumeric(Trim(Mid(strLinha, 10, 10))) Then
        varDouble = CDbl(Trim(Mid(strLinha, 10, 10))) / 100
    Else
        varDouble = 0
    End If
    '
    rs("preco") = CStr(varDouble)
    '
    'If IsNumeric(Trim(Mid(strLinha, 20, 6))) Then
    '    varDouble = CDbl(Trim(Mid(strLinha, 20, 6))) / 100
    'Else
    '    varDouble = 0
    'End If
    ''
    'rs("desconto") = CStr(varDouble)
    '
    rs.Update
    '
End Sub
'
'Filler 51 - Leitura - Retorno de Pedidos
'
Sub LeArquivoIncluiPedido(ByVal strLinha As String)
    '
    Dim varDouble As Double
    Dim mObservacao As String
    '
    'ExecSQL "CREATE TABLE pedido("
    '
    ' codigo_vendedor="00000"
    ' codigo_cliente VARCHAR(5),
    ' numero_pedido_interno="000000"
    ' numero_pedido_externo ="000000"
    ' pedido_cliente ="0000000000"
    ' data_emissao ="00000000"
    ' hora_emissao ="000000"
    ' data_entrega ="00000000"
    ' acrescimo_valor ="0000000000"
    ' desconto_valor ="0000000000"
    ' forma_pgto ="00"
    ' condicao_pgto ="00"
    ' tipo_movimento="0"
    ' status ="A"
    ' observacao=""
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
    rs("codigo_cliente") = Trim(Mid(strLinha, 3, 5))
    '
    rs("numero_pedido_interno") = Trim(Mid(strLinha, 8, 6))
    '
    rs("numero_pedido_externo") = Trim(Mid(strLinha, 14, 6))
    '
    rs("pedido_cliente") = Trim(Mid(strLinha, 20, 10))
    '
    rs("data_emissao") = Trim(Mid(strLinha, 30, 8))
    '
    rs("hora_emissao") = Trim(Mid(strLinha, 38, 6))
    '
    rs("data_entrega") = Trim(Mid(strLinha, 44, 8))
    '
    If IsNumeric(Trim(Mid(strLinha, 52, 10))) Then
        varDouble = CDbl(Trim(Mid(strLinha, 52, 10))) / 100
    Else
        varDouble = 0
    End If
    '
    rs("acrescimo_valor") = CStr(varDouble)
    '
    If IsNumeric(Trim(Mid(strLinha, 62, 10))) Then
        varDouble = CDbl(Trim(Mid(strLinha, 62, 10))) / 100
    Else
        varDouble = 0
    End If
    '
    rs("desconto_valor") = CStr(varDouble)
    '
    rs("forma_pgto") = Trim(Mid(strLinha, 72, 2))
    '
    rs("condicao_pgto") = Trim(Mid(strLinha, 74, 2))
    '
    rs("tipo_movimento") = Trim(Mid(strLinha, 76, 1))
    '
    rs("status") = LCase(Trim(Mid(strLinha, 77, 1)))
    '
    mObservacao = ""
    '
    rs("observacao") = mObservacao
    '
    rs.Update
    '
End Sub
'
'Filler 32 - Leitura - Retorno de Itens de Pedidos
'
Sub LeArquivoIncluiItenPedido(ByVal strLinha As String)
    '
    Dim varDouble As Double
    '
    rs("numero_pedido") = Trim(Mid(strLinha, 3, 6))
    rs("codigo_produto") = Trim(Mid(strLinha, 9, 6))
    rs("qtd_pedida") = Trim(Mid(strLinha, 15, 8))
    rs("qtd_faturada") = Trim(Mid(strLinha, 23, 8))
    '
    If IsNumeric(Trim(Mid(strLinha, 31, 10))) Then
       varDouble = CDbl(Trim(Mid(strLinha, 31, 10))) / 100
    Else
        varDouble = 0
    End If
    '
    rs("valor_unitario") = CStr(varDouble)
    '
    rs("desconto") = Trim(Mid(strLinha, 41, 5))
    '
    rs.Update
    '
End Sub
'
'Filler 33 - Leitura - Retorno de Observação de de Pedidos
'
Sub LeArquivoIncluiOBSPedido(ByVal strLinha As String)
    '
    Dim varDouble As Double
    '
    rs("observacao") = strObservacaoPed
    '
    rs.Update
    '
    ' If rs.State = 1 Then rs.Close
    '
End Sub
'
'Filler 34 - Leitura - Títulos em Aberto
'
Sub LeArquivoIncluiTituloAberto(ByVal strLinha As String)
    Dim strString As String
    Dim strVal As Double
    '
    rs("codigo_cliente") = Trim(Mid(strLinha, 3, 5))
    rs("numero_documento") = Trim(Mid(strLinha, 8, 20))
    rs("data_vencimento") = Trim(Mid(strLinha, 28, 8))
    rs("vencimento_data") = Mid(Trim(Mid(strLinha, 28, 8)), 5, 4) + Mid(Trim(Mid(strLinha, 28, 8)), 3, 2) + Mid(Trim(Mid(strLinha, 28, 8)), 1, 2)
    '
    ' MsgBox "Data:" & Trim(Mid(strLinha, 28, 8)), vbOKOnly + vbCritical, App.Title
    '
    If IsNumeric(Trim(Mid(strLinha, 36, 10))) Then
       '
       strVal = CDbl(Trim(Mid(strLinha, 36, 10)))
       '
    Else
       '
       strVal = 0
       '
    End If
    '
    strVal = strVal / 100
    '
    strString = CStr(strVal)
    '
    rs("valor") = strString
    '
    rs.Update
    '
    ' If rs.State = 1 Then rs.Close
    '
End Sub
'
'Filler 38 - Leitura - Justificativas de Não Venda
'
Sub LeArquivoIncluiNaoVenda(ByVal strLinha As String)
    '
    rs("codigo") = Trim(Mid(strLinha, 3, 3))
    rs("tipo") = Trim(Mid(strLinha, 6, 1))
    rs("descricao") = Trim(Mid(strLinha, 7, 20))
    rs.Update
    '
End Sub
'
'Filler 40 - Leitura - Descontos de Canal
'
Sub LeArquivoIncluiDescontoCanal(ByVal strLinha As String)
    Dim varDouble As Double
    '
    rs("sub_brand") = Trim(Mid(strLinha, 3, 3))
    rs("cliente") = Trim(Mid(strLinha, 6, 5))
    '
    If IsNumeric(Trim(Mid(strLinha, 11, 4))) Then
        varDouble = CDbl(Trim(Mid(strLinha, 11, 4))) / 100
    Else
        varDouble = 0
    End If
    '
    rs("desconto") = CStr(varDouble)
    '
    rs.Update
    '
End Sub
'
'Filler 44 - Leitura - Observações de Clientes
'
'
'Filler 91 - Leitura - Corpo de Mensagens (Corpo)
'
Sub LeArquivoIncluiCorpoMensagem(ByVal strLinha As String)
    '
    strMensagem = strMensagem & Trim(Mid(strLinha, 3, cstLarguraMensagem))
    '
End Sub
'
'Filler 92 - Leitura - Capa de Mensagens (Capa)
'
Sub LeArquivoIncluiMensagem(ByVal strLinha As String)
    '
    rs.Close
    '
    rs.Open "SELECT * FROM mensagens ;", CONN, adOpenDynamic, adLockOptimistic
    '
    rs.AddNew
    '
    ' Tipo de Mensagem: Recebida
    '
    rs("tipo_mensagem") = "1"
    '
    rs("codigo_vendedor_origem") = Trim(Mid(strLinha, 3, 5))
    rs("codigo_vendedor_destino") = Trim(Mid(strLinha, 8, 5))
    rs("assunto") = Trim(Mid(strLinha, 13, 20))
    rs("data") = Trim(Mid(strLinha, 39, 8))
    rs("hora") = Trim(Mid(strLinha, 47, 6))
    rs("status") = "R"
    '
    If Trim(strMensagem) <> "" Then
       '
       Do While InStr(1, strMensagem, "@%", vbTextCompare) > 0
          '
          strMensagem = Mid(strMensagem, 1, InStr(1, strMensagem, "@%", vbTextCompare) - 1) _
          & Chr(13) & Chr(10) & Mid(strMensagem, InStr(1, strMensagem, "@%", vbTextCompare) _
          + 2, Len(strMensagem) - InStr(1, strMensagem, "@%", vbTextCompare) + 1)
          '
          ' MsgBox "mlinhaObs(*): *" & mObservacao & "* len=" & Len(mObservacao), vbOKOnly + vbCritical, App.Title
          '
       Loop
       '
    End If
    '
    rs("mensagem") = strMensagem
    '
    rs.Update
    '
    strMensagem = vbNullString
    '
    ' If rs.State = 1 Then rs.Close
    '
End Sub
'
'Filler 93 - Leitura - Corpo de Relatórios (Corpo)
'
Sub LeArquivoIncluiCorpoRelatorio(ByVal strLinha As String)
  '
  strRelatorio = strRelatorio & Mid(strLinha, 3, cstLarguraRelatorio) & Chr(13) & Chr(10)
  '
End Sub
'
'Filler 94 - Leitura - Capa de Relatórios (Capa)
'
Sub LeArquivoIncluiRelatorio(ByVal strLinha As String)
    '
    rs.Close
    '
    rs.Open "SELECT * FROM relatorios ;", CONN, adOpenDynamic, adLockOptimistic
    '
    rs.AddNew
    '
    rs("descricao") = Trim(Mid(strLinha, 8, 25))
    '
    If Trim(strRelatorio) <> "" Then
       '
       Do While InStr(1, strRelatorio, "@%", vbTextCompare) > 0
          '
          strRelatorio = Mid(strRelatorio, 1, InStr(1, strRelatorio, "@%", vbTextCompare) - 1) _
          & Chr(13) & Chr(10) & Mid(strRelatorio, InStr(1, strRelatorio, "@%", vbTextCompare) _
          + 2, Len(strRelatorio) - InStr(1, strRelatorio, "@%", vbTextCompare) + 1)
          '
          ' MsgBox "mlinhaObs(*): *" & mObservacao & "* len=" & Len(mObservacao), vbOKOnly + vbCritical, App.Title
          '
       Loop
       '
    End If
    '
    rs("relatorio") = strRelatorio
    '
    rs("data") = Trim(Mid(strLinha, 39, 8))
    rs("hora") = Trim(Mid(strLinha, 47, 6))
    '
    rs.Update
    '
    strRelatorio = vbNullString
    '
    ' If rs.State = 1 Then rs.Close
    '
End Sub
'
'
'
Public Sub GravaArquivos()
    '
    On Error Resume Next
    '
    Dim strDia As String
    Dim strMes As String
    Dim strAno As String
    Dim strMensagem As String
    Dim intNumeroAtual As Integer
    '
    Dim mLinha As String
    Dim mlinhaItens As String
    Dim mlinhaObs As String
    Dim mlinhaBuffer As String
    Dim mObservacao As String
    Dim mMensagem As String
    '
    Dim rs, rsaux
    '
    Screen.MousePointer = 11
    AddStatus "Abrindo arquivo... aguarde"
    '
    connOpen
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    Set rsaux = CreateObject("ADOCE.Recordset.3.0")
    '
    strDia = Trim(Day(Now))
    strMes = Trim(Month(Now))
    strAno = Right(Trim(Year(Now)), 2)
    '
    If Len(strDia) = 1 Then strDia = "0" & Trim(strDia)
    If Len(strMes) = 1 Then strMes = "0" & Trim(strMes)
    '
    If frmStart.FileSystem.Dir(strPath & "\" & Trim(strAno) & Trim(strMes) & Trim(strDia) & ".out") <> "" Then
       frmLogin.File.Open strPath & "\" & Trim(strAno) & Trim(strMes) & Trim(strDia) & ".out", fsModeInput, fsAccessRead, fsLockReadWrite
       intNumeroAtual = CInt(frmLogin.File.LineInputString)
       frmLogin.File.Close
       frmStart.FileSystem.Kill strPath & "\" & Trim(strAno) & Trim(strMes) & Trim(strDia) & ".out"
    Else
        intNumeroAtual = 0
    End If
    '
    If frmStart.FileSystem.Dir("\my documents") <> "" Then
        frmLogin.File.Open "\my documents\T" & Trim(strAno) & Trim(strMes) & Trim(strDia) & CStr(intNumeroAtual) & "." & Right(usrCodigoVendedor, 3), fsModeOutput, fsAccessWrite
    ElseIf frmStart.FileSystem.Dir("\meus documentos") <> "" Then
        frmLogin.File.Open "\meus documentos\T" & Trim(strAno) & Trim(strMes) & Trim(strDia) & CStr(intNumeroAtual) & "." & Right(usrCodigoVendedor, 3), fsModeOutput, fsAccessWrite
    End If
    '
    'Filler 00 - Gravação - vendedor atual
    '
    rs.Open "SELECT * FROM vendedor WHERE codigo_vendedor='" & usrCodigoVendedor & "';", CONN, adOpenForwardOnly, adLockReadOnly
    '
    If rs.RecordCount > 0 Then
       '
       frmLogin.File.LinePrint _
       "00" & strDia & strMes & strAno & RetornaHoraString(CStr(Hour(Now)), CStr(Minute(Now)), CStr(Second(Now))) & _
       RetornaStringEspacos(usrCodigoVendedor, 5) & RetornaStringEspacos(rs("senha"), 6)
       '
    End If
    '
    AddStatus "Incluindo vendedor"
    If rs.State = 1 Then rs.Close
    '
    'Filler 11 - Gravação- cliente com status = D (digitado) e M (Modificado) - 01
    '
    rs.Open "SELECT * FROM clientes WHERE status='D' or status='M' ;", CONN, adOpenDynamic, adLockOptimistic
    '
    If rs.RecordCount > 0 Then
       '
       Do Until rs.EOF
          '
          AddStatus "Incluindo Cliente"
          '
          mLinha = ""
          mLinha = mLinha & "11" ' 01
          mLinha = mLinha & RetornaStringEspacos(rs("codigo_cliente"), 5)
          mLinha = mLinha & RetornaStringEspacos(rs("nome_fantasia"), 20)
          mLinha = mLinha & RetornaStringEspacos(rs("razao_social"), 40)
          mLinha = mLinha & RetornaStringEspacos(rs("endereco_entrega"), 40)
          mLinha = mLinha & RetornaStringEspacos(rs("cep_entrega"), 8)
          mLinha = mLinha & RetornaStringEspacos(rs("bairro_entrega"), 20)
          mLinha = mLinha & RetornaStringEspacos(rs("cidade_entrega"), 5)
          mLinha = mLinha & RetornaStringEspacos(rs("endereco_cobranca"), 40)
          mLinha = mLinha & RetornaStringEspacos(rs("cep_cobranca"), 8)
          mLinha = mLinha & RetornaStringEspacos(rs("bairro_cobranca"), 20)
          mLinha = mLinha & RetornaStringEspacos(rs("cidade_cobranca"), 5)
          '
          mLinha = mLinha & "  "
          '
          mLinha = mLinha & RetornaStringEspacos(rs("telefone"), 12)
          mLinha = mLinha & RetornaStringEspacos(rs("fax"), 12)
          mLinha = mLinha & RetornaStringEspacos(rs("email"), 30)
          mLinha = mLinha & RetornaStringEspacos(rs("www"), 30)
          mLinha = mLinha & RetornaStringEspacos(rs("data_fundacao"), 8)
          mLinha = mLinha & rs("predio_proprio")
          mLinha = mLinha & RetornaStringEspacos(rs("referencia_bancaria_1"), 20)
          mLinha = mLinha & RetornaStringEspacos(rs("referencia_bancaria_2"), 20)
          mLinha = mLinha & RetornaStringEspacos(rs("referencia_comercial_1"), 20)
          mLinha = mLinha & RetornaStringEspacos(rs("referencia_comercial_2"), 20)
          mLinha = mLinha & RetornaStringEspacos(rs("contato"), 25)
          mLinha = mLinha & RetornaStringEspacos(rs("cnpjmf"), 14)
          mLinha = mLinha & RetornaStringEspacos(rs("incricao_estadual"), 20)
          mLinha = mLinha & RetornaStringEspacos(rs("cpf"), 14)
          mLinha = mLinha & RetornaStringEspacos(rs("rg"), 20)
          mLinha = mLinha & RetornaStringEspacos(rs("ramo_atividade"), 3)
          mLinha = mLinha & RetornaStringEspacos(rs("status"), 1)
          '
          frmLogin.File.LinePrint mLinha
          '
          rs.MoveNext
          '
       Loop
       '
       rs.MoveFirst
       '
       Do Until rs.EOF
          '
          rs.Delete
          rs.MoveFirst
          '
       Loop
       '
    End If
    '
    If rs.State = 1 Then rs.Close
    '
    'Filler 16 - Gravação - Pedido com status = D (digitado)
    '
    rs.Open "SELECT * FROM pedido WHERE status='D';", CONN, adOpenDynamic, adLockOptimistic
    '
    If rs.RecordCount > 0 Then
       '
       Do Until rs.EOF
          '
          AddStatus "Incluindo Pedido"
          '
          mLinha = ""
          mLinha = mLinha & "16" ' 02
          '
          ' mlinha = mlinha & RetornaStringEspacos(rs("numero_pedido_interno"), 6)
          '
          mLinha = mLinha & RetornaStringZeros(rs("numero_pedido_interno"), 6, 0, False)
          '
          mLinha = mLinha & RetornaStringEspacos(rs("codigo_cliente"), 5)
          '
          mLinha = mLinha & RetornaStringEspacos(rs("pedido_cliente"), 10)
          '
          ' mlinha = mlinha & RetornaStringEspacos(rs("data_emissao"), 8)
          '
          mLinha = mLinha & RetornaStringZeros(rs("data_emissao"), 8, 0, False)
          '
          ' mlinha = mlinha & RetornaStringEspacos(rs("hora_emissao"), 6)
          '
          mLinha = mLinha & RetornaStringZeros(rs("hora_emissao"), 6, 0, False)
          '
          ' mlinha = mlinha & RetornaStringEspacos(rs("data_entrega"), 8)
          '
          mLinha = mLinha & RetornaStringZeros(rs("data_entrega"), 8, 0, False)
          '
          ' mlinha = mlinha & RetornaStringEspacos(rs("acrescimo_valor"), 10)
          '
          mLinha = mLinha & RetornaStringZeros(rs("acrescimo_valor"), 10, 2, False)
          '
          ' mlinha = mlinha & RetornaStringEspacos(rs("desconto_valor"), 10)
          '
          mLinha = mLinha & RetornaStringZeros(rs("desconto_valor"), 10, 2, False)
          '
          mLinha = mLinha & RetornaStringEspacos(rs("forma_pgto"), 2)
          '
          mLinha = mLinha & RetornaStringEspacos(rs("condicao_pgto"), 2)
          '
          mLinha = mLinha & rs("tipo_movimento")
          '
          mLinha = mLinha & rs("status")
          '
          frmLogin.File.LinePrint mLinha
          '
          rsaux.Open "SELECT * FROM itens_pedido WHERE numero_pedido='" & rs("numero_pedido_interno") & "';", CONN, adOpenForwardOnly, adLockReadOnly
          '
          If rsaux.RecordCount > 0 Then
             '
             AddStatus "Incluindo Itens do Pedido"
             '
             Do Until rsaux.EOF
                '
                mlinhaItens = ""
                mlinhaItens = mlinhaItens & "21" ' 03
                mlinhaItens = mlinhaItens & RetornaStringZeros(rsaux("numero_pedido"), 6, 0, False)
                mlinhaItens = mlinhaItens & RetornaStringEspacos(rsaux("codigo_produto"), 6)
                mlinhaItens = mlinhaItens & RetornaStringZeros(rsaux("qtd_pedida"), 8, 0, False)
                '
                ' mlinhaItens = mlinhaItens & RetornaStringEspacos(rsaux("valor_unitario"), 10) ' , False)
                '
                mlinhaItens = mlinhaItens & RetornaStringZeros(rsaux("valor_unitario"), 10, 2, False)
                '
                mlinhaItens = mlinhaItens & RetornaStringZeros(rsaux("desconto"), 6, 2, False)
                '
                frmLogin.File.LinePrint mlinhaItens
                '
                rsaux.MoveNext
                '
             Loop
             '
             rsaux.Close
             '
          End If
          '
          AddStatus "Incluindo Observações de Pedido"
          '
          '====================================================
          '
          '
          '==================================================== 04
          '
          mlinhaBuffer = "26" & RetornaStringZeros(rs("numero_pedido_interno"), 6, 0, False) & RetornaStringEspacos(rs("codigo_cliente"), 5) & "00"
          '
          mObservacao = rs("Observacao")
          '
          ' Troca line feed e feed back
          '
          Do While InStr(1, mObservacao, Chr(13) & Chr(10), vbTextCompare) > 0
             '
             mObservacao = Mid(mObservacao, 1, InStr(1, mObservacao, Chr(13) & Chr(10), vbTextCompare) - 1) & "%@" _
             & Mid(mObservacao, InStr(1, mObservacao, Chr(13) & Chr(10), vbTextCompare) + 2, Len(mObservacao) - InStr(1, mObservacao, Chr(13) & Chr(10), vbTextCompare) + 1)
             '
             ' MsgBox "mlinhaObs(*): *" & mObservacao & "* len=" & Len(mObservacao), vbOKOnly + vbCritical, App.Title
             '
          Loop
          '
          ' MsgBox "Observação: *" & mObservacao & "* len=" & Len(mObservacao), vbOKOnly + vbCritical, App.Title
          '
          ' Pica o memo em pedacinhos de 80 bytes
          '
          Do While True
             '
             If Len(mObservacao) > 80 Then
                '
                mlinhaObs = mlinhaBuffer & Mid(mObservacao, 1, 80)
                '
                mObservacao = Mid(mObservacao, 81, Len(mObservacao) - 80)
                '
             Else
                '
                mlinhaObs = mlinhaBuffer & mObservacao
                '
                mObservacao = ""
                '
             End If
             '
             frmLogin.File.LinePrint mlinhaObs
             '
             If mObservacao = "" Then Exit Do
             '
          Loop
          '
          rs.MoveNext
          '
       Loop
       '
       rs.MoveFirst
       '
    End If
    '
    If rs.State = 1 Then rs.Close
    '
    'Filler 31 - Gravação - Roteiro Percorrido - 05
    '
    rs.Open "SELECT * FROM roteiro_percorrido;", CONN, adOpenDynamic, adLockOptimistic
    '
    If rs.RecordCount > 0 Then
       '
       Do Until rs.EOF
          '
          mLinha = ""
          mLinha = mLinha & "31" ' 05
          '
          ' roteiro_percorrido
          '
          ' codigo_vendedor_atual VARCHAR(5)
          ' data_base VARCHAR(8)
          ' codigo_cliente VARCHAR(5)
          ' data_visita VARCHAR(8)
          ' hora_visita VARCHAR(6)
          ' motivo_nao_visita VARCHAR(3)
          ' visita_extra VARCHAR(1)
          ' automatico VARCHAR(1)
          ' status VARCHAR(1)
          '
          mLinha = mLinha & RetornaStringZeros(rs("data_base"), 8, 0, False)
          mLinha = mLinha & RetornaStringEspacos(rs("codigo_cliente"), 5)
          mLinha = mLinha & RetornaStringZeros(rs("data_visita"), 8, 0, False)
          mLinha = mLinha & RetornaStringZeros(rs("hora_visita"), 6, 0, False)
          mLinha = mLinha & RetornaStringEspacos(rs("motivo_nao_visita"), 3)
          mLinha = mLinha & RetornaStringEspacos(rs("visita_extra"), 1)
          mLinha = mLinha & RetornaStringEspacos(rs("automatico"), 1)
          mLinha = mLinha & RetornaStringEspacos(rs("dia_visita"), 1)
          mLinha = mLinha & RetornaStringEspacos(rs("status"), 1)
          '
          frmLogin.File.LinePrint mLinha
          '
          rs.MoveNext
          '
       Loop
       '
    End If
    '
    If rs.State = 1 Then rs.Close
    '
    '===============================================================================================
    '
    ' Filler 36/41 - Gravação - Mensagem
    '
    AddStatus "Incluindo Mensagens"
    '
    rs.Open "SELECT * FROM mensagens;", CONN, adOpenDynamic, adLockOptimistic
    '
    If rs.RecordCount > 0 Then
       '
       Do Until rs.EOF
          '
          mlinhaBuffer = "36"
          '
          mMensagem = rs("mensagem")
          '
          ' Troca line feed e feed back
          '
          Do While InStr(1, mMensagem, Chr(13) & Chr(10), vbTextCompare) > 0
             '
             mMensagem = Mid(mMensagem, 1, InStr(1, mMensagem, Chr(13) & Chr(10), vbTextCompare) - 1) & "%@" _
             & Mid(mMensagem, InStr(1, mMensagem, Chr(13) & Chr(10), vbTextCompare) + 2, Len(mMensagem) - InStr(1, mMensagem, Chr(13) & Chr(10), vbTextCompare) + 1)
             '
             ' MsgBox "mlinhaObs(*): *" & mMensagem & "* len=" & Len(mMensagem), vbOKOnly + vbCritical, App.Title
             '
          Loop
          '
          ' MsgBox "Observação: *" & mMensagem & "* len=" & Len(mMensagem), vbOKOnly + vbCritical, App.Title
          '
          ' Pica o memo em pedacinhos de 80 bytes
          '
          Do While True
             '
             If Len(mMensagem) > 80 Then
                '
                mlinhaObs = mlinhaBuffer & Mid(mMensagem, 1, 80)
                '
                mMensagem = Mid(mMensagem, 81, Len(mMensagem) - 80)
                '
             Else
                '
                mlinhaObs = mlinhaBuffer & mMensagem
                '
                mMensagem = ""
                '
             End If
             '
             frmLogin.File.LinePrint mlinhaObs
             '
             If mMensagem = "" Then Exit Do
             '
          Loop
          '
          '
          '
          ' "41"
          '
          frmLogin.File.LinePrint "41" & RetornaStringEspacos(rs("codigo_vendedor_origem"), 5) & RetornaStringEspacos(rs("codigo_vendedor_destino"), 5) & RetornaStringEspacos(rs("assunto"), 20) & "000000" & RetornaStringZeros(rs("data"), 8, 0, False) & RetornaStringZeros(rs("hora"), 6, 0, False)
          '
          rs.MoveNext
          '
       Loop
       '
    End If
    '
    If rs.State = 1 Then rs.Close
    '
    '===================================================================================================
    '
    'Filler 99 - Gravação - Fim de Arquivo
    '
    frmLogin.File.LinePrint "99"
    '
    'Fecha o arquivo
    '
    frmLogin.File.Close
    '
    frmLogin.File.Open strPath & "\" & Trim(strAno) & Trim(strMes) & Trim(strDia) & ".out", fsModeOutput, fsAccessWrite, fsLockReadWrite
    '
    frmLogin.File.LinePrint CStr(intNumeroAtual + 1)
    '
    frmLogin.File.Close
    '
    Set rs = Nothing
    '
    connClose
    '
    AddStatus "Arquivo Gerado com sucesso"
    '
    Screen.MousePointer = 0
    '
End Sub
'
'
'
'=====================================================================================
'                                        Funcções
'=====================================================================================
Function RetornaStringEspacos(ByVal strStringAtual As String, intTamanhoTotal As Integer) As String
    If Len(strStringAtual) > intTamanhoTotal Then
        strStringAtual = Left(strStringAtual, intTamanhoTotal)
    ElseIf Len(strStringAtual) = intTamanhoTotal Then
        strStringAtual = strStringAtual
    Else
        For I = Len(strStringAtual) To intTamanhoTotal - 1
            strStringAtual = strStringAtual & " "
        Next
    End If
    RetornaStringEspacos = strStringAtual
End Function

Function RetornaStringZeros(ByVal strAtual As String, ByVal intNumeroZeros As Integer, ByVal intNumeroCasas As Integer, ByVal bolPoeMenos As Boolean) As String
    '
    Dim mPosicaoReal As Integer
    '
    RetornaStringZeros = vbNullString
    '
    strAtual = Trim(strAtual)
    '
    '========================= RETIRA R$ e virgula do PRECO UNITÁRIO ==============================
    '
    ' MsgBox "strAtual(1): *" & Trim(strAtual) & "*", vbOKOnly + vbCritical, App.Title
    '
    mPosicaoReal = InStr(1, strAtual, "R$", vbTextCompare)
    '
    ' MsgBox "Posicao: *" & CStr(mPosicaoReal) & "*", vbOKOnly + vbCritical, App.Title
    '
    If mPosicaoReal > 0 Then
       '
       If mPosicaoReal = 1 Then
          '
          strAtual = Mid(strAtual, 3, Len(strAtual) - 2)
          '
       Else
          '
          strAtual = Mid(strAtual, 1, mPosicaoReal - 1) _
          & Mid(strAtual, mPosicaoReal + 2, Len(strAtual) - (mPosicaoReal + 2))
          '
       End If
    '
    End If
    '
    ' MsgBox "strAtual(2): *" & Trim(strAtual) & "*", vbOKOnly + vbCritical, App.Title
    '
    While InStr(1, Trim(strAtual), ",", vbTextCompare) < 1
       '
       strAtual = strAtual & "," & String(intNumeroCasas, "0")
       '
       ' MsgBox "strAtual(3): *" & Trim(strAtual) & "*", vbOKOnly + vbCritical, App.Title
       '
    Wend
    '
    strAtual = strAtual & String(intNumeroCasas, "0")
    '
    mPosicaoReal = InStr(1, Trim(strAtual), ",", vbTextCompare)
    '
    strAtual = Mid(strAtual, 1, mPosicaoReal + intNumeroCasas)
    '
    mPosicaoReal = InStr(1, Trim(strAtual), ",", vbTextCompare)
    '
    ' MsgBox "Posicao: *" & CStr(mPosicaoReal) & "*", vbOKOnly + vbCritical, App.Title
    '
    If mPosicaoReal > 0 Then
       '
       If mPosicaoReal = 1 Then
          '
          strAtual = Mid(strAtual, 2, Len(strAtual) - 1)
          '
       Else
          '
          strAtual = Mid(strAtual, 1, mPosicaoReal - 1) _
          & Mid(strAtual, mPosicaoReal + 1, Len(strAtual) - (mPosicaoReal))
          '
       End If
       '
    End If
    '
    ' MsgBox "strAtual(4): *" & Trim(strAtual) & "*", vbOKOnly + vbCritical, App.Title
    '
    '===================================================================================================
    '
    If Len(strAtual) > 0 Then
       '
       For I = 1 To Len(strAtual)
           '
           If IsNumeric(Mid(strAtual, I, 1)) Then
              '
              RetornaStringZeros = RetornaStringZeros & Mid(strAtual, I, 1)
              '
           Else
              '
              RetornaStringZeros = String(intNumeroZeros, "0")
              '
              Exit For
              '
           End If
           '
       Next
       '
    Else
       '
       RetornaStringZeros = String(intNumeroZeros, "0")
       '
    End If
    '
    If Len(RetornaStringZeros) > intNumeroZeros Then
       '
       RetornaStringZeros = Left(RetornaStringZeros, intNumeroZeros)
       '
    ElseIf Len(RetornaStringZeros) = intNumeroZeros Then
       '
       RetornaStringZeros = RetornaStringZeros
       '
    Else
       '
       For I = Len(RetornaStringZeros) To intNumeroZeros - 1
           '
           RetornaStringZeros = "0" & RetornaStringZeros
           '
       Next
       '
    End If
    '
    If bolPoeMenos Then
       '
       RetornaStringZeros = Right(RetornaStringZeros, Len(RetornaStringZeros) - 1)
       '
       RetornaStringZeros = "-" & RetornaStringZeros
       '
    End If
    '
    ' MsgBox "strAtual(5): *" & Trim(RetornaStringZeros) & "*", vbOKOnly + vbCritical, App.Title
    '
End Function
