Attribute VB_Name = "modTabelaPrecos"
Option Explicit

Dim strColuna01 As String
Dim strColuna02 As String
Dim strColuna03 As String
Dim strColuna04 As String
Dim strColuna05 As String

Sub SetaGridCaptionTabela()

    Dim rs
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    rs.Open "SELECT * FROM descricao_tabela ORDER BY codigo_tabela;", CONN, adOpenForwardOnly, adLockReadOnly
    '
    frmTabela.GridCtrl1.Rows = 1
    frmTabela.GridCtrl1.Cols = 6
    '
    frmTabela.GridCtrl1.Row = 0
    frmTabela.GridCtrl1.Col = 0
    frmTabela.GridCtrl1.CellFontSize = 7
    frmTabela.GridCtrl1.CellBackColor = &HC0C0C0
    frmTabela.GridCtrl1.CellFontBold = True
    frmTabela.GridCtrl1.CellFontBold = flexAlignLeftCenter
    '
    frmTabela.GridCtrl1.Row = 0
    frmTabela.GridCtrl1.Col = 1
    frmTabela.GridCtrl1.CellFontSize = 7
    frmTabela.GridCtrl1.CellBackColor = &HC0C0C0
    frmTabela.GridCtrl1.CellFontBold = True
    frmTabela.GridCtrl1.CellFontBold = flexAlignLeftCenter
    '
    frmTabela.GridCtrl1.Row = 0
    frmTabela.GridCtrl1.Col = 2
    frmTabela.GridCtrl1.CellFontSize = 7
    frmTabela.GridCtrl1.CellBackColor = &HC0C0C0
    frmTabela.GridCtrl1.CellFontBold = True
    frmTabela.GridCtrl1.Row = 0
    frmTabela.GridCtrl1.Col = 3
    frmTabela.GridCtrl1.CellFontSize = 7
    frmTabela.GridCtrl1.CellBackColor = &HC0C0C0
    frmTabela.GridCtrl1.CellFontBold = True
    frmTabela.GridCtrl1.Row = 0
    frmTabela.GridCtrl1.Col = 4
    frmTabela.GridCtrl1.CellFontSize = 7
    frmTabela.GridCtrl1.CellBackColor = &HC0C0C0
    frmTabela.GridCtrl1.CellFontBold = True
    frmTabela.GridCtrl1.Row = 0
    frmTabela.GridCtrl1.Col = 5
    frmTabela.GridCtrl1.CellFontSize = 7
    frmTabela.GridCtrl1.CellBackColor = &HC0C0C0
    frmTabela.GridCtrl1.CellFontBold = True
    '
    frmTabela.GridCtrl1.ColWidth(0) = 2500
    frmTabela.GridCtrl1.ColWidth(1) = 700
    frmTabela.GridCtrl1.ColWidth(2) = 660
    frmTabela.GridCtrl1.ColWidth(3) = 660
    frmTabela.GridCtrl1.ColWidth(4) = 660
    frmTabela.GridCtrl1.ColWidth(5) = 660
    '
    frmTabela.GridCtrl1.TextMatrix(0, 0) = "Produto"
    '
    If rs.RecordCount > 0 Then
        strColuna01 = rs("codigo_tabela")
        frmTabela.GridCtrl1.TextMatrix(0, 1) = Mid(Trim(rs("descricao")), 1, 7)
    Else
        frmTabela.GridCtrl1.TextMatrix(0, 1) = "-"
    End If
    rs.MoveNext
    If Not rs.EOF Then
        strColuna02 = rs("codigo_tabela")
        frmTabela.GridCtrl1.TextMatrix(0, 2) = Mid(Trim(rs("descricao")), 1, 7)
    Else
        frmTabela.GridCtrl1.TextMatrix(0, 2) = "-"
    End If
    rs.MoveNext
    If Not rs.EOF Then
        strColuna03 = rs("codigo_tabela")
        frmTabela.GridCtrl1.TextMatrix(0, 3) = Mid(Trim(rs("descricao")), 1, 7)
    Else
        frmTabela.GridCtrl1.TextMatrix(0, 3) = "-"
    End If
    rs.MoveNext
    If Not rs.EOF Then
        strColuna04 = rs("codigo_tabela")
        frmTabela.GridCtrl1.TextMatrix(0, 4) = Mid(Trim(rs("descricao")), 1, 7)
    Else
        frmTabela.GridCtrl1.TextMatrix(0, 4) = "-"
    End If
    rs.MoveNext
    If Not rs.EOF Then
        strColuna05 = rs("codigo_tabela")
        frmTabela.GridCtrl1.TextMatrix(0, 5) = Mid(Trim(rs("descricao")), 1, 7)
    Else
        frmTabela.GridCtrl1.TextMatrix(0, 5) = "-"
    End If
    '
    If 1 = 0 Then
       '
    frmTabela.GridCtrl.Rows = 0
    frmTabela.GridCtrl.Cols = 6
    '
    frmTabela.GridCtrl.Row = 0
    frmTabela.GridCtrl.Col = 0
    frmTabela.GridCtrl.CellFontSize = 7
    frmTabela.GridCtrl.CellBackColor = &HC0C0C0
    frmTabela.GridCtrl.CellFontBold = True
    frmTabela.GridCtrl.CellFontBold = flexAlignLeftCenter
    '
    frmTabela.GridCtrl.Row = 0
    frmTabela.GridCtrl.Col = 1
    frmTabela.GridCtrl.CellFontSize = 7
    frmTabela.GridCtrl.CellBackColor = &HC0C0C0
    frmTabela.GridCtrl.CellFontBold = True
    frmTabela.GridCtrl.CellFontBold = flexAlignLeftCenter
    '
    frmTabela.GridCtrl.Row = 0
    frmTabela.GridCtrl.Col = 2
    frmTabela.GridCtrl.CellFontSize = 7
    frmTabela.GridCtrl.CellBackColor = &HC0C0C0
    frmTabela.GridCtrl.CellFontBold = True
    frmTabela.GridCtrl.Row = 0
    frmTabela.GridCtrl.Col = 3
    frmTabela.GridCtrl.CellFontSize = 7
    frmTabela.GridCtrl.CellBackColor = &HC0C0C0
    frmTabela.GridCtrl.CellFontBold = True
    frmTabela.GridCtrl.Row = 0
    frmTabela.GridCtrl.Col = 4
    frmTabela.GridCtrl.CellFontSize = 7
    frmTabela.GridCtrl.CellBackColor = &HC0C0C0
    frmTabela.GridCtrl.CellFontBold = True
    frmTabela.GridCtrl.Row = 0
    frmTabela.GridCtrl.Col = 5
    frmTabela.GridCtrl.CellFontSize = 7
    frmTabela.GridCtrl.CellBackColor = &HC0C0C0
    frmTabela.GridCtrl.CellFontBold = True
    '
    End If
    '
    frmTabela.GridCtrl.ColWidth(0) = 2500
    frmTabela.GridCtrl.ColWidth(1) = 700
    frmTabela.GridCtrl.ColWidth(2) = 660
    frmTabela.GridCtrl.ColWidth(3) = 660
    frmTabela.GridCtrl.ColWidth(4) = 660
    frmTabela.GridCtrl.ColWidth(5) = 660
    '
    ' frmTabela.GridCtrl.TextMatrix(0, 0) = "Produto"
    '
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    '
End Sub

Function EncheTabelaPrecos(ByVal letra) As Integer
  '
  EncheTabelaPrecos = 3
  '
  Screen.MousePointer = 11
  '
  frmTabela.GridCtrl1.Visible = False
  frmTabela.GridCtrl.Visible = False
  '
  frmTabela.cboDescricao.Visible = False
  '
  ' MsgBox "Agora inicio:" & Now, vbOKOnly + vbCritical, App.Title
  '
  Dim rsEncheTabelaPrecos, rsTabela
  Dim strFormataNumero As String
  Dim I As Integer
  '
  Set rsEncheTabelaPrecos = CreateObject("ADOCE.Recordset.3.0")
  Set rsTabela = CreateObject("ADOCE.Recordset.3.0")
  '
  connOpen
  '
  rsEncheTabelaPrecos.Open "SELECT * FROM produtos WHERE status='A' AND descricao like '" & letra & "%' ORDER BY descricao;", CONN, adOpenForwardOnly, adLockReadOnly
  '
  If rsEncheTabelaPrecos.RecordCount > 0 Then
     '
     ' frmTabela.cboDescricao.Clear
     '
     SetaGridCaptionTabela
     '
     I = 0
     '
     Do Until rsEncheTabelaPrecos.EOF
        '
        frmTabela.cboDescricao.AddItem rsEncheTabelaPrecos("descricao")
        '
        frmTabela.GridCtrl.Rows = frmTabela.GridCtrl.Rows + 1
        '
        frmTabela.GridCtrl.Row = frmTabela.GridCtrl.Rows - 1
        frmTabela.GridCtrl.Col = 0
        frmTabela.GridCtrl.CellFontSize = 8
        frmTabela.GridCtrl.CellFontBold = False
        '
        frmTabela.GridCtrl.TextMatrix(frmTabela.GridCtrl.Rows - 1, 0) = UCase(Mid(rsEncheTabelaPrecos("descricao"), 1, 1)) & LCase(Mid(rsEncheTabelaPrecos("descricao"), 2, 39))
        '
        '
        '=============================== Retira somente o numero da formatação ==================================
        '
        '
        frmTabela.GridCtrl.Col = 1
        frmTabela.GridCtrl.CellFontSize = 7
        frmTabela.GridCtrl.CellFontBold = False
        '
        frmTabela.GridCtrl.TextMatrix(frmTabela.GridCtrl.Rows - 1, 1) = _
        Mid( _
        FormatCurrency(CDbl(rsEncheTabelaPrecos("preco1")), 2, vbTrue, vbTrue, vbTrue) _
               , 3, Len( _
        FormatCurrency(CDbl(rsEncheTabelaPrecos("preco1")), 2, vbTrue, vbTrue, vbTrue) _
                ) _
                - 2)
        '
        frmTabela.GridCtrl.Col = 2
        frmTabela.GridCtrl.CellFontSize = 7
        frmTabela.GridCtrl.CellFontBold = False
        '
        frmTabela.GridCtrl.TextMatrix(frmTabela.GridCtrl.Rows - 1, 2) = _
        Mid( _
        FormatCurrency(CDbl(rsEncheTabelaPrecos("preco2")), 2, vbTrue, vbTrue, vbTrue) _
               , 3, Len( _
        FormatCurrency(CDbl(rsEncheTabelaPrecos("preco2")), 2, vbTrue, vbTrue, vbTrue) _
                ) _
                - 2)
        '
        frmTabela.GridCtrl.Col = 3
        frmTabela.GridCtrl.CellFontSize = 7
        frmTabela.GridCtrl.CellFontBold = False
        '
        frmTabela.GridCtrl.TextMatrix(frmTabela.GridCtrl.Rows - 1, 3) = _
        Mid( _
        FormatCurrency(CDbl(rsEncheTabelaPrecos("preco3")), 2, vbTrue, vbTrue, vbTrue) _
               , 3, Len( _
        FormatCurrency(CDbl(rsEncheTabelaPrecos("preco3")), 2, vbTrue, vbTrue, vbTrue) _
                ) _
                - 2)
        '
        frmTabela.GridCtrl.Col = 4
        frmTabela.GridCtrl.CellFontSize = 7
        frmTabela.GridCtrl.CellFontBold = False
        '
        frmTabela.GridCtrl.TextMatrix(frmTabela.GridCtrl.Rows - 1, 4) = _
        Mid( _
        FormatCurrency(CDbl(rsEncheTabelaPrecos("preco4")), 2, vbTrue, vbTrue, vbTrue) _
               , 3, Len( _
        FormatCurrency(CDbl(rsEncheTabelaPrecos("preco4")), 2, vbTrue, vbTrue, vbTrue) _
                ) _
                - 2)
        '
        frmTabela.GridCtrl.Col = 5
        frmTabela.GridCtrl.CellFontSize = 7
        frmTabela.GridCtrl.CellFontBold = False
        '
        frmTabela.GridCtrl.TextMatrix(frmTabela.GridCtrl.Rows - 1, 5) = _
        Mid( _
        FormatCurrency(CDbl(rsEncheTabelaPrecos("preco5")), 2, vbTrue, vbTrue, vbTrue) _
               , 3, Len( _
        FormatCurrency(CDbl(rsEncheTabelaPrecos("preco5")), 2, vbTrue, vbTrue, vbTrue) _
                ) _
                - 2)
        '
        '========================================================================================================
        '
        'frmTabela.GridCtrl.TextMatrix(frmTabela.GridCtrl.Rows - 1, 2) = rsEncheTabelaPrecos("preco2")
        'frmTabela.GridCtrl.TextMatrix(frmTabela.GridCtrl.Rows - 1, 3) = rsEncheTabelaPrecos("preco3")
        'frmTabela.GridCtrl.TextMatrix(frmTabela.GridCtrl.Rows - 1, 4) = rsEncheTabelaPrecos("preco4")
        'frmTabela.GridCtrl.TextMatrix(frmTabela.GridCtrl.Rows - 1, 5) = rsEncheTabelaPrecos("preco5")
        '
        rsEncheTabelaPrecos.MoveNext
        '
        I = I + 1
        '
        If I > rsEncheTabelaPrecos.RecordCount Then Exit Do
        '
     Loop
     '
     ' The Sort property is unavailable at design time. It is write-only at run time.
     '
     ' The Sort property always sorts entire rows. To specify the range to be sorted,
     ' set the Row and RowSel properties. If Row and RowSel are the same, the Grid control
     ' sorts all nonfixed rows.
     '
     ' The Col and ColSel properties determine the keys used for sorting. Keys are always
     ' sorted from left to right. For example, if Col = 3 and ColSel = 1, the sort is done
     ' according to the contents of column 1, then 2, then 3.
     '
     '
     frmTabela.GridCtrl.Row = frmTabela.GridCtrl.Rows - 1
     frmTabela.GridCtrl.RowSel = frmTabela.GridCtrl.Rows - 1
     '
     frmTabela.GridCtrl.Col = 0
     frmTabela.GridCtrl.ColSel = 0
     '
     frmTabela.GridCtrl.Sort = flexSortStringNoCaseAscending
     '
     EncheTabelaPrecos = 1
     '
  Else
     '
     MsgBox "Não existem dados a serem mostrados", vbOKOnly + vbCritical, App.Title
     '
     EncheTabelaPrecos = 2
     '
  End If
  '
  rsEncheTabelaPrecos.Close
  '
  connClose
  '
  Set rsEncheTabelaPrecos = Nothing
  '
  frmTabela.GridCtrl.Row = 0
  frmTabela.GridCtrl.Col = 0
  '
  ' MsgBox "Agora final:" & Now, vbOKOnly + vbCritical, App.Title
  '
  frmTabela.GridCtrl.Visible = True
  frmTabela.GridCtrl1.Visible = True
  '
  Screen.MousePointer = 0
  '
  frmTabela.cboDescricao.Visible = True
  '
End Function

