Attribute VB_Name = "modDatabaseCdPgto"
Option Explicit

Public Sub EncheFormularioCdPgto(strParametro As String)
    Screen.MousePointer = 11
    Dim rs
    Dim strCodigo As String
    '
    connOpen
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    '
    rs.Open "SELECT * FROM clientes WHERE codigo_cliente='" & Trim(strParametro) & "';", CONN, adOpenFowardOnly, adLockReadOnly
    '
    GridCtrl.Row = 0
    GridCtrl.Col = 0
    GridCtrl.CellBackColor = &HC0C0C0
    GridCtrl.CellFontBold = True
    GridCtrl.Row = 0
    GridCtrl.Col = 1
    GridCtrl.CellBackColor = &HC0C0C0
    GridCtrl.CellFontBold = True
    GridCtrl.Row = 0
    GridCtrl.Col = 2
    GridCtrl.CellBackColor = &HC0C0C0
    GridCtrl.CellFontBold = True
    GridCtrl.Row = 0
    GridCtrl.Col = 3
    GridCtrl.CellBackColor = &HC0C0C0
    GridCtrl.CellFontBold = True
    GridCtrl.Row = 0
    GridCtrl.Col = 4
    GridCtrl.CellBackColor = &HC0C0C0
    GridCtrl.CellFontBold = True
    GridCtrl.Row = 0
    GridCtrl.Col = 5
    GridCtrl.CellBackColor = &HC0C0C0
    GridCtrl.CellFontBold = True
    GridCtrl.ColWidth(0) = 2500
    GridCtrl.ColWidth(1) = 1500
    GridCtrl.ColWidth(2) = 1500
    GridCtrl.ColWidth(3) = 1500
    GridCtrl.ColWidth(4) = 1500
    GridCtrl.ColWidth(5) = 1500
    GridCtrl.TextMatrix(0, 0) = "Produto"
    GridCtrl.TextMatrix(0, 1) = "Tab1"
    GridCtrl.TextMatrix(0, 2) = "Tab2"
    GridCtrl.TextMatrix(0, 3) = "Tab3"
    GridCtrl.TextMatrix(0, 4) = "Tab4"
    GridCtrl.TextMatrix(0, 5) = "Tab5"
    '
    Do Until rs.EOF
       '
       frmContas.GridCtrl.Rows = frmContas.GridCtrl.Rows + 1
       frmTabela.GridCtrl.TextMatrix(frmContas.GridCtrl.Rows - 1, 0) = rs("numero_documento")
       valTotal = valTotal + CDbl(rs("valor"))
       '
       rs.MoveNext
       '
    Loop
    If rs.State = 1 Then rs.Close
    connClose
    Set rs = Nothing
    Screen.MousePointer = 0
End Sub
