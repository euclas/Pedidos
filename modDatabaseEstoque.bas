Attribute VB_Name = "modDatabaseEstoque"
Option Explicit

Public Sub EncheLstProdutos()
    Dim rs As adoce.Recordset
    Set rs = New adoce.Recordset
    connOpen
    rs.Open "SELECT codigo_produto,descricao FROM produtos ORDER BY codigo_produto"
    On Error Resume Next
    frmEstoque.lstProdutos.Clear
    Do Until rs.EOF
        frmEstoque.lstProdutos.AddItem rs("codigo_produto") & " - " & rs("descricao")
        rs.MoveNext
    Loop
    On Error GoTo 0
    connClose
    rs.Close
    Set rs = Nothing
End Sub
