Attribute VB_Name = "modMain"
Option Explicit
'
Const cstLarguraRelatorio = 50
Const cstLarguraObsCliente = 80
Const cstLarguraMensagem = 80
Const cstLarguraRosto = 80
'
Const opcClear = 0
Const opcInclusao = 1
Const opcAlteracao = 2
Const opcConsulta = 3
Const opcExclusao = 4
'
' &HC0FFC0 - VerdeClaro
'
' Const VerdeClaro = &HC0FFC0
'
' &HC0C0FF - Vermelho Claro
'
Const Branco = &HFFFFFF
Const VerdeClaro = &H80FF80
Const AmareloClaro = &H80FFFF
Const VermelhoClaro = &H8080FF    ' &HC0C0FF ' &H80C0FF
'
'
'
Const Verde = &HFF00&
Const Azul = &HFF0000
Const Vermelho = &HFF&
Const Amarelo = &HFFFF&
'
Public I As Integer
Public mLinhaGrid As Integer
Public mExecutou As Boolean
'
Public usrCodigoVendedor As String
Public usrCodigoCliente As String
Public usrCodigoPedido As String
Public usrTempoConectar As String
'
Public usrhabilitar_desconto_item As Boolean
Public usrhabilitar_desconto_pedido As Boolean
Public usrhabilitar_edicao_preco As Boolean
Public usrhabilitar_acrescimo As Boolean
'
Public usrhabilitar_cobranca_titulo As Boolean
'
Public bolModoManutencao As Boolean
Public bolMensagemStatus As Boolean
'
Public strPath As String
Public strPrograma As String
Public mStatusCliente As String
'
Public strContraSEnha As String
'
Public bolEntrada As Boolean
Public bolEstoque As Boolean
'
Public mPosicaoInicio As Integer
Public mPosicaoTamanho As Integer
Public mPosicao As Integer
Public mASCII As Integer
'
Dim mRoteiroExtra As Integer
Dim strFantasia As String
'
Public IntIncrDataEst As Integer
Public IntIncrDataPed As Integer
'
Public mProcessa As Boolean
'
' Variavel publica onde fica o Cliente Corrente/Ultimo Cliente
'
Public mUltimoCliente As String
'
' Constantes para as funcoes: GetWindowLong e SetWindowLong
'
Const GWL_STYLE = -16
Const ES_NUMBER = 8192
Const ES_UPPERCASE = 8
Const ES_LOWERCASE = 16
'
'
'
Declare Function SHFullScreen Lib "aygshell.dll" (ByVal hwndRequester As Long, ByVal dwState As Long) As Integer
'
Declare Function GetWindowLong Lib "Coredll" Alias "GetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "Coredll" Alias "SetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'
'
'
Sub Main()
    '
    bolMensagemStatus = True
    bolEstoque = False
    '
    bolEntrada = False
    '
    frmLogin.Show
    '
End Sub
'
Public Function VerificaUsuario(ByVal strUserName As String, ByVal strSenha As String) As Integer
   '
   Dim rs
   '
   '0 = usuário ok
   '1 = senha inválida
   '2 = usuário não cadastrado
   '4 = Tempo máximo ultrapassado
   '
   strUserName = LCase(strUserName)
   '
   '======== Apaga o input panel
   '
   SHFullScreen frmStart.hwnd, 8
   SHFullScreen frmStart.hwnd, 12
   '
   frmStart.Refresh
   '
   '
   'frmStart.Frame1.Visible = True
   '
   'frmStart.PictureBox1.Refresh
   '
   'frmStart.Refresh
   '
   If LCase(Trim(strUserName)) = "manutencao" Then
      '
      bolModoManutencao = True
      '
      '======== Apaga o input panel
      '
      SHFullScreen frmStart.hwnd, 8
      SHFullScreen frmStart.hwnd, 12
      '
      frmStart.Refresh
      '
      frmStart.Frame1.Visible = True
      '
      frmStart.Frame1.Refresh
      '
      frmStart.PictureBox1.Refresh
      '
      frmStart.Refresh
      '
      VerificaUsuario = 0
      '
   Else
      '
      '======== Apaga o input panel
      '
      SHFullScreen frmStart.hwnd, 8
      SHFullScreen frmStart.hwnd, 12
      '
      frmStart.Refresh
      '
      frmStart.Frame1.Visible = False
      '
      frmStart.Frame1.Refresh
      '
      frmStart.PictureBox1.Refresh
      '
      frmStart.Refresh
      '
      bolModoManutencao = False
      '
      '
      '======================== verifica se mesmo que o banco exista se existe usuário
      '
      '
      If frmStart.FileSystem.Dir(strPath & "\base.cdb") = "" Then
         '
         MsgBox "Banco de Dados Inexistente. Não pode processar.", vbCritical + vbOKOnly, App.Title
         '
         App.End
         '
      End If
      '
      connOpen
      '
      'On Error Resume Next
      '
      Set rs = CreateObject("ADOCE.Recordset.3.0")
      '
      rs.Open "SELECT * FROM vendedor;", CONN, adOpenStatic, adLockReadOnly
      '
      If rs.RecordCount = 0 Then
         '
         MsgBox "Não existe usuários para logar.", vbCritical + vbOKOnly, App.Title
         '
         App.End
         '
      End If
      '
      Set rs = CreateObject("ADOCE.Recordset.3.0")
      '
      strUserName = LCase(strUserName)
      '
      rs.Open "SELECT * FROM vendedor WHERE nome='" & Trim(strUserName) & "';", CONN, adOpenStatic, adLockReadOnly
      '
      If rs.RecordCount = 0 Then
          '
          VerificaUsuario = 2
          '
          Exit Function
          '
      Else
           If Trim(strSenha) = Trim(rs("senha")) Then
              '
              ' vendedor
              '
              ' codigo_vendedor VARCHAR(5)
              ' senha VARCHAR(6)
              ' nome VARCHAR(50)
              ' aceita_pedido_bloq VARCHAR(1)
              ' contra_senha VARCHAR(1)
              ' codigo_proximo_cliente VARCHAR(5)
              ' extra1 VARCHAR(1)
              ' extra2 VARCHAR(1)
              ' numero_proximo_pedido VARCHAR(8)
              ' tempo_maximo VARCHAR(3)
              ' habilitar_desconto_item VARCHAR(1)
              ' habilitar_desconto_pedido VARCHAR(1)
              ' habilitar_edicao_preco VARCHAR(1)
              ' habilitar_cobranca_titulo VARCHAR(1)
              ' empresa VARCHAR(2)
              ' filial VARCHAR(2)
              ' data_cortetitulos VARCHAR(8)
              ' mensagem TEXT
              ' status VARCHAR(1))
              '
               usrCodigoVendedor = rs("codigo_vendedor")
               usrCodigoCliente = rs("codigo_proximo_cliente")
               usrCodigoPedido = rs("numero_proximo_pedido")
               usrTempoConectar = rs("tempo_maximo")
               '
               usrhabilitar_desconto_item = False
               usrhabilitar_desconto_pedido = False
               usrhabilitar_edicao_preco = False
               usrhabilitar_acrescimo = False
               '
               usrhabilitar_cobranca_titulo = False
               '
               If rs("habilitar_desconto_item") = "S" Then usrhabilitar_desconto_item = True
               If rs("habilitar_desconto_pedido") = "S" Then usrhabilitar_desconto_pedido = True
               If rs("habilitar_edicao_preco") = "S" Then usrhabilitar_edicao_preco = True
               If rs("habilitar_acrescimo") = "S" Then usrhabilitar_acrescimo = True
               '
               If rs("habilitar_cobranca_titulo") = "S" Then usrhabilitar_cobranca_titulo = True
               '
               If frmLogin.FileSystem.Dir(strPath & "\last.txt") <> "" Then
                   '
                   Dim MyDate As Date
                   Dim MyDateAux As Date
                   Dim MyAux As String
                   '
                   frmLogin.File.Open strPath & "\last.txt", fsModeInput, fsAccessRead
                   '
                   MyAux = Trim(frmLogin.File.LineInputString)
                   '
                   frmLogin.File.Close
                   '
                   MyDateAux = DateSerial(Mid(MyAux, 5, 4), Mid(MyAux, 3, 2), Mid(MyAux, 1, 2))
                   '
                   If Len(usrTempoConectar) > 0 And IsNumeric(usrTempoConectar) Then
                       MyDate = Now - (CInt(usrTempoConectar) + 1)
                   Else
                       MyDate = Now - 1
                   End If
                   '
                   If MyDateAux < MyDate Then
                      frmValidarVendedor.txtCodigoClienteAtual.Text = "12345"
                      frmValidarVendedor.txtCodigoVendedorAtual.Text = usrCodigoVendedor
                      frmValidarVendedor.txtDataAtual.Text = Mid(RetornaDataString(Now), 1, 4) + Mid(RetornaDataString(Now), 7, 2)
                      VerificaUsuario = 4
                   Else
                       usrTempoConectar = CStr(CInt(MyDateAux - MyDate))
                       '
                       frmStart.txtMensagem.Text = rs("mensagem")
                       '
                       VerificaUsuario = 0
                       '
                   End If
                   '
               Else
                   '
                   '
                   VerificaUsuario = 0
                   '
               End If
           Else
               VerificaUsuario = 1
           End If
      End If
      '
      If rs.State = 1 Then rs.Close
      '
      Set rs = Nothing
      '
      connClose
      '
      On Error GoTo 0
      '
   End If
   '
End Function

Public Sub AcertaVendedor(valParametro As Integer)
    '
    '1- codigo_proximo_cliente++
    '2- numero_proximo_pedido++
    '
    Dim rs
    '
    connOpen
    '
    Set rs = CreateObject("ADOCE.Recordset.3.0")
    '
    rs.Open "SELECT * FROM vendedor WHERE codigo_vendedor='" & Trim(usrCodigoVendedor) & "';", CONN, adOpenDynamic, adLockOptimistic
    '
    Select Case valParametro
        Case 1
            rs("codigo_proximo_cliente") = usrCodigoCliente
        Case 2
            rs("numero_proximo_pedido") = usrCodigoPedido
    End Select
    '
    rs.Update
    '
    If rs.State = 1 Then rs.Close
    '
    connClose
    '
    Set rs = Nothing
    '
End Sub

Public Sub Foco(ByVal myTextBox As TextBox)
    '
    On Error Resume Next
    '
    myTextBox.SelStart = 0
    myTextBox.SelLength = Len(myTextBox.Text)
    '
End Sub

Public Sub setamenu(myMenu As MenuBar, ByVal myFrmAtual As Form)
    Dim FileMenu As MenuBarMenu
    Dim FormMenu As MenuBarMenu
    Set FileMenu = myMenu.Controls.AddMenu("Sistema")
    '
    If bolModoManutencao = True Then
       FileMenu.Items.Add , , "Sair"
    Else
        If Not myFrmAtual.Name = "frmClientes" Then FileMenu.Items.Add , , "Clientes"
        If Not myFrmAtual.Name = "frmContas" Then FileMenu.Items.Add , , "Contas a Receber"
        If Not myFrmAtual.Name = "frmSistematica" Then FileMenu.Items.Add , , "Sistemática de Visitação"
        If Not myFrmAtual.Name = "frmHistorico" Then FileMenu.Items.Add , , "Histórico de Pedidos"
        FileMenu.Items.Add , , , mbrMenuSeparator
        If Not myFrmAtual.Name = "frmTabela" Then FileMenu.Items.Add , , "Tabela de Preços"
        If Not myFrmAtual.Name = "frmEstoque" Then FileMenu.Items.Add , , "Contagem de Estoque"
        If Not myFrmAtual.Name = "frmPedido" Then FileMenu.Items.Add , , "Pedido"
        FileMenu.Items.Add , , , mbrMenuSeparator
        If Not myFrmAtual.Name = "frmMensagens" Then FileMenu.Items.Add , , "Mensagens"
        If Not myFrmAtual.Name = "frmRelatorios" Then FileMenu.Items.Add , , "Relatórios"
        FileMenu.Items.Add , , , mbrMenuSeparator
        FileMenu.Items.Add , , "Sair"
        Set FormMenu = myMenu.Controls.AddMenu("Opções")
        Select Case myFrmAtual.Name
            Case "frmClientes"
                FormMenu.Items.Add , , "Incluir Cliente"
                FormMenu.Items.Add , , "Editar Cliente"
                'FormMenu.Items.Add , , "Excluir Cliente"
                FormMenu.Items.Add , , , mbrMenuSeparator
            Case "frmPedido"
                FormMenu.Items.Add , , "Novo Pedido"
                FormMenu.Items.Add , , "Editar Pedido"
                FormMenu.Items.Add , , "Excluir Pedido"
                FormMenu.Items.Add , , , mbrMenuSeparator
            Case "frmMensagens"
                FormMenu.Items.Add , , "Salvar Nova Mensagem"
                FormMenu.Items.Add , , "Cancelar Nova Mensagem"
                FormMenu.Items.Add , , , mbrMenuSeparator
            Case "frmGerenciador"
                FormMenu.Items.Add , , "Exportar Arquivo"
                FormMenu.Items.Add , , , mbrMenuSeparator
            Case "frmEstoque"
                ' FormMenu.Items.Add , , "Salvar dados de Hoje"
                ' FormMenu.Items.Add , , "Adicionar ao Pedido"
                ' FormMenu.Items.Add , , , mbrMenuSeparator
        End Select
        If Not myFrmAtual.Name = "frmGerenciador" Then FormMenu.Items.Add , , "Gerenciar Arquivos"
        If Not myFrmAtual.Name = "frmGerenciador" Then FormMenu.Items.Add , , , mbrMenuSeparator
        FormMenu.Items.Add , , "Sobre"
    End If
End Sub

Public Sub ExecutaMenu(ByVal myCaption As String, ByVal myFrmAtual As Form)
    '
    Dim FileMenu As MenuBarMenu
    Dim retVal As Boolean
    '
    Select Case myCaption
    Case "Adicionar ao Pedido"
         '
         frmPedido.fraItens.Visible = True
         '
         If frmPedido.fraItens.Visible = True Then
            '
            frmPedido.cboProdutos.Text = frmEstoque.fraHistorico.Caption
            frmPedido.txtQtdPedida.Text = frmEstoque.txtHp.Text
            '
         Else
            '
            ExecutaMenu "Novo Pedido", frmPedido
            '
            frmPedido.cboPedidoCliente.Text = frmEstoque.cboClientes.Text
            frmPedido.cboProdutos.Text = frmEstoque.fraHistorico.Caption
            '
            frmPedido.txtQtdPedida.Text = frmEstoque.txtHp.Text
            '
         End If
         '
         Verificador = "2"
         '
         strNomeFantasia = frmPedido.cboPedidoCliente.Text
         '
         retVal = VerificaStatusPedido(frmPedido.cboProdutos.List(frmPedido.cboProdutos.ListIndex))
         '
         frmPedido.txtNomeProdutoPromo.Text = frmPedido.cboProdutos.List(frmPedido.cboProdutos.ListIndex)
         '
         frmPedido.txtPrecoUnit.Text = frmPedido.txtPrecoOriginal.Text
         '
         frmPedido.TabStrip.Tabs.Item(1).Selected = True
         '
    Case "Salvar dados de Hoje"
            If Len(frmEstoque.cboClientes.Text) <= 0 Then
                MsgBox "Selecione um cliente para poder cadastrar.", vbOKOnly + vbCritical, App.Title
                Exit Sub
            End If
            If frmEstoque.GridCtrlProduto.Row = 0 Then
                MsgBox "Selecione um produto para poder o cadastrar.", vbOKOnly + vbCritical, App.Title
                Exit Sub
            End If
            If Not IsNumeric(frmEstoque.txtHe.Text) Then
                MsgBox "Digite um valor válido para o estoque atual.", vbOKOnly + vbCritical, App.Title
                Exit Sub
            End If
            Screen.MousePointer = 11
            Dim strCodigoCliente As String
            Dim strCodigoProduto As String
            Dim bolPrimeiro As String
            Dim rs
            connOpen
            Set rs = CreateObject("ADOCE.Recordset.3.0")
            rs.Open "SELECT * FROM clientes WHERE nome_fantasia='" & Trim(frmEstoque.cboClientes.Text) & "';", CONN, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount > 0 Then
                strCodigoCliente = rs("codigo_cliente")
            Else
                MsgBox "Cliente inválido ou não cadastrado.", vbOKOnly + vbCritical, App.Title
                Screen.MousePointer = 0
                Exit Sub
            End If
            If rs.State = 1 Then rs.Close
            strCodigoProduto = frmEstoque.GridCtrlProduto.TextMatrix(frmEstoque.GridCtrlProduto.Row, 0)
            rs.Open "SELECT * FROM estoque WHERE codigo_cliente='" & Trim(strCodigoCliente) & "' AND codigo_produto='" & strCodigoProduto & "' ORDER BY data;", CONN, adOpenDynamic, adLockOptimistic
            If rs.RecordCount >= 8 Then
                Do Until rs.RecordCount = 7
                    rs.MoveFirst
                    rs.Delete
                Loop
            End If
            If rs.RecordCount = 0 Then
                bolPrimeiro = True
            Else
                bolPrimeiro = False
            End If
            rs.AddNew
            rs("codigo_cliente") = strCodigoCliente
            rs("codigo_produto") = strCodigoProduto
            rs("data") = RetornaDataString(Now)
            rs("estoque") = frmEstoque.txtHe.Text
            If bolPrimeiro Then
                rs("sugestao") = "0"
            Else
                rs("sugestao") = "-"
            End If
            rs.Update
            If rs.State = 1 Then rs.Close
            connClose
            Set rs = Nothing
            Screen.MousePointer = 0
            MsgBox "Dados de hoje adicionados com sucesso.", vbOKOnly + vbInformation, App.Title
            '
        Case "Cancelar Alterações no Pedido"
            '
            frmPedido.MenuBar.Controls.Clear
            setamenu frmPedido.MenuBar, frmPedido
            '
            LimpaFormularioPEdido
            '
            frmPedido.TabStrip.Tabs.Item(1).Selected = True
            '
        Case "Salvar Alterações no Pedido"
            '
            Screen.MousePointer = 11
            '
            If EditarPedido Then
               '
               MsgBox "Pedido alterado com sucesso.", vbOKOnly + vbInformation, App.Title
               '
               frmPedido.MenuBar.Controls.Clear
               '
               setamenu frmPedido.MenuBar, frmPedido
               '
               LimpaFormularioPEdido
               '
            Else
               '
               MsgBox "Ocorreu um erro ao editar o pedido.", vbCritical + vbOKOnly, App.Title
               '
            End If
            '
            Screen.MousePointer = 0
            '
            frmPedido.TabStrip.Tabs.Item(1).Selected = True
            '
        Case "Editar Pedido"
            '
            frmPedido.txtDescItem.Enabled = True
            '
            frmPedido.txtDesconto.Enabled = True
            '
            frmPedido.txtPrecoUnit.Enabled = True
            '
            frmPedido.txtAcrescimo.Enabled = True
            '
            If usrhabilitar_desconto_item = False Then frmPedido.txtDescItem.Enabled = False
            '
            If usrhabilitar_desconto_pedido = False Then frmPedido.txtDesconto.Enabled = False
            '
            If usrhabilitar_edicao_preco = False Then frmPedido.txtPrecoUnit.Enabled = False
            '
            If usrhabilitar_acrescimo = False Then frmPedido.txtAcrescimo.Enabled = False
            '
            If Len(frmPedido.txtNumeroPedido.Text) > 0 Then
               '
               Screen.MousePointer = 11
               frmPedido.MenuBar.Controls.Clear
               Set FileMenu = frmPedido.MenuBar.Controls.AddMenu("Opções")
               FileMenu.Items.Add , , "Salvar Alterações no Pedido"
               FileMenu.Items.Add , , "Cancelar Alterações no Pedido"
               '
               connOpen
               '
               EncheComboClientes frmPedido.cboPedidoCliente, 2
               '
               EncheCombosPedidos
               '
               connClose
               '
               frmPedido.cboCPagto.Visible = True
               frmPedido.cboFPagto.Visible = True
               frmPedido.cboPedidoCliente.Visible = True
               frmPedido.cboTmov.Visible = True
               '
               frmPedido.cboCPagto.Text = frmPedido.txtCPagto.Text
               frmPedido.cboFPagto.Text = frmPedido.txtFPgto.Text
               frmPedido.cboPedidoCliente.Text = frmPedido.txtCliente.Text
               frmPedido.cboTmov.Text = frmPedido.txtTMvto.Text
               '
               frmPedido.txtQtdPedida.Text = 0 ' vbNullString
               frmPedido.txtPrecoUnit.Text = 0 ' vbNullString
               frmPedido.txtDescItem.Text = 0 ' vbNullString
               frmPedido.txtTotal.Text = 0 ' vbNullString
               '
               Screen.MousePointer = 0
               '
               frmPedido.TabStrip.Tabs.Item(1).Selected = True
               '
            Else
               MsgBox "Selecione um pedido no histórico de pedidos para poder realizar esta operação.", vbOKOnly + vbCritical, App.Title
            End If
            '
        Case "Excluir Pedido"
            '
            If Len(frmPedido.txtNumeroPedido.Text) > 0 Then
                If MsgBox("Deseja realmente excluir este pedido ?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                    If ExcluiPedido Then
                        MsgBox "Pedido excluido com sucesso.", vbOKOnly + vbInformation, App.Title
                    Else
                        MsgBox "Não foi possível excluir este pedido.", vbOKOnly + vbCritical, App.Title
                    End If
                End If
            Else
                MsgBox "Selecione um pedido no histórico de pedidos para poder realizar esta operação.", vbOKOnly + vbCritical, App.Title
            End If
            '
            frmPedido.TabStrip.Tabs.Item(1).Selected = True
            '
        Case "Exportar Arquivo"
            '
            GravaArquivos
            '
        Case "Sair"
            '
            If MsgBox("Deseja realmente encerrar o sistema ?", vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
            '
            App.End
            '
        Case "Sobre"
            '
            frmAbout.Show vbModal, myFrmAtual
            '
        Case "Clientes"
            '
            If myFrmAtual.Name = "frmCliente" Then Exit Sub
            Screen.MousePointer = 11
            frmClientes.Show
            myFrmAtual.Hide
            '
        Case "Contas a Receber"
            '
            If myFrmAtual.Name = "frmContas" Then Exit Sub
            Screen.MousePointer = 11
            frmContas.Show
            myFrmAtual.Hide
            '
        Case "Contagem de Estoque"
            '
            If myFrmAtual.Name = "frmEstoque" Then Exit Sub
            Screen.MousePointer = 11
            frmEstoque.Show
            myFrmAtual.Hide
            '
        Case "Gerenciar Arquivos"
            '
            If myFrmAtual.Name = "frmGerenciador" Then Exit Sub
            Screen.MousePointer = 11
            frmGerenciador.Show
            myFrmAtual.Hide
            '
        Case "Sistemática de Visitação"
            '
            If myFrmAtual.Name = "frmSistematica" Then Exit Sub
            Screen.MousePointer = 11
            frmSistematica.Show
            myFrmAtual.Hide
            '
        Case "Histórico de Pedidos"
            '
            If myFrmAtual.Name = "frmHistorico" Then Exit Sub
            Screen.MousePointer = 11
            frmHistorico.Show
            myFrmAtual.Hide
            '
        Case "Tabela de Preços"
            '
            'If MsgBox("Este procedimento poderá demorar alguns minutos na primeira carga dependendo da quantidade de produtos cadastrados. Deseja continuar ?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                '
                If myFrmAtual.Name = "frmTabela" Then Exit Sub
                Screen.MousePointer = 11
                '
                frmTabela.Show
                '
                myFrmAtual.Hide
                '
            'End If
            '
        Case "Pedido"
            '
            If myFrmAtual.Name = "frmPedido" Then Exit Sub
            '
            Screen.MousePointer = 11
            '
            frmPedido.Show
            '
            myFrmAtual.Hide
            '
        Case "Mensagens"
            '
            If myFrmAtual.Name = "frmMensagens" Then Exit Sub
            Screen.MousePointer = 11
            frmMensagens.Show
            myFrmAtual.Hide
            '
        Case "Relatórios"
            '
            If myFrmAtual.Name = "frmRelatorios" Then Exit Sub
            Screen.MousePointer = 11
            frmRelatorios.Show
            myFrmAtual.Hide
            '
        Case "Incluir Cliente"
            '
            Screen.MousePointer = 11
            '
            LimpaControlesCliente opcClear
            '
            LimpaControlesCliente opcInclusao
            '
            frmClientes.MenuBar.Controls.Clear
            '
            Set FileMenu = frmClientes.MenuBar.Controls.AddMenu("Opções")
            '
            FileMenu.Items.Add , , "Salvar Novo Cliente"
            FileMenu.Items.Add , , "Cancelar Novo Cliente"
            '
            usrCodigoCliente = CStr(CInt(usrCodigoCliente) + 1)
            '
            Select Case Len(Trim(usrCodigoCliente))
            Case 1
                 usrCodigoCliente = "0000" & Trim(usrCodigoCliente)
            Case 2
                 usrCodigoCliente = "000" & Trim(usrCodigoCliente)
            Case 3
                 usrCodigoCliente = "00" & Trim(usrCodigoCliente)
            Case 4
                 usrCodigoCliente = "0" & Trim(usrCodigoCliente)
            Case 5
                 usrCodigoCliente = Trim(usrCodigoCliente)
            Case Else
                 usrCodigoCliente = Left(usrCodigoCliente, 5)
            End Select
            '
            frmClientes.cboCodigo.Text = usrCodigoCliente
            frmClientes.txtCodigo.Text = usrCodigoCliente
            '
            frmClientes.TabStrip.Tabs.Item(1).Selected = True
            '
            frmClientes.txtFantasia.SetFocus
            '
            Screen.MousePointer = 0
            '
        Case "Editar Cliente"
            '
            If Len(Trim(frmClientes.cboCodigo.Text)) <= 0 Then
               '
               MsgBox "É necessário selecionar um cliente para executar essa operação.", vbOKOnly + vbCritical, App.Title
               '
               Exit Sub
               '
            End If
            '
            LimpaControlesCliente opcAlteracao
            '
            frmClientes.txtCodigo.Text = frmClientes.cboCodigo.Text
            frmClientes.txtFantasia.Text = frmClientes.cboFantasia.Text
            frmClientes.txtRazaoSocial.Text = frmClientes.cboRSocial.Text
            '
            frmClientes.cboCidadeCobranca.Text = frmClientes.txtCobrancaCidade.Text
            frmClientes.cboCidadeEndereco.Text = frmClientes.txtCidade.Text
            frmClientes.cboAtividade.Text = frmClientes.txtAtividade.Text
            '
            frmClientes.MenuBar.Controls.Clear
            '
            Set FileMenu = frmClientes.MenuBar.Controls.AddMenu("Opções")
            '
            FileMenu.Items.Add , , "Salvar Alterações do Cliente"
            FileMenu.Items.Add , , "Cancelar Alterações do Cliente"
            '
            frmClientes.TabStrip.Tabs.Item(1).Selected = True
            '
            frmClientes.txtCNPJMF.SetFocus
            '
        Case "Cancelar Novo Cliente"
            '
            Screen.MousePointer = 11
            frmClientes.MenuBar.Controls.Clear
            setamenu frmClientes.MenuBar, frmClientes
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
            LimpaControlesCliente opcClear
            '
            LimpaControlesCliente opcConsulta
            '
            Screen.MousePointer = 0
            '
        Case "Salvar Novo Cliente"
            '
            'If Len(Trim(frmClientes.txtCNPJMF.Text)) > 3 Then
            '    If Not VerificaCGC(frmClientes.txtCNPJMF.Text) Then
            '        MsgBox "CNPJ Inválido !!!", vbCritical + vbOKOnly, App.Title
            '        Exit Sub
            '    End If
            'End If
            '
            'If Len(Trim(frmClientes.txtCPF.Text)) > 3 Then
            '    If Not VerificaCPF(frmClientes.txtCPF.Text) Then
            '        MsgBox "CPF Inválido !!!", vbCritical + vbOKOnly, App.Title
            '        Exit Sub
            '    End If
            'End If
            '
            If (Trim(frmClientes.txtFantasia.Text) = "") Then
               '
               MsgBox "Nome Fantasia inválido:(" & Trim(frmClientes.txtFantasia.Text) & ")", vbOKOnly + vbCritical, App.Title
               '
               Exit Sub
               '
            End If
            '
            If (Trim(frmClientes.txtRazaoSocial.Text) = "") Then
               '
               MsgBox "Razão Social inválida:(" & Trim(frmClientes.txtRazaoSocial.Text) & ")", vbOKOnly + vbCritical, App.Title
               '
               Exit Sub
               '
            End If
            '
            mStatusCliente = "D"
            '
            retVal = IncluiCliente
            '
            If retVal = True Then
               '
               MsgBox "Cliente cadastrado com sucesso !!!", vbOKOnly + vbInformation, App.Title
               '
            Else
               '
               MsgBox "Ocorreu um erro ao tentar cadastrar o cliente.", vbOKOnly + vbCritical, App.Title
               '
            End If
            '
            Screen.MousePointer = 11
            '
            frmClientes.MenuBar.Controls.Clear
            '
            setamenu frmClientes.MenuBar, frmClientes
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
            LimpaControlesCliente opcClear
            '
            LimpaControlesCliente opcConsulta
            '
            Screen.MousePointer = 0
            '
        Case "Excluir Cliente"
            '
            ExcluiCliente
            '
        Case "Novo Pedido"
            '
            frmPedido.txtDescItem.Enabled = True
            '
            frmPedido.txtDesconto.Enabled = True
            '
            frmPedido.txtPrecoUnit.Enabled = True
            '
            frmPedido.txtAcrescimo.Enabled = True
            '
            If usrhabilitar_desconto_item = False Then frmPedido.txtDescItem.Enabled = False
            '
            If usrhabilitar_desconto_pedido = False Then frmPedido.txtDesconto.Enabled = False
            '
            If usrhabilitar_edicao_preco = False Then frmPedido.txtPrecoUnit.Enabled = False
            '
            If usrhabilitar_acrescimo = False Then frmPedido.txtAcrescimo.Enabled = False
            '
            Screen.MousePointer = 11
            '
            frmPedido.MenuBar.Controls.Clear
            '
            Set FileMenu = frmPedido.MenuBar.Controls.AddMenu("Opções")
            '
            FileMenu.Items.Add , , "Salvar Novo Pedido"
            FileMenu.Items.Add , , "Cancelar Novo Pedido"
            '
            connOpen
            '
            EncheComboClientes frmPedido.cboPedidoCliente, 2
            '
            EncheCombosPedidos
            '
            connClose
            '
            LimpaFormularioPEdido
            '
            frmPedido.cboCPagto.Visible = True
            frmPedido.cboFPagto.Visible = True
            frmPedido.cboPedidoCliente.Visible = True
            frmPedido.cboTmov.Visible = True
            '
            frmPedido.txtQtdPedida.Text = 0 ' vbNullString
            frmPedido.txtPrecoUnit.Text = 0 ' vbNullString
            frmPedido.txtDescItem.Text = 0 ' vbNullString
            frmPedido.txtTotal.Text = 0 ' vbNullString
            '
            usrCodigoPedido = usrCodigoPedido + 1
            '
            Screen.MousePointer = 0
            '
            frmPedido.TabStrip.Tabs.Item(1).Selected = True
            '
            frmPedido.txtNumeroPedido.Text = usrCodigoPedido
            '
            frmPedido.txtEntrega.Text = Mid(RetornaDataString(CDate(Now)), 1, 2) & "/" & Mid(RetornaDataString(CDate(Now)), 3, 2) & "/" & Mid(RetornaDataString(CDate(Now)), 5, 4)
            '
        Case "Salvar Novo Pedido"
            '
            Screen.MousePointer = 11
            '
            If IncluiPedido Then
               '
               MsgBox "Pedido incluso com sucesso.", vbOKOnly + vbInformation, App.Title
               '
               AcertaVendedor 2
               '
            Else
               '
               MsgBox "Ocorreu um erro ao cadastrar o pedido.", vbCritical + vbOKOnly, App.Title
               '
            End If
            '
            frmPedido.MenuBar.Controls.Clear
            setamenu frmPedido.MenuBar, frmPedido
            '
            LimpaFormularioPEdido
            '
            Screen.MousePointer = 0
            '
            frmPedido.TabStrip.Tabs.Item(1).Selected = True
            '
        Case "Cancelar Novo Pedido"
            '
            frmPedido.MenuBar.Controls.Clear
            '
            setamenu frmPedido.MenuBar, frmPedido
            '
            LimpaFormularioPEdido
            '
            frmPedido.TabStrip.Tabs.Item(1).Selected = True
            '
        Case "Cancelar Alterações do Cliente"
            '
            Screen.MousePointer = 11
            '
            LimpaControlesCliente opcClear
            '
            LimpaControlesCliente opcConsulta
            '
            frmClientes.MenuBar.Controls.Clear
            setamenu frmClientes.MenuBar, frmClientes
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
            Screen.MousePointer = 0
            '
        Case "Salvar Alterações do Cliente"
            '
            'If Len(Trim(frmClientes.txtCNPJMF.Text)) > 3 Then
            '    If Not VerificaCGC(frmClientes.txtCNPJMF.Text) Then
            '        MsgBox "CNPJ Inválido !!!", vbCritical + vbOKOnly, App.Title
            '        Exit Sub
            '    End If
            'End If
            'If Len(Trim(frmClientes.txtCPF.Text)) > 3 Then
            '    If Not VerificaCPF(frmClientes.txtCPF.Text) Then
            '        MsgBox "CPF Inválido !!!", vbCritical + vbOKOnly, App.Title
            '        Exit Sub
            '    End If
            'End If
            '
            Dim msai As Boolean
            '
            If (Trim(frmClientes.txtFantasia.Text) = "") Then
               '
               MsgBox "Nome Fantasia inválido:(" & Trim(frmClientes.txtFantasia.Text) & ")", vbOKOnly + vbCritical, App.Title
               '
               Exit Sub
               '
            End If
            '
            If (Trim(frmClientes.txtRazaoSocial.Text) = "") Then
               '
               MsgBox "Razão Social inválida:(" & Trim(frmClientes.txtRazaoSocial.Text) & ")", vbOKOnly + vbCritical, App.Title
               '
               Exit Sub
               '
            End If
            '
            msai = ValidaDataCliente(frmClientes.txtDataFundacao.Text)
            '
            If (Trim(frmClientes.txtDataFundacao.Text) = "") Or (Trim(frmClientes.txtDataFundacao.Text) = "**/**/****") Or (Trim(frmClientes.txtDataFundacao.Text) = "//") Then msai = True
            '
            If msai = False Then
               '
               MsgBox "(1) - Data de Fundação não é uma Data Válida:(" & frmClientes.txtDataFundacao.Text & ")" & CInt(mValida), vbOKOnly + vbCritical, App.Title
               '
               OpenCloseCalendar Now, frmClientes.txtDataFundacao.Text
               '
            Else
               '
               mStatusCliente = "M"
               '
               retVal = EditarCliente(frmClientes.txtCodigo.Text)
               '
               If retVal = True Then
                  '
                  MsgBox "Cliente alterado com sucesso !!!", vbOKOnly + vbInformation, App.Title
                  '
               Else
                  MsgBox "Ocorreu um erro ao tentar alterar o cliente.", vbOKOnly + vbCritical, App.Title
               End If
               '
               Screen.MousePointer = 11
               '
               frmClientes.MenuBar.Controls.Clear
               '
               setamenu frmClientes.MenuBar, frmClientes
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
               Screen.MousePointer = 0
               '
               LimpaControlesCliente opcClear
               '
               LimpaControlesCliente opcConsulta
               '
               frmClientes.TabStrip.Tabs(1).Selected = True
               '
            End If
            '
        Case "Salvar Nova Mensagem"
            '
            If Len(frmMensagens.txtNAssunto.Text) <= 0 Then
               '
               MsgBox "Digite o assunto da mensagem.", vbOKOnly + vbCritical, App.Title
               '
               frmMensagens.txtNAssunto.Visible = True
               frmMensagens.txtNAssunto.ZOrder vbBringToFront ' vbSendToBack
               '
               Exit Sub
               '
            End If
            '
            If Len(frmMensagens.txtNMensagem.Text) <= 0 Then
               '
               MsgBox "Digite o texto da mensagem.", vbOKOnly + vbCritical, App.Title
               '
               Exit Sub
               '
            End If
            '
            If SalvaNovaMensagem = False Then
               '
               MsgBox "Ocorreu um erro ao cadastrar a mensagem para o vendedor.", vbOKOnly + vbCritical, App.Title
               '
            Else
               '
               MsgBox "Mensagem cadastrada com sucesso.", vbOKOnly + vbInformation, App.Title
               '
            End If
            '
            On Error Resume Next
            '
            frmMensagens.cboNDestinatario.Text = vbNullString
            frmMensagens.txtNAssunto.Text = vbNullString
            frmMensagens.txtNMensagem.Text = vbNullString
            '
            frmMensagens.cboNAssunto.Clear
            '
            frmMensagens.cboNAssunto.Visible = True
            '
            On Error GoTo 0
            '
        Case "Cancelar Nova Mensagem"
            '
            On Error Resume Next
            '
            frmMensagens.cboNDestinatario.Text = vbNullString
            frmMensagens.txtNAssunto.Text = vbNullString
            frmMensagens.txtNMensagem.Text = vbNullString
            frmMensagens.cboNAssunto.Clear
            frmMensagens.cboNAssunto.Visible = True
            '
            On Error GoTo 0
            '
    End Select
End Sub

Public Sub AddStatus(ByVal MyStatus As String)
    If bolMensagemStatus Then
       If Trim(frmStart.txtStatus.Text) <> Trim(MyStatus) Then frmStart.txtStatus.Text = MyStatus
    Else
       If Trim(frmGerenciador.txtStatus.Text) <> Trim(MyStatus) Then frmGerenciador.txtStatus.Text = MyStatus
    End If
    '
End Sub

Public Sub Progresso(ByVal intProgress As Integer, ByVal intComprimento As Integer, ByVal intTamanho As Integer, ByVal intPassagem As Integer)
   '
   If intProgress = 1 Then
      '
      ' fazendo um progressbar com um picturebox.
      '
      ' PictureBox1.DrawLine 0, 0, X, H, Color, True, True
      '
      ' Onde:
      '
      ' H é = a altura da PictureBox.
      ' x é = numero de pontos do comprimento da barra.
      '
      frmStart.PictureBox1.DrawLine 0, 0, intComprimento, 350, &HFF00&, True, True
      '
      frmStart.PictureBox1.DrawLine intComprimento + 1, 0, intTamanho, 350, &HFF&, True, True
      '
      frmStart.PictureBox1.ForeColor = &H80000005
      '
      frmStart.PictureBox1.DrawText CStr(intPassagem) & " %", 1600, 0
      '
      frmStart.Frame1.Refresh
      '
      frmStart.PictureBox1.Refresh
      '
    End If
    '
    If intProgress = 2 Then
      '
      frmGerenciador.PictureBox1.DrawLine 0, 0, intComprimento, 350, &HFF00&, True, True
      '
      frmGerenciador.PictureBox1.DrawLine intComprimento + 1, 0, intTamanho, 350, &HFF&, True, True
      '
      frmGerenciador.PictureBox1.ForeColor = &H80000005
      '
      frmGerenciador.PictureBox1.DrawText CStr(intPassagem) & " %", 1600, 0
      '
      frmGerenciador.Frame1.Refresh
      '
      frmGerenciador.PictureBox1.Refresh
      '
    End If
    '
End Sub

Public Sub LimpaControlesCliente(opcStatus As Integer)
    '
    On Error Resume Next
    '
    Select Case opcStatus
    '
    Case opcConsulta
         '
         ' Pagina Dados
         '
         frmClientes.TabStrip.Tabs.Item(1).Selected = True
         '
         frmClientes.cboCodigo.Visible = True
         frmClientes.cboFantasia.Visible = True
         frmClientes.cboRSocial.Visible = True
         '
         frmClientes.cboCodigo.Enabled = True
         frmClientes.cboFantasia.Enabled = True
         frmClientes.cboRSocial.Enabled = True
         '
         frmClientes.txtCodigo.Visible = True
         frmClientes.txtCodigo.Enabled = False
         '
         frmClientes.txtFantasia.Visible = False
         frmClientes.txtRazaoSocial.Visible = False
         '
         frmClientes.txtFantasia.Enabled = False
         frmClientes.txtRazaoSocial.Enabled = False
         '
         frmClientes.txtCNPJMF.Enabled = False
         frmClientes.txtIEST.Enabled = False
         frmClientes.txtCPF.Enabled = False
         frmClientes.txtRG.Enabled = False
         '
         ' Pagina Endereco
         '
         frmClientes.TabStrip.Tabs.Item(2).Selected = True
         '
         frmClientes.txtEntrega.Enabled = False
         frmClientes.txtBairro.Enabled = False
         frmClientes.cboCidadeEndereco.Visible = False
         frmClientes.cboCidadeEndereco.Enabled = False
         '
         frmClientes.txtCidade.Visible = True
         frmClientes.txtCidade.Enabled = False
         '
         frmClientes.txtCEP.Enabled = False
         frmClientes.txtContato.Enabled = False
         frmClientes.txtEmail.Enabled = False
         frmClientes.txtWWW.Enabled = False
         '
         ' Pagina Cobranca
         '
         frmClientes.TabStrip.Tabs.Item(3).Selected = True
         '
         frmClientes.txtCobranca.Enabled = False
         frmClientes.txtCobrancaBairro.Enabled = False
         '
         frmClientes.cboCidadeCobranca.Visible = False
         frmClientes.cboCidadeCobranca.Enabled = False
         '
         frmClientes.txtCobrancaCidade.Visible = True
         frmClientes.txtCobrancaCidade.Enabled = False
         '
         frmClientes.txtCobrancaCEP.Enabled = False
         '
         frmClientes.txtTelefone.Enabled = False
         frmClientes.txtFax.Enabled = False
         '
         ' Pagina Referencias
         '
         frmClientes.TabStrip.Tabs.Item(4).Selected = True
         '
         frmClientes.cboAtividade.Visible = False
         frmClientes.cboAtividade.Enabled = False
         '
         frmClientes.txtAtividade.Visible = True
         frmClientes.txtAtividade.Enabled = False
         '
         frmClientes.optNao.Enabled = False
         frmClientes.optSim.Enabled = False
         '
         frmClientes.txtRefBanc01.Enabled = False
         frmClientes.txtRefBanc02.Enabled = False
         frmClientes.txtRefPess01.Enabled = False
         frmClientes.txtRefPess02.Enabled = False
         '
         frmClientes.txtDataFundacao.Enabled = False
         '
         ' Pagina Observações
         '
         frmClientes.TabStrip.Tabs.Item(5).Selected = True
         '
         frmClientes.txtObs.Enabled = False
         '
         frmClientes.TabStrip.Tabs.Item(1).Selected = True
         '
    Case opcInclusao
         '
         ' Pagina Dados
         '
         frmClientes.TabStrip.Tabs.Item(1).Selected = True
         '
         frmClientes.cboCodigo.Visible = False
         frmClientes.cboFantasia.Visible = False
         frmClientes.cboRSocial.Visible = False
         '
         frmClientes.cboCodigo.Enabled = False
         frmClientes.cboFantasia.Enabled = False
         frmClientes.cboRSocial.Enabled = False
         '
         frmClientes.txtCodigo.Visible = True
         frmClientes.txtCodigo.Enabled = False
         '
         frmClientes.txtFantasia.Visible = True
         frmClientes.txtRazaoSocial.Visible = True
         '
         frmClientes.txtFantasia.Enabled = True
         frmClientes.txtRazaoSocial.Enabled = True
         '
         frmClientes.txtCNPJMF.Enabled = True
         frmClientes.txtIEST.Enabled = True
         frmClientes.txtCPF.Enabled = True
         frmClientes.txtRG.Enabled = True
         '
         ' Pagina Endereco
         '
         frmClientes.TabStrip.Tabs.Item(2).Selected = True
         '
         frmClientes.txtEntrega.Enabled = True
         frmClientes.txtBairro.Enabled = True
         '
         frmClientes.cboCidadeEndereco.Visible = True
         frmClientes.cboCidadeEndereco.Enabled = True
         '
         frmClientes.txtCidade.Visible = False
         frmClientes.txtCidade.Enabled = False
         '
         frmClientes.txtCEP.Enabled = True
         frmClientes.txtContato.Enabled = True
         frmClientes.txtEmail.Enabled = True
         frmClientes.txtWWW.Enabled = True
         '
         ' Pagina Cobranca
         '
         frmClientes.TabStrip.Tabs.Item(3).Selected = True
         '
         frmClientes.txtCobranca.Enabled = True
         frmClientes.txtCobrancaBairro.Enabled = True
         '
         frmClientes.cboCidadeCobranca.Visible = True
         frmClientes.cboCidadeCobranca.Enabled = True
         '
         frmClientes.txtCobrancaCidade.Visible = False
         frmClientes.txtCobrancaCidade.Enabled = False
         '
         frmClientes.txtCobrancaCEP.Enabled = True
         '
         frmClientes.txtTelefone.Enabled = True
         frmClientes.txtFax.Enabled = True
         '
         ' Pagina Referencias
         '
         frmClientes.TabStrip.Tabs.Item(4).Selected = True
         '
         frmClientes.cboAtividade.Visible = True
         frmClientes.cboAtividade.Enabled = True
         '
         frmClientes.txtAtividade.Visible = False
         frmClientes.txtAtividade.Enabled = False
         '
         frmClientes.optNao.Enabled = True
         frmClientes.optSim.Enabled = True
         '
         frmClientes.txtRefBanc01.Enabled = True
         frmClientes.txtRefBanc02.Enabled = True
         frmClientes.txtRefPess01.Enabled = True
         frmClientes.txtRefPess02.Enabled = True
         '
         frmClientes.txtDataFundacao.Enabled = True
         '
         ' Pagina Observações
         '
         frmClientes.TabStrip.Tabs.Item(5).Selected = True
         '
         frmClientes.txtObs.Enabled = True
         '
         frmClientes.TabStrip.Tabs.Item(1).Selected = True
         '
    Case opcAlteracao
         '
         ' Pagina Dados
         '
         frmClientes.TabStrip.Tabs.Item(1).Selected = True
         '
         frmClientes.cboCodigo.Visible = False
         frmClientes.cboFantasia.Visible = False
         frmClientes.cboRSocial.Visible = False
         '
         frmClientes.cboCodigo.Enabled = False
         frmClientes.cboFantasia.Enabled = False
         frmClientes.cboRSocial.Enabled = False
         '
         frmClientes.txtCodigo.Visible = True
         frmClientes.txtCodigo.Enabled = False
         '
         frmClientes.txtFantasia.Visible = True
         frmClientes.txtRazaoSocial.Visible = True
         '
         frmClientes.txtFantasia.Enabled = True
         frmClientes.txtRazaoSocial.Enabled = True
         '
         frmClientes.txtCNPJMF.Enabled = True
         frmClientes.txtIEST.Enabled = True
         frmClientes.txtCPF.Enabled = True
         frmClientes.txtRG.Enabled = True
         '
         ' Pagina Endereco
         '
         frmClientes.TabStrip.Tabs.Item(2).Selected = True
         '
         frmClientes.txtEntrega.Enabled = True
         frmClientes.txtBairro.Enabled = True
         '
         frmClientes.cboCidadeEndereco.Visible = True
         frmClientes.cboCidadeEndereco.Enabled = True
         '
         frmClientes.txtCidade.Visible = False
         frmClientes.txtCidade.Enabled = False
         '
         frmClientes.txtCEP.Enabled = True
         frmClientes.txtContato.Enabled = True
         frmClientes.txtEmail.Enabled = True
         frmClientes.txtWWW.Enabled = True
         '
         ' Pagina Cobranca
         '
         frmClientes.TabStrip.Tabs.Item(3).Selected = True
         '
         frmClientes.txtCobranca.Enabled = True
         frmClientes.txtCobrancaBairro.Enabled = True
         '
         frmClientes.cboCidadeCobranca.Visible = True
         frmClientes.cboCidadeCobranca.Enabled = True
         '
         frmClientes.txtCobrancaCidade.Visible = False
         frmClientes.txtCobrancaCidade.Enabled = False
         '
         frmClientes.txtCobrancaCEP.Enabled = True
         '
         frmClientes.txtTelefone.Enabled = True
         frmClientes.txtFax.Enabled = True
         '
         ' Pagina Referencias
         '
         frmClientes.TabStrip.Tabs.Item(4).Selected = True
         '
         frmClientes.cboAtividade.Visible = True
         frmClientes.cboAtividade.Enabled = True
         '
         frmClientes.txtAtividade.Visible = False
         frmClientes.txtAtividade.Enabled = False
         '
         frmClientes.optNao.Enabled = True
         frmClientes.optSim.Enabled = True
         '
         frmClientes.txtRefBanc01.Enabled = True
         frmClientes.txtRefBanc02.Enabled = True
         frmClientes.txtRefPess01.Enabled = True
         frmClientes.txtRefPess02.Enabled = True
         '
         frmClientes.txtDataFundacao.Enabled = True
         '
         ' Pagina Observações
         '
         frmClientes.TabStrip.Tabs.Item(5).Selected = True
         '
         frmClientes.txtObs.Enabled = True
         '
         frmClientes.TabStrip.Tabs.Item(1).Selected = True
         '
    Case opcClear
         '
         frmClientes.txtCodigo.Text = ""
         '
         frmClientes.cboCodigo.Text = ""
         frmClientes.cboFantasia.Text = ""
         frmClientes.cboRSocial.Text = ""
    '
         frmClientes.txtFantasia.Text = ""
         frmClientes.txtRazaoSocial.Text = ""
         frmClientes.txtBairro.Text = ""
         frmClientes.txtCEP.Text = ""
         frmClientes.txtCidade.Text = ""
         frmClientes.txtCNPJMF.Text = ""
         frmClientes.txtCobranca.Text = ""
         frmClientes.txtCobrancaBairro.Text = ""
         frmClientes.txtCobrancaCEP.Text = ""
         frmClientes.txtCobrancaCidade.Text = ""
         frmClientes.txtContato.Text = ""
         frmClientes.txtCPF.Text = ""
         frmClientes.txtDataFundacao.Text = ""
         frmClientes.txtEmail.Text = ""
         frmClientes.txtEmail.Text = ""
         frmClientes.txtEntrega.Text = ""
         frmClientes.txtFax.Text = ""
         frmClientes.txtIEST.Text = ""
         frmClientes.txtObs.Text = ""
         frmClientes.txtRefBanc01.Text = ""
         frmClientes.txtRefBanc02.Text = ""
         frmClientes.txtRefPess01.Text = ""
         frmClientes.txtRefPess02.Text = ""
         frmClientes.txtRG.Text = ""
         frmClientes.txtTelefone.Text = ""
         frmClientes.txtWWW.Text = ""
         frmClientes.optNao.Value = False
         frmClientes.optSim.Value = True
         '
         frmClientes.txtCidade.Text = ""
         frmClientes.txtCidadeCobranca.Text = ""
         frmClientes.txtAtividade.Text = ""
         '
         frmClientes.TabStrip.Tabs.Item(1).Selected = True
         '
    End Select
    '
    On Error GoTo 0
    '
End Sub

Public Function MontaNomeArquivo(ByVal strMontaNomeArquivo As String) As String
    '
    If bolModoManutencao = False Then
       '
         Dim strDia As String
         Dim strMes As String
         Dim strAno As String
         Dim intNumeroAtual As Integer
         '
         strDia = CStr(Day(Now))
         strMes = CStr(Month(Now))
         strAno = Right(CStr(Year(Now)), 2)
         '
         If Len(Trim(strDia)) = 1 Then strDia = "0" & Trim(strDia)
         If Len(Trim(strMes)) = 1 Then strMes = "0" & Trim(strMes)
         '
         If frmStart.FileSystem.Dir(strPath & "\" & Trim(strAno) & Trim(strMes) & Trim(strDia) & ".in") <> "" Then
            '
            frmLogin.File.Open strPath & "\" & Trim(strAno) & Trim(strMes) & Trim(strDia) & ".in", fsModeInput, fsAccessRead, fsLockReadWrite
            intNumeroAtual = CInt(frmLogin.File.LineInputString)
            frmLogin.File.Close
            '
            'frmStart.FileSystem.Kill App.Path & "\" & Trim(strAno) & Trim(strMes) & Trim(strDia) & ".in"
            '
         Else
            intNumeroAtual = 0
         End If
         '
         MontaNomeArquivo = Trim(strMontaNomeArquivo) & "\R" & Trim(strAno) & Trim(strMes) & Trim(strDia) & CStr(intNumeroAtual) & "." & Trim(Right(usrCodigoVendedor, 3))
         '
         If frmStart.FileSystem.Dir(MontaNomeArquivo) <> "" Then
             frmLogin.File.Open strPath & "\" & Trim(strAno) & Trim(strMes) & Trim(strDia) & ".in", fsModeOutput, fsAccessWrite, fsLockReadWrite
             frmLogin.File.LinePrint CStr(intNumeroAtual + 1)
             frmLogin.File.Close
         End If
         '
    Else
        MontaNomeArquivo = Trim(strMontaNomeArquivo) & "\I0000000.000"
    End If
    '
End Function

Public Function RetornaDataString(ByVal TheDate As Date) As String
    Dim strDay As String
    Dim strMonth As String
    Dim strYear As String
    '
    strDay = CStr(Day(TheDate))
    strMonth = CStr(Month(TheDate))
    strYear = CStr(Year(TheDate))
    '
    If Len(Trim(strDay)) = 1 Then strDay = "0" & Trim(strDay)
    '
    If Len(Trim(strMonth)) = 1 Then strMonth = "0" & Trim(strMonth)
    '
    RetornaDataString = Trim(strDay) & Trim(strMonth) & Trim(strYear)
    '
End Function

Public Function RetornaHoraString(ByVal TheHora As String, TheMinuto As String, TheSegundo As String) As String
  '
  ' RetornaHoraString(CStr(Hour(Now)), CStr(Minute(Now)), CStr(Second(Now)))
  '
  Select Case Len(TheHora)
  Case 0
       TheHora = "00"
  Case 1
       TheHora = "0" & TheHora
  Case 2
  Case 3
  Case 4
  Case 5
  End Select
  '
  Select Case Len(TheMinuto)
  Case 0
       TheHora = "00"
  Case 1
       TheMinuto = "0" & TheMinuto
  Case 2
  Case 3
  Case 4
  Case 5
  End Select
  '
  Select Case Len(TheSegundo)
  Case 0
       TheHora = "00"
  Case 1
       TheSegundo = "0" & TheSegundo
  Case 2
  Case 3
  Case 4
  Case 5
  End Select
  '
  RetornaHoraString = Trim(TheHora) & Trim(TheMinuto) & Trim(TheSegundo)
  '
End Function

Public Function Primos() As Boolean
   Dim A(118) As String, PE(200) As Integer
   Dim V As String, C As String, P As String, D As String
   Dim D1 As String, D2 As String, D3 As String
   Dim K As String, W As String, X As String, Y As String, Z As String
   Dim O As String
   Dim Q As String, R As String, S As String, T As String
   Dim PW As String
   Dim E As Integer
   Dim I As Integer
   Dim C0 As Integer, C1 As Integer, C2 As Integer, C3 As Integer, C4 As Integer
   Dim C5 As Integer, C6 As Integer, C7 As Integer, C8 As Integer, C9 As Integer
   Dim PRIMO As Integer, M As Integer
   '
   A(1) = "H ": A(2) = "HE": A(3) = "LI": A(4) = "BE": A(5) = "B ": A(6) = "C ": A(7) = "N "
   A(8) = "0 ": A(9) = "F ": A(10) = "NE": A(11) = "NA": A(12) = "MG": A(13) = "AL"
   A(14) = "SI": A(15) = "P ": A(16) = "S ": A(17) = "CL": A(18) = "AR": A(19) = "K "
   A(20) = "CA": A(21) = "SC": A(22) = "TI": A(23) = "V ": A(24) = "CR": A(25) = "MN"
   A(26) = "FE": A(27) = "CO": A(28) = "NI": A(29) = "CU": A(30) = "ZN": A(31) = "GA"
   A(32) = "GE": A(33) = "AS": A(34) = "SE": A(35) = "BR": A(36) = "KR": A(37) = "RB"
   A(38) = "SR": A(39) = "Y ": A(40) = "ZR": A(41) = "NB": A(42) = "MO": A(43) = "TC"
   A(44) = "RU": A(45) = "RH": A(46) = "PD": A(47) = "AG": A(48) = "CD": A(49) = "IN"
   A(50) = "SN": A(51) = "SB": A(52) = "TE": A(53) = "I ": A(54) = "XE": A(55) = "CS"
   A(56) = "BA": A(57) = "LA": A(58) = "CE": A(59) = "PR": A(60) = "ND": A(61) = "PM"
   A(62) = "SM": A(63) = "EU": A(64) = "GD": A(65) = "TB": A(66) = "DY": A(67) = "HO"
   A(68) = "ER": A(69) = "TM": A(70) = "YB": A(71) = "LU": A(72) = "HF": A(73) = "TA"
   A(74) = "W ": A(75) = "RE": A(76) = "OS": A(77) = "IR": A(78) = "PT": A(79) = "AU"
   A(80) = "HG": A(81) = "TL": A(82) = "PB": A(83) = "BI": A(84) = "PO": A(85) = "AT"
   A(86) = "RN": A(87) = "FR": A(88) = "RA": A(89) = "AC": A(90) = "TH": A(91) = "PA"
   A(92) = "U": A(93) = "NP": A(94) = "PU": A(95) = "AM": A(96) = "CM": A(97) = "BK"
   A(98) = "CF": A(99) = "ES": A(100) = "FM": A(101) = "MD": A(102) = "NO": A(103) = "LR"
   A(104) = "RF": A(105) = "DB": A(106) = "SG": A(107) = "BH": A(108) = "HS": A(109) = "MT"
   A(110) = "UUN": A(111) = "UUU": A(112) = "UUB": A(113) = "UUT": A(114) = "UUQ": A(115) = "UUP"
   A(116) = "UUH": A(117) = "UUS": A(118) = "UUO"
   '
   S = ""
   R = ""
   '
   ' ==================================================================
   ' ENTRADA DE: ULTIMO ALGARISMO DO ANO + MES + DIA FORMATO: AMMDD
   ' ==================================================================
   '
   ' V = Vendedor / C = Cliente / P = Pedido / D = Data
   '
   V = Right("00000" + Trim(frmValidaCliente.txtCodigoVendedorAtual.Text), 5)
   C = Right("00000" + Trim(frmValidaCliente.txtCodigoClienteAtual.Text), 5)
   P = Right("00000" + Trim(frmValidaCliente.txtNumeroPedido.Text), 5)
   D = Right("000000" + Trim(frmValidaCliente.txtDataAtual.Text), 6)
   '
   frmValidaCliente.txtCodigoVendedorAtual.Text = V
   frmValidaCliente.txtCodigoClienteAtual.Text = C
   frmValidaCliente.txtNumeroPedido.Text = P
   '
   frmValidaCliente.Label7.Caption = V
   frmValidaCliente.Label8.Caption = C
   frmValidaCliente.Label9.Caption = P
   '
   D1 = Mid(frmValidaCliente.txtDataAtual.Text, 6, 1)
   D2 = Mid(frmValidaCliente.txtDataAtual.Text, 3, 2)
   D3 = Mid(frmValidaCliente.txtDataAtual.Text, 1, 2)
   '
   D = D1 + D2 + D3
   '
   frmValidaCliente.Label10.Caption = D
   '
   ' ==================================
   ' TRANSPOE A MATRIZ
   ' ==================================
   '
   K = Mid(D, 5, 1) + Mid(P, 5, 1) + Mid(C, 5, 1) + Mid(V, 5, 1)
   W = Mid(D, 4, 1) + Mid(P, 4, 1) + Mid(C, 4, 1) + Mid(V, 4, 1)
   X = Mid(D, 3, 1) + Mid(P, 3, 1) + Mid(C, 3, 1) + Mid(V, 3, 1)
   Y = Mid(D, 2, 1) + Mid(P, 2, 1) + Mid(C, 2, 1) + Mid(V, 2, 1)
   Z = Mid(D, 1, 1) + Mid(P, 1, 1) + Mid(C, 1, 1) + Mid(V, 1, 1)
   '
   O = K + W + X + Y + Z
   '
   C2 = 0
   '
   ' =========================
   '   CRIA MATRIZ DE PRIMOS
   ' =========================
   '
   For C0 = 1 To 530
       '
       ' PRIMO = 1 SE O NUMERO EM C0 FOR PRIMO
       '
       PRIMO = 1
       '
       For C1 = 2 To (C0 - 1)
           '
           ' PRIMO = 0 SE NUMERO EM C0 NÃO FOR PRIMO...
           '
           If Int(C0 / C1) = C0 / C1 Then PRIMO = 0: Exit For 'Isso serve para abandonar o Loop...
           '
       Next C1
       '
       ' ACUMULA NA MATRIZ PE OS PRIMEIROS 100 NUMEROS PRIMOS ENCONTRADOS
       '
       If PRIMO = 1 Then
          '
          C2 = C2 + 1: PE(C2) = C0
          '
       End If
       '
       If C2 >= 100 Then Exit For
       '
   Next C0
   '
   '
   'For C8 = 1 To 100
   '    '
   '     MsgBox CStr(C8) + " = " + CStr(PE(C8))
   '    '
   'Next C8
   '
   ' ============================================
   '    TROCA CARACTERES 2 A 2 COM SEUS PRIMOS
   ' ============================================
   '
   E = 1
   '
   For C4 = 1 To Len(O) Step 2
       '
       Q = Mid(O, C4, 2)
       '
       If Q = "00" Then
          '
          T = "Extra Ordem=" + CStr(E) + "  X=" + Q + " Primo=" + CStr(PE(CInt(Q)))
          '
          ' MsgBox T
          '
          S = S + " 00"
          R = R + " 00"
          '
       Else
          '
          T = "Ordem=" + CStr(E) + "  X=" + Q + " Primo=" + CStr(PE(CInt(Q)))
          '
          ' MsgBox T
          '
          S = S + " " + CStr(PE(CInt(Q)))
          '
          R = R + " " + A(CInt(Q))
          '
       End If
       '
       E = E + 1
       '
   Next C4
   '
   '=======================================================================
   '     MOSTRA SENHA E CONTRA-SENHA
   ' A SENHA E CALCULADA E MOSTRADA, A CONTRA-SENHA E SOMENTE CALCULADA
   ' UMA VEZ TRANSMITIDO A SENHA PARA A EMPRESA O PROGRAMA DESKTOP VAI
   ' CALCULAR A CONTRA-SENHA COM BASE NA SENHA E O RESULTADO SERA PASSADO
   ' PARA O USUARIO DO POCKET QUE DIGITARA A CONTRA-SENHA E ASSIM LIBERANDO
   ' O FORMULARIO PARA A DIGITACAO DE NOVO PEDIDO.
   '=======================================================================
   '
   frmValidaCliente.txtSenhaGerada.Text = S
   '
   PW = 0 ': VARIAVEL ONDE SERA COLOCADA A CONTRA-SENHA
   '
   For C7 = 1 To Len(R)
       '
       PW = PW + Asc(Mid(R, C7, 1)) * Asc(Mid(S, C7, 1))
       '
   Next C7
   '
   frmValidaCliente.txtContraSenha.Text = PW
   '
   strContraSEnha = PW
   '
End Function

Public Function Primos02() As Boolean
   '
   Dim A(118) As String, PE(200) As Integer
   Dim V As String, C As String, P As String, D As String
   Dim D1 As String, D2 As String, D3 As String
   Dim K As String, W As String, X As String, Y As String, Z As String
   Dim O As String
   Dim Q As String, R As String, S As String, T As String
   Dim PW As String
   Dim E As Integer
   Dim I As Integer
   Dim C0 As Integer, C1 As Integer, C2 As Integer, C3 As Integer, C4 As Integer
   Dim C5 As Integer, C6 As Integer, C7 As Integer, C8 As Integer, C9 As Integer
   Dim PRIMO As Integer, M As Integer
   '
   A(1) = "H ": A(2) = "HE": A(3) = "LI": A(4) = "BE": A(5) = "B ": A(6) = "C ": A(7) = "N "
   A(8) = "0 ": A(9) = "F ": A(10) = "NE": A(11) = "NA": A(12) = "MG": A(13) = "AL"
   A(14) = "SI": A(15) = "P ": A(16) = "S ": A(17) = "CL": A(18) = "AR": A(19) = "K "
   A(20) = "CA": A(21) = "SC": A(22) = "TI": A(23) = "V ": A(24) = "CR": A(25) = "MN"
   A(26) = "FE": A(27) = "CO": A(28) = "NI": A(29) = "CU": A(30) = "ZN": A(31) = "GA"
   A(32) = "GE": A(33) = "AS": A(34) = "SE": A(35) = "BR": A(36) = "KR": A(37) = "RB"
   A(38) = "SR": A(39) = "Y ": A(40) = "ZR": A(41) = "NB": A(42) = "MO": A(43) = "TC"
   A(44) = "RU": A(45) = "RH": A(46) = "PD": A(47) = "AG": A(48) = "CD": A(49) = "IN"
   A(50) = "SN": A(51) = "SB": A(52) = "TE": A(53) = "I ": A(54) = "XE": A(55) = "CS"
   A(56) = "BA": A(57) = "LA": A(58) = "CE": A(59) = "PR": A(60) = "ND": A(61) = "PM"
   A(62) = "SM": A(63) = "EU": A(64) = "GD": A(65) = "TB": A(66) = "DY": A(67) = "HO"
   A(68) = "ER": A(69) = "TM": A(70) = "YB": A(71) = "LU": A(72) = "HF": A(73) = "TA"
   A(74) = "W ": A(75) = "RE": A(76) = "OS": A(77) = "IR": A(78) = "PT": A(79) = "AU"
   A(80) = "HG": A(81) = "TL": A(82) = "PB": A(83) = "BI": A(84) = "PO": A(85) = "AT"
   A(86) = "RN": A(87) = "FR": A(88) = "RA": A(89) = "AC": A(90) = "TH": A(91) = "PA"
   A(92) = "U": A(93) = "NP": A(94) = "PU": A(95) = "AM": A(96) = "CM": A(97) = "BK"
   A(98) = "CF": A(99) = "ES": A(100) = "FM": A(101) = "MD": A(102) = "NO": A(103) = "LR"
   A(104) = "RF": A(105) = "DB": A(106) = "SG": A(107) = "BH": A(108) = "HS": A(109) = "MT"
   A(110) = "UUN": A(111) = "UUU": A(112) = "UUB": A(113) = "UUT": A(114) = "UUQ": A(115) = "UUP"
   A(116) = "UUH": A(117) = "UUS": A(118) = "UUO"
   '
   ' ==================================================================
   ' ENTRADA DE: ULTIMO ALGARISMO DO ANO + MES + DIA FORMATO: AMMDD
   ' ==================================================================
   '
   ' V = Vendedor / C = Cliente / P = Pedido / D = Data
   '
   S = ""
   R = ""
   '
   V = Right("00000" + Trim(frmValidarVendedor.txtCodigoVendedorAtual.Text), 5)
   C = Right("00000" + Trim(frmValidarVendedor.txtCodigoClienteAtual.Text), 5)
   P = Right("00000" + Trim(frmValidarVendedor.txtNumeroPedido.Text), 5)
   '
   D = Right("000000" + Trim(frmValidarVendedor.txtDataAtual.Text), 6)
   '
   frmValidarVendedor.txtCodigoVendedorAtual.Text = V
   frmValidarVendedor.txtCodigoClienteAtual.Text = C
   frmValidarVendedor.txtNumeroPedido.Text = P
   '
   frmValidarVendedor.Label7.Caption = V
   frmValidarVendedor.Label8.Caption = C
   frmValidarVendedor.Label9.Caption = P
   '
   D1 = Mid(frmValidarVendedor.txtDataAtual.Text, 6, 1)
   D2 = Mid(frmValidarVendedor.txtDataAtual.Text, 3, 2)
   D3 = Mid(frmValidarVendedor.txtDataAtual.Text, 1, 2)
   '
   D = D1 + D2 + D3
   '
   frmValidarVendedor.Label10.Caption = D
   '
   ' ==================================
   ' TRANSPOE A MATRIZ
   ' ==================================
   '
   K = Mid(D, 5, 1) + Mid(P, 5, 1) + Mid(C, 5, 1) + Mid(V, 5, 1)
   W = Mid(D, 4, 1) + Mid(P, 4, 1) + Mid(C, 4, 1) + Mid(V, 4, 1)
   X = Mid(D, 3, 1) + Mid(P, 3, 1) + Mid(C, 3, 1) + Mid(V, 3, 1)
   Y = Mid(D, 2, 1) + Mid(P, 2, 1) + Mid(C, 2, 1) + Mid(V, 2, 1)
   Z = Mid(D, 1, 1) + Mid(P, 1, 1) + Mid(C, 1, 1) + Mid(V, 1, 1)
   '
   O = K + W + X + Y + Z
   '
   C2 = 0
   '
   ' =========================
   '   CRIA MATRIZ DE PRIMOS
   ' =========================
   '
   For C0 = 1 To 530
       '
       ' PRIMO = 1 SE O NUMERO EM C0 FOR PRIMO
       '
       PRIMO = 1
       '
       For C1 = 2 To (C0 - 1)
           '
           ' PRIMO = 0 SE NUMERO EM C0 NÃO FOR PRIMO...
           '
           If Int(C0 / C1) = C0 / C1 Then PRIMO = 0: Exit For 'Isso serve para abandonar o Loop...
           '
       Next C1
       '
       ' ACUMULA NA MATRIZ PE OS PRIMEIROS 100 NUMEROS PRIMOS ENCONTRADOS
       '
       If PRIMO = 1 Then
          '
          C2 = C2 + 1: PE(C2) = C0
          '
       End If
       '
       If C2 >= 100 Then Exit For
       '
   Next C0
   '
   '
   'For C8 = 1 To 100
   '    '
   '     MsgBox CStr(C8) + " = " + CStr(PE(C8))
   '    '
   'Next C8
   '
   ' ============================================
   '    TROCA CARACTERES 2 A 2 COM SEUS PRIMOS
   ' ============================================
   '
   E = 1
   '
   For C4 = 1 To Len(O) Step 2
       '
       Q = Mid(O, C4, 2)
       '
       If Q = "00" Then
          '
          T = "Extra Ordem=" + CStr(E) + "  X=" + Q + " Primo=" + CStr(PE(CInt(Q)))
          '
          ' MsgBox T
          '
          S = S + " 00"
          R = R + " 00"
          '
       Else
          '
          T = "Ordem=" + CStr(E) + "  X=" + Q + " Primo=" + CStr(PE(CInt(Q)))
          '
          ' MsgBox T
          '
          S = S + " " + CStr(PE(CInt(Q)))
          '
          R = R + " " + A(CInt(Q))
          '
       End If
       '
       E = E + 1
       '
   Next C4
   '
   '=======================================================================
   '     MOSTRA SENHA E CONTRA-SENHA
   ' A SENHA E CALCULADA E MOSTRADA, A CONTRA-SENHA E SOMENTE CALCULADA
   ' UMA VEZ TRANSMITIDO A SENHA PARA A EMPRESA O PROGRAMA DESKTOP VAI
   ' CALCULAR A CONTRA-SENHA COM BASE NA SENHA E O RESULTADO SERA PASSADO
   ' PARA O USUARIO DO POCKET QUE DIGITARA A CONTRA-SENHA E ASSIM LIBERANDO
   ' O FORMULARIO PARA A DIGITACAO DE NOVO PEDIDO.
   '=======================================================================
   '
   frmValidarVendedor.txtSenhaGerada.Text = S
   '
   PW = 0 ': VARIAVEL ONDE SERA COLOCADA A CONTRA-SENHA
   '
   For C7 = 1 To Len(R)
       '
       PW = PW + Asc(Mid(R, C7, 1)) * Asc(Mid(S, C7, 1))
       '
   Next C7
   '
   frmValidarVendedor.txtCodigoClienteAtual.Text = PW
   '
   frmValidarVendedor.txtContraSenha.Text = PW
   '
   strContraSEnha = PW
   '
End Function

Public Function VerificaCGC(strCGC As String) As Boolean
    Dim D1 As Integer, d4 As Integer, xx As Integer, nCount As Integer, fator As Integer, digito1 As Integer, digito2 As Integer
    Dim resto As Double
    Dim Check As String
    D1 = 0
    d4 = 0
    xx = 1
    For nCount = 1 To (Len(strCGC) - 2)
        If xx < 5 Then
            fator = 6 - xx
        Else
            fator = 14 - xx
        End If
        D1 = D1 + CInt(Mid(strCGC, nCount, 1)) * fator
        If xx < 6 Then
            fator = 7 - xx
        Else
            fator = 15 - xx
        End If
        d4 = d4 + CInt(Mid(strCGC, nCount, 1)) * fator
        xx = xx + 1
    Next
    resto = D1 Mod 11
    If resto < 2 Then
        digito1 = 0
    Else
        digito1 = 11 - resto
    End If
    d4 = d4 + 2 * digito1
    resto = d4 Mod 11
    If resto < 2 Then
        digito2 = 0
    Else
        digito2 = 11 - resto
    End If
    Check = CStr(digito1) & CStr(digito2)
    If Trim(Check) <> Mid(strCGC, Len(strCGC) - 1, 2) Then
        VerificaCGC = False
    Else
        VerificaCGC = True
    End If
End Function

Public Function VerificaCPF(strCPF As String) As Boolean
    Dim D1 As Integer, d4 As Integer, xx As Integer, nCount As Integer, resto As Integer, digito1 As Integer, digito2 As Integer
    Dim Check As String
    strCPF = Trim(strCPF)
    D1 = 0
    d4 = 0
    xx = 1
    For nCount = 1 To Len(strCPF) - 2
        If Mid(strCPF, nCount, 1) <> "-" And Mid(strCPF, nCount, 1) <> "." Then
            D1 = D1 + (11 - xx) * CInt(Mid(strCPF, nCount, 1))
            d4 = d4 + (12 - xx) * CInt(Mid(strCPF, nCount, 1))
            xx = xx + 1
        End If
    Next
    resto = D1 Mod 11
    If resto < 2 Then
       digito1 = 0
    Else
       digito1 = 11 - resto
    End If
    d4 = d4 + 2 * digito1
    resto = d4 Mod 11
    If resto < 2 Then
       digito2 = 0
    Else
       digito2 = 11 - resto
    End If
    Check = CStr(digito1) & CStr(digito2)
    If Check <> Mid(strCPF, Len(strCPF) - 1, 2) Then
       VerificaCPF = False
    Else
       VerificaCPF = True
    End If
End Function

Function OnlyNumericKeys(KeyAscii As Integer)
  '
  ' MsgBox "Key=" & CStr(KeyAscii)
  '
  Select Case KeyAscii
  Case 8, 44, 46, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 61
       '
       'KeyAscii=8 is a Backspace
       'KeyAscii=44 is a ","
       'KeyAscii=46 is a "."
       'KeyAscii=61 is a "="
       '
       If KeyAscii = 46 Then
          '
          OnlyNumericKeys = 44
          '
       Else
          '
          OnlyNumericKeys = KeyAscii
          '
       End If
       '
  Case Else
       '
       KeyAscii = 0 'Reject everything else
       '
       OnlyNumericKeys = 0
       '
  End Select
  '
End Function

Public Sub Elementos(valParametro As Integer)
   '
   A(1) = "H "
   A(2) = "HE"
   A(3) = "LI"
   A(4) = "BE"
   A(5) = "B "
   A(6) = "C "
   A(7) = "N "
   A(8) = "0 "
   A(9) = "F "
   A(10) = "NE"
   A(11) = "NA"
   A(12) = "MG"
   A(13) = "AL"
   A(14) = "SI"
   A(15) = "P "
   A(16) = "S "
   A(17) = "CL"
   A(18) = "AR"
   A(19) = "K "
   A(20) = "CA"
   A(21) = "SC"
   A(22) = "TI"
   A(23) = "V "
   A(24) = "CR"
   A(25) = "MN"
   A(26) = "FE"
   A(27) = "CO"
   A(28) = "NI"
   A(29) = "CU"
   A(30) = "ZN"
   A(31) = "GA"
   A(32) = "GE"
   A(33) = "AS"
   A(34) = "SE"
   A(35) = "BR"
   A(36) = "KR"
   A(37) = "RB"
   A(38) = "SR"
   A(39) = "Y "
   A(40) = "ZR"
   A(41) = "NB"
   A(42) = "MO"
   A(43) = "TC"
   A(44) = "RU"
   A(45) = "RH"
   A(46) = "PD"
   A(47) = "AG"
   A(48) = "CD"
   A(49) = "IN"
   A(50) = "SN"
   A(51) = "SB"
   A(52) = "TE"
   A(53) = "I "
   A(54) = "XE"
   A(55) = "CS"
   A(56) = "BA"
   A(57) = "LA"
   A(58) = "CE"
   A(59) = "PR"
   A(60) = "ND"
   A(61) = "PM"
   A(62) = "SM"
   A(63) = "EU"
   A(64) = "GD"
   A(65) = "TB"
   A(66) = "DY"
   A(67) = "HO"
   A(68) = "ER"
   A(69) = "TM"
   A(70) = "YB"
   A(71) = "LU"
   A(72) = "HF"
   A(73) = "TA"
   A(74) = "W"
   A(75) = "RE"
   A(76) = "OS"
   A(77) = "IR"
   A(78) = "PT"
   A(79) = "AU"
   A(80) = "HG"
   A(81) = "TL"
   A(82) = "PB"
   A(83) = "BI"
   A(84) = "PO"
   A(85) = "AT"
   A(86) = "RN"
   A(87) = "FR"
   A(88) = "RA"
   A(89) = "AC"
   A(90) = "TH"
   A(91) = "PA"
   A(92) = "U"
   A(93) = "NP"
   A(94) = "PU"
   A(95) = "AM"
   A(96) = "CM"
   A(97) = "BK"
   A(98) = "CF"
   A(99) = "ES"
   A(100) = "FM"
   A(101) = "MD"
   A(102) = "NO"
   A(103) = "LR"
   A(104) = "RF"
   A(105) = "DB"
   A(106) = "SG"
   A(107) = "BH"
   A(108) = "HS"
   A(109) = "MT"
   A(110) = "UUN"
   A(111) = "UUU"
   A(112) = "UUB"
   A(113) = "UUT"
   A(114) = "UUQ"
   A(115) = "UUP"
   A(116) = "UUH"
   A(117) = "UUS"
   A(118) = "UUO"
End Sub

