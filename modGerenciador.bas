Attribute VB_Name = "modGerenciador"
Option Explicit
'
Public IntIncrDataGer As Integer
Public TxtDataGer As String

Public Sub GerenciaArquivos()
    Dim strDia As String
    Dim strMes As String
    Dim strAno As String
    Screen.MousePointer = 11
    Dim bolAchouAlgo As Boolean
    bolAchouAlgo = False
    frmGerenciador.lblR00.Visible = False
    frmGerenciador.lblR01.Visible = False
    frmGerenciador.lblR02.Visible = False
    frmGerenciador.lblR03.Visible = False
    frmGerenciador.lblR04.Visible = False
    frmGerenciador.lblR05.Visible = False
    frmGerenciador.lblR06.Visible = False
    frmGerenciador.lblR07.Visible = False
    frmGerenciador.lblR08.Visible = False
    frmGerenciador.lblR09.Visible = False
    frmGerenciador.lblE00.Visible = False
    frmGerenciador.lblE01.Visible = False
    frmGerenciador.lblE02.Visible = False
    frmGerenciador.lblE03.Visible = False
    frmGerenciador.lblE04.Visible = False
    frmGerenciador.lblE05.Visible = False
    frmGerenciador.lblE06.Visible = False
    frmGerenciador.lblE07.Visible = False
    frmGerenciador.lblE08.Visible = False
    frmGerenciador.lblE09.Visible = False
    If Len(Trim(frmGerenciador.txtData.Text)) <> 10 Then
       MsgBox "A data deverá ser digitada com dois digitos para o dia, uma barra, dois dias para o mês, uma barra e quatro dígitos para o ano. Assim 30/6/1977 deverá ser digitado como 30/06/1977. Datas como 6/2/2002 ou 10/3/77 são inválidas", vbOKOnly + vbCritical, App.Title
       frmGerenciador.txtData.Text = vbNullString
       Exit Sub
    End If
    strDia = Trim(Left(Trim(frmGerenciador.txtData.Text), 2))
    strMes = Trim(Mid(Trim(frmGerenciador.txtData.Text), 4, 2))
    strAno = Trim(Right(Trim(frmGerenciador.txtData.Text), 2))
    'Lendo arquivo in
    For I = 0 To 9
        If frmLogin.FileSystem.Dir(strPath & "\R" & strAno & strMes & strDia & Trim(CStr(I)) & "." & Right(usrCodigoVendedor, 3)) <> "" Then
            Select Case I
                Case 0
                    bolAchouAlgo = True
                    frmGerenciador.lblR00.BackColor = Vermelho
                    frmGerenciador.lblR00.Visible = True
                Case 1
                    bolAchouAlgo = True
                    frmGerenciador.lblR01.BackColor = Vermelho
                    frmGerenciador.lblR01.Visible = True
                Case 2
                    bolAchouAlgo = True
                    frmGerenciador.lblR02.BackColor = Vermelho
                    frmGerenciador.lblR02.Visible = True
                Case 3
                    bolAchouAlgo = True
                    frmGerenciador.lblR03.BackColor = Vermelho
                    frmGerenciador.lblR03.Visible = True
                Case 4
                    bolAchouAlgo = True
                    frmGerenciador.lblR04.BackColor = Vermelho
                    frmGerenciador.lblR04.Visible = True
                Case 5
                    bolAchouAlgo = True
                    frmGerenciador.lblR05.BackColor = Vermelho
                    frmGerenciador.lblR05.Visible = True
                Case 6
                    bolAchouAlgo = True
                    frmGerenciador.lblR06.BackColor = Vermelho
                    frmGerenciador.lblR06.Visible = True
                Case 7
                    bolAchouAlgo = True
                    frmGerenciador.lblR07.BackColor = Vermelho
                    frmGerenciador.lblR07.Visible = True
                Case 8
                    bolAchouAlgo = True
                    frmGerenciador.lblR08.BackColor = Vermelho
                    frmGerenciador.lblR08.Visible = True
                Case 9
                    bolAchouAlgo = True
                    frmGerenciador.lblR09.BackColor = Vermelho
                    frmGerenciador.lblR09.Visible = True
            End Select
        End If
        If frmLogin.FileSystem.Dir(strPath & "\-R" & strAno & strMes & strDia & Trim(CStr(I)) & "." & Right(usrCodigoVendedor, 3)) <> "" Then
            Select Case I
                Case 0
                    bolAchouAlgo = True
                    frmGerenciador.lblR00.BackColor = Azul
                    frmGerenciador.lblR00.Visible = True
                Case 1
                    bolAchouAlgo = True
                    frmGerenciador.lblR01.BackColor = Azul
                    frmGerenciador.lblR01.Visible = True
                Case 2
                    bolAchouAlgo = True
                    frmGerenciador.lblR02.BackColor = Azul
                    frmGerenciador.lblR02.Visible = True
                Case 3
                    bolAchouAlgo = True
                    frmGerenciador.lblR03.BackColor = Azul
                    frmGerenciador.lblR03.Visible = True
                Case 4
                    bolAchouAlgo = True
                    frmGerenciador.lblR04.BackColor = Azul
                    frmGerenciador.lblR04.Visible = True
                Case 5
                    bolAchouAlgo = True
                    frmGerenciador.lblR05.BackColor = Azul
                    frmGerenciador.lblR05.Visible = True
                Case 6
                    bolAchouAlgo = True
                    frmGerenciador.lblR06.BackColor = Azul
                    frmGerenciador.lblR06.Visible = True
                Case 7
                    bolAchouAlgo = True
                    frmGerenciador.lblR07.BackColor = Azul
                    frmGerenciador.lblR07.Visible = True
                Case 8
                    bolAchouAlgo = True
                    frmGerenciador.lblR08.BackColor = Azul
                    frmGerenciador.lblR08.Visible = True
                Case 9
                    bolAchouAlgo = True
                    frmGerenciador.lblR09.BackColor = Azul
                    frmGerenciador.lblR09.Visible = True
            End Select
        End If
        If frmLogin.FileSystem.Dir(strPath & "\=R" & strAno & strMes & strDia & Trim(CStr(I)) & "." & Right(usrCodigoVendedor, 3)) <> "" Then
            If frmLogin.FileSystem.Dir(strPath & "\-R" & strAno & strMes & strDia & Trim(CStr(I)) & "." & Right(usrCodigoVendedor, 3)) <> "" Then frmLogin.FileSystem.Kill strPath & "\-R" & strAno & strMes & strDia & Trim(CStr(I)) & "." & Right(usrCodigoVendedor, 3)
            Select Case I
                Case 0
                    bolAchouAlgo = True
                    frmGerenciador.lblR00.BackColor = Verde
                    frmGerenciador.lblR00.Visible = True
                Case 1
                    bolAchouAlgo = True
                    frmGerenciador.lblR01.BackColor = Verde
                    frmGerenciador.lblR01.Visible = True
                Case 2
                    bolAchouAlgo = True
                    frmGerenciador.lblR02.BackColor = Verde
                    frmGerenciador.lblR02.Visible = True
                Case 3
                    bolAchouAlgo = True
                    frmGerenciador.lblR03.BackColor = Verde
                    frmGerenciador.lblR03.Visible = True
                Case 4
                    bolAchouAlgo = True
                    frmGerenciador.lblR04.BackColor = Verde
                    frmGerenciador.lblR04.Visible = True
                Case 5
                    bolAchouAlgo = True
                    frmGerenciador.lblR05.BackColor = Verde
                    frmGerenciador.lblR05.Visible = True
                Case 6
                    bolAchouAlgo = True
                    frmGerenciador.lblR06.BackColor = Verde
                    frmGerenciador.lblR06.Visible = True
                Case 7
                    bolAchouAlgo = True
                    frmGerenciador.lblR07.BackColor = Verde
                    frmGerenciador.lblR07.Visible = True
                Case 8
                    bolAchouAlgo = True
                    frmGerenciador.lblR08.BackColor = Verde
                    frmGerenciador.lblR08.Visible = True
                Case 9
                    bolAchouAlgo = True
                    frmGerenciador.lblR09.BackColor = Verde
                    frmGerenciador.lblR09.Visible = True
            End Select
        End If
    Next
    'Lendo arquivo out
    For I = 0 To 9
        If frmLogin.FileSystem.Dir(strPath & "\T" & strAno & strMes & strDia & Trim(CStr(I)) & "." & Right(usrCodigoVendedor, 3)) <> "" Then
            Select Case I
                Case 0
                    bolAchouAlgo = True
                    frmGerenciador.lblE00.BackColor = Azul
                    frmGerenciador.lblE00.Visible = True
                Case 1
                    bolAchouAlgo = True
                    frmGerenciador.lblE01.BackColor = Azul
                    frmGerenciador.lblE01.Visible = True
                Case 2
                    bolAchouAlgo = True
                    frmGerenciador.lblE02.BackColor = Azul
                    frmGerenciador.lblE02.Visible = True
                Case 3
                    bolAchouAlgo = True
                    frmGerenciador.lblE03.BackColor = Azul
                    frmGerenciador.lblE03.Visible = True
                Case 4
                    bolAchouAlgo = True
                    frmGerenciador.lblE04.BackColor = Azul
                    frmGerenciador.lblE04.Visible = True
                Case 5
                    bolAchouAlgo = True
                    frmGerenciador.lblE05.BackColor = Azul
                    frmGerenciador.lblE05.Visible = True
                Case 6
                    bolAchouAlgo = True
                    frmGerenciador.lblE06.BackColor = Azul
                    frmGerenciador.lblE06.Visible = True
                Case 7
                    bolAchouAlgo = True
                    frmGerenciador.lblE07.BackColor = Azul
                    frmGerenciador.lblE07.Visible = True
                Case 8
                    bolAchouAlgo = True
                    frmGerenciador.lblE08.BackColor = Azul
                    frmGerenciador.lblE08.Visible = True
                Case 9
                    bolAchouAlgo = True
                    frmGerenciador.lblE09.BackColor = Azul
                    frmGerenciador.lblE09.Visible = True
            End Select
        End If
        If frmLogin.FileSystem.Dir(strPath & "\=T" & strAno & strMes & strDia & Trim(CStr(I)) & "." & Right(usrCodigoVendedor, 3)) <> "" Then
            If frmLogin.FileSystem.Dir(strPath & "\T" & strAno & strMes & strDia & Trim(CStr(I)) & "." & Right(usrCodigoVendedor, 3)) <> "" Then frmLogin.FileSystem.Kill strPath & "\T" & strAno & strMes & strDia & Trim(CStr(I)) & "." & Right(usrCodigoVendedor, 3)
            Select Case I
                Case 0
                    bolAchouAlgo = True
                    frmGerenciador.lblE00.BackColor = Verde
                    frmGerenciador.lblE00.Visible = True
                Case 1
                    bolAchouAlgo = True
                    frmGerenciador.lblE01.BackColor = Verde
                    frmGerenciador.lblE01.Visible = True
                Case 2
                    bolAchouAlgo = True
                    frmGerenciador.lblE02.BackColor = Verde
                    frmGerenciador.lblE02.Visible = True
                Case 3
                    bolAchouAlgo = True
                    frmGerenciador.lblE03.BackColor = Verde
                    frmGerenciador.lblE03.Visible = True
                Case 4
                    bolAchouAlgo = True
                    frmGerenciador.lblE04.BackColor = Verde
                    frmGerenciador.lblE04.Visible = True
                Case 5
                    bolAchouAlgo = True
                    frmGerenciador.lblE05.BackColor = Verde
                    frmGerenciador.lblE05.Visible = True
                Case 6
                    bolAchouAlgo = True
                    frmGerenciador.lblE06.BackColor = Verde
                    frmGerenciador.lblE06.Visible = True
                Case 7
                    bolAchouAlgo = True
                    frmGerenciador.lblE07.BackColor = Verde
                    frmGerenciador.lblE07.Visible = True
                Case 8
                    bolAchouAlgo = True
                    frmGerenciador.lblE08.BackColor = Verde
                    frmGerenciador.lblE08.Visible = True
                Case 9
                    bolAchouAlgo = True
                    frmGerenciador.lblE09.BackColor = Verde
                    frmGerenciador.lblE09.Visible = True
            End Select
        End If
    Next
    Screen.MousePointer = 0
    '
    If bolAchouAlgo = False Then
       '
       MsgBox "Não foi encontrado nenhum arquivo neste dia.", vbOKOnly + vbInformation, App.Title
       '
    Else
       '
       frmGerenciador.Frame1.Visible = True
       '
    End If
    '
End Sub

Public Function MontaNomeGerenciador(ByVal strBotao As String) As String
    If CInt(strBotao) < 10 Then
       MontaNomeGerenciador = strPath & "\R" & Trim(Right(frmGerenciador.txtData.Text, 2)) & Trim(Mid(frmGerenciador.txtData.Text, 4, 2)) & Trim(Left(frmGerenciador.txtData.Text, 2)) & Trim(strBotao) & "." & Trim(Right(usrCodigoVendedor, 3))
    Else
       MontaNomeGerenciador = strPath & "\T" & Trim(Right(frmGerenciador.txtData.Text, 2)) & Trim(Mid(frmGerenciador.txtData.Text, 4, 2)) & Trim(Left(frmGerenciador.txtData.Text, 2)) & Trim(CStr(CInt(strBotao) - 10)) & "." & Trim(Right(usrCodigoVendedor, 3))
    End If
    '
    ' MsgBox "Nome do Arquivo:" & MontaNomeGerenciador, vbOKOnly + vbCritical, App.Title
    '
End Function
