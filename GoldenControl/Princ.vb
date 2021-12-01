Public Class Princ

    Shared _RetEnv As Boolean = False

    Public Shared Property RetEnv() As String
        Get
            Return _RetEnv
        End Get
        Set(ByVal value As String)
            _RetEnv = value
        End Set
    End Property

    Private Function SemAspas(ByVal Texto As String) As String
        Return Texto.Trim("""", Chr(147), Chr(148))
    End Function


    Private Sub Princ_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = My.Application.Info.ProductName & " - " & Icraft.IcftBase.VersaoApl & " - " & My.Application.Info.CompanyName & " - 2013"
        Icraft.IcftBase.Form.Conteudo(Me, "txt") = Icraft.IcftBase.NZ(Icraft.IcftBase.RegAplKey("GoldenControl", "Princ"), "")
        Mais.Carrega()
        Dim Args As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs()
        Dim help As Boolean = False
        For z As Integer = 0 To Args.Count - 1 Step 2
            Dim Coma As String = SemAspas(Args(z)).ToLower

            Select Case Coma
                Case "-verinterv"
                    txtInterv.Text = SemAspas(Args(z + 1))
                    txtVerGolden.Checked = True
                Case "-help"
                    Help = True
                Case "-usuario"
                    txtReplUsu.Text = SemAspas(Args(z + 1))
                Case "-senha"
                    txtReplSenha.Text = SemAspas(Args(z + 1))
                Case "-siteidip"
                    Dim It() As String = Split(SemAspas(Args(z + 1)), ";")
                    For zz As Integer = 0 To Math.Min(9, It.Count - 1)
                        Dim txt As TextBox
                        If zz Mod 2 = 0 Then
                            txt = GroupBox1.Controls.Find("txtReplId" & Int(zz / 2) + 1, False)(0)
                        Else
                            txt = GroupBox1.Controls.Find("txtReplIp" & Int(zz / 2) + 1, False)(0)
                        End If
                        txt.Text = It(zz)
                    Next
                Case "-sitesenhasysconn"
                    Dim It() As String = Split(SemAspas(Args(z + 1)), ";")
                    For zz As Integer = 0 To Math.Min(9, It.Count - 1)
                        Dim txt As TextBox
                        If zz Mod 2 = 0 Then
                            txt = Mais.TabPage1.Controls.Find("txtSenhaSys" & Int(zz / 2) + 1, False)(0)
                        Else
                            txt = Mais.TabPage1.Controls.Find("txtConn" & Int(zz / 2) + 1, False)(0)
                        End If
                        txt.Text = It(zz)
                    Next
                Case "-emailde"
                    Mais.txtDe.Text = SemAspas(Args(z + 1))
                    Mais.txtEnvia.Checked = True
                Case "-emailpara"
                    Mais.txtPara.Text = SemAspas(Args(z + 1))
                    Mais.txtEnvia.Checked = True
                Case "-emailassunto"
                    Mais.txtAssunto.Text = SemAspas(Args(z + 1))
                    Mais.txtEnvia.Checked = True
                Case "-emailsmtp"
                    Mais.txtSmtp.Text = SemAspas(Args(z + 1))
                    Mais.txtEnvia.Checked = True
                Case "-arquivolog"
                    Mais.txtArquivo.Text = SemAspas(Args(z + 1))
                    Mais.txtGravaLog.Checked = True
            End Select
        Next

        AtualizaVerGold()
    End Sub

    Sub Salva()
        Icraft.IcftBase.RegAplKey("GoldenControl", "Princ") = Icraft.IcftBase.Form.Conteudo(Me, "txt")
    End Sub

    Public Function DosShell(ByVal Comando As String, ByVal Evento As EventHandler) As String
        Dim Temp As String = System.IO.Path.GetTempFileName()
        Dim p As Process = New Process
        p.StartInfo.FileName = "CMD"
        p.StartInfo.Arguments = "/C """ & Comando & """ > " & Temp
        p.StartInfo.RedirectStandardOutput = False
        p.StartInfo.UseShellExecute = True
        p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
        p.Start()
        While Not p.HasExited
            Application.DoEvents()
            If Not IsNothing(Evento) Then
                Evento.Invoke(Me, Nothing)
            End If
            p.WaitForExit(10)
        End While
        p.Dispose()
        p.Close()
        Dim Ret As String = Icraft.IcftBase.CarregaArqTxt(Temp)
        Kill(Temp)
        Return Ret
    End Function

    Dim ExecMomento As Date = Nothing
    Sub ExecComa()
        lblComa.Text = "Executando... Aguarde: " & Icraft.IcftBase.ExibeSegs(DateDiff(DateInterval.Second, ExecMomento, Now), Icraft.IcftBase.ExibeSegsOpc.hh_mm_ss)
    End Sub

    Private Sub btnAtualiza_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAtualiza.Click
        ExecMomento = Now

        lblTempo.Text = "Comando em " & Format(ExecMomento, "dd/MM/yyyy") & " " & Format(ExecMomento, " HH:mm:ss")

        If Not Icraft.IcftBase.TemNaLista(txtComandos.Text, txtComando.Text) Then
            txtComandos.Text &= IIf(txtComandos.Text <> "", ";", "") & txtComando.Text
        End If
        AtualizaComandoCombo()

        For z As Integer = 1 To 5
            If IP(z) <> "" Then
                Exec(z)
            End If
        Next
        lblComa.Text = "Concluído. Tempo decorrido: " & Icraft.IcftBase.ExibeSegs(DateDiff(DateInterval.Second, ExecMomento, Now), Icraft.IcftBase.ExibeSegsOpc.hh_mm_ss)
        Salva()
    End Sub


    Sub Exec(ByVal z As Integer)
        Dim ERROS As String = ""
        Dim Comando As String = """" & Icraft.IcftBase.FileExpr("~/PSEXEC.EXE") & """ /accepteula \\" & IP(z) & " -u " & txtReplUsu.Text & " -p " & txtReplSenha.Text & " " & Conv(txtComando.Text, z)
        Dim SAIDA As String = DosShell(Comando, AddressOf ExecComa)

        Buffer(z).Text = SAIDA
        If Len(SAIDA) > 80 Then
            SAIDA = Microsoft.VisualBasic.Left(SAIDA, 80) & " ..."
        End If
        Result(z).Text = "==>" & Icraft.IcftBase.TrocaTexto(SAIDA, vbCrLf, " ", "  ", " ")
    End Sub


    Function IP(ByVal Num As Integer) As String
        Return CType(Icraft.IcftBase.Form.FindControl(Me, "txtReplIP" & Num), TextBox).Text()
    End Function

    Function SenhaSys(ByVal Num As Integer) As String
        Return CType(Icraft.IcftBase.Form.FindControl(Mais, "txtSenhaSys" & Num), TextBox).Text()
    End Function

    Function Conn(ByVal Num As Integer) As String
        Return CType(Icraft.IcftBase.Form.FindControl(Mais, "txtConn" & Num), TextBox).Text()
    End Function

    Function ReplID(ByVal Num As Integer) As String
        Return CType(Icraft.IcftBase.Form.FindControl(Me, "txtReplID" & Num), TextBox).Text()
    End Function

    Function Result(ByVal Num As Integer) As Label
        Return CType(Icraft.IcftBase.Form.FindControl(Me, "Result" & Num), Label)
    End Function

    Function Buffer(ByVal Num As Integer) As Label
        Return CType(Icraft.IcftBase.Form.FindControl(Me, "B" & Num), Label)
    End Function

    Function BStatus(ByVal Num As Integer) As Label
        Return CType(Icraft.IcftBase.Form.FindControl(Me, "BStatus" & Num), Label)
    End Function

    Private Sub Result_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Result1.Click, Result2.Click, Result3.Click, Result4.Click, Result5.Click
        Dim n As Integer = Val(Mid(sender.name, 7))
        Det.Text = Buffer(n).Text
    End Sub

    Public Sub AtualizaComandoCombo()
        txtComando.Items.Clear()
        For Each Item As String In Split(txtComandos.Text, ";")
            txtComando.Items.Add(Item)
        Next
    End Sub

    Private Sub txtComando_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtComando.Enter
        AtualizaComandoCombo()
    End Sub

    Dim ExecStatusMomento As Date = Nothing
    Sub ExecStatus()
        lblStatus.Text = "Executando... Aguarde: " & Icraft.IcftBase.ExibeSegs(DateDiff(DateInterval.Second, ExecStatusMomento, Now), Icraft.IcftBase.ExibeSegsOpc.hh_mm_ss)
    End Sub

    Sub VerStatus(ByVal z As Integer)
        'RetEnv = False
        If IP(z) <> "" Then

            Dim ERROS As String = ""
            Dim Comando As String = """" & Icraft.IcftBase.FileExpr("~/PSEXEC.EXE") & """ /accepteula \\" & IP(z) & IIf(txtReplUsu.Text <> "", " -u " & txtReplUsu.Text, "") & IIf(txtReplSenha.Text <> "", " -p " & txtReplSenha.Text, "") & " C:\ORACLE\GGC\GG.BAT STATUS"
            Dim SAIDA As String = DosShell(Comando, AddressOf ExecStatus)

            Dim IT As New Dictionary(Of String, String)

            For Each LINHA As String In Split(SAIDA, vbCrLf)
                LINHA = Icraft.IcftBase.ReplRepl(LINHA & " MANAGER", "  ", " ")
                Dim AR() As String = Split(LINHA, " ")

                If LINHA Like "MANAGER *" Then
                    IT.Add(AR(2), AR(1))
                    Log &= IIf(Log <> "", vbCrLf, "") & IP(z) & " " & AR(2) & " ==> " & AR(1) & IIf(AR(1) <> "RUNNING", " [erro]", "")
                ElseIf LINHA Like "EXTRACT *" Then
                    IT.Add(AR(2), AR(1))
                    Log &= IIf(Log <> "", vbCrLf, "") & IP(z) & " " & AR(2) & " ==> " & AR(1) & IIf(AR(1) <> "RUNNING", " [erro]", "")
                ElseIf LINHA Like "REPLICAT *" Then
                    IT.Add(AR(2), AR(1))
                    Log &= IIf(Log <> "", vbCrLf, "") & IP(z) & " " & AR(2) & " ==> " & AR(1) & IIf(AR(1) <> "RUNNING", " [erro]", "")
                End If
            Next

            BStatus(z).Text = "...  " & SAIDA
            BStatus(z).Visible = True

            Dim x As Integer = 0
            Dim lbl As Windows.Forms.Label = Nothing


            For Each Chave As String In IT.Keys
                lbl = New Windows.Forms.Label
                lbl.Name = "lbl_" & z
                AddHandler lbl.Click, AddressOf ItemClick

                If IT(Chave) = "RUNNING" Then
                    lbl.BackColor = Color.Cyan
                Else
                    lbl.BackColor = Color.Red
                    Princ.RetEnv = True
                End If

                lbl.Height = 18
                lbl.Text = Chave
                lbl.Top = (z - 1) * 26 + 10
                lbl.Width = 70
                lbl.Left = x * 72
                x += 1
                lbl.Padding = New Windows.Forms.Padding(2, 2, 0, 0)
                pnlStatusResult.Controls.Add(lbl)
            Next

        End If

    End Sub

    Dim Log As String = ""

    Private Sub btnBusca_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBusca.Click

        Dim Type = sender.GetType()
        If Type.Name = "Button" Then
            RetEnv = False
        ElseIf Type.Name = "String" Then
            RetEnv = True
        End If
        ExecStatusMomento = Now

        lblTempo.Text = "GoldenGate Status em " & Format(ExecStatusMomento, "dd/MM/yyyy") & " " & Format(ExecStatusMomento, " HH:mm:ss")

        pnlStatusResult.Controls.Clear()

        Log = ""
        For z As Integer = 1 To 5
            VerStatus(z)
        Next
        lblStatus.Text = "Concluído. Tempo decorrido: " & Icraft.IcftBase.ExibeSegs(DateDiff(DateInterval.Second, ExecStatusMomento, Now), Icraft.IcftBase.ExibeSegsOpc.hh_mm_ss)

        Dim ARQ As String = ""
        If Not Mais.txtGravaLog.Checked Then
            ARQ = System.IO.Path.GetTempFileName()
        Else
            ARQ = Icraft.IcftBase.FileExpr(System.IO.Path.GetPathRoot(Mais.txtArquivo.Text), System.IO.Path.GetFileNameWithoutExtension(Mais.txtArquivo.Text) & ".jpg")
        End If

        tbExec.SelectedTab = tbGold

        ControleParaImagem(pnlPrinc).Save(ARQ, System.Drawing.Imaging.ImageFormat.Jpeg)

        Log = "Verificação em " & Me.Text & vbCrLf & Format(Now, "yyyy-MM-dd HH:mm:ss") & vbCrLf & Log & vbCrLf

        Dim MSG As String = "<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbCrLf & _
        "<head>" & vbCrLf & _
        "    <title>Untitled Page</title>" & vbCrLf & _
        "</head>" & vbCrLf & _
        "<body>" & vbCrLf & _
        "" & vbCrLf & _
        "    <p>" & vbCrLf & _
        "        INTERMESA - BANCO DE DADOS - ESTADO DE CONEXÕES</p>" & vbCrLf & _
        "    <p>" & vbCrLf & _
        "        Intercraft Solutions Informática Ltda - 2012</p>" & vbCrLf & _
        "    <p>" & vbCrLf & _
        "        Estado atual:</p>" & vbCrLf & _
        "    <p>" & vbCrLf & _
        "        <img src=""" & ARQ & """ /></p>" & vbCrLf & _
        "" & vbCrLf & Replace(vbCrLf & Icraft.IcftBase.Entifica(Log), vbCrLf, "<br/>") & _
        "</body>" & vbCrLf & _
        "</html>" & vbCrLf & _
        ""

        If Not Mais.txtEnvia.Checked And Not Mais.txtGravaLog.Checked Then
            Mais.txtEnvia.Checked = True
        End If

        Try
            If RetEnv = True And Mais.txtEnvia.Checked Then
                Dim de As String = Icraft.IcftBase.NZV(Mais.txtDe.Text, "suporte@icraft.com.br")
                Dim para As String = Icraft.IcftBase.NZV(Mais.txtPara.Text, "web@icraft.com.br")
                Dim assunto As String = Icraft.IcftBase.NZV(Mais.txtAssunto.Text, "Intermesa - Banco de Dados - Estado atual de conexão")
                Dim smtp As String = Icraft.IcftBase.NZV(Mais.txtSmtp.Text, "SMTPI")
                Dim Ret As String = Icraft.IcftBase.EnviaEmail(de, para, assunto, MSG, , smtp, , , , , , True)
                If Ret <> "" Then
                    Log &= IIf(Log <> "", vbCrLf, "") & "[erro] Falha no envio da mensagem: " & Ret
                End If
            End If
        Catch
        End Try

        If Not Mais.txtGravaLog.Checked Then
            Kill(ARQ)
        Else
            Dim ArqW As New System.IO.StreamWriter(Mais.txtArquivo.Text)
            ArqW.WriteLine(Log)
            ArqW.Close()

        End If

        Salva()
    End Sub

    Private Sub BStatus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BStatus1.Click, BStatus2.Click, BStatus3.Click, BStatus4.Click, BStatus5.Click
        Dim n As Integer = Mid(sender.name, 8)
        Det.Text = BStatus(n).Text
    End Sub

    Private Sub txtVerGolden_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVerGolden.CheckedChanged
        AtualizaVerGold()
        Salva()
    End Sub

    Dim MostraUlt As Date = Nothing
    Dim EmExec As Boolean = False
    Private Sub tmrVerGolden_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrVerGolden.Tick
        Application.DoEvents()
        If Not EmExec Then
            EmExec = True
            If MostraUlt = CDate(Nothing) Then
                MostraUlt = Now
            End If
            Dim Segs As Integer = DateDiff(DateInterval.Second, MostraUlt, Now)
            Dim Limit As Integer = Val(Microsoft.VisualBasic.Left(txtInterv.Text, 3)) * 60 + Val(Microsoft.VisualBasic.Right(txtInterv.Text, 2))
            If Limit = 0 Then

                txtVerGolden.Checked = False
                AtualizaVerGold()

                MsgBox("Antes de ativar informe o intervalo no formato MM:SS.")
                Exit Sub
            End If
            Dim Dif As Integer = Limit - Segs
            If Dif < 0 Then
                btnBusca_Click("OK", Nothing)
                MostraUlt = Now
            End If
            lblFalta.Visible = True
            lblFalta.Text = "Falta " & Icraft.IcftBase.ExibeSegs(Dif, Icraft.IcftBase.ExibeSegsOpc.hh_mm_ss)
            EmExec = False
            Application.DoEvents()
        End If
    End Sub

    Sub AtualizaVerGold()
        tmrVerGolden.Enabled = txtVerGolden.Checked
        If Not tmrVerGolden.Enabled Then
            lblFalta.Visible = False
        End If
        MostraUlt = Now
        Application.DoEvents()
    End Sub

    Private Sub txtInterv_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtInterv.Validating
        Salva()
    End Sub

    Function ControleParaImagem(ByVal Controle As Windows.Forms.Control) As Bitmap
        Dim bmp As New Bitmap(Controle.Width, Controle.Height)
        pnlPrinc.DrawToBitmap(bmp, Controle.ClientRectangle)
        Return bmp
    End Function

    Private Sub btnMais_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMais.Click
        Mais.Visible = True
        Mais.Show()
    End Sub

    Private Sub txt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtReplIP1.TextChanged, txtReplIP2.TextChanged, txtReplIP3.TextChanged, txtReplIP4.TextChanged, txtReplIP5.TextChanged, txtReplID1.TextChanged, txtReplID2.TextChanged, txtReplID3.TextChanged, txtReplID4.TextChanged, txtReplID5.TextChanged
        If Mais.Visible Then
            Mais.Atualiza()
        End If
    End Sub

    Function Conv(ByVal Param As String, ByVal Item As Integer) As String
        Param = Replace(Param, "[:ID]", ReplID(Item))
        Param = Replace(Param, "[:IP]", IP(Item))
        Param = Replace(Param, "[:SenhaSys]", SenhaSys(Item))
        Param = Replace(Param, "[:Conexão]", Conn(Item))
        Return Param
    End Function

    Private Sub TexAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lip.Click, lid.Click, lsenha.Click, lconn.Click
        txtComando.Text &= CType(sender, Label).Text
    End Sub

    Private Sub tbExec_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbExec.SelectedIndexChanged
        BStatus1.Visible = tbExec.SelectedTab Is tbGold And BStatus(1).Text <> ""
        BStatus2.Visible = tbExec.SelectedTab Is tbGold And BStatus(2).Text <> ""
        BStatus3.Visible = tbExec.SelectedTab Is tbGold And BStatus(3).Text <> ""
        BStatus4.Visible = tbExec.SelectedTab Is tbGold And BStatus(4).Text <> ""
        BStatus5.Visible = tbExec.SelectedTab Is tbGold And BStatus(5).Text <> ""
    End Sub

    Private Sub Det_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Det.MouseDoubleClick
        Dim Arq As String = System.IO.Path.GetTempFileName()
        Icraft.IcftBase.GravaLog(Arq, Det.Text, True)
        Shell("notepad.exe """ & Arq & """", AppWinStyle.NormalFocus)
        Kill(Arq)
    End Sub

    Private Sub btnExec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExec1.Click, btnExec2.Click, btnExec3.Click, btnExec4.Click, btnExec5.Click
        Dim z As Integer = Val(Mid(sender.name, 8))
        If tbExec.SelectedIndex = 0 Then
            ExecMomento = Now
            Exec(z)
            lblComa.Text = "Concluído. Tempo decorrido: " & Icraft.IcftBase.ExibeSegs(DateDiff(DateInterval.Second, ExecMomento, Now), Icraft.IcftBase.ExibeSegsOpc.hh_mm_ss)
        Else
            ExecStatusMomento = Now
            VerStatus(z)
            lblStatus.Text = "Concluído. Tempo decorrido: " & Icraft.IcftBase.ExibeSegs(DateDiff(DateInterval.Second, ExecStatusMomento, Now), Icraft.IcftBase.ExibeSegsOpc.hh_mm_ss)
        End If
    End Sub

    Private Sub ItemClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ExecClickMomento = Now
        Dim z As Integer = Split(sender.name, "_")(1)
        Dim Comando As String = """" & Icraft.IcftBase.FileExpr("~/PSEXEC.EXE") & """ /accepteula \\" & IP(z) & " -u " & txtReplUsu.Text & " -p " & txtReplSenha.Text & " C:\ORACLE\GGC\GG.BAT REPORT " & sender.TEXT
        Det.Text = DosShell(Comando, AddressOf ExecClick)
        lblStatus.Text = "Concluído. Tempo decorrido: " & Icraft.IcftBase.ExibeSegs(DateDiff(DateInterval.Second, ExecClickMomento, Now), Icraft.IcftBase.ExibeSegsOpc.hh_mm_ss)
    End Sub

    Dim ExecClickMomento As Date = Nothing
    Sub ExecClick()
        lblStatus.Text = "Buscando Report... Aguarde: " & Icraft.IcftBase.ExibeSegs(DateDiff(DateInterval.Second, ExecClickMomento, Now), Icraft.IcftBase.ExibeSegsOpc.hh_mm_ss)
    End Sub

    Private Sub btnTravar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If MsgBox("Bloquear base?") = MsgBoxResult.Yes Then
            txtComando.Text = "@(ECHO SELECT 'FOI' FROM DUAL||EXIT)|SQLPLUS IM@IMRJ"
        End If
    End Sub

    Private Sub btnLiberar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If MsgBox("Liberar base?") = MsgBoxResult.Yes Then
            txtComando.Text = ""
        End If
    End Sub

    Private Sub btnAjuda_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAjuda.Click
        frmHelp.Show()
        frmHelp.Focus()
    End Sub
End Class
