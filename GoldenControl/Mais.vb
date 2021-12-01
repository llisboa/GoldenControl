Public Class Mais

    Sub Carrega()
        Icraft.IcftBase.Form.Conteudo(Me, "txt") = Icraft.IcftBase.NZ(Icraft.IcftBase.RegAplKey("GoldenControl", "Mais"), "")
    End Sub

    Sub Grava()
        Icraft.IcftBase.RegAplKey("GoldenControl", "Mais") = Icraft.IcftBase.Form.Conteudo(Me, "txt")
    End Sub

    Private Sub btnFechar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFechar.Click
        Me.Close()
    End Sub

    Private Sub Mais_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Carrega()
        Atualiza()
    End Sub

    Private Sub btnGravar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGravar.Click
        Grava()
        Me.Close()
    End Sub

    Sub Atualiza()

        For z As Integer = 1 To 5
            CType(Controls.Find("ReplID" & z, True)(0), TextBox).Text = CType(Princ.Controls.Find("TxtReplID" & z, True)(0), TextBox).Text
            CType(Controls.Find("ReplIP" & z, True)(0), TextBox).Text = CType(Princ.Controls.Find("TxtReplIP" & z, True)(0), TextBox).Text
        Next

        lstOpc.Items.Clear()
        For Each Item As String In Split(Princ.txtComandos.Text, ";")
            lstOpc.Items.Add(Item)
        Next
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Atualiza()
    End Sub


    Private Sub btnFech_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFech.Click
        Me.Close()
    End Sub

    Private Sub btnEliminar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEliminar.Click
        Try
            lstOpc.Items.RemoveAt(lstOpc.SelectedIndex)
            Princ.txtComandos.Text = ""
            For z As Integer = 0 To lstOpc.Items.Count - 1
                Princ.txtComandos.Text &= IIf(Princ.txtComandos.Text <> "", ";", "") & lstOpc.Items(z)
            Next

            Princ.AtualizaComandoCombo()
            Princ.Salva()
        Catch
        End Try
    End Sub

    Private Sub btnCancela_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancela.Click
        Me.Close()
    End Sub

    Private Sub btnGrava_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGrava.Click
        Grava()
        Me.Close()
    End Sub
End Class