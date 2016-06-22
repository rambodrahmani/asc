Public Class NomeUnita

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If TextBox1.Text <> "" Then
            My.Settings.NomeUnita = TextBox1.Text
            My.Settings.NomeUtente = "admin"
            My.Settings.PassUtente = "admin"
            Utente.Show()
            Me.Close()
        Else
            MsgBox("Il campo " & """" & "Nome Unità" & """" & " non può essere lasciato vuoto.", MsgBoxStyle.Critical, "Asc Server")
        End If

    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            If TextBox1.Text <> "" Then
                My.Settings.NomeUnita = TextBox1.Text
                My.Settings.NomeUtente = "admin"
                My.Settings.PassUtente = "admin"
                utente.Show()
                Me.Close()
            Else
                MsgBox("Il campo " & """" & "Nome Unità" & """" & " non può essere lasciato vuoto.", MsgBoxStyle.Critical, "Asc Server")
            End If
        End If

    End Sub

End Class