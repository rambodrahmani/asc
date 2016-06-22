Public Class NomeUnita

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If TextBox1.Text <> "" Then
            My.Settings.NomeUnita = TextBox1.Text
            If My.Settings.x = 35 Then
                My.Settings.x = 0
                MsgBox("Nome Unità cambiata con successo.", MsgBoxStyle.Information, "Asc Client - Nome Unità")
                Me.Close()
            Else
                Form1.Show()
                Me.Close()
            End If
        Else
            MsgBox("Il campo " & """" & "Nome Unità" & """" & " non può essere lasciato vuoto.", MsgBoxStyle.Critical, "Asc Server")
        End If

    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            Button1.PerformClick()
        End If

    End Sub

    Private Sub NomeUnita_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If My.Settings.x = 35 Then
            Button1.Text = "Salva"
        End If

    End Sub

End Class