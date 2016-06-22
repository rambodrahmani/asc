Public Class Accesso

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            If TextBox1.Text = My.Settings.NomeUtente Then
                If TextBox2.Text = My.Settings.PasswordUtente Then
                    Accesso()
                Else
                    MsgBox("La Password inserita è errata. Impossibile continuare.", MsgBoxStyle.Information, "Asc Exchange - Accesso Errato")
                End If
            Else
                MsgBox("Il Nome Utente inserito è errato. Impossibile continuare.", MsgBoxStyle.Information, "Asc Exchange - Accesso Errato")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub Accesso()

        Try
            If My.Settings.x = 1 Then
                My.Settings.x = 0
                Utenza.Show()
                Me.Close()
            ElseIf My.Settings.x = 2 Then
                My.Settings.x = 0
                Settings.Show()
                Me.Close()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button1.PerformClick()
        End If

    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button1.PerformClick()
        End If

    End Sub
End Class