Public Class Utenza

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            If TextBox1.Text <> "" Then
                If TextBox2.Text <> "" Then
                    My.Settings.NomeUtente = TextBox1.Text
                    My.Settings.PasswordUtente = TextBox2.Text
                    MsgBox("Nome Utente e Password salvati correttamente.", MsgBoxStyle.Information, "Asc Exchange - Utenza")
                    If My.Settings.x = 1 Then
                        My.Settings.x = 0
                        Me.Close()
                    Else
                        Form1.LoadForm()
                        Me.Close()
                    End If
                Else
                    MsgBox("Non è stata inserita alcuna Password, inserirne una per poter continuare.", MsgBoxStyle.Information, "Asc Exchange - Password Mancante")
                End If
            Else
                MsgBox("Non è stato inserito alcun Nome Utente, inserirne uno per poter continuare.", MsgBoxStyle.Information, "Asc Exchange - Nome Utente Mancante")
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