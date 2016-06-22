Public Class LogIn

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            If TextBox1.Text = My.Settings.NomeUtente Then
                If TextBox2.Text = My.Settings.PassUtente Then
                    Main.Show()
                    Me.Close()
                Else
                    MsgBox("Password utente errata, acesso non acconsetito", MsgBoxStyle.Information, "Asc Server - Errore Identificazione")
                End If
            Else
                MsgBox("Nome utente errato, acesso non acconsetito", MsgBoxStyle.Information, "Asc Server - Errore Identificazione")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button1.PerformClick()
        End If

    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button1.PerformClick()
        End If

    End Sub

    Private Sub LogIn_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If My.Settings.NomeUtente = "admin" And My.Settings.PassUtente = "admin" Then
            Dim Message = MsgBox("Il nome utente e la password sono impostati su " & """" & "admin" & """" & ", si desidera impostare password e nome utente ora?", MsgBoxStyle.YesNo, "Asc Server - Utente")
            If Message = 6 Then
                Utente.Show()
                Me.Close()
            End If
        End If

    End Sub

End Class