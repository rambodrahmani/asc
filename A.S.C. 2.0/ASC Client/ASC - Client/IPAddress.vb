Public Class IPAddress

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            If TextBox1.Text <> "" Then
                My.Settings.IPAddress = TextBox1.Text
                Me.Close()
            Else
                MsgBox("Inserire l'indirizzo IP nella textbox per poter continuare!", MsgBoxStyle.Information, "Asc Client - Indirizzo IP")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub IPAddress_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.TopMost = True
            TextBox1.Text = My.Settings.IPAddress
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            Button1.PerformClick()
        End If

    End Sub
End Class