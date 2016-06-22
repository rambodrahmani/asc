Public Class Utente

    Public Sub NewUserCredentials()

        Try
            If TextBox1.Text <> "" Then
                If TextBox2.Text <> "" Then
                    If TextBox3.Text <> "" Then
                        My.Settings.NomeUtente = TextBox1.Text
                        My.Settings.PassUtente = TextBox2.Text
                        My.Settings.NomeUnita = TextBox3.Text
                        LogIn.Show()
                        Me.Close()
                    Else
                        MsgBox("Scegleire un nome per l'unità per poter continuare.", MsgBoxStyle.Information, "Asc Server - Nome Unità")
                    End If
                Else
                    MsgBox("Scegleire una password utente per poter continuare.", MsgBoxStyle.Information, "Asc Server - Utente")
                End If
            Else
                MsgBox("Scegleire un nome utente per poter continuare.", MsgBoxStyle.Information, "Asc Server - Utente")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        NewUserCredentials()

    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            NewUserCredentials()
        End If

    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown

        If e.KeyCode = Keys.Enter Then
            NewUserCredentials()
        End If

    End Sub

    Private Sub Utente_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        TextBox3.Text = My.Settings.NomeUnita

    End Sub
End Class