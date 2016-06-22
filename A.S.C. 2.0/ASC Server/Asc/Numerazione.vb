Public Class Numerazione

    Dim datagiorno As String

    Private Sub Preferenze_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            TextBox1.Text = My.Settings.NumEmail2.ToString("D4")
            TextBox2.Text = My.Settings.NumEmail.ToString("D4")
            TextBox3.Text = My.Settings.NumEmailIntelligence.ToString("D4")
        Catch ex As Exception

        End Try
        datagiorno = My.Settings.DataOggi

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            My.Settings.NumEmail2 = Val(TextBox1.Text)
            My.Settings.NumEmail = Val(TextBox2.Text)
            My.Settings.NumEmailIntelligence = Val(TextBox3.Text)
            Me.Close()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Try
            Dim filescollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno)
            MsgBox("Il software ha rilevato che il numero progressivo per gli arrivi debba essere: " & (filescollection.Count - 1).ToString("D4"), MsgBoxStyle.Information, "Asc Server - Numerazione Progressiva")
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Try
            Dim filescollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno)
            MsgBox("Il software ha rilevato che il numero progressivo per le partenze debba essere: " & (filescollection.Count - 1).ToString("D4"), MsgBoxStyle.Information, "Asc Server - Numerazione Progressiva")
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Try
            Dim filescollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno)
            MsgBox("Il software ha rilevato che il numero progressivo per l'intelligence debba essere: " & (filescollection.Count - 1).ToString("D4"), MsgBoxStyle.Information, "Asc Server - Numerazione Progressiva")
        Catch ex As Exception

        End Try

    End Sub
End Class