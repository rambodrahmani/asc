Public Class Data

    Private Sub Data_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Dim MySplitter() As String = Split(My.Settings.DataOggi, ".")
            For r = 0 To UBound(MySplitter)
                If r = MySplitter.Count - 1 Then
                    TextBox1.Text += MySplitter(r).ToString
                Else
                    TextBox1.Text += MySplitter(r).ToString & "/"
                End If
            Next
            TextBox2.Text = My.Settings.JlDate
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            If TextBox1.Text <> "" And TextBox2.Text <> "" Then
                Dim AlertMessage = MsgBox("Procedendo verrà modifica la data utilizzata dal software, si è sicuri di voler procedere?", MsgBoxStyle.YesNo, "Asc Server - Modifica Data")
                If AlertMessage = 6 Then
                    My.Settings.DataOggi = ""
                    My.Settings.JlDate = 0
                    Dim MySplitter() As String = Split(TextBox1.Text, "/")
                    For r = 0 To UBound(MySplitter)
                        If r = MySplitter.Count - 1 Then
                            My.Settings.DataOggi += MySplitter(r).ToString
                        Else
                            My.Settings.DataOggi += (MySplitter(r).ToString) & "."
                        End If
                    Next
                    My.Settings.JlDate = TextBox2.Text
                    Main.Label4.Text = My.Settings.JlDate
                    Main.SaveOthersProgressive()
                    Me.Close()
                End If
            Else
                MsgBox("Una delle due date non è stata inserita, impossibile continuare.", MsgBoxStyle.Critical, "Asc Server - Modifca Data Errata")
            End If
        Catch ex As Exception
            MsgBox("Errore durante la modifica della data. Errore rilevato: " & ex.Message, MsgBoxStyle.Critical, "Asc Server - Errore Critico")
        End Try

    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            Button1.PerformClick()
        End If

    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown

        If e.KeyCode = Keys.Enter Then
            Button1.PerformClick()
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Me.Close()

    End Sub
End Class