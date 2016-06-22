Public Class ConnectedClientsList

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        Me.Close()

    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click

        Try
            If ListBox1.SelectedItem.ToString <> "" Then
                Uffici.Show()
                UfficiAdd.Show()
                UfficiAdd.TopMost = True
                UfficiAdd.TextBox2.Text = ListBox1.SelectedItem.ToString
            Else
                MsgBox("Selezionare un client da aggiungere agli uffici per poter continuare.", MsgBoxStyle.Information, "Asc Server - Selezionare Client")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click

        Try
            If ListBox1.SelectedItem.ToString <> "" Then
                Uffici.Show()
                ContactsAdd.Show()
                ContactsAdd.TopMost = True
                ContactsAdd.TextBox1.Text = ListBox1.SelectedItem.ToString
            Else
                MsgBox("Selezionare un client da aggiungere agli uffici per poter continuare.", MsgBoxStyle.Information, "Asc Server - Selezionare Client")
            End If
        Catch ex As Exception

        End Try

    End Sub
End Class