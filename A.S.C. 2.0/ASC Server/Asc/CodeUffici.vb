Public Class CodeUffici

    Dim LastSelectedFolder As String = ""

    Private Sub ListView1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListView1.MouseDoubleClick

        If e.Button = Windows.Forms.MouseButtons.Left Then
            Try
                For Each ListViewSelectedItem As ListViewItem In ListView1.SelectedItems
                    LastSelectedFolder = ListViewSelectedItem.Text.ToString
                    Main.client.Send(2, "MessagesRequestFor|" & ListViewSelectedItem.Text.ToString & "|")
                    ListView1.Items.Clear()
                    Label1.Visible = True
                Next
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        Me.Close()

    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click

        If Main.MenuItem42.Text = "Disconnetti" Then
            Try
                LastSelectedFolder = ""
                Main.client.Send(2, "ServerRequestsCode")
            Catch ex As Exception

            End Try
            ListView1.Items.Clear()
        Else
            MsgBox("Asc Server non risulta connesso ad Asc Echange, impossibile continuare. Connettersi e riprovare.", MsgBoxStyle.Information, "Asc Server - Visualizza Code")
        End If

    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click

        Try
            Dim SelectedItemText As String = ""
            For Each ListViewSelectedItem As ListViewItem In ListView1.SelectedItems
                SelectedItemText = ListViewSelectedItem.Text.ToString
                ListView1.Items.RemoveAt(ListViewSelectedItem.Index)
            Next
            Main.client.Send(2, "DeleteMessageRequestfor|" & LastSelectedFolder & "\" & SelectedItemText & ".txt|")
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click

        Try
            Dim SelectedItemText As String = ""
            For Each ListViewSelectedItem As ListViewItem In ListView1.SelectedItems
                SelectedItemText = ListViewSelectedItem.Text.ToString
                ListView1.Items.RemoveAt(ListViewSelectedItem.Index)
            Next
            Main.client.Send(2, "DeleteFolderRequestfor|" & SelectedItemText & "|")
        Catch ex As Exception

        End Try

    End Sub
End Class