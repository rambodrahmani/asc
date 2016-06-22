Public Class ConnectedClientsList

    Dim IndexToRemove As Integer = 1
    Dim counter As Integer = 0

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        Me.Close()

    End Sub

    Public Sub Removeitems()

        Timer1.Stop()
        Try
            For r = 0 To ListBox1.Items.Count - 1
                ListBox1.SelectedIndex = IndexToRemove
                ListBox1.Items.Remove(ListBox1.SelectedItem)
                IndexToRemove += 1
            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        counter += 1
        If counter > 5 Then
            Removeitems()
        End If

    End Sub

    Private Sub ConnectedClientsList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Timer1.Start()

    End Sub
End Class