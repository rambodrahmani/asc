Public Class SearchResults

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        Me.Close()

    End Sub

    Private Sub SearchResults_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        TextBox1.SelectionStart = 0
        TextBox1.SelectionLength = 0

    End Sub
End Class