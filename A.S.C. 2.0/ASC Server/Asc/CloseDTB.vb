Public Class CloseDTB

    Private Sub CloseDTB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Timer1.Start()

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Me.Hide()
        Try
            If My.Settings.CloseDTBNow = 2 Then
                My.Settings.CloseDTBNow = 0
                DATABASE_PARTENZE.Close()
                Me.Close()
            End If
        Catch ex As Exception

        End Try

    End Sub
End Class