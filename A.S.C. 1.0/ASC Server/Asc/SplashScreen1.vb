Public NotInheritable Class SplashScreen1

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Me.TopMost = False
        LogIn.Show()
        Me.Close()

    End Sub

End Class
