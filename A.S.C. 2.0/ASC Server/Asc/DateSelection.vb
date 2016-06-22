Public Class DateSelection

    Dim datadacercare As String
    Dim datadef As String

    Private Sub MonthCalendar1_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged

        Try
            datadacercare = ""
            datadef = MonthCalendar1.SelectionStart
            If datadef(0) = "0" Then
                datadacercare += datadef(1) & "."
            Else
                datadacercare += datadef(0) & datadef(1) & "."
            End If
            If datadef(3) = "0" Then
                datadacercare += datadef(4) & "."
            Else
                datadacercare += datadef(3) & datadef(4) & "."
            End If
            datadacercare += datadef(6) & datadef(7) & datadef(8) & datadef(9)
            My.Settings.SelectedDate = datadacercare
        Catch ex As Exception

        End Try

        Try
            If My.Settings.DTBArriviActive = 1 Then
                DATABASE_ARRIVI.Close()
                DATABASE_ARRIVI.Show()
            End If
            If My.Settings.DTBPartenzeActive = 1 Then
                DATABASE_PARTENZE.Close()
                DATABASE_PARTENZE.Show()
            End If
            If My.Settings.DTBIntelligenceActive = 1 Then
                DATABASE_INTELLIGENCE.Close()
                DATABASE_INTELLIGENCE.Show()
            End If
        Catch ex As Exception

        End Try

        Me.Hide()

    End Sub

    Private Sub DateSelection_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        My.Settings.SelectedDate = ""

    End Sub

    Private Sub DateSelection_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        My.Settings.SelectedDate = ""

    End Sub

End Class