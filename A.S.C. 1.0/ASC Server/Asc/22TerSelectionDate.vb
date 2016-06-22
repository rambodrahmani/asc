Public Class TerSelectionDate

    Dim datadacercare As String
    Dim datadef As String

    Private Sub MonthCalendar1_DateChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged

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
            My.Settings.SelectedDate2 = datadacercare
            Label2.Text = My.Settings.SelectedDate2
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TarSelectionDate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        My.Settings.SelectedDate2 = ""
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
            My.Settings.SelectedDate2 = datadacercare
            Label2.Text = My.Settings.SelectedDate2
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        DATABASE_ARRIVI.Show()
        CloseDTB.Show()
        Me.Close()

    End Sub

End Class