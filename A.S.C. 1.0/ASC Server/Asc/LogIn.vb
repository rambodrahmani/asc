Public Class LogIn

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            If TextBox1.Text = My.Settings.NomeUtente Then
                If TextBox2.Text = My.Settings.PassUtente Then
                    Main.Show()
                    Me.Close()
                Else
                    MsgBox("Password utente errata, acesso non acconsetito", MsgBoxStyle.Information, "Asc - Errore Identificazione")
                End If
            Else
                MsgBox("Nome utente errato, acesso non acconsetito", MsgBoxStyle.Information, "Asc - Errore Identificazione")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button1.PerformClick()
        End If

    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button1.PerformClick()
        End If

    End Sub

    Private Sub LogIn_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim ActiveProcessesCounter As Integer = 0
        Dim ActiveProcessesList() As Process = Diagnostics.Process.GetProcesses
        Dim SingleProcess As Process
        For Each SingleProcess In ActiveProcessesList
            If SingleProcess.ProcessName.Contains("Asc") Then
                ActiveProcessesCounter += 1
            End If
        Next

        If ActiveProcessesCounter > 1 Then
            MsgBox("Impossibile avviare Asc, è già in esecuzione un processo Asc e non è possibile utilizzarne due contemporaneamente. Chiusura.", _
       MsgBoxStyle.Critical, "Asc - Processo già attivo")
            Me.Close()
        End If

    End Sub

End Class