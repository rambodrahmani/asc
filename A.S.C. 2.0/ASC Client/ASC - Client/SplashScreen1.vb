Public NotInheritable Class SplashScreen1

    Dim DataOggi As String = Date.Now.Day.ToString & "." & Date.Now.Month.ToString & "." & Date.Now.Year.ToString

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        If My.Settings.NomeUnita = "" Then
            Me.TopMost = False
            NomeUnita.Show()
            Me.Close()
        Else
            Me.TopMost = False
            Form1.Show()
            Me.Close()
        End If

    End Sub

    Private Sub SplashScreen1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Dim ActiveProcessesCounter As Integer = 0
            Dim ActiveProcessesList() As Process = Diagnostics.Process.GetProcesses
            Dim SingleProcess As Process
            For Each SingleProcess In ActiveProcessesList
                If SingleProcess.ProcessName.Contains("ASC - Client") Then
                    ActiveProcessesCounter += 1
                End If
            Next

            If ActiveProcessesCounter > 1 Then
                Timer1.Stop()
                MsgBox("Impossibile avviare Asc Client, è già in esecuzione un processo Asc Client e non è possibile utilizzarne due contemporaneamente. Chiusura.", _
           MsgBoxStyle.Critical, "Asc Client - Processo già attivo")
                Me.Close()
            Else
                Try
                    Form1.Text = "Asc - Client 2.0    " & "[MACCHINA: " & My.Computer.Name.ToString & "]"
                Catch ex As Exception

                End Try
                Try
                    If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti") Then
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Ricevuti")
                    End If

                    If System.IO.Directory.Exists(Application.StartupPath & "\temp") Then
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\temp")
                    End If
                Catch ex As Exception

                End Try

                Timer1.Start()
                Form1.CreateMessagesFolder()

                Try
                    If System.IO.Directory.Exists(Application.StartupPath & "\File Inviati in ASC") Then
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\File Inviati in ASC")
                    End If

                    If System.IO.Directory.Exists(Application.StartupPath & "\File Inviati in ASC\" & DataOggi) Then
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\File Inviati in ASC\" & DataOggi)
                    End If
                Catch ex As Exception

                End Try
            End If
        Catch ex As Exception

        End Try

    End Sub

End Class
