Public NotInheritable Class SplashScreen1

    Dim scrivi As System.IO.StreamWriter
    Dim leggi As System.IO.StreamReader
    Dim letto As String = ""

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Me.TopMost = False

        Try
            If My.Settings.NomeUnita = "" Then
                MsgBox("Per poter utilizzare Asc Server bisogna prima scegliere il nome dell'unità.", MsgBoxStyle.Information, "Asc Server - Nome Unità")
                NomeUnita.Show()
            Else
                LogIn.Show()
            End If
        Catch ex As Exception

        End Try

        Me.Close()

    End Sub

    Private Sub SplashScreen1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim ActiveProcessesCounter As Integer = 0
        Try
            Dim ActiveProcessesList() As Process = Diagnostics.Process.GetProcesses
            Dim SingleProcess As Process
            For Each SingleProcess In ActiveProcessesList
                If SingleProcess.ProcessName = "Asc" Then
                    ActiveProcessesCounter += 1
                End If
            Next
        Catch ex As Exception

        End Try

        If ActiveProcessesCounter > 1 Then
            Try
                Timer1.Stop()
                MsgBox("Impossibile avviare Asc, è già in esecuzione un processo Asc e non è possibile utilizzarne due contemporaneamente. Chiusura.", _
           MsgBoxStyle.Critical, "Asc Server - Processo già attivo")
                Me.Close()
            Catch ex As Exception

            End Try
        Else
            Timer1.Start()
            Try
                If System.IO.Directory.Exists(Application.StartupPath & "\Memorized") Then
                    If System.IO.File.Exists(Application.StartupPath & "\Memorized\DataChiusura.txt") Then
                        Try
                            leggi = System.IO.File.OpenText(Application.StartupPath & "\Memorized\DataChiusura.txt")
                            letto = leggi.ReadToEnd
                            leggi.Close()
                            My.Settings.DataIeri = letto
                        Catch ex As Exception

                        End Try
                    End If
                    If System.IO.File.Exists(Application.StartupPath & "\Memorized\JDChiusura.txt") Then
                        Try
                            leggi = System.IO.File.OpenText(Application.StartupPath & "\Memorized\JDChiusura.txt")
                            letto = leggi.ReadToEnd
                            leggi.Close()
                            My.Settings.JlDate = letto
                        Catch ex As Exception

                        End Try
                    End If
                    If System.IO.File.Exists(Application.StartupPath & "\Memorized\NrProgressivoArriviChiusura.txt") Then
                        Try
                            leggi = System.IO.File.OpenText(Application.StartupPath & "\Memorized\NrProgressivoArriviChiusura.txt")
                            letto = leggi.ReadToEnd
                            leggi.Close()
                            My.Settings.NumEmail = Val(letto)
                        Catch ex As Exception

                        End Try
                    End If
                    If System.IO.File.Exists(Application.StartupPath & "\Memorized\NrProgressivoPartenzeChiusura.txt") Then
                        Try
                            leggi = System.IO.File.OpenText(Application.StartupPath & "\Memorized\NrProgressivoPartenzeChiusura.txt")
                            letto = leggi.ReadToEnd
                            leggi.Close()
                            My.Settings.NumEmail2 = Val(letto)
                        Catch ex As Exception

                        End Try
                    End If
                    If System.IO.File.Exists(Application.StartupPath & "\Memorized\NrProgressivoIntellChiusura.txt") Then
                        Try
                            leggi = System.IO.File.OpenText(Application.StartupPath & "\Memorized\NrProgressivoIntellChiusura.txt")
                            letto = leggi.ReadToEnd
                            leggi.Close()
                            My.Settings.NumEmailIntelligence = Val(letto)
                        Catch ex As Exception

                        End Try
                    End If
                    If System.IO.File.Exists(Application.StartupPath & "\Memorized\NrProgressivoUfficiChiusura.txt") Then
                        Try
                            leggi = System.IO.File.OpenText(Application.StartupPath & "\Memorized\NrProgressivoUfficiChiusura.txt")
                            letto = leggi.ReadToEnd
                            leggi.Close()
                            My.Settings.Ufficio = Val(letto)
                        Catch ex As Exception

                        End Try
                    End If
                End If
            Catch ex As Exception

            End Try
        End If

    End Sub
End Class
