Public Class Chiusure

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        Try
            If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                My.Settings.PercorsoChiusura = FolderBrowserDialog1.SelectedPath.ToString
                If FolderBrowserDialog1.SelectedPath.ToString.Length > 50 Then
                    Label2.Text = ""
                    For r = 0 To 45
                        Label2.Text += FolderBrowserDialog1.SelectedPath(r).ToString
                    Next
                    Label2.Text += "..."
                Else
                    Label2.Text = FolderBrowserDialog1.SelectedPath.ToString
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Chiusure_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Try
            If My.Settings.PercorsoChiusura.Length > 50 Then
                For r = 0 To 45
                    Label2.Text += My.Settings.PercorsoChiusura(r).ToString
                Next
                Label2.Text += "..."
            Else
                Label2.Text = My.Settings.PercorsoChiusura
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        Try
            If My.Settings.DataUltimaChiusura <> My.Settings.DataOggi Then
                If My.Settings.PercorsoChiusura <> "" Then
                    Dim lettore As System.IO.StreamReader = System.IO.File.OpenText(Application.StartupPath & "\Email\DATABASEARRIVI\dateusate.txt")
                    Dim letto As String = lettore.ReadToEnd
                    lettore.Close()
                    Dim MySplitter() As String = Split(letto, vbCrLf)
                    If MySplitter.Count > 2 Then
                    Else
                        MsgBox("Non sono presenti sufficienti date lavorative per poter eseguire la chiusura.", MsgBoxStyle.Critical, "Asc Server - Chiusure")
                        Exit Sub
                    End If
                    Dim DataSelezionata As String = MySplitter(MySplitter.Count - 3).ToString
                    Dim InstantSplitter() As String = Split(DataSelezionata & ".", ".")
                    Dim ConvertitoreMese As String = ""
                    If Val(InstantSplitter(1).ToString) = 1 Then
                        ConvertitoreMese = "Gennaio"
                    ElseIf Val(InstantSplitter(1).ToString) = 2 Then
                        ConvertitoreMese = "Febbraio"
                    ElseIf Val(InstantSplitter(1).ToString) = 3 Then
                        ConvertitoreMese = "Marzo"
                    ElseIf Val(InstantSplitter(1).ToString) = 4 Then
                        ConvertitoreMese = "Aprile"
                    ElseIf Val(InstantSplitter(1).ToString) = 5 Then
                        ConvertitoreMese = "Maggio"
                    ElseIf Val(InstantSplitter(1).ToString) = 6 Then
                        ConvertitoreMese = "Giugno"
                    ElseIf Val(InstantSplitter(1).ToString) = 7 Then
                        ConvertitoreMese = "Luglio"
                    ElseIf Val(InstantSplitter(1).ToString) = 8 Then
                        ConvertitoreMese = "Agosto"
                    ElseIf Val(InstantSplitter(1).ToString) = 9 Then
                        ConvertitoreMese = "Settembre"
                    ElseIf Val(InstantSplitter(1).ToString) = 10 Then
                        ConvertitoreMese = "Ottobre"
                    ElseIf Val(InstantSplitter(1).ToString) = 11 Then
                        ConvertitoreMese = "Novembre"
                    ElseIf Val(InstantSplitter(1).ToString) = 12 Then
                        ConvertitoreMese = "Dicembre"
                    End If
                    If System.IO.Directory.Exists(My.Settings.PercorsoChiusura & "\Chiusure\RACCOLTA ASC\ANNO " & InstantSplitter(2).ToString & "\" & ConvertitoreMese.ToUpper & "\" & InstantSplitter(0).ToString & "\DATABASEARRIVI") Then
                    Else
                        System.IO.Directory.CreateDirectory(My.Settings.PercorsoChiusura & "\Chiusure\RACCOLTA ASC\ANNO " & InstantSplitter(2).ToString & "\" & ConvertitoreMese.ToUpper & "\" & InstantSplitter(0).ToString & "\DATABASEARRIVI")
                    End If
                    My.Computer.FileSystem.CopyDirectory(Application.StartupPath & "\Email\DATABASEARRIVI\" & MySplitter(MySplitter.Count - 3).ToString, My.Settings.PercorsoChiusura & "\Chiusure\RACCOLTA ASC\ANNO " & InstantSplitter(2).ToString & "\" & ConvertitoreMese.ToUpper & "\" & InstantSplitter(0).ToString & "\DATABASEARRIVI")

                    lettore = System.IO.File.OpenText(Application.StartupPath & "\Email\DATABASEPARTENZE\dateusate.txt")
                    letto = lettore.ReadToEnd
                    lettore.Close()
                    MySplitter = Split(letto, vbCrLf)
                    If MySplitter.Count > 2 Then
                    Else
                        MsgBox("Non sono presenti sufficienti date lavorative per poter eseguire la chiusura.", MsgBoxStyle.Critical, "Asc Server - Chiusure")
                        Exit Sub
                    End If
                    DataSelezionata = MySplitter(MySplitter.Count - 3).ToString
                    InstantSplitter = Split(DataSelezionata & ".", ".")
                    ConvertitoreMese = ""
                    If Val(InstantSplitter(1).ToString) = 1 Then
                        ConvertitoreMese = "Gennaio"
                    ElseIf Val(InstantSplitter(1).ToString) = 2 Then
                        ConvertitoreMese = "Febbraio"
                    ElseIf Val(InstantSplitter(1).ToString) = 3 Then
                        ConvertitoreMese = "Marzo"
                    ElseIf Val(InstantSplitter(1).ToString) = 4 Then
                        ConvertitoreMese = "Aprile"
                    ElseIf Val(InstantSplitter(1).ToString) = 5 Then
                        ConvertitoreMese = "Maggio"
                    ElseIf Val(InstantSplitter(1).ToString) = 6 Then
                        ConvertitoreMese = "Giugno"
                    ElseIf Val(InstantSplitter(1).ToString) = 7 Then
                        ConvertitoreMese = "Luglio"
                    ElseIf Val(InstantSplitter(1).ToString) = 8 Then
                        ConvertitoreMese = "Agosto"
                    ElseIf Val(InstantSplitter(1).ToString) = 9 Then
                        ConvertitoreMese = "Settembre"
                    ElseIf Val(InstantSplitter(1).ToString) = 10 Then
                        ConvertitoreMese = "Ottobre"
                    ElseIf Val(InstantSplitter(1).ToString) = 11 Then
                        ConvertitoreMese = "Novembre"
                    ElseIf Val(InstantSplitter(1).ToString) = 12 Then
                        ConvertitoreMese = "Dicembre"
                    End If
                    If System.IO.Directory.Exists(My.Settings.PercorsoChiusura & "\Chiusure\RACCOLTA ASC\ANNO " & InstantSplitter(2).ToString & "\" & ConvertitoreMese.ToUpper & "\" & InstantSplitter(0).ToString & "\DATABASEPARTENZE") Then
                    Else
                        System.IO.Directory.CreateDirectory(My.Settings.PercorsoChiusura & "\Chiusure\RACCOLTA ASC\ANNO " & InstantSplitter(2).ToString & "\" & ConvertitoreMese.ToUpper & "\" & InstantSplitter(0).ToString & "\DATABASEPARTENZE")
                    End If
                    My.Computer.FileSystem.CopyDirectory(Application.StartupPath & "\Email\DATABASEPARTENZE\" & MySplitter(MySplitter.Count - 3).ToString, My.Settings.PercorsoChiusura & "\Chiusure\RACCOLTA ASC\ANNO " & InstantSplitter(2).ToString & "\" & ConvertitoreMese.ToUpper & "\" & InstantSplitter(0).ToString & "\DATABASEPARTENZE")
                    My.Settings.DataUltimaChiusura = My.Settings.DataOggi
                    MsgBox("Chiusura messaggi per la data " & DataSelezionata & " effettuata con successo.", MsgBoxStyle.Information, "Asc Server - Chiusure")
                    Me.Close()
                Else
                    MsgBox("Non è ancora stata scelta una cartella per effettuare le chiusure, impossibile proseguire..", MsgBoxStyle.Information, "Asc Server - Chiusure")
                End If
            Else
                MsgBox("La chiusura dei messaggi risulta già essere stata effettuata.", MsgBoxStyle.Information, "Asc Server - Chiusure")
            End If
        Catch ex As Exception

        End Try

    End Sub
End Class