Public Class SelezioneIndirizzi

    Private Sub SelezioneIndirizzi_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        CaricaUffici()
        CaricaIndirizzi()

    End Sub

    Public Sub CaricaUffici()

        Try
            Dim leggi As System.IO.StreamReader = System.IO.File.OpenText(Application.StartupPath & "\Uffici\listauffici.txt")
            Dim letto As String = leggi.ReadToEnd
            leggi.Close()
            Dim MySplitter() As String = Split(letto, vbCrLf)
            For r = 0 To UBound(MySplitter)
                If MySplitter(r).ToString <> "" Then
                    ListView1.Items.Add(MySplitter(r).ToString)
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Sub CaricaIndirizzi()

        Try
            ListView3.Items.Clear()
            If System.IO.File.Exists(Application.StartupPath & "\Uffici\indirizziselezionati.txt") Then
                Dim leggi As System.IO.StreamReader = System.IO.File.OpenText(Application.StartupPath & "\Uffici\indirizziselezionati.txt")
                Dim letto As String = leggi.ReadToEnd
                leggi.Close()
                Dim MySplitter() As String = Split(letto, vbCrLf)
                For r = 0 To UBound(MySplitter)
                    If MySplitter(r).ToString <> "" Then
                        ListView3.Items.Add(MySplitter(r).ToString)
                    End If
                Next
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ListView1.SelectedIndexChanged

        Try
            ListView2.Items.Clear()
            For Each elementoselezionato As ListViewItem In ListView1.SelectedItems
                Dim leggi As System.IO.StreamReader = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & elementoselezionato.Index.ToString & ".txt")
                Dim letto As String = leggi.ReadToEnd
                leggi.Close()
                Dim MySplitter() As String = Split(letto, "#")
                For r = 1 To UBound(MySplitter)
                    If MySplitter(r).ToString <> "" Then
                        ListView2.Items.Add(MySplitter(r).ToString)
                    End If
                Next
            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        Try
            If ListView2.SelectedItems.Count = 1 Then
                If System.IO.File.Exists(Application.StartupPath & "\Uffici\indirizziselezionati.txt") Then
                    Dim scrivi As System.IO.StreamWriter = System.IO.File.AppendText(Application.StartupPath & "\Uffici\indirizziselezionati.txt")
                    scrivi.WriteLine(Label2.Text)
                    scrivi.Close()
                    MsgBox("Indirizzo aggiunto alla lista " & """" & "Indirizzi Non Classificati" & """" & ".", MsgBoxStyle.Information, "Asc Server - Selezione Indirizzi")
                    CaricaIndirizzi()
                Else
                    Dim scrivi As System.IO.StreamWriter = System.IO.File.CreateText(Application.StartupPath & "\Uffici\indirizziselezionati.txt")
                    scrivi.WriteLine(Label2.Text)
                    scrivi.Close()
                    MsgBox("Indirizzo aggiunto alla lista " & """" & "Indirizzi Non Classificati" & """" & ".", MsgBoxStyle.Information, "Asc Server - Selezione Indirizzi")
                    CaricaIndirizzi()
                End If
            Else
                MsgBox("Non è stato selezionato alcun indirizzo da aggiungere alla lista " & """" & "Indirizzi Non Classificati" & """" & ".", MsgBoxStyle.Information, "Asc Server - Selezione Indirizzi")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ListView2_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ListView2.SelectedIndexChanged

        Try
            For Each elementoselezionato As ListViewItem In ListView2.SelectedItems
                Label2.Text = elementoselezionato.Text.ToString
            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        Try
            For Each elementoselezionato As ListViewItem In ListView3.SelectedItems
                ListView3.Items.RemoveAt(elementoselezionato.Index)
            Next

            If ListView1.Items.Count < 0 Then
                For Each elementolistview As ListViewItem In ListView3.Items
                    Dim scrivi As System.IO.StreamWriter = System.IO.File.CreateText(Application.StartupPath & "\Uffici\indirizziselezionati.txt")
                    scrivi.WriteLine(elementolistview.Text)
                    scrivi.Close()
                Next

                CaricaIndirizzi()
            Else
                System.IO.File.Delete(Application.StartupPath & "\Uffici\indirizziselezionati.txt")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ListView3_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ListView3.SelectedIndexChanged

        Try
            For Each elementoselezionato As ListViewItem In ListView3.SelectedItems
                Label2.Text = elementoselezionato.Text.ToString
            Next
        Catch ex As Exception

        End Try

    End Sub
End Class