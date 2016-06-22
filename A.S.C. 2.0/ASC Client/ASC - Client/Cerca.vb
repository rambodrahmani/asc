﻿Imports System.IO.File

Public Class Cerca

    Dim datagiorno As String
    Dim datagiorni() As String
    Dim datadacercare As String
    Dim datadef As String
    Dim scrivi As System.IO.StreamWriter
    Dim leggi As System.IO.StreamReader
    Dim letto As String
    Dim x As Integer = 0
    Dim y As Integer = 0
    Dim intell As Integer = 0
    Dim DTBArrivi As Integer = 0
    Dim DTBPartenze As Integer = 0
    Dim DTBIntelligence As Integer = 0
    Dim DefinitiveSearchesString As String = ""
    Dim MeseDaCercare As String = ""
    Dim juliandatedacercare As Integer = 0

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        DefinitiveSearchesString = ""
        CheckRequirements()

    End Sub

    Public Sub CheckRequirements()

        Try
            If TextBox1.Text = "" Then
                MsgBox("Scrivi il testo da cercare.  Impossibile continuare.", MsgBoxStyle.Information, "Asc Client - Ricerca")
                Exit Sub
            End If

            If ComboBox1.SelectedItem.ToString = "" Then
                MsgBox("Selezionare la modalità di ricerca. Impossibile continuare.", MsgBoxStyle.Information, "Asc Client - Ricerca")
                Exit Sub
            End If

            If DTBArrivi = 0 And DTBPartenze = 0 And DTBIntelligence = 0 Then
                MsgBox("Selezionare almeno un database in cui cercare. Impossibile continuare.", MsgBoxStyle.Information, "Asc Client - Ricerca")
                Exit Sub
            End If

            If CheckBox4.Checked = False And CheckBox5.Checked = False Then
                MsgBox("Selezionare almeno un tipo di messaggistica in cui cercare. Impossibile continuare.", MsgBoxStyle.Information, "Asc Client - Ricerca")
                Exit Sub
            End If

            If ComboBox1.SelectedItem.ToString = "Ricerca Mensile" Then
                If CheckBox5.Checked = True Then
                    MsgBox("Con la modalità di ricerca " & """" & "Ricerca Mensile" & """" & " non è possibile svolgere la ricerca tra i messaggi Ricevuti. Impossibile continuare.", MsgBoxStyle.Information, "Asc Client - Ricerca")
                    Exit Sub
                End If
            End If

            If ComboBox1.SelectedItem.ToString = "Ricerca Mensile" Then
                If CheckBox4.Checked = True Then
                    If ListView1.SelectedItems.Count > 0 Then
                        Dim ListviewItem As ListViewItem
                        Dim ListviewSelectedText As String = ""
                        For Each ListviewItem In ListView1.SelectedItems
                            ListviewSelectedText = ListviewItem.Text.ToString
                        Next
                        Dim SelectedMonthNumber As Integer = 0
                        If ListviewSelectedText = "Gennaio" Then
                            SelectedMonthNumber = 1
                            Dim DirectoryCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\File Inviati in ASC")
                            For r = 0 To UBound(DirectoryCollection)
                                Dim DirectoryName As New System.IO.DirectoryInfo(DirectoryCollection(r).ToString)
                                Dim MySplitter() As String = Split(DirectoryName.Name.ToString, ".")
                                If MySplitter(1).ToString = SelectedMonthNumber Then
                                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(DirectoryCollection(r).ToString)
                                    x = FilesCollection.Count - 1
                                    For z = 0 To x
                                        leggi = System.IO.File.OpenText(FilesCollection(z).ToString)
                                        letto = leggi.ReadToEnd
                                        leggi.Close()
                                        If letto.Contains(TextBox1.Text) Then
                                            DefinitiveSearchesString += vbCrLf & letto & vbCrLf
                                        End If
                                    Next
                                End If
                            Next
                        ElseIf ListviewSelectedText = "Febbraio" Then
                            SelectedMonthNumber = 2
                            Dim DirectoryCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\File Inviati in ASC")
                            For r = 0 To UBound(DirectoryCollection)
                                Dim DirectoryName As New System.IO.DirectoryInfo(DirectoryCollection(r).ToString)
                                Dim MySplitter() As String = Split(DirectoryName.Name.ToString, ".")
                                If MySplitter(1).ToString = SelectedMonthNumber Then
                                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(DirectoryCollection(r).ToString)
                                    x = FilesCollection.Count - 1
                                    For z = 0 To x
                                        leggi = System.IO.File.OpenText(FilesCollection(z).ToString)
                                        letto = leggi.ReadToEnd
                                        leggi.Close()
                                        If letto.Contains(TextBox1.Text) Then
                                            DefinitiveSearchesString += vbCrLf & letto & vbCrLf
                                        End If
                                    Next
                                End If
                            Next
                        ElseIf ListviewSelectedText = "Marzo" Then
                            SelectedMonthNumber = 3
                            Dim DirectoryCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\File Inviati in ASC")
                            For r = 0 To UBound(DirectoryCollection)
                                Dim DirectoryName As New System.IO.DirectoryInfo(DirectoryCollection(r).ToString)
                                Dim MySplitter() As String = Split(DirectoryName.Name.ToString, ".")
                                If MySplitter(1).ToString = SelectedMonthNumber Then
                                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(DirectoryCollection(r).ToString)
                                    x = FilesCollection.Count - 1
                                    For z = 0 To x
                                        leggi = System.IO.File.OpenText(FilesCollection(z).ToString)
                                        letto = leggi.ReadToEnd
                                        leggi.Close()
                                        If letto.Contains(TextBox1.Text) Then
                                            DefinitiveSearchesString += vbCrLf & letto & vbCrLf
                                        End If
                                    Next
                                End If
                            Next
                        ElseIf ListviewSelectedText = "Aprile" Then
                            SelectedMonthNumber = 4
                            Dim DirectoryCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\File Inviati in ASC")
                            For r = 0 To UBound(DirectoryCollection)
                                Dim DirectoryName As New System.IO.DirectoryInfo(DirectoryCollection(r).ToString)
                                Dim MySplitter() As String = Split(DirectoryName.Name.ToString, ".")
                                If MySplitter(1).ToString = SelectedMonthNumber Then
                                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(DirectoryCollection(r).ToString)
                                    x = FilesCollection.Count - 1
                                    For z = 0 To x
                                        leggi = System.IO.File.OpenText(FilesCollection(z).ToString)
                                        letto = leggi.ReadToEnd
                                        leggi.Close()
                                        If letto.Contains(TextBox1.Text) Then
                                            DefinitiveSearchesString += vbCrLf & letto & vbCrLf
                                        End If
                                    Next
                                End If
                            Next
                        ElseIf ListviewSelectedText = "Maggio" Then
                            SelectedMonthNumber = 5
                            Dim DirectoryCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\File Inviati in ASC")
                            For r = 0 To UBound(DirectoryCollection)
                                Dim DirectoryName As New System.IO.DirectoryInfo(DirectoryCollection(r).ToString)
                                Dim MySplitter() As String = Split(DirectoryName.Name.ToString, ".")
                                If MySplitter(1).ToString = SelectedMonthNumber Then
                                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(DirectoryCollection(r).ToString)
                                    x = FilesCollection.Count - 1
                                    For z = 0 To x
                                        leggi = System.IO.File.OpenText(FilesCollection(z).ToString)
                                        letto = leggi.ReadToEnd
                                        leggi.Close()
                                        If letto.Contains(TextBox1.Text) Then
                                            DefinitiveSearchesString += vbCrLf & letto & vbCrLf
                                        End If
                                    Next
                                End If
                            Next
                        ElseIf ListviewSelectedText = "Giugno" Then
                            SelectedMonthNumber = 6
                            Dim DirectoryCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\File Inviati in ASC")
                            For r = 0 To UBound(DirectoryCollection)
                                Dim DirectoryName As New System.IO.DirectoryInfo(DirectoryCollection(r).ToString)
                                Dim MySplitter() As String = Split(DirectoryName.Name.ToString, ".")
                                If MySplitter(1).ToString = SelectedMonthNumber Then
                                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(DirectoryCollection(r).ToString)
                                    x = FilesCollection.Count - 1
                                    For z = 0 To x
                                        leggi = System.IO.File.OpenText(FilesCollection(z).ToString)
                                        letto = leggi.ReadToEnd
                                        leggi.Close()
                                        If letto.Contains(TextBox1.Text) Then
                                            DefinitiveSearchesString += vbCrLf & letto & vbCrLf
                                        End If
                                    Next
                                End If
                            Next
                        ElseIf ListviewSelectedText = "Luglio" Then
                            SelectedMonthNumber = 7
                            Dim DirectoryCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\File Inviati in ASC")
                            For r = 0 To UBound(DirectoryCollection)
                                Dim DirectoryName As New System.IO.DirectoryInfo(DirectoryCollection(r).ToString)
                                Dim MySplitter() As String = Split(DirectoryName.Name.ToString, ".")
                                If MySplitter(1).ToString = SelectedMonthNumber Then
                                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(DirectoryCollection(r).ToString)
                                    x = FilesCollection.Count - 1
                                    For z = 0 To x
                                        leggi = System.IO.File.OpenText(FilesCollection(z).ToString)
                                        letto = leggi.ReadToEnd
                                        leggi.Close()
                                        If letto.Contains(TextBox1.Text) Then
                                            DefinitiveSearchesString += vbCrLf & letto & vbCrLf
                                        End If
                                    Next
                                End If
                            Next
                        ElseIf ListviewSelectedText = "Agosto" Then
                            SelectedMonthNumber = 8
                            Dim DirectoryCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\File Inviati in ASC")
                            For r = 0 To UBound(DirectoryCollection)
                                Dim DirectoryName As New System.IO.DirectoryInfo(DirectoryCollection(r).ToString)
                                Dim MySplitter() As String = Split(DirectoryName.Name.ToString, ".")
                                If MySplitter(1).ToString = SelectedMonthNumber Then
                                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(DirectoryCollection(r).ToString)
                                    x = FilesCollection.Count - 1
                                    For z = 0 To x
                                        leggi = System.IO.File.OpenText(FilesCollection(z).ToString)
                                        letto = leggi.ReadToEnd
                                        leggi.Close()
                                        If letto.Contains(TextBox1.Text) Then
                                            DefinitiveSearchesString += vbCrLf & letto & vbCrLf
                                        End If
                                    Next
                                End If
                            Next
                        ElseIf ListviewSelectedText = "Settembre" Then
                            SelectedMonthNumber = 9
                            Dim DirectoryCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\File Inviati in ASC")
                            For r = 0 To UBound(DirectoryCollection)
                                Dim DirectoryName As New System.IO.DirectoryInfo(DirectoryCollection(r).ToString)
                                Dim MySplitter() As String = Split(DirectoryName.Name.ToString, ".")
                                If MySplitter(1).ToString = SelectedMonthNumber Then
                                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(DirectoryCollection(r).ToString)
                                    x = FilesCollection.Count - 1
                                    For z = 0 To x
                                        leggi = System.IO.File.OpenText(FilesCollection(z).ToString)
                                        letto = leggi.ReadToEnd
                                        leggi.Close()
                                        If letto.Contains(TextBox1.Text) Then
                                            DefinitiveSearchesString += vbCrLf & letto & vbCrLf
                                        End If
                                    Next
                                End If
                            Next
                        ElseIf ListviewSelectedText = "Ottobre" Then
                            SelectedMonthNumber = 10
                            Dim DirectoryCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\File Inviati in ASC")
                            For r = 0 To UBound(DirectoryCollection)
                                Dim DirectoryName As New System.IO.DirectoryInfo(DirectoryCollection(r).ToString)
                                Dim MySplitter() As String = Split(DirectoryName.Name.ToString, ".")
                                If MySplitter(1).ToString = SelectedMonthNumber Then
                                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(DirectoryCollection(r).ToString)
                                    x = FilesCollection.Count - 1
                                    For z = 0 To x
                                        leggi = System.IO.File.OpenText(FilesCollection(z).ToString)
                                        letto = leggi.ReadToEnd
                                        leggi.Close()
                                        If letto.Contains(TextBox1.Text) Then
                                            DefinitiveSearchesString += vbCrLf & letto & vbCrLf
                                        End If
                                    Next
                                End If
                            Next
                        ElseIf ListviewSelectedText = "Novembre" Then
                            SelectedMonthNumber = 11
                            Dim DirectoryCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\File Inviati in ASC")
                            For r = 0 To UBound(DirectoryCollection)
                                Dim DirectoryName As New System.IO.DirectoryInfo(DirectoryCollection(r).ToString)
                                Dim MySplitter() As String = Split(DirectoryName.Name.ToString, ".")
                                If MySplitter(1).ToString = SelectedMonthNumber Then
                                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(DirectoryCollection(r).ToString)
                                    x = FilesCollection.Count - 1
                                    For z = 0 To x
                                        leggi = System.IO.File.OpenText(FilesCollection(z).ToString)
                                        letto = leggi.ReadToEnd
                                        leggi.Close()
                                        If letto.Contains(TextBox1.Text) Then
                                            DefinitiveSearchesString += vbCrLf & letto & vbCrLf
                                        End If
                                    Next
                                End If
                            Next
                        ElseIf ListviewSelectedText = "Dicembre" Then
                            SelectedMonthNumber = 12
                            Dim DirectoryCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\File Inviati in ASC")
                            For r = 0 To UBound(DirectoryCollection)
                                Dim DirectoryName As New System.IO.DirectoryInfo(DirectoryCollection(r).ToString)
                                Dim MySplitter() As String = Split(DirectoryName.Name.ToString, ".")
                                If MySplitter(1).ToString = SelectedMonthNumber Then
                                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(DirectoryCollection(r).ToString)
                                    x = FilesCollection.Count - 1
                                    For z = 0 To x
                                        leggi = System.IO.File.OpenText(FilesCollection(z).ToString)
                                        letto = leggi.ReadToEnd
                                        leggi.Close()
                                        If letto.Contains(TextBox1.Text) Then
                                            DefinitiveSearchesString += vbCrLf & letto & vbCrLf
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Else
                        MsgBox("Selezionare il mese in cui cercare.", MsgBoxStyle.Information, "Asc Client - Ricerca")
                    End If
                End If
            End If

            If CheckBox4.Checked = True Then
                Dim FilesCollection2() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\File Inviati in ASC\" & datadacercare)
                x = FilesCollection2.Count - 1
                Try
                    For r = 0 To x
                        leggi = System.IO.File.OpenText(FilesCollection2(r).ToString)
                        letto = leggi.ReadToEnd
                        leggi.Close()
                        If letto.Contains(TextBox1.Text) Then
                            DefinitiveSearchesString += vbCrLf & letto & vbCrLf
                        End If
                    Next
                Catch ex As Exception

                End Try
            End If

            If DTBArrivi = 1 Then
                SearchDTBArrivi()
                Exit Sub
            ElseIf DTBPartenze = 1 Then
                SearchDTBPartenze()
                Exit Sub
            ElseIf DTBIntelligence = 1 Then
                SearchDTBIntelligence()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Errore durante il processo di ricerca, messaggio errore " & ex.Message, MsgBoxStyle.Information, "Asc Client - Errore Ricerca")
        End Try

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

        Try
            If CheckBox1.Checked = True Then
                DTBArrivi = 1
            Else
                DTBArrivi = 0
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged

        Try
            If CheckBox2.Checked = True Then
                DTBPartenze = 1
            Else
                DTBPartenze = 0
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged

        Try
            If CheckBox3.Checked = True Then
                DTBIntelligence = 1
            Else
                DTBIntelligence = 0
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        Try
            If ComboBox1.SelectedIndex = 0 Then
                ListView1.Visible = False
                MonthCalendar1.Visible = True
                Label1.Text = "Data selezionata in cui cercare:"
                Label2.Text = datadacercare
            ElseIf ComboBox1.SelectedIndex = 1 Then
                ListView1.Visible = True
                MonthCalendar1.Visible = False
                Label1.Text = "Mese selezionato in cui cercare:"
                Label2.Text = MeseDaCercare
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged

        Try
            Dim selecteditem As ListViewItem
            For Each selecteditem In ListView1.SelectedItems
                MeseDaCercare = selecteditem.Text.ToString
                Label2.Text = MeseDaCercare
            Next
        Catch ex As Exception

        End Try

    End Sub

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
            Label2.Text = datadacercare
            juliandatedacercare = MonthCalendar1.SelectionStart.DayOfYear
        Catch ex As Exception

        End Try

    End Sub

    Public Sub SearchDTBArrivi()

        Try
            If ComboBox1.SelectedItem.ToString = "Ricerca Giornaliera" Then
                If CheckBox5.Checked = True Then
                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Messaggi\" & juliandatedacercare.ToString("D3") & "\Arrivi")
                    x = FilesCollection.Count - 1
                    Try
                        For r = 0 To x
                            leggi = System.IO.File.OpenText(FilesCollection(r).ToString)
                            letto = leggi.ReadToEnd
                            leggi.Close()
                            If letto.Contains(TextBox1.Text) Then
                                DefinitiveSearchesString += vbCrLf & letto & vbCrLf
                            End If
                        Next
                    Catch ex As Exception

                    End Try
                End If
            End If

            DTBArrivi = 0
            If DTBPartenze = 1 Then
                SearchDTBPartenze()
                Exit Sub
            ElseIf DTBIntelligence = 1 Then
                SearchDTBIntelligence()
                Exit Sub
            Else
                ShowMessageText()
            End If
        Catch ex As Exception
            MsgBox("Errore durante il processo di ricerca, messaggio errore " & ex.Message, MsgBoxStyle.Information, "Asc Client - Errore Ricerca")
        End Try

    End Sub

    Public Sub SearchDTBPartenze()

        Try
            If ComboBox1.SelectedItem.ToString = "Ricerca Giornaliera" Then
                If CheckBox5.Checked = True Then
                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Messaggi\" & juliandatedacercare.ToString("D3") & "\Partenze")
                    x = FilesCollection.Count - 1
                    Try
                        For r = 0 To x
                            leggi = System.IO.File.OpenText(FilesCollection(r).ToString)
                            letto = leggi.ReadToEnd
                            leggi.Close()
                            If letto.Contains(TextBox1.Text) Then
                                DefinitiveSearchesString += vbCrLf & letto & vbCrLf
                            End If
                        Next
                    Catch ex As Exception

                    End Try
                End If
            End If

            DTBPartenze = 0
            If DTBArrivi = 1 Then
                SearchDTBArrivi()
                Exit Sub
            ElseIf DTBIntelligence = 1 Then
                SearchDTBIntelligence()
                Exit Sub
            Else
                ShowMessageText()
            End If
        Catch ex As Exception
            MsgBox("Errore durante il processo di ricerca, messaggio errore " & ex.Message, MsgBoxStyle.Information, "Asc Client - Errore Ricerca")
        End Try

    End Sub

    Public Sub SearchDTBIntelligence()

        Try
            If ComboBox1.SelectedItem.ToString = "Ricerca Giornaliera" Then
                If CheckBox5.Checked = True Then
                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Messaggi\" & juliandatedacercare.ToString("D3") & "\Intelligence")
                    x = FilesCollection.Count - 1
                    Try
                        For r = 0 To x
                            leggi = System.IO.File.OpenText(FilesCollection(r).ToString)
                            letto = leggi.ReadToEnd
                            leggi.Close()
                            If letto.Contains(TextBox1.Text) Then
                                DefinitiveSearchesString += vbCrLf & letto & vbCrLf
                            End If
                        Next
                    Catch ex As Exception

                    End Try
                End If
            End If

            DTBIntelligence = 0
            If DTBArrivi = 1 Then
                SearchDTBArrivi()
                Exit Sub
            ElseIf DTBPartenze = 1 Then
                SearchDTBPartenze()
                Exit Sub
            Else
                ShowMessageText()
            End If
        Catch ex As Exception
            MsgBox("Errore durante il processo di ricerca, messaggio errore " & ex.Message, MsgBoxStyle.Information, "Asc Client - Errore Ricerca")
        End Try

    End Sub

    Public Sub ShowMessageText()

        Try
            If DefinitiveSearchesString <> "" Then
                SearchResults.TextBox1.Text = DefinitiveSearchesString
                SearchResults.Show()
                CheckBox1.Checked = False
                CheckBox2.Checked = False
                CheckBox3.Checked = False
            Else
                MsgBox("La ricerca non ha prodotto alcun risultato.", MsgBoxStyle.Information, "Asc Client - Ricerca")
                CheckBox1.Checked = False
                CheckBox2.Checked = False
                CheckBox3.Checked = False
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Cerca_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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
            Label2.Text = datadacercare
            juliandatedacercare = MonthCalendar1.SelectionStart.DayOfYear
        Catch ex As Exception

        End Try

    End Sub

End Class