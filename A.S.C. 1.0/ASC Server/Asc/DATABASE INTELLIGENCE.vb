Public Class DATABASE_INTELLIGENCE

    Dim GetMousePosition As Point
    Dim Breaker As Integer = 0
    Dim TerPrintPageNum As Integer = 0
    Dim DatagridviewRowToPrint As String = ""
    Dim Print22TerAvailable As Integer = 0
    Dim NumFoglioStampa As Integer = 0
    Dim stoppami As Integer = 0
    Dim numeroriga As Integer = 0
    Dim righelette() As String
    Dim dastampare As String
    Dim stringNr As String
    Dim maildasalvare As String
    Dim datadefinitiva() As String
    Dim datadacercare As String
    Dim riga As Integer
    Dim contMailInt As Integer
    Dim contMail() As String
    Dim contatore As Integer = 0
    Dim datagiorno As String
    Dim scrivi As System.IO.StreamWriter
    Dim leggi As System.IO.StreamReader
    Dim letto As String
    Dim info() As String
    Dim numNr As Integer = 0
    Dim numEmail As Integer = 1
    Dim numOgg As Integer = 2
    Dim numData As Integer = 3
    Dim numAll As Integer = 4
    Dim numGDO As Integer = 5
    Dim numFM As Integer = 6
    Declare Auto Function mouse_event Lib "user32.dll" (ByVal dwflags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal dwData As Integer, ByVal dwExtraInfo As Integer) As Integer
    Const MOUSEEVENTF_LEFTDOWN As Integer = 2
    Const MOUSEEVENTF_LEFTUP As Integer = 4
    Const MOUSEEVENTF_RIGHTDOWN As Integer = 8
    Const MOUSEEVENTF_RIGHTUP As Integer = 16

    Public Sub UseMessage()

        Dim cell As DataGridViewCell
        Try
            For Each cell In DataGridView1.SelectedCells
                riga = cell.RowIndex
                datadacercare = DataGridView1.Rows(riga).Cells(3).Value.ToString
                datadefinitiva = Split(datadacercare, " ")
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Email\INTELLIGENCE\" & datadefinitiva(0) & "\emailinviate" & (cell.RowIndex + 1).ToString("D3") & "INTEL.txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                Main.TextBox3.Text = letto
                Main.TopMost = True
                Main.TopMost = False
            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Sub DATABASE_INTELLIGENCE_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        My.Settings.DTBIntelligenceActive = 0
        My.Settings.Intelligence22Ter = ""

    End Sub

    Private Sub DATABASE_INTELLIGENCE_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            My.Settings.DTBIntelligenceActive = 1
            datagiorno = My.Settings.DataOggi
            DataGridView1.Rows.Add(1)
            LoadList()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub LoadList()

        If My.Settings.Intelligence22Ter = "" Then
            If My.Settings.SelectedDate = "" Then

                Try
                    If System.IO.File.Exists(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt") Then
                        leggi = System.IO.File.OpenText(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                        letto = leggi.ReadToEnd
                        leggi.Close()
                        info = Split(letto, "#")
                        contMail = Split(letto, vbCrLf)
                        contMailInt = contMail.Count
                        LoadNr()
                    ElseIf letto = "" Then
                        MsgBox("Non sono state inviate, nella data selezionata, email da visualizzare!", MsgBoxStyle.Information)
                    Else
                        MsgBox("Non sono state inviate, nella data selezionata, email da visualizzare!", MsgBoxStyle.Information)
                    End If
                Catch ex As Exception

                End Try

            Else

                Try
                    If System.IO.File.Exists(Application.StartupPath & "\Email\INTELLIGENCE\" & My.Settings.SelectedDate & "\datiemailinviate" & My.Settings.SelectedDate & ".txt") Then
                        leggi = System.IO.File.OpenText(Application.StartupPath & "\Email\INTELLIGENCE\" & My.Settings.SelectedDate & "\datiemailinviate" & My.Settings.SelectedDate & ".txt")
                        letto = leggi.ReadToEnd
                        leggi.Close()
                        info = Split(letto, "#")
                        contMail = Split(letto, vbCrLf)
                        contMailInt = contMail.Count
                        LoadNr()
                    ElseIf letto = "" Then
                        MsgBox("Non sono state inviate, nella data selezionata, email da visualizzare!", MsgBoxStyle.Information)
                    Else
                        MsgBox("Non sono state inviate, nella data selezionata, email da visualizzare!", MsgBoxStyle.Information)
                    End If
                Catch ex As Exception

                End Try

            End If
        Else
            Load22TerList()
        End If

    End Sub

    Public Sub Load22TerList()

        If My.Settings.Intelligence22Ter <> "" Then
            Try
                If System.IO.File.Exists(Application.StartupPath & "\Email\INTELLIGENCE\" & My.Settings.Intelligence22Ter & "\datiemailinviate" & My.Settings.Intelligence22Ter & ".txt") Then
                    leggi = System.IO.File.OpenText(Application.StartupPath & "\Email\INTELLIGENCE\" & My.Settings.Intelligence22Ter & "\datiemailinviate" & My.Settings.Intelligence22Ter & ".txt")
                    letto = leggi.ReadToEnd
                    leggi.Close()
                    info = Split(letto, "#")
                    contMail = Split(letto, vbCrLf)
                    contMailInt = contMail.Count
                    Print22TerAvailable = 1
                    LoadNr()
                ElseIf letto = "" Then
                    MsgBox("Non sono state inviate, nella data selezionata, email da visualizzare!", MsgBoxStyle.Information, "Asc - Database Partenze")
                End If
            Catch ex As Exception

            End Try
        End If

        If DataGridView1.Rows.Count < 3 Then
            TerPrintPageNum = My.Settings.TerPrintPageNum
            stoppami = 0
            Start22TerPrint()
        End If

    End Sub

    Public Sub LoadNr()

        Try
            For r = 0 To UBound(info) - 1
                DataGridView1.Rows(r).Cells(0).Value = info(numNr)
                numNr += 7
                DataGridView1.Rows.Add(1)
                If r > contMailInt - 3 Then
                    LoadEmail()
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Sub LoadEmail()

        Try
            For r = 0 To UBound(info) - 1
                Dim stringaemail As String = ""
                Dim sceltaemail() As String
                sceltaemail = Split(info(numEmail), "/")
                For z = 0 To UBound(sceltaemail)
                    stringaemail += sceltaemail(z) & ", "
                Next
                DataGridView1.Rows(r).Cells(1).Value = stringaemail
                numEmail += 7
                If r > contMailInt - 3 Then
                    LoadSubject()
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Sub LoadSubject()

        Try
            For r = 0 To UBound(info) - 1
                DataGridView1.Rows(r).Cells(2).Value = info(numOgg)
                numOgg += 7
                If r > contMailInt - 3 Then
                    LoadDate()
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Sub LoadDate()

        Try
            For r = 0 To UBound(info) - 1
                DataGridView1.Rows(r).Cells(3).Value = info(numData)
                numData += 7
                If r > contMailInt - 3 Then
                    LoadAll()
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Sub LoadAll()

        Try
            For r = 0 To UBound(info) - 1
                DataGridView1.Rows(r).Cells(4).Value = info(numAll)
                numAll += 7
                If r > contMailInt - 3 Then
                    LoadRiga1()
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Sub LoadRiga1()

        Try
            For r = 0 To UBound(info) - 1
                DataGridView1.Rows(r).Cells(5).Value = info(numGDO)
                numGDO += 7
                If r > contMailInt - 3 Then
                    LoadRiga2()
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Sub LoadRiga2()

        Try
            For r = 0 To UBound(info) - 1
                DataGridView1.Rows(r).Cells(6).Value = info(numFM)
                numFM += 7
            Next
        Catch ex As Exception

        End Try

        If Print22TerAvailable = 1 Then
            Print22TerList()
        End If

    End Sub

    Public Sub Print22TerList()

        Try
            scrivi = System.IO.File.CreateText(Application.StartupPath & "\22TarPrintPage" & TerPrintPageNum & ".txt")
            scrivi.Write("                          22-Ter del: " & My.Settings.Intelligence22Ter & vbCrLf & vbCrLf & "Nr. Progressivo" & "         " & "GDO Messaggio:" & "        " & "FM Messaggio:" & vbCrLf)
            scrivi.Close()
        Catch ex As Exception

        End Try

        For r = 0 To DataGridView1.Rows.Count - 3
            Try
                Dim DefinitiveGDO As String = ""
                Try
                    Dim GDOtoPrint As String = DataGridView1.Rows(r).Cells(5).Value.ToString
                    For z = 0 To 20
                        DefinitiveGDO += GDOtoPrint(z)
                    Next
                Catch ex As Exception

                End Try
                DatagridviewRowToPrint = DataGridView1.Rows(r).Cells(0).Value.ToString & "          " & DefinitiveGDO & "      " & DataGridView1.Rows(r).Cells(6).Value.ToString
                scrivi = System.IO.File.AppendText(Application.StartupPath & "\22TarPrintPage" & TerPrintPageNum & ".txt")
                scrivi.WriteLine(DatagridviewRowToPrint)
                scrivi.Close()
                If r = My.Settings.NumRighePerPagina + stoppami Then
                    stoppami = stoppami + My.Settings.NumRighePerPagina
                    TerPrintPageNum += 1
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\22TarPrintPage" & TerPrintPageNum & ".txt")
                    scrivi.Write("22-Ter del: " & My.Settings.Intelligence22Ter & vbCrLf & "Nr. Progressivo" & "    " & "GDO Messaggio:" & "    " & "FM Messaggio:" & vbCrLf)
                    scrivi.Close()
                End If
                If r = DataGridView1.Rows.Count - 3 Then
                    Start22TerPrint()
                End If
            Catch ex As Exception

            End Try
        Next

    End Sub

    Public Sub Start22TerPrint()

        Try
            For r = 0 To TerPrintPageNum
                leggi = System.IO.File.OpenText(Application.StartupPath & "\22TarPrintPage" & r & ".txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                dastampare = letto
                PrintDocument1.Print()
            Next
        Catch ex As Exception

        End Try

        Try
            For r = 0 To TerPrintPageNum
                System.IO.File.Delete(Application.StartupPath & "\22TarPrintPage" & r & ".txt")
            Next
            stoppami = 0
            TerPrintPageNum = 0
            Me.Close()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub cercaemail()

        Dim cell As DataGridViewCell
        Try
            For Each cell In DataGridView1.SelectedCells
                riga = cell.RowIndex
                datadacercare = DataGridView1.Rows(riga).Cells(3).Value.ToString
                datadefinitiva = Split(datadacercare, " ")
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Email\INTELLIGENCE\" & datadefinitiva(0) & "\emailinviate" & (cell.RowIndex + 1).ToString("D3") & "INTEL.txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                MessageText.TextBox1.Text = letto
                My.Settings.OpenedFilePath = Application.StartupPath & "\Email\INTELLIGENCE\" & datadefinitiva(0) & "\emailinviate" & (cell.RowIndex + 1).ToString("D3") & "INTEL.txt"
                MessageText.Show()
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Sub salvaemail()

        If DataGridView1.SelectedCells.Count > 1 Then
            MsgBox("Devi selezionare un'unica casella, quella contenete il numero progressivo dell'email che vuoi salvare.", MsgBoxStyle.Information, "Asc - Errore Selezione Casella")
        Else
            Dim cell As DataGridViewCell
            Try
                For Each cell In DataGridView1.SelectedCells
                    riga = cell.RowIndex
                    stringNr = DataGridView1.Rows(riga).Cells(0).Value.ToString
                    datadacercare = DataGridView1.Rows(riga).Cells(3).Value.ToString
                    datadefinitiva = Split(datadacercare, " ")
                    leggi = System.IO.File.OpenText(Application.StartupPath & "\Email\INTELLIGENCE\" & datadefinitiva(0) & "\emailinviate" & (cell.RowIndex + 1).ToString("D3") & "INTEL.txt")
                    letto = leggi.ReadToEnd
                    leggi.Close()
                    maildasalvare = letto
                    SaveFileDialog1.Filter = "File Testo | *.txt"
                    SaveFileDialog1.ShowDialog()
                    scrivi = System.IO.File.CreateText(SaveFileDialog1.FileName)
                    scrivi.Write(maildasalvare)
                    scrivi.Close()
                Next
            Catch ex As Exception

            End Try
        End If

    End Sub

    Public Sub stampaemail()

        If DataGridView1.SelectedCells.Count > 1 Then
            MsgBox("Devi selezionare un'unica casella, quella contenete il numero progressivo dell'email che vuoi salvare.", MsgBoxStyle.Information, "Asc - Errore Selezione Casella")
        Else
            Dim cell As DataGridViewCell
            Try
                For Each cell In DataGridView1.SelectedCells
                    riga = cell.RowIndex
                    stringNr = DataGridView1.Rows(riga).Cells(0).Value.ToString
                    datadacercare = DataGridView1.Rows(riga).Cells(3).Value.ToString
                    datadefinitiva = Split(datadacercare, " ")
                    leggi = System.IO.File.OpenText(Application.StartupPath & "\Email\INTELLIGENCE\" & datadefinitiva(0) & "\emailinviate" & (cell.RowIndex + 1).ToString("D3") & "INTEL.txt")
                    letto = leggi.ReadToEnd
                    leggi.Close()
                    maildasalvare = letto
                    righelette = Split(maildasalvare, vbCrLf)
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\fogliostampa" & NumFoglioStampa & ".txt")
                    scrivi.Close()
                    For r = 0 To UBound(righelette)
                        scrivi = System.IO.File.AppendText(Application.StartupPath & "\fogliostampa" & NumFoglioStampa & ".txt")
                        scrivi.Write(righelette(numeroriga) & vbCrLf)
                        scrivi.Close()
                        numeroriga += 1
                        If r = My.Settings.NumRighePerPagina + stoppami Then
                            stoppami = stoppami + My.Settings.NumRighePerPagina
                            NumFoglioStampa += 1
                            scrivi = System.IO.File.CreateText(Application.StartupPath & "\fogliostampa" & NumFoglioStampa & ".txt")
                            scrivi.Close()
                        End If
                    Next
                    Stampa()
                Next
            Catch ex As Exception

            End Try
        End If

    End Sub

    Public Sub Stampa()

        Try
            For r = 0 To NumFoglioStampa
                leggi = System.IO.File.OpenText(Application.StartupPath & "\fogliostampa" & r & ".txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                dastampare = letto
                PrintDocument1.Print()
            Next
        Catch ex As Exception

        End Try

        Try
            For r = 0 To NumFoglioStampa
                System.IO.File.Delete(Application.StartupPath & "\fogliostampa" & r & ".txt")
                NumFoglioStampa = 0
                numeroriga = 0
                stoppami = 0
                Me.TopMost = True
            Next
        Catch ex As Exception

        End Try

        Main.TopMost = False
        Me.TopMost = False

    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        e.Graphics.DrawString(dastampare, Main.TextBox3.Font, Brushes.Black, New System.Drawing.RectangleF(My.Settings.Top, My.Settings.Left, My.Settings.Width, My.Settings.Height))
        e.HasMorePages = False

    End Sub

    Private Sub VisualizzaEmailToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VisualizzaEmailToolStripMenuItem.Click

        cercaemail()

    End Sub

    Private Sub SalvaEmailToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalvaEmailToolStripMenuItem.Click

        salvaemail()

    End Sub

    Private Sub StampaEmailToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StampaEmailToolStripMenuItem.Click

        stampaemail()

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick

        cercaemail()

    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick

        cercaemail()

    End Sub

    Private Sub DataGridView1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseDown

        Try
            GetMousePosition = MousePosition
            If e.Button = Windows.Forms.MouseButtons.Right Then
                Timer1.Start()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Breaker += 1
        If Breaker = 1 Then
            Try
                Dim x As Integer = MousePosition.X
                Dim y As Integer = MousePosition.Y
                mouse_event(MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
                mouse_event(MOUSEEVENTF_LEFTUP, x, y, 0, 0)
            Catch ex As Exception

            End Try
        ElseIf Breaker = 2 Then
            Try
                ContextMenuStrip1.Show(GetMousePosition)
                Breaker = 0
                Timer1.Stop()
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub UtilizzaMessaggioToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UtilizzaMessaggioToolStripMenuItem.Click

        UseMessage()

    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click

        salvaemail()

    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click

        stampaemail()

    End Sub

    Private Sub MenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem10.Click

        Me.Close()

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        DateSelection.Show()

    End Sub

    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click

        INTEL22TerSelectionDate.Show()
        Me.Close()

    End Sub
End Class

