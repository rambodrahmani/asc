Imports System.IO, System.Net, System.Net.Sockets

Public Class Form1

    Dim InvioAnnullato As Integer = 0
    Dim IsFirstChkLastBT As Integer = 0
    Dim IsFisrChkLastBT As Integer = 0
    Dim IsFirstChkBtIsRegular As Integer = 0
    Dim IsFirstSelectBTasLastLine As Integer = 0
    Dim IsFirstChkBTcounter As Integer = 0
    Dim IsFirstChkFollowingSICline As Integer = 0
    Dim IsFirstChkGDOIsRegular As Integer = 0
    Dim IsFirstChkFmIsRegular As Integer = 0
    Dim FinalSectionNr As Integer = 0
    Dim RecreateTxtBox3 As String = ""
    '*** Dichiarazione keybd_event inizio ***
    Declare Sub keybd_event Lib "user32" _
        (ByVal bVk As Byte, ByVal bScan As Byte, _
         ByVal dwFlags As Integer, ByVal dwExtraInfo As Integer)
    '*** Dichiarazione keybd_event fine ***
    Dim scrivi As System.IO.StreamWriter
    Dim dastampare As String = ""
    Dim Client As TcpClient
    Dim sWriter As StreamWriter
    Dim NickPrefix As Integer = New Random().Next(1111, 9999)
    Delegate Sub _xUpdate(ByVal Str As String)
    Dim RigheTextBox() As String
    Dim NumFoglioStampa As Integer = 0
    Dim stoppami As Integer = 0
    Dim numeroriga As Integer = 0
    Dim leggi As System.IO.StreamReader
    Dim letto As String = ""
    Dim DataOggi As String = Date.Now.Day.ToString & "." & Date.Now.Month.ToString & "." & Date.Now.Year.ToString

    Sub Read(ByVal ar As IAsyncResult)

        Try
            xUpdate(New StreamReader(Client.GetStream).ReadLine)
            Client.GetStream.BeginRead(New Byte() {0}, 0, 0, AddressOf Read, Nothing)
        Catch ex As Exception
            Try
                ToolStripButton1.Text = "Connect"
            Catch ex1 As Exception

            End Try
            MsgBox("Errore connessione: " & ex.Message, MsgBoxStyle.Critical, "Client - Disconnesso")
            Exit Sub
        End Try

    End Sub

    Private Sub Send(ByVal Str As String)

        Try
            sWriter = New StreamWriter(Client.GetStream)
            sWriter.WriteLine(Str)
            sWriter.Flush()
        Catch ex As Exception
            InvioAnnullato = 1
            MsgBox("Errore invio: " & ex.Message, MsgBoxStyle.Critical, "Asc - Errore Invio")
            CreateTempFiles()
            Exit Sub
        End Try

    End Sub

    Sub xUpdate(ByVal Str As String)

        Try
            If InvokeRequired Then
                Invoke(New _xUpdate(AddressOf xUpdate), Str)
            Else
                TextBox1.AppendText(Str & vbNewLine)
            End If
        Catch ex As Exception
            MsgBox("Errore: " & ex.Message, MsgBoxStyle.Critical, "Asc - Errore")
        End Try

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click

        Try
            Client = New TcpClient(ToolStripTextBox1.Text, CInt(ToolStripTextBox2.Text))
            Client.GetStream.BeginRead(New Byte() {0}, 0, 0, New AsyncCallback(AddressOf Read), Nothing)
            ToolStripButton1.Text = "Disconnect"
        Catch ex As Exception
            ToolStripButton1.Text = "Connect"
            MsgBox("Impossibile connetersi al server. Codice Errore Rilevato: " & ex.Message, MsgBoxStyle.Critical, "ASC - Connection Error")
        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            OpenFileDialog1.Filter = "Text File|*.txt"
            OpenFileDialog1.FileName = ""
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim leggi As System.IO.StreamReader
                Dim letto As String
                leggi = System.IO.File.OpenText(OpenFileDialog1.FileName)
                letto = leggi.ReadToEnd
                leggi.Close()
                TextBox2.Text = letto
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub CreateTempFiles()

        Try
            Dim scrivi As System.IO.StreamWriter
            If System.IO.File.Exists(Application.StartupPath & "\File Inviati in ASC\" & DataOggi & "\" & TextBox3.Text & ".txt") Then
                For r = 0 To 300
                    If System.IO.File.Exists(Application.StartupPath & "\File Inviati in ASC\" & DataOggi & "\" & TextBox3.Text & r & ".txt") Then
                    Else
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\File Inviati in ASC\" & DataOggi & "\" & TextBox3.Text & r & ".txt")
                        scrivi.Write(TextBox2.Text)
                        scrivi.Close()
                        Exit For
                    End If
                Next
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\File Inviati in ASC\" & DataOggi & "\" & TextBox3.Text & ".txt")
                scrivi.Write(TextBox2.Text)
                scrivi.Close()
            End If
        Catch ex As Exception
            MsgBox("Errore nel salvataggio del file, ex message: " & ex.Message, MsgBoxStyle.Information, "Asc - Salvataggio File")
        End Try

        Try
            Dim scrivi As System.IO.StreamWriter
            If System.IO.File.Exists(Application.StartupPath & "\temp\" & TextBox3.Text & ".txt") Then
                For r = 0 To 300
                    If System.IO.File.Exists(Application.StartupPath & "\temp\" & TextBox3.Text & r & ".txt") Then
                    Else
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\temp\" & TextBox3.Text & r & ".txt")
                        scrivi.Write(TextBox2.Text)
                        scrivi.Close()
                        Exit For
                    End If
                Next
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\temp\" & TextBox3.Text & ".txt")
                scrivi.Write(TextBox2.Text)
                scrivi.Close()
            End If

            MsgBox("Il messaggio è stato inserito nella cartella temp. L'invio verrà eseguito automaticamente appena possibile.", MsgBoxStyle.Information, "Asc - Errore Invio")
        Catch ex As Exception
            MsgBox("Errore nel salvataggio del file temporaneo, ex message: " & ex.Message, MsgBoxStyle.Information, "Asc - Salvataggio File")
        End Try

        TextBox2.Clear()
        TextBox3.Clear()

    End Sub

    Public Sub SendToAsc()

        Try
            Try
                Dim mysplitter() As String
                mysplitter = Split(TextBox2.Text, vbCrLf)
                Dim definitivestring As String = ""
                For r = 0 To UBound(mysplitter)
                    definitivestring += mysplitter(r) & "|"
                Next
                Send("Email da salvare: |" & TextBox3.Text & "|" & definitivestring)
            Catch ex As Exception

            End Try

            If InvioAnnullato = 0 Then
                Try
                    Dim scrivi As System.IO.StreamWriter
                    If System.IO.File.Exists(Application.StartupPath & "\File Inviati in ASC\" & DataOggi & "\" & TextBox3.Text & ".txt") Then
                        For r = 0 To 300
                            If System.IO.File.Exists(Application.StartupPath & "\File Inviati in ASC\" & DataOggi & "\" & TextBox3.Text & r & ".txt") Then
                            Else
                                scrivi = System.IO.File.CreateText(Application.StartupPath & "\File Inviati in ASC\" & DataOggi & "\" & TextBox3.Text & r & ".txt")
                                scrivi.Write(TextBox2.Text)
                                scrivi.Close()
                                Exit For
                            End If
                        Next
                    Else
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\File Inviati in ASC\" & DataOggi & "\" & TextBox3.Text & ".txt")
                        scrivi.Write(TextBox2.Text)
                        scrivi.Close()
                    End If
                Catch ex As Exception
                    MsgBox("Errore nel salvataggio del file, ex message: " & ex.Message, MsgBoxStyle.Information, "Asc - Salvataggio File")
                End Try

                TextBox2.Clear()
                TextBox3.Clear()
                MsgBox("File inviato correttamente al server, recarsi in ASC con la copia firmata del messaggio GIA' PROTOCCOLATO.", MsgBoxStyle.Information, "Asc Client - Trasferimento Files")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ChkCharsErrorsPresence()

        Try
            Dim SelectionPointStart As Integer = 0
            Dim MySplitter() As String = Split(TextBox2.Text, vbCrLf)
            For r = 0 To UBound(MySplitter)
                If MySplitter(r).ToString.Length > 69 Then
                    MsgBox("Riga " & (r + 1).ToString & " non conforme allo standard di compilazione, prego correggere.", MsgBoxStyle.Information, "Asc - Errore Compilazione")
                    TextBox2.SelectionStart = TextBox2.Text.IndexOf(MySplitter(r).ToString)
                    TextBox2.SelectionLength = MySplitter(r).ToString.Length
                    TextBox2.HideSelection = False
                    Exit Sub
                End If
            Next
            SendToAsc()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ChkLastBT()

        Try
            If IsFirstChkLastBT = 0 Then
                IsFirstChkLastBT = 1
                If TextBox2.Lines(TextBox2.Lines.Count - 2).ToString = "BT" Then
                    ChkCharsErrorsPresence()
                Else
                    MsgBox("Errore battitura rilevato, il messaggio deve terminare con la stringa " & """" & "BT" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc - Errore Battitura BT")
                    IsFirstChkLastBT = 1
                    Exit Sub
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ChkBtIsRegular()

        Try
            If IsFirstChkBtIsRegular = 0 Then
                IsFirstChkBtIsRegular = 1
                Dim MyTxtBoxSpliiter() As String = Split(TextBox2.Text, vbCrLf)
                Dim BTcounter As Integer = 0
                Dim ClassificaturaPresente As Integer = 0
                For r = 0 To UBound(MyTxtBoxSpliiter)
                    If MyTxtBoxSpliiter(r).ToString.Contains("BT") Then
                        Dim MySplitter() As String = Split(MyTxtBoxSpliiter(r).ToString & " ", " ")
                        For z = 0 To UBound(MySplitter)
                            If MySplitter(z).ToString = "" Or MySplitter(z).ToString = " " Then
                                If z = MySplitter.Count - 1 Then
                                    If BTcounter = 0 Then
                                        BTcounter += 1
                                        If MyTxtBoxSpliiter(r + 1).ToString.Contains("R I S E R V A T O") Or _
                                            MyTxtBoxSpliiter(r + 1).ToString.Contains("R I S E R V A T I S S I M O") Or _
                                            MyTxtBoxSpliiter(r + 1).ToString.Contains("S E G R E T O") Or _
                                            MyTxtBoxSpliiter(r + 1).ToString.Contains("C O N F I D E N T I A L") Or _
                                            MyTxtBoxSpliiter(r + 1).ToString.Contains("R E S T R I C T E D") Or _
                                            MyTxtBoxSpliiter(r + 1).ToString.Contains("S E C R E T") Or _
                                            MyTxtBoxSpliiter(r + 1).ToString.Contains("N A T O R E S T R I C T E D") Or _
                                            MyTxtBoxSpliiter(r + 1).ToString.Contains("N A T O C O N F I D E N T I A L") Or _
                                            MyTxtBoxSpliiter(r + 1).ToString.Contains("N A T O S E C R E T") Then
                                            ClassificaturaPresente = 1
                                        Else
                                            Exit For
                                        End If
                                    ElseIf BTcounter = 1 Then
                                        If ClassificaturaPresente = 1 Then
                                            If MyTxtBoxSpliiter(r - 1).ToString.Contains("DISTRUGGERE") Or _
                                                MyTxtBoxSpliiter(r - 1).ToString.Contains("DESTROY") Or _
                                                MyTxtBoxSpliiter(r - 1).ToString.Contains("DISTRUZIONE") Then
                                                If MyTxtBoxSpliiter(r - 2).ToString.Contains("PDC:") Or MyTxtBoxSpliiter(r - 2).ToString.Contains("POC:") Then
                                                    BTcounter += 1
                                                    Exit For
                                                Else
                                                    Dim AlertMessage = MsgBox("PDC/POC mancante. Si desidera continaure senza inserire il PDC/POC?", MsgBoxStyle.YesNo, "Server - Data distruzione")
                                                    If AlertMessage = 6 Then
                                                        ChkLastBT()
                                                    Else
                                                        Exit Sub
                                                        IsFirstChkBtIsRegular = 1
                                                    End If
                                                End If
                                            Else
                                                MsgBox("Data distruzione mancante. Prego inserire la Data di Distruzione.", MsgBoxStyle.Information, "Server - Data distruzione")
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If MySplitter(z).ToString <> "BT" Then
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                Next
                ChkLastBT()
            End If
        Catch ex As Exception

        End Try

    End Sub
    Public Sub SelectBTasLastLine()

        Try
            If IsFirstSelectBTasLastLine = 0 Then
                IsFirstSelectBTasLastLine = 1
                Dim BtCounter As Integer = 0
                Dim Textbox3Splitter() As String = Split(TextBox2.Text, vbCrLf)
                TextBox2.Clear()
                For r = 0 To UBound(Textbox3Splitter)
                    If Textbox3Splitter(r).ToString.Contains("BT") Then
                        Dim MySplitter() As String = Split(Textbox3Splitter(r).ToString & " ", " ")
                        For z = 0 To UBound(MySplitter)
                            If MySplitter(z).ToString = "" Or MySplitter(z).ToString = " " Then
                                If z = MySplitter.Count - 1 Then
                                    If BtCounter = 0 Then
                                        TextBox2.Text += Textbox3Splitter(r).ToString & vbCrLf
                                        BtCounter += 1
                                    Else
                                        TextBox2.Text += Textbox3Splitter(r).ToString & vbCrLf
                                        ChkBtIsRegular()
                                        Exit Sub
                                    End If
                                End If
                            Else
                                If MySplitter(z).ToString <> "BT" Then
                                    TextBox2.Text += Textbox3Splitter(r).ToString & vbCrLf
                                    Exit For
                                End If
                            End If
                        Next
                    Else
                        TextBox2.Text += Textbox3Splitter(r).ToString & vbCrLf
                    End If
                Next
            End If
        Catch ex As Exception

        End Try

    End Sub
    Public Sub ChkBTcounter()

        Try
            If IsFirstChkBTcounter = 0 Then
                IsFirstChkBTcounter = 1
                Dim BtCounter As Integer = 0
                Dim Textbox3Splitter() As String = Split(TextBox2.Text, vbCrLf)
                For r = 0 To UBound(Textbox3Splitter)
                    If Textbox3Splitter(r).ToString.Contains("BT") Then
                        Dim MySplitter() As String = Split(Textbox3Splitter(r).ToString & " ", " ")
                        For z = 0 To UBound(MySplitter)
                            If MySplitter(z).ToString = "" Or MySplitter(z).ToString = " " Then
                                If z = MySplitter.Count - 1 Then
                                    BtCounter += 1
                                End If
                            Else
                                If MySplitter(z).ToString <> "BT" Then
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                Next
                If BtCounter = 2 Then
                    SelectBTasLastLine()
                Else
                    MsgBox("Conteggio BT non conforme, numero BT presenti non conforme.", MsgBoxStyle.Information, "Asc - BT Non Conforme")
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ChkFollowingSICline()

        Try
            If IsFirstChkFollowingSICline = 0 Then
                IsFirstChkFollowingSICline = 1
                Dim MySplitter() As String = Split(TextBox2.Text, vbCrLf)
                For r = 0 To UBound(MySplitter)
                    If MySplitter(r).ToString.Contains("SIC") Then
                        If MySplitter(r + 1).ToString.Contains("GIORGIO") Then
                            If MySplitter(r + 1).ToString.Contains("NAVE SAN GIORGIO") Or MySplitter(r + 1).ToString.Contains("ITS SAN GIORGIO") Then
                                ChkBTcounter()
                            Else
                                MsgBox("Errore battitura rilevato, sostituire " & """" & MySplitter(r + 1).ToString & """" & " con " & """" _
                                   & "NAVE SAN GIORGIO" & """" & " oppure con " & """" & "ITS SAN GIORGIO" & """" & ". Invio Annullato.", MsgBoxStyle.Information, "Asc - Errore Battitura")
                                Exit Sub
                            End If
                        Else
                            ChkBTcounter()
                        End If
                    End If
                Next
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ChkGDOIsRegular()

        Try
            If IsFirstChkGDOIsRegular = 0 Then
                IsFirstChkGDOIsRegular = 1
                Dim mysplitter() As String = Split(TextBox2.Text, vbCrLf)
                Dim mysplittertostring As String = mysplitter(0).ToString
                Dim GDOsplitter() As String = Split(mysplittertostring & " ", " ")
                If GDOsplitter(0).ToString & GDOsplitter(1).ToString = "Z O" Or GDOsplitter(0).ToString & GDOsplitter(1).ToString = "Z P" Or _
                    GDOsplitter(0).ToString & GDOsplitter(1).ToString = "Z R" Or GDOsplitter(0).ToString & GDOsplitter(1).ToString = "O P" Or _
                    GDOsplitter(0).ToString & GDOsplitter(1).ToString = "O R" Or GDOsplitter(0).ToString & GDOsplitter(1).ToString = "P R" Then
                    ChkFollowingSICline()
                Else
                    If GDOsplitter(0).ToString = "Z" Or GDOsplitter(0).ToString = "O" Or GDOsplitter(0).ToString = "P" Or GDOsplitter(0).ToString = "R" Then
                        ChkFollowingSICline()
                    Else
                        MsgBox("GDO non comforme. Invio annullato.", MsgBoxStyle.Information, "Asc - GDO Errato")
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ChkFmIsRegular()

        Try
            If IsFirstChkFmIsRegular = 0 Then
                IsFirstChkFmIsRegular = 1
                Dim MySplitter() As String = Split(TextBox2.Text, vbCrLf)
                If MySplitter(1).ToString.Contains("GIORGIO") Then
                    If MySplitter(1).ToString = "FM ITS SAN GIORGIO" Or MySplitter(1).ToString = "FM NAVE SAN GIORGIO" Then
                        ChkGDOIsRegular()
                    Else
                        MsgBox("FM non conforme. Invio annullato.", MsgBoxStyle.Information, "Asc - FM Non Conforme")
                        Exit Sub
                    End If
                Else
                    ChkGDOIsRegular()
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Try
            InvioAnnullato = 0
            IsFirstChkLastBT = 0
            IsFisrChkLastBT = 0
            IsFirstChkBtIsRegular = 0
            IsFirstSelectBTasLastLine = 0
            IsFirstChkBTcounter = 0
            IsFirstChkFollowingSICline = 0
            IsFirstChkGDOIsRegular = 0
            IsFirstChkFmIsRegular = 0
        Catch ex As Exception

        End Try
        Try
            If TextBox3.Text <> "" Then
                TextBox2.Text = TextBox2.Text.Replace("°", "^")
                TextBox2.Text = TextBox2.Text.Replace("%", "pct")
                ChkFmIsRegular()
            Else
                MsgBox("Impossibile inviare messaggio ad ASC, scegliere un nome per il file e riprovare.", MsgBoxStyle.Information, "Asc Client - Nome file mancante")
                Exit Sub
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub LoadTheClientLoadForm()

        If System.IO.Directory.Exists(Application.StartupPath & "\temp") Then
        Else
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\temp")
        End If

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

        Try
            Client = New TcpClient(ToolStripTextBox1.Text, CInt(ToolStripTextBox2.Text))
            Client.GetStream.BeginRead(New Byte() {0}, 0, 0, New AsyncCallback(AddressOf Read), Nothing)
            ToolStripButton1.Text = "Disconnect"
        Catch ex As Exception
            ToolStripButton1.Text = "Connect"
            MsgBox("Impossibile connetersi al server. Codice Errore Rilevato: " & ex.Message, MsgBoxStyle.Critical, "ASC - Connection Error")
        End Try

        AutoSender.Start()

    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Try
            PrintSett.Close()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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
                MsgBox("Impossibile avviare Asc, è già in esecuzione un processo Asc e non è possibile utilizzarne due contemporaneamente. Chiusura.", _
           MsgBoxStyle.Critical, "Asc - Processo già attivo")
                Me.Close()
            Else
                LoadTheClientLoadForm()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

        Try
            If TextBox1.Text.Contains("ServerRequestsConnectedClientList") Then
                TextBox1.Text = ""
                Send("ConnectedClientListAnswearing:|" & My.Computer.Name & "|")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox2.Focus()
        End If

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        If TextBox3.Text <> "" And TextBox2.Text <> "" Then
            Try
                Dim scrivi As System.IO.StreamWriter
                SaveFileDialog1.FileName = TextBox3.Text
                If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    scrivi = System.IO.File.CreateText(SaveFileDialog1.FileName)
                    scrivi.Write(TextBox2.Text)
                    scrivi.Close()
                End If
            Catch ex As Exception
                MsgBox("Errore nel salvataggio del file, ex message: " & ex.Message, MsgBoxStyle.Information, "Asc - Salvataggio File")
            End Try
        Else
            If TextBox3.Text = "" Then
                MsgBox("Devi scegliere un nome per il file da salvare.", MsgBoxStyle.Information, "Asc - Salvataggio File")
            ElseIf TextBox2.Text = "" Then
                MsgBox("Devi scrivere del testo per il messaggio da salvare.", MsgBoxStyle.Information, "Asc - Salvataggio File")
            End If
        End If

    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click

        Me.Close()

    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        e.Graphics.DrawString(dastampare, TextBox2.Font, Brushes.Black, New System.Drawing.RectangleF(My.Settings.Top, My.Settings.Left, My.Settings.Width, My.Settings.Height))
        e.HasMorePages = False

    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click

        If TextBox2.Text <> "" Then
            If PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                If My.Settings.SavePrintSett = 1 Then
                    Stampa()
                Else
                    My.Settings.x = 9
                    PrintSett.Show()
                End If
            End If
        Else
            MsgBox("Testo da stampara mancate, inserire qualcosa da stampare.", MsgBoxStyle.Information, "Asc Client - Errore stampa")
        End If

    End Sub

    Public Sub Stampa()

        Try
            RigheTextBox = Split(TextBox2.Text, vbCrLf)
        Catch ex As Exception

        End Try

        Try
            scrivi = System.IO.File.CreateText(Application.StartupPath & "\fogliostampa" & NumFoglioStampa & ".txt")
            scrivi.Close()
            For r = 0 To UBound(RigheTextBox)
                scrivi = System.IO.File.AppendText(Application.StartupPath & "\fogliostampa" & NumFoglioStampa & ".txt")
                scrivi.Write(RigheTextBox(numeroriga) & vbCrLf)
                scrivi.Close()
                numeroriga += 1
                If r = My.Settings.NumRighePerPagina + stoppami Then
                    stoppami = stoppami + My.Settings.NumRighePerPagina
                    NumFoglioStampa += 1
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\fogliostampa" & NumFoglioStampa & ".txt")
                    scrivi.Close()
                End If
            Next
        Catch ex As Exception

        End Try

        IniziaStampa()

    End Sub

    Public Sub IniziaStampa()

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
            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        If TextBox3.Text <> "" And TextBox2.Text <> "" Then
            Try
                Dim scrivi As System.IO.StreamWriter
                SaveFileDialog1.Filter = "Text File|*.txt"
                SaveFileDialog1.FileName = TextBox3.Text
                If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    scrivi = System.IO.File.CreateText(SaveFileDialog1.FileName)
                    scrivi.Write(TextBox2.Text)
                    scrivi.Close()
                End If
            Catch ex As Exception
                MsgBox("Errore nel salvataggio del file, ex message: " & ex.Message, MsgBoxStyle.Information, "Asc - Salvataggio File")
            End Try
        Else
            If TextBox3.Text = "" Then
                MsgBox("Devi scegliere un nome per il file da salvare.", MsgBoxStyle.Information, "Asc - Salvataggio File")
            ElseIf TextBox2.Text = "" Then
                MsgBox("Devi scrivere del testo per il messaggio da salvare.", MsgBoxStyle.Information, "Asc - Salvataggio File")
            End If
        End If

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        If TextBox2.Text <> "" Then
            If PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                If My.Settings.SavePrintSett = 1 Then
                    Stampa()
                Else
                    My.Settings.x = 9
                    PrintSett.Show()
                End If
            End If
        Else
            MsgBox("Testo da stampare mancate, inserire qualcosa da stampare.", MsgBoxStyle.Information, "Asc Client - Errore stampa")
        End If

    End Sub


    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        TextBox4.Clear()

    End Sub

    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click

        TextBox2.SelectAll()

    End Sub

    Private Sub MenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem8.Click

        TextBox2.Cut()

    End Sub

    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem9.Click

        TextBox2.Copy()

    End Sub

    Private Sub MenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem10.Click

        TextBox2.Paste()

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click

        Try
            OpenFileDialog1.Filter = "Text File|*.txt"
            OpenFileDialog1.FileName = ""
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim leggi As System.IO.StreamReader
                Dim letto As String
                leggi = System.IO.File.OpenText(OpenFileDialog1.FileName)
                letto = leggi.ReadToEnd
                leggi.Close()
                TextBox4.Text = letto
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        Try
            SaveFileDialog1.Filter = "Text File| *.txt"
            SaveFileDialog1.FileName = ""
            If SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                scrivi = System.IO.File.CreateText(SaveFileDialog1.FileName)
                scrivi.Write(TextBox4.Text)
                scrivi.Close()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub Stampa2()

        Try
            RigheTextBox = Split(TextBox4.Text, vbCrLf)
        Catch ex As Exception

        End Try

        Try
            scrivi = System.IO.File.CreateText(Application.StartupPath & "\fogliostampa" & NumFoglioStampa & ".txt")
            scrivi.Close()
            For r = 0 To UBound(RigheTextBox)
                scrivi = System.IO.File.AppendText(Application.StartupPath & "\fogliostampa" & NumFoglioStampa & ".txt")
                scrivi.Write(RigheTextBox(numeroriga) & vbCrLf)
                scrivi.Close()
                numeroriga += 1
                If r = My.Settings.NumRighePerPagina + stoppami Then
                    stoppami = stoppami + My.Settings.NumRighePerPagina
                    NumFoglioStampa += 1
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\fogliostampa" & NumFoglioStampa & ".txt")
                    scrivi.Close()
                End If
            Next
        Catch ex As Exception

        End Try

        IniziaStampa()

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click

        Try
            If TextBox4.Text <> "" Then
                If PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                    If My.Settings.SavePrintSett = 1 Then
                        Stampa2()
                    Else
                        My.Settings.x = 10
                        PrintSett.Show()
                    End If
                End If
            Else
                MsgBox("Testo da stampare mancate, inserire qualcosa da stampare.", MsgBoxStyle.Information, "Asc Client - Errore stampa")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

        Try
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click

        TextBox2.Text = TextBox4.Text
        TextBox4.Clear()
        TextBox3.Focus()

    End Sub

    Private Sub TextBox2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.GotFocus

        If My.Computer.Keyboard.CapsLock = False Then
            keybd_event(Windows.Forms.Keys.Capital, 0, 0, 0)
        End If

    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown

        If My.Computer.Keyboard.CapsLock = False Then
            keybd_event(Windows.Forms.Keys.Capital, 0, 0, 0)
        End If

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

        Try
            If TextBox2.Text.Contains("°") Then
                TextBox2.Text = TextBox2.Text.Replace("°", "^")
                TextBox2.SelectionStart = TextBox2.TextLength
            ElseIf TextBox2.Text.Contains("%") Then
                TextBox2.Text = TextBox2.Text.Replace("%", "pct")
                TextBox2.SelectionStart = TextBox2.TextLength
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.GotFocus

        If My.Computer.Keyboard.CapsLock = False Then
            keybd_event(Windows.Forms.Keys.Capital, 0, 0, 0)
        End If

    End Sub

    Private Sub TextBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyDown

        If My.Computer.Keyboard.CapsLock = False Then
            keybd_event(Windows.Forms.Keys.Capital, 0, 0, 0)
        End If

    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged

        Try
            If TextBox4.Text.Contains("°") Then
                TextBox4.Text = TextBox4.Text.Replace("°", "^")
                TextBox4.SelectionStart = TextBox4.TextLength
            ElseIf TextBox4.Text.Contains("%") Then
                TextBox4.Text = TextBox4.Text.Replace("%", "pct")
                TextBox4.SelectionStart = TextBox4.TextLength
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub AutoSender_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoSender.Tick

        AutoSendProcess()

    End Sub

    Public Sub AutoSendProcess()

        If ToolStripButton1.Text = "Connect" Then
            Try
                Client = New TcpClient(ToolStripTextBox1.Text, CInt(ToolStripTextBox2.Text))
                Client.GetStream.BeginRead(New Byte() {0}, 0, 0, New AsyncCallback(AddressOf Read), Nothing)
                ToolStripButton1.Text = "Disconnect"
            Catch ex As Exception
                ToolStripButton1.Text = "Connect"
                Exit Sub
            End Try
        End If

        Try
            Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\temp")
            If FilesCollection.Count > 0 Then
                AutoSender.Interval = "5000"
                For r = 0 To UBound(FilesCollection)
                    Dim FileName As New FileInfo(FilesCollection(r).ToString)
                    Dim NameSplitter() As String = Split(FileName.Name.ToString, ".txt")
                    leggi = System.IO.File.OpenText(FilesCollection(r).ToString)
                    letto = leggi.ReadToEnd
                    leggi.Close()
                    Dim mysplitter() As String
                    mysplitter = Split(letto, vbCrLf)
                    Dim definitivestring As String = ""
                    For z = 0 To UBound(mysplitter)
                        definitivestring += mysplitter(z).ToString & "|"
                    Next
                    AutoSend("Email da salvare: |" & NameSplitter(0).ToString & "|" & definitivestring)
                    System.IO.File.Delete(FilesCollection(r).ToString)
                    Exit Sub
                Next
            Else
                AutoSender.Interval = "15000"
                Exit Sub
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub AutoSend(ByVal Str As String)

        Try
            sWriter = New StreamWriter(Client.GetStream)
            sWriter.WriteLine(Str)
            sWriter.Flush()
        Catch ex As Exception

        End Try

    End Sub

End Class
