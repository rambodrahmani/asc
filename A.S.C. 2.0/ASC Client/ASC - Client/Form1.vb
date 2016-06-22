Imports System.IO, System.Net, System.Net.Sockets, System.IO.Directory
Imports ASC___Client.tcpConnection

Public Class Form1

    Dim InvioAnnullato As Integer = 0
    Dim IsFirstChkLastBT As Integer = 0
    Dim IsFirstControlloPOC As Integer = 0
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
    Dim RigheTextBox() As String
    Dim NumFoglioStampa As Integer = 0
    Dim stoppami As Integer = 0
    Dim numeroriga As Integer = 0
    Dim leggi As System.IO.StreamReader
    Dim letto As String = ""
    Dim DataOggi As String = Date.Now.Day.ToString & "." & Date.Now.Month.ToString & "." & Date.Now.Year.ToString
    Dim FilesNameContainer As String = ""
    Dim MSGIDdef As String
    Dim MSGID() As String
    Dim MeOldWidth As Double = 0
    Dim MeOldHeihgt As Double = 0
    Dim Resizing As Boolean = False
    Dim SelectedItem As String = ""
    Dim MessaggioRicevuto As String = ""
    Dim AutomaticOpen As Boolean = False

#Region "DichiarazioneClient"
    Private WithEvents client As tcpConnection
    Enum Requests As Byte
        DataFile = 1
        StringMessage = 2
    End Enum
#End Region

#Region "Eventi Client"

    Public Sub client_Connect(ByVal sender As tcpConnection) Handles client.Connect

        'specificare cosa avviene quando il client si connette

    End Sub

    Public Sub client_Disconnect(ByVal sender As tcpConnection) Handles client.Disconnect

        'specificare cosa avviene quando il client si disconnette

    End Sub

    Private Sub client_StringReceived(ByVal Sender As tcpConnection, ByVal msgTag As Byte, ByVal message As String) Handles client.StringReceived

        'This is where the client will send us requests for file data using our 
        ' predefined message tags
        Debug.Print("String Received from Client: " & message)
        Select Case msgTag
            Case Requests.DataFile
                Sender.SendFile(msgTag, "Path per salvataggio")
            Case Requests.StringMessage
                MessaggioRicevuto = message
        End Select

    End Sub

    Private Sub client_DataReceived(ByVal Sender As tcpConnection, ByVal msgTag As Byte, _
        ByVal mstream As System.IO.MemoryStream) Handles client.DataReceived
        'This code is run in a seperate thread from the thread that started the form
        'so we must either handle any control access in a special thread-safe way
        'or ignore illegal cross thread calls
        Select Case msgTag
            Case Requests.DataFile
                'file data, save to a local file
                If System.IO.Directory.Exists(Application.StartupPath & "\temp") Then
                    If System.IO.File.Exists(Application.StartupPath & "\temp\temp.txt") Then
                        For r = 1 To 9999
                            If System.IO.File.Exists(Application.StartupPath & "\temp\temp(" & r.ToString & ").txt") Then
                            Else
                                SaveFile(Application.StartupPath & "\temp\temp(" & r.ToString & ").txt", mstream)
                                Exit Sub
                            End If
                        Next
                    Else
                        SaveFile(Application.StartupPath & "\temp\temp.txt", mstream)
                    End If
                Else
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\temp")
                    If System.IO.File.Exists(Application.StartupPath & "\temp\temp.txt") Then
                        For r = 1 To 9999
                            If System.IO.File.Exists(Application.StartupPath & "\temp\temp(" & r.ToString & ").txt") Then
                            Else
                                SaveFile(Application.StartupPath & "\temp\temp(" & r.ToString & ").txt", mstream)
                                Exit Sub
                            End If
                        Next
                    Else
                        SaveFile(Application.StartupPath & "\temp\temp.txt", mstream)
                    End If
                End If
        End Select

    End Sub

    Private Sub SaveFile(ByVal FilePath As String, ByVal mstream As System.IO.MemoryStream)

        'save file to path specified
        Try
            Dim FS As New FileStream(FilePath, IO.FileMode.Create, IO.FileAccess.Write)
            mstream.WriteTo(FS)
            mstream.Flush()
            mstream.Close()
            FS.Close()
        Catch ex As Exception

        End Try

    End Sub

#End Region

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        ConnectClient()

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

    Public Sub SendToAsc()

        Try
            Try
                Dim scrittore As System.IO.StreamWriter
                scrittore = System.IO.File.CreateText(Application.StartupPath & "\temp.txt")
                scrittore.Write("MessaggioDalClient" & vbCrLf)
                scrittore.Write(TextBox3.Text & vbCrLf)
                scrittore.Write(TextBox2.Text)
                scrittore.Close()
                client.SendFile(Requests.DataFile, Application.StartupPath & "\temp.txt")
                System.IO.File.Delete(Application.StartupPath & "\temp.txt")
            Catch ex As Exception
                MsgBox("Errore durante l'invio dell'email al server, messaggio ex: " & ex.Message, MsgBoxStyle.Critical, "Asc Server - Errore Invio")
                InvioAnnullato = 1
                Exit Sub
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
                    MsgBox("Errore nel salvataggio del file, messaggio errore " & ex.Message, MsgBoxStyle.Information, "Asc Client - Salvataggio File")
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
                    MsgBox("Riga " & (r + 1).ToString & " non conforme allo standard di compilazione, prego correggere.", MsgBoxStyle.Information, "Asc Client - Errore Compilazione")
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
                    MsgBox("Errore battitura rilevato, il messaggio deve terminare con la stringa " & """" & "BT" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Errore Battitura BT")
                    IsFirstChkLastBT = 1
                    Exit Sub
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ControlloPOC()

        Try
            If IsFirstControlloPOC = 0 Then
                IsFirstControlloPOC = 1
                If TextBox2.Text.Contains("PDC:") Or TextBox2.Text.Contains("POC:") Then
                    ChkLastBT()
                Else
                    Dim MySplitter2() As String = Split(TextBox2.Text, vbCrLf)
                    If MySplitter2(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                        Or MySplitter2(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                        Or MySplitter2(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                        Dim AlertMessage = MsgBox("PDC/POC mancante. Si desidera continaure senza inserire il PDC/POC?", MsgBoxStyle.YesNo, "Server - Data distruzione")
                        If AlertMessage = 6 Then
                            ChkLastBT()
                        Else
                            IsFirstControlloPOC = 1
                            Exit Sub
                        End If
                    Else
                        ChkLastBT()
                    End If
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
                                            MyTxtBoxSpliiter(r + 1).ToString.Contains("T O P S E C R E T") Or _
                                            MyTxtBoxSpliiter(r + 1).ToString.Contains("N A T O T O P S E C R E T") Or _
                                            MyTxtBoxSpliiter(r + 1).ToString.Contains("N A T O R E S T R I C T E D") Or _
                                            MyTxtBoxSpliiter(r + 1).ToString.Contains("N A T O C O N F I D E N T I A L") Or _
                                            MyTxtBoxSpliiter(r + 1).ToString.Contains("N A T O S E C R E T") Then
                                            ClassificaturaPresente = 1
                                            If MyTxtBoxSpliiter(r + 2).ToString.Contains("SIC") Then
                                            Else
                                                MsgBox("Non è stata rilevata alcuna riga SIC, impossibile continuare.", MsgBoxStyle.Critical, "Asc Client - Sic Mancante")
                                                Exit Sub
                                            End If
                                            If MyTxtBoxSpliiter(r + 3).ToString.Contains("NAVE " & My.Settings.NomeUnita) _
                                                Or MyTxtBoxSpliiter(r + 3).ToString.Contains("SMG " & My.Settings.NomeUnita) _
                                                Or MyTxtBoxSpliiter(r + 3).ToString.Contains("ITS " & My.Settings.NomeUnita) Then
                                            Else
                                                MsgBox("La linea del protocollo è sbagliata. Invio annullato.", MsgBoxStyle.Critical, "Asc Client - Sic Mancante")
                                                Exit Sub
                                            End If
                                        ElseIf MyTxtBoxSpliiter(r + 1).ToString.Contains("UNCLASSIFIED") Or _
                                        MyTxtBoxSpliiter(r + 1).ToString.Contains("UNCLASS") Or _
                                        MyTxtBoxSpliiter(r + 1).ToString.Contains("NATO UNCLASS") Or _
                                        MyTxtBoxSpliiter(r + 1).ToString.Contains("UNCLAS") Or _
                                        MyTxtBoxSpliiter(r + 1).ToString.Contains("NATO UNCLAS") Or _
                                        MyTxtBoxSpliiter(r + 1).ToString.Contains("NATO UNCLASSIFIED") Or _
                                        MyTxtBoxSpliiter(r + 1).ToString.Contains("NON CLASSIFICATO") Then
                                            If MyTxtBoxSpliiter(r + 3).ToString.Contains("NAVE " & My.Settings.NomeUnita) _
                                                Or MyTxtBoxSpliiter(r + 3).ToString.Contains("SMG " & My.Settings.NomeUnita) _
                                                Or MyTxtBoxSpliiter(r + 3).ToString.Contains("ITS " & My.Settings.NomeUnita) Then
                                            Else
                                                MsgBox("La linea del protocollo è sbagliata. Invio annullato.", MsgBoxStyle.Critical, "Asc Client - Sic Mancante")
                                                Exit Sub
                                            End If
                                            If MyTxtBoxSpliiter(r + 2).ToString.Contains("SIC") Then
                                                ControlloPOC()
                                                Exit Sub
                                            Else
                                                MsgBox("Non è stata rilevata alcuna riga SIC, impossibile continuare.", MsgBoxStyle.Critical, "Asc Client - Sic Mancante")
                                                Exit Sub
                                            End If
                                        Else
                                            MsgBox("Classifica di segretezza mancante. Impossibile continuare.", MsgBoxStyle.Critical, "Asc Client - Classifica Segretezza Mancante")
                                            Exit Sub
                                        End If
                                    ElseIf BTcounter = 1 Then
                                        If ClassificaturaPresente = 1 Then
                                            If MyTxtBoxSpliiter(r - 1).ToString.Contains("DISTRUGGERE") Or _
                                                MyTxtBoxSpliiter(r - 1).ToString.Contains("DESTROY") Or _
                                                MyTxtBoxSpliiter(r - 1).ToString.Contains("DISTRUZIONE") Then
                                                ControlloPOC()
                                            Else
                                                MsgBox("Data distruzione mancante. Prego inserire la Data di Distruzione.", MsgBoxStyle.Critical, "Server - Data distruzione")
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
                ControlloPOC()
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
                    MsgBox("Conteggio BT non conforme, numero BT presenti non conforme.", MsgBoxStyle.Critical, "Asc Client - BT Non Conforme")
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
                        If MySplitter(r + 1).ToString.Contains(My.Settings.NomeUnita) Then
                            If MySplitter(r + 1).ToString.Contains("NAVE " & My.Settings.NomeUnita) _
                                Or MySplitter(r + 1).ToString.Contains("ITS " & My.Settings.NomeUnita) _
                                Or MySplitter(r + 1).ToString.Contains("SMG " & My.Settings.NomeUnita) Then
                                ChkBTcounter()
                                Exit For
                            Else
                                MsgBox("Errore battitura rilevato, sostituire " & """" & MySplitter(r + 1).ToString & """" & " con " & """" _
                                   & "NAVE/SMG " & My.Settings.NomeUnita & """" & " oppure con " & """" & "ITS " & My.Settings.NomeUnita & """" & ". Invio Annullato.", MsgBoxStyle.Information, "Asc Server - Errore Battitura")
                                Exit Sub
                            End If
                        Else
                            Dim RigaTrovata As String = MySplitter(r + 1).ToString
                            Dim NomeUnitaConfronto As String = My.Settings.NomeUnita
                            Dim Counter As Integer = 0
                            If RigaTrovata.Contains("ITS ") Then
                                For z = 0 To RigaTrovata.Length - 5
                                    If z < RigaTrovata.Length And z < NomeUnitaConfronto.Length Then
                                        If RigaTrovata(z + 4) = NomeUnitaConfronto(z) Then
                                            Counter += 1
                                        End If
                                    End If
                                Next
                            ElseIf RigaTrovata.Contains("NAVE ") Then
                                For z = 0 To RigaTrovata.Length - 6
                                    If z < RigaTrovata.Length And z < NomeUnitaConfronto.Length Then
                                        If RigaTrovata(z + 5) = NomeUnitaConfronto(z) Then
                                            Counter += 1
                                        End If
                                    End If
                                Next
                            ElseIf RigaTrovata.Contains("SMG ") Then
                                For z = 0 To RigaTrovata.Length - 5
                                    If z < RigaTrovata.Length And z < NomeUnitaConfronto.Length Then
                                        If RigaTrovata(z + 4) = NomeUnitaConfronto(z) Then
                                            Counter += 1
                                        End If
                                    End If
                                Next
                            End If
                            If Counter >= 4 Then
                                Dim AlertMessage = MsgBox("Durante il controllo ortografico è stata rilevata una possibilità di errore nella riga protocollo, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Server - Possibile Errore")
                                If AlertMessage = 6 Then
                                    Exit For
                                Else
                                    Exit Sub
                                End If
                            Else
                                Exit For
                            End If
                        End If
                    Else
                        If r = MySplitter.Count - 1 Then
                            MsgBox("Non è stata rilevata alcuna riga SIC, impossibile continuare.", MsgBoxStyle.Information, "Asc Server - Sic Mancante")
                            Exit Sub
                        End If
                    End If
                Next
                ChkBTcounter()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Sub ChkGDOIsRegular()

        Try
            If IsFirstChkGDOIsRegular = 0 Then
                IsFirstChkGDOIsRegular = 1
                Dim mysplitter() As String = Split(TextBox2.Text, vbCrLf)
                Dim mysplittertostring As String = mysplitter(0).ToString
                If mysplittertostring.Length > 3 Then
                    Dim GDOtrelettere As String = mysplittertostring(0) & mysplittertostring(1) & mysplittertostring(2)
                    If GDOtrelettere = "Z O" Or GDOtrelettere = "Z P" Or GDOtrelettere = "Z R" Or GDOtrelettere = "O P" Or GDOtrelettere = "O R" _
                        Or GDOtrelettere = "P R" Then
                        Dim MyGDOsplitter() As String = Split(mysplittertostring, " ")
                        If MyGDOsplitter(0).ToString.Length + MyGDOsplitter(1).ToString.Length = 2 And MyGDOsplitter(2).ToString.Length = 7 _
                            And MyGDOsplitter(3).ToString.Length = 3 And MyGDOsplitter(4).ToString.Length = 2 Then
                            If mysplittertostring.Contains("ZDS") Or mysplittertostring.Contains("ZDK") Or mysplittertostring.Contains("ZFG") Or mysplittertostring.Contains("ZWL") _
                                        Or mysplittertostring.Contains("ZFD") Or mysplittertostring.Contains("ZWN") Then
                                ChkFollowingSICline()
                                Exit Sub
                            End If
                            If Val(mysplittertostring(4).ToString & mysplittertostring(5).ToString) > Val(Date.Now.Day) Then
                                Dim AlertMessage = MsgBox("Data GDO non conforme: giorno superiore al corrente, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Client - Data GDO Errata")
                                If AlertMessage = 6 Then
                                    If Date.Now.Month = 1 Then
                                        If MyGDOsplitter(3).ToString = "GEN" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "GEN" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 2 Then
                                        If MyGDOsplitter(3).ToString = "FEB" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "FEB" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 3 Then
                                        If MyGDOsplitter(3).ToString = "MAR" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "MAR" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 4 Then
                                        If MyGDOsplitter(3).ToString = "APR" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "APR" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 5 Then
                                        If MyGDOsplitter(3).ToString = "MAG" Or MyGDOsplitter(3).ToString = "MAY" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "MAG" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 6 Then
                                        If MyGDOsplitter(3).ToString = "GIU" Or MyGDOsplitter(3).ToString = "JUN" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "GIU" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 7 Then
                                        If MyGDOsplitter(3).ToString = "LUG" Or MyGDOsplitter(3).ToString = "JUL" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "LUG" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 8 Then
                                        If MyGDOsplitter(3).ToString = "AGO" Or MyGDOsplitter(3).ToString = "AUG" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "AGO" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 9 Then
                                        If MyGDOsplitter(3).ToString = "SET" Or MyGDOsplitter(3).ToString = "SEP" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "SET" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 10 Then
                                        If MyGDOsplitter(3).ToString = "OTT" Or MyGDOsplitter(3).ToString = "OCT" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "OTT" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 11 Then
                                        If MyGDOsplitter(3).ToString = "NOV" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "NOV" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 12 Then
                                        If MyGDOsplitter(3).ToString = "DIC" Or MyGDOsplitter(3).ToString = "DEC" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "DIC" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    End If
                                Else
                                    Exit Sub
                                End If
                            ElseIf Val(mysplittertostring(4).ToString & mysplittertostring(5).ToString) < Val(Date.Now.Day) Then
                                If mysplittertostring.Contains("ZDS") Or mysplittertostring.Contains("ZDK") Or mysplittertostring.Contains("ZFG") Or mysplittertostring.Contains("ZWL") _
                                        Or mysplittertostring.Contains("ZFD") Or mysplittertostring.Contains("ZWN") Then
                                    ChkFollowingSICline()
                                    Exit Sub
                                End If
                                Dim AlertMessage = MsgBox("Data GDO non conforme: giorno inferiore al corrente, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Client - Data GDO Errata")
                                If AlertMessage = 6 Then
                                    If Date.Now.Month = 1 Then
                                        If MyGDOsplitter(3).ToString = "GEN" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "GEN" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 2 Then
                                        If MyGDOsplitter(3).ToString = "FEB" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "FEB" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 3 Then
                                        If MyGDOsplitter(3).ToString = "MAR" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "MAR" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 4 Then
                                        If MyGDOsplitter(3).ToString = "APR" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "APR" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 5 Then
                                        If MyGDOsplitter(3).ToString = "MAG" Or MyGDOsplitter(3).ToString = "MAY" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "MAG" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 6 Then
                                        If MyGDOsplitter(3).ToString = "GIU" Or MyGDOsplitter(3).ToString = "JUN" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "GIU" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 7 Then
                                        If MyGDOsplitter(3).ToString = "LUG" Or MyGDOsplitter(3).ToString = "JUL" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "LUG" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 8 Then
                                        If MyGDOsplitter(3).ToString = "AGO" Or MyGDOsplitter(3).ToString = "AUG" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "AGO" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 9 Then
                                        If MyGDOsplitter(3).ToString = "SET" Or MyGDOsplitter(3).ToString = "SEP" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "SET" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 10 Then
                                        If MyGDOsplitter(3).ToString = "OTT" Or MyGDOsplitter(3).ToString = "OCT" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "OTT" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 11 Then
                                        If MyGDOsplitter(3).ToString = "NOV" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "NOV" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 12 Then
                                        If MyGDOsplitter(3).ToString = "DIC" Or MyGDOsplitter(3).ToString = "DEC" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "DIC" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    End If
                                Else
                                    Exit Sub
                                End If
                            Else
                                If mysplittertostring.Contains("ZDS") Or mysplittertostring.Contains("ZDK") Or mysplittertostring.Contains("ZFG") Or mysplittertostring.Contains("ZWL") _
                                        Or mysplittertostring.Contains("ZFD") Or mysplittertostring.Contains("ZWN") Then
                                    ChkFollowingSICline()
                                    Exit Sub
                                End If
                                If Date.Now.Month = 1 Then
                                    If MyGDOsplitter(3).ToString = "GEN" Then
                                        ChkFollowingSICline()
                                    Else
                                        MsgBox("Mese del GDO errato, il mese corretto è " & """" & "GEN" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                        Exit Sub
                                    End If
                                ElseIf Date.Now.Month = 2 Then
                                    If MyGDOsplitter(3).ToString = "FEB" Then
                                        ChkFollowingSICline()
                                    Else
                                        MsgBox("Mese del GDO errato, il mese corretto è " & """" & "FEB" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                        Exit Sub
                                    End If
                                ElseIf Date.Now.Month = 3 Then
                                    If MyGDOsplitter(3).ToString = "MAR" Then
                                        ChkFollowingSICline()
                                    Else
                                        MsgBox("Mese del GDO errato, il mese corretto è " & """" & "MAR" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                        Exit Sub
                                    End If
                                ElseIf Date.Now.Month = 4 Then
                                    If MyGDOsplitter(3).ToString = "APR" Then
                                        ChkFollowingSICline()
                                    Else
                                        MsgBox("Mese del GDO errato, il mese corretto è " & """" & "APR" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                        Exit Sub
                                    End If
                                ElseIf Date.Now.Month = 5 Then
                                    If MyGDOsplitter(3).ToString = "MAG" Or MyGDOsplitter(3).ToString = "MAY" Then
                                        ChkFollowingSICline()
                                    Else
                                        MsgBox("Mese del GDO errato, il mese corretto è " & """" & "MAG" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                        Exit Sub
                                    End If
                                ElseIf Date.Now.Month = 6 Then
                                    If MyGDOsplitter(3).ToString = "GIU" Or MyGDOsplitter(3).ToString = "JUN" Then
                                        ChkFollowingSICline()
                                    Else
                                        MsgBox("Mese del GDO errato, il mese corretto è " & """" & "GIU" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                        Exit Sub
                                    End If
                                ElseIf Date.Now.Month = 7 Then
                                    If MyGDOsplitter(3).ToString = "LUG" Or MyGDOsplitter(3).ToString = "JUL" Then
                                        ChkFollowingSICline()
                                    Else
                                        MsgBox("Mese del GDO errato, il mese corretto è " & """" & "LUG" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                        Exit Sub
                                    End If
                                ElseIf Date.Now.Month = 8 Then
                                    If MyGDOsplitter(3).ToString = "AGO" Or MyGDOsplitter(3).ToString = "AUG" Then
                                        ChkFollowingSICline()
                                    Else
                                        MsgBox("Mese del GDO errato, il mese corretto è " & """" & "AGO" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                        Exit Sub
                                    End If
                                ElseIf Date.Now.Month = 9 Then
                                    If MyGDOsplitter(3).ToString = "SET" Or MyGDOsplitter(3).ToString = "SEP" Then
                                        ChkFollowingSICline()
                                    Else
                                        MsgBox("Mese del GDO errato, il mese corretto è " & """" & "SET" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                        Exit Sub
                                    End If
                                ElseIf Date.Now.Month = 10 Then
                                    If MyGDOsplitter(3).ToString = "OTT" Or MyGDOsplitter(3).ToString = "OCT" Then
                                        ChkFollowingSICline()
                                    Else
                                        MsgBox("Mese del GDO errato, il mese corretto è " & """" & "OTT" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                        Exit Sub
                                    End If
                                ElseIf Date.Now.Month = 11 Then
                                    If MyGDOsplitter(3).ToString = "NOV" Then
                                        ChkFollowingSICline()
                                    Else
                                        MsgBox("Mese del GDO errato, il mese corretto è " & """" & "NOV" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                        Exit Sub
                                    End If
                                ElseIf Date.Now.Month = 12 Then
                                    If MyGDOsplitter(3).ToString = "DIC" Or MyGDOsplitter(3).ToString = "DEC" Then
                                        ChkFollowingSICline()
                                    Else
                                        MsgBox("Mese del GDO errato, il mese corretto è " & """" & "DIC" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Else
                            MsgBox("GDO non conforme. Invio annullato.", MsgBoxStyle.Information, "Asc Client - GDO Errato")
                            Exit Sub
                        End If
                    Else
                        If mysplittertostring(0) = "R" Or mysplittertostring(0) = "P" Or mysplittertostring(0) = "O" Or mysplittertostring(0) = "Z" Then
                            Dim MyGDOsplitter() As String = Split(mysplittertostring, " ")
                            If MyGDOsplitter(0).ToString.Length = 1 And MyGDOsplitter(1).ToString.Length = 7 _
                                And MyGDOsplitter(2).ToString.Length = 3 And MyGDOsplitter(3).ToString.Length = 2 Then
                                If mysplittertostring.Contains("ZDS") Or mysplittertostring.Contains("ZDK") Or mysplittertostring.Contains("ZFG") Or mysplittertostring.Contains("ZWL") _
                                        Or mysplittertostring.Contains("ZFD") Or mysplittertostring.Contains("ZWN") Then
                                    ChkFollowingSICline()
                                    Exit Sub
                                End If
                                If Val(mysplittertostring(2).ToString & mysplittertostring(3).ToString) > Val(Date.Now.Day) Then
                                    Dim AlertMessage = MsgBox("Data GDO non conforme: giorno superiore al corrente, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Client - Data GDO Errata")
                                    If AlertMessage = 6 Then
                                        If Date.Now.Month = 1 Then
                                            If MyGDOsplitter(2).ToString = "GEN" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "GEN" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 2 Then
                                            If MyGDOsplitter(2).ToString = "FEB" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "FEB" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 3 Then
                                            If MyGDOsplitter(2).ToString = "MAR" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "MAR" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 4 Then
                                            If MyGDOsplitter(2).ToString = "APR" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "APR" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 5 Then
                                            If MyGDOsplitter(2).ToString = "MAG" Or MyGDOsplitter(2).ToString = "MAY" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "MAG" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 6 Then
                                            If MyGDOsplitter(2).ToString = "GIU" Or MyGDOsplitter(2).ToString = "JUN" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "GIU" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 7 Then
                                            If MyGDOsplitter(2).ToString = "LUG" Or MyGDOsplitter(2).ToString = "JUL" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "LUG" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 8 Then
                                            If MyGDOsplitter(2).ToString = "AGO" Or MyGDOsplitter(2).ToString = "AUG" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "AGO" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 9 Then
                                            If MyGDOsplitter(2).ToString = "SET" Or MyGDOsplitter(2).ToString = "SEP" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "SET" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 10 Then
                                            If MyGDOsplitter(2).ToString = "OTT" Or MyGDOsplitter(2).ToString = "OCT" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "OTT" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 11 Then
                                            If MyGDOsplitter(2).ToString = "NOV" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "NOV" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 12 Then
                                            If MyGDOsplitter(2).ToString = "DIC" Or MyGDOsplitter(2).ToString = "DEC" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "DIC" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        Exit Sub
                                    End If
                                ElseIf Val(mysplittertostring(2).ToString & mysplittertostring(3).ToString) < Val(Date.Now.Day) Then
                                    If mysplittertostring.Contains("ZDS") Or mysplittertostring.Contains("ZDK") Or mysplittertostring.Contains("ZFG") Or mysplittertostring.Contains("ZWL") _
                                        Or mysplittertostring.Contains("ZFD") Or mysplittertostring.Contains("ZWN") Then
                                        ChkFollowingSICline()
                                        Exit Sub
                                    End If
                                    Dim AlertMessage = MsgBox("Data GDO non conforme: giorno inferiore al corrente, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Client - Data GDO Errata")
                                    If AlertMessage = 6 Then
                                        If Date.Now.Month = 1 Then
                                            If MyGDOsplitter(2).ToString = "GEN" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "GEN" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 2 Then
                                            If MyGDOsplitter(2).ToString = "FEB" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "FEB" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 3 Then
                                            If MyGDOsplitter(2).ToString = "MAR" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "MAR" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 4 Then
                                            If MyGDOsplitter(2).ToString = "APR" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "APR" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 5 Then
                                            If MyGDOsplitter(2).ToString = "MAG" Or MyGDOsplitter(2).ToString = "MAY" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "MAG" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 6 Then
                                            If MyGDOsplitter(2).ToString = "GIU" Or MyGDOsplitter(2).ToString = "JUN" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "GIU" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 7 Then
                                            If MyGDOsplitter(2).ToString = "LUG" Or MyGDOsplitter(2).ToString = "JUL" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "LUG" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 8 Then
                                            If MyGDOsplitter(2).ToString = "AGO" Or MyGDOsplitter(2).ToString = "AUG" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "AGO" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 9 Then
                                            If MyGDOsplitter(2).ToString = "SET" Or MyGDOsplitter(2).ToString = "SEP" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "SET" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 10 Then
                                            If MyGDOsplitter(2).ToString = "OTT" Or MyGDOsplitter(2).ToString = "OCT" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "OTT" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 11 Then
                                            If MyGDOsplitter(2).ToString = "NOV" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "NOV" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        ElseIf Date.Now.Month = 12 Then
                                            If MyGDOsplitter(2).ToString = "DIC" Or MyGDOsplitter(2).ToString = "DEC" Then
                                                ChkFollowingSICline()
                                            Else
                                                MsgBox("Mese del GDO errato, il mese corretto è " & """" & "DIC" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                        Exit Sub
                                    End If
                                Else
                                    If mysplittertostring.Contains("ZDS") Or mysplittertostring.Contains("ZDK") Or mysplittertostring.Contains("ZFG") Or mysplittertostring.Contains("ZWL") _
                                        Or mysplittertostring.Contains("ZFD") Or mysplittertostring.Contains("ZWN") Then
                                        ChkFollowingSICline()
                                        Exit Sub
                                    End If
                                    If Date.Now.Month = 1 Then
                                        If MyGDOsplitter(2).ToString = "GEN" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "GEN" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 2 Then
                                        If MyGDOsplitter(2).ToString = "FEB" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "FEB" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 3 Then
                                        If MyGDOsplitter(2).ToString = "MAR" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "MAR" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 4 Then
                                        If MyGDOsplitter(2).ToString = "APR" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "APR" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 5 Then
                                        If MyGDOsplitter(2).ToString = "MAG" Or MyGDOsplitter(2).ToString = "MAY" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "MAG" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 6 Then
                                        If MyGDOsplitter(2).ToString = "GIU" Or MyGDOsplitter(2).ToString = "JUN" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "GIU" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 7 Then
                                        If MyGDOsplitter(2).ToString = "LUG" Or MyGDOsplitter(2).ToString = "JUL" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "LUG" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 8 Then
                                        If MyGDOsplitter(2).ToString = "AGO" Or MyGDOsplitter(2).ToString = "AUG" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "AGO" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 9 Then
                                        If MyGDOsplitter(2).ToString = "SET" Or MyGDOsplitter(2).ToString = "SEP" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "SET" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 10 Then
                                        If MyGDOsplitter(2).ToString = "OTT" Or MyGDOsplitter(2).ToString = "OCT" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "OTT" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 11 Then
                                        If MyGDOsplitter(2).ToString = "NOV" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "NOV" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    ElseIf Date.Now.Month = 12 Then
                                        If MyGDOsplitter(2).ToString = "DIC" Or MyGDOsplitter(2).ToString = "DEC" Then
                                            ChkFollowingSICline()
                                        Else
                                            MsgBox("Mese del GDO errato, il mese corretto è " & """" & "DIC" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Client - Mese GDO Errato")
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Else
                                MsgBox("GDO non conforme. Invio annullato.", MsgBoxStyle.Information, "Asc Client - GDO Errato")
                                Exit Sub
                            End If
                        Else
                            MsgBox("Qualifica di precedenza GDO errata. Invio annullato.", MsgBoxStyle.Information, "Asc Client - Qualifica Errata")
                            Exit Sub
                        End If
                    End If
                Else
                    If mysplittertostring(0) = "R" Or mysplittertostring(0) = "P" Or mysplittertostring(0) = "O" Or mysplittertostring(0) = "Z" Then
                        ChkFollowingSICline()
                    Else
                        MsgBox("Qualifica di precedenza GDO errata. Invio annullato.", MsgBoxStyle.Information, "Asc Client - Qualifica Errata")
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox("GDO non conforme. Invio annullato. Errore segnalato: " & ex.Message.ToString, MsgBoxStyle.Information, "Asc Client - GDO Errato")
            Exit Sub
        End Try

    End Sub

    Public Sub TrovaMSGID()

        Try
            MSGIDdef = ""
            MSGID = Split(TextBox2.Text, vbCrLf)
            For r = 0 To MSGID.Count - 1
                If MSGID(r).Contains("MSGID") Or MSGID(r).Contains("OGGETTO") Or MSGID(r).Contains("SUBJ") _
                    Or MSGID(r).Contains("SUBJECT") Or MSGID(r).Contains("ARG") Then
                    MSGIDdef = MSGID(r)
                End If
            Next
        Catch ex As Exception

        End Try

        If MSGIDdef <> "" Then
            ChkGDOIsRegular()
        Else
            Dim AlertMessage = MsgBox("Non è stato trovato alcun campo MSGID, si desidera proseguire comunque?", MsgBoxStyle.YesNo, "Asc Client - MSGID Mancante")
            If AlertMessage = 6 Then
                ChkGDOIsRegular()
            End If
        End If

    End Sub

    Public Sub ChkFmIsRegular()

        Try
            If IsFirstChkFmIsRegular = 0 Then
                IsFirstChkFmIsRegular = 1
                Dim MySplitter() As String = Split(TextBox2.Text, vbCrLf)
                If MySplitter(1).ToString.Contains(My.Settings.NomeUnita) Then
                    If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                        Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                        Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                        ChkGDOIsRegular()
                    Else
                        MsgBox("FM non conforme. Invio annullato.", MsgBoxStyle.Information, "Asc Server - FM Non Conforme")
                        Exit Sub
                    End If
                Else
                    MsgBox("FM non conforme. Invio annullato.", MsgBoxStyle.Information, "Asc Server - FM Non Conforme")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
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
            IsFirstControlloPOC = 0
        Catch ex As Exception

        End Try
        Try
            If MenuItem18.Text = "Connetti" Then
                MsgBox("Asc Client non risulta connesso ad Asc Exchange. Impossibile continuare.", MsgBoxStyle.Critical, "Asc Client - Errore Invio")
            Else
                Try
                    client.Send(Requests.StringMessage, "VerificaConnettivitàClient")
                Catch ex As Exception
                    MsgBox("Asc Client non risulta connesso ad Asc Exchange. Impossibile continuare.", MsgBoxStyle.Critical, "Asc Client - Errore Invio")
                    MenuItem18.Text = "Connetti"
                    Exit Sub
                End Try
                If TextBox3.Text <> "" Then
                    TextBox2.Text = TextBox2.Text.Replace("°", "^")
                    TextBox2.Text = TextBox2.Text.Replace("%", "pct")
                    ChkFmIsRegular()
                Else
                    MsgBox("Impossibile inviare messaggio ad ASC, scegliere un nome per il file e riprovare.", MsgBoxStyle.Information, "Asc Client - Nome file mancante")
                    Exit Sub
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub CreateMessagesFolder()

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi")
            End If

            My.Settings.JDOggi = Date.Now.DayOfYear

            If My.Settings.JDOggi = 0 Or My.Settings.JDIeri = 0 Then
                My.Settings.JDOggi = Date.Now.DayOfYear
                My.Settings.JDIeri = My.Settings.JDOggi
                If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3")) Then
                Else
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3"))
                End If
                If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Partenze") Then
                Else
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Partenze")
                End If
                If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Arrivi") Then
                Else
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Arrivi")
                End If
                If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Intelligence") Then
                Else
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Intelligence")
                End If
                If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\SVC") Then
                Else
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\SVC")
                End If
            Else
                If My.Settings.JDIeri > My.Settings.JDOggi Then
                    My.Settings.JDOggi = Date.Now.DayOfYear
                    My.Settings.JDIeri = My.Settings.JDOggi
                    If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3")) Then
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3"))
                    End If
                    If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Partenze") Then
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Partenze")
                    End If
                    If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Arrivi") Then
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Arrivi")
                    End If
                    If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Intelligence") Then
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Intelligence")
                    End If
                    If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\SVC") Then
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\SVC")
                    End If
                Else
                    For r = 0 To 365
                        If My.Settings.JDOggi <> My.Settings.JDIeri Then
                            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3")) Then
                            Else
                                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3"))
                            End If
                            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Partenze") Then
                            Else
                                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Partenze")
                            End If
                            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Arrivi") Then
                            Else
                                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Arrivi")
                            End If
                            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Intelligence") Then
                            Else
                                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Intelligence")
                            End If
                            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\SVC") Then
                            Else
                                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\SVC")
                            End If
                            My.Settings.JDIeri += 1
                        Else
                            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3")) Then
                            Else
                                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3"))
                            End If
                            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Partenze") Then
                            Else
                                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Partenze")
                            End If
                            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Arrivi") Then
                            Else
                                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Arrivi")
                            End If
                            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Intelligence") Then
                            Else
                                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\Intelligence")
                            End If
                            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\SVC") Then
                            Else
                                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & My.Settings.JDIeri.ToString("D3") & "\SVC")
                            End If
                            Exit For
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
        End Try

    End Sub

    Public Sub ConnectClient()

        If MenuItem18.Text = "Connetti" Then
            If MenuItem18.Text <> "" Then
                Try
                    client = New tcpConnection(My.Settings.IPAddress, CInt(My.Settings.PortaConnessione))
                    MenuItem18.Text = "Disconnetti"
                    client.Send(Requests.StringMessage, "NewClientConnected|" & My.Computer.Name.ToString & "|")
                    If Timer2.Enabled = False Then
                        Timer2.Start()
                    End If
                    If Timer1.Enabled = False Then
                        Timer1.Start()
                    End If
                    Try
                        RectangleShape1.Width = 93
                        RectangleShape1.BackColor = Color.Green
                        RectangleShape1.BorderColor = Color.Green
                    Catch ex As Exception

                    End Try
                    Timer4.Start()
                Catch ex As Exception
                    MenuItem18.Text = "Connetti"
                    MsgBox("Impossibile connetersi al server. Codice Errore Rilevato: " & ex.Message, MsgBoxStyle.Critical, "Asc Client - Errore Connessione")
                End Try
            Else
                MsgBox("Non è stato possibile tentare la connessione poichè non è stato ancora inserito alcun Indirizzo IP. Impostare L'indirizzo IP per poter continuare.", MsgBoxStyle.Information, "Asc Client - Indirizzo IP Mancante")
                IPAddress.Show()
            End If
        ElseIf MenuItem18.Text = "Disconnetti" Then
            Try
                client.Send(Requests.StringMessage, "ClientDisconnected" & "|" & My.Computer.Name.ToString & "|")
            Catch ex As Exception

            End Try
            Try
                MenuItem18.Text = "Connetti"
            Catch ex As Exception

            End Try
            Try
                client.Send(RequestTags.Disconnect)
                client.client.GetStream.Close()
                client.client.Close()
                client.client = Nothing
            Catch ex As Exception

            End Try
            Try
                If Timer2.Enabled = True Then
                    Timer2.Stop()
                End If
                If Timer1.Enabled = True Then
                    Timer1.Stop()
                End If
            Catch ex As Exception

            End Try
            Try
                RectangleShape1.Width = 79
                RectangleShape1.BackColor = Color.Red
                RectangleShape1.BorderColor = Color.Red
            Catch ex As Exception

            End Try
            Timer4.Start()
        End If

    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Try
            If client.client.Connected = True Then
                If client.client.Available > 0 Then
                    Dim AlertMessage = MsgBox("Il client sta ricevendo files da Asc Exchange, se viene chiuso in questo momento, i files andranno persi per sempre, si desidera continuare? [Sconsigliato]", MsgBoxStyle.YesNo, "Asc Exchange - Chiusura")
                    If AlertMessage = 6 Then
                        Try
                            client.Send(Requests.StringMessage, "ClientDisconnected" & "|" & My.Computer.Name.ToString & "|")
                        Catch ex As Exception
                        End Try

                        Try
                            client.Send(RequestTags.Disconnect)
                            client.client.GetStream.Close()
                            client.client.Close()
                            client.client = Nothing
                        Catch ex As Exception

                        End Try

                        Try
                            If Me.WindowState = FormWindowState.Maximized Then
                                SaveNewSize()
                            End If

                            Me.Hide()
                        Catch ex As Exception

                        End Try

                        Try
                            If Timer2.Enabled = True Then
                                Timer2.Stop()
                            End If
                            If Timer3.Enabled = True Then
                                Timer3.Stop()
                            End If
                        Catch ex As Exception
                        End Try

                        Try
                            PrintSett.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            SearchResults.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            Cerca.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            SplashScreen1.Close()
                        Catch ex As Exception
                        End Try
                    Else
                        e.Cancel = True
                    End If
                Else
                    Try
                        client.Send(Requests.StringMessage, "ClientDisconnected" & "|" & My.Computer.Name.ToString & "|")
                    Catch ex As Exception
                    End Try

                    Try
                        client.Send(RequestTags.Disconnect)
                        client.client.GetStream.Close()
                        client.client.Close()
                        client.client = Nothing
                    Catch ex As Exception

                    End Try

                    Try
                        If Me.WindowState = FormWindowState.Maximized Then
                            SaveNewSize()
                        End If

                        Me.Hide()
                    Catch ex As Exception

                    End Try

                    Try
                        If Timer2.Enabled = True Then
                            Timer2.Stop()
                        End If
                        If Timer3.Enabled = True Then
                            Timer3.Stop()
                        End If
                    Catch ex As Exception
                    End Try

                    Try
                        PrintSett.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        SearchResults.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        Cerca.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        SplashScreen1.Close()
                    Catch ex As Exception
                    End Try
                End If
            Else
                Try
                    If Me.WindowState = FormWindowState.Maximized Then
                        SaveNewSize()
                    End If

                    Me.Hide()
                Catch ex As Exception

                End Try

                Try
                    If Timer2.Enabled = True Then
                        Timer2.Stop()
                    End If
                    If Timer3.Enabled = True Then
                        Timer3.Stop()
                    End If
                Catch ex As Exception
                End Try

                Try
                    PrintSett.Close()
                Catch ex As Exception
                End Try

                Try
                    SearchResults.Close()
                Catch ex As Exception
                End Try

                Try
                    Cerca.Close()
                Catch ex As Exception
                End Try

                Try
                    SplashScreen1.Close()
                Catch ex As Exception
                End Try
            End If
        Catch ex1 As Exception
            Try
                If Me.WindowState = FormWindowState.Maximized Then
                    SaveNewSize()
                End If

                Me.Hide()
            Catch ex As Exception

            End Try

            Try
                If Timer2.Enabled = True Then
                    Timer2.Stop()
                End If
                If Timer3.Enabled = True Then
                    Timer3.Stop()
                End If
            Catch ex As Exception
            End Try

            Try
                PrintSett.Close()
            Catch ex As Exception
            End Try

            Try
                SearchResults.Close()
            Catch ex As Exception
            End Try

            Try
                Cerca.Close()
            Catch ex As Exception
            End Try

            Try
                SplashScreen1.Close()
            Catch ex As Exception
            End Try
        End Try

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If My.Settings.IPAddress = "" Then
            MsgBox("Non è ancora stato inserito alcun indirizzo IP, per un corretto funzionamento di Asc Client si consiglia di inserirlo adesso.", MsgBoxStyle.Information, "Asc Client - Indirizzo IP Mancante")
            IPAddress.Show()
            IPAddress.TopMost = True
        End If

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\Memorized") Then
                If System.IO.File.Exists(Application.StartupPath & "\Memorized\LastSize.txt") Then
                    Dim Reader As System.IO.StreamReader
                    Dim letto As String = ""
                    Reader = System.IO.File.OpenText(Application.StartupPath & "\Memorized\LastSize.txt")
                    letto = Reader.ReadToEnd
                    If letto = "MAX" Then
                        Me.WindowState = FormWindowState.Maximized
                    Else
                        Dim MySplitter() As String = Split(letto, "|")
                        Me.Width = Val(MySplitter(0).ToString)
                        Me.Height = Val(MySplitter(1).ToString)
                    End If
                End If
            End If
        Catch ex As Exception

        End Try

        Try
            ResizeForm()
            RadioButton1.Checked = True
        Catch ex As Exception

        End Try

        Try
            Timer3.Start()
        Catch ex As Exception

        End Try

        Try
            ConnectClient()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ResizeForm()

        Try
            If Me.WindowState <> FormWindowState.Minimized Then
                If MeOldWidth <> Me.Width Or MeOldHeihgt <> Me.Height Then
                    MeOldWidth = Me.Width
                    MeOldHeihgt = Me.Height

                    '526 : 1200 = 43.83333333333333 : 100
                    Dim LarghezzaGrpBox2 As Double = 43.833333333333329
                    GroupBox2.Width = ((Me.Width * LarghezzaGrpBox2) / 100)
                    GroupBox2.Height = (Me.Height - 80)

                    TextBox2.Width = (GroupBox2.Width - 17)
                    TextBox2.Height = (GroupBox2.Height - 152)

                    TextBox3.Width = (GroupBox2.Width - 139)

                    PictureBox1.Left = (TextBox3.Right + 25)

                    GroupBox1.Left = GroupBox2.Right + 6
                    GroupBox1.Width = (Me.Width - (GroupBox2.Width + 45))
                    GroupBox1.Height = (Me.Height - 80)

                    TextBox4.Width = (GroupBox1.Width - 16)
                    TextBox4.Height = (GroupBox1.Height - 152)

                    ListBox1.Width = (GroupBox1.Width - 354)
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub RefreshListView()

        Try
            FilesNameContainer = ""
            If RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False And RadioButton4.Checked = False _
                And RadioButton5.Checked = False Then
                Dim StringToBeAdded As String
                If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti") Then
                    ListBox1.Items.Clear()
                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Messaggi Ricevuti")
                    For r = 0 To UBound(FilesCollection)
                        If FilesCollection(r).ToString.Contains("SVC") Then
                            Dim FileInfo As New FileInfo(FilesCollection(r).ToString)
                            Dim Splitter() As String = Split(FileInfo.Name.ToString, ".txt")
                            ListBox1.Items.Add(Splitter(0).ToString)
                            FilesNameContainer += FilesCollection(r).ToString & "|"
                        Else
                            leggi = System.IO.File.OpenText(FilesCollection(r).ToString)
                            letto = leggi.ReadToEnd
                            leggi.Close()
                            Dim FileInfo As New FileInfo(FilesCollection(r).ToString)
                            If FileInfo.Name.ToString.Contains("- ") Then
                                Dim Splitter() As String = Split(letto, vbCrLf)
                                StringToBeAdded = Splitter(0).ToString & "  " & Splitter(1).ToString
                                If StringToBeAdded <> "" Then
                                    ListBox1.Items.Add("- " & StringToBeAdded)
                                    FilesNameContainer += FilesCollection(r).ToString & "|"
                                End If
                            Else
                                Dim Splitter() As String = Split(letto, vbCrLf)
                                StringToBeAdded = Splitter(0).ToString & "  " & Splitter(1).ToString
                                If StringToBeAdded <> "" Then
                                    ListBox1.Items.Add(StringToBeAdded)
                                    FilesNameContainer += FilesCollection(r).ToString & "|"
                                End If
                            End If
                        End If
                    Next
                End If
                Dim ItemsCounter As Integer = 0
                For r = 0 To ListBox1.Items.Count - 1
                    If ListBox1.Items(r).ToString = SelectedItem Then
                        ItemsCounter += 1
                    End If
                Next
                If ItemsCounter < 2 Then
                    For r = 0 To ListBox1.Items.Count - 1
                        If ListBox1.Items(r).ToString = SelectedItem Then
                            ListBox1.SetSelected(r, True)
                            Exit For
                        End If
                    Next
                End If
                RefreshLabelsNumber()
            Else
                RefreshLstViewWithSelection()
            End If
        Catch ex As Exception

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
                MsgBox("Errore nel salvataggio del file, messaggio errore " & ex.Message, MsgBoxStyle.Information, "Asc Client - Salvataggio File")
            End Try
        Else
            If TextBox3.Text = "" Then
                MsgBox("Devi scegliere un nome per il file da salvare.", MsgBoxStyle.Information, "Asc Client - Salvataggio File")
            ElseIf TextBox2.Text = "" Then
                MsgBox("Devi scrivere del testo per il messaggio da salvare.", MsgBoxStyle.Information, "Asc Client - Salvataggio File")
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
                MsgBox("Errore nel salvataggio del file, messaggio errore " & ex.Message, MsgBoxStyle.Information, "Asc Client - Salvataggio File")
            End Try
        Else
            If TextBox3.Text = "" Then
                MsgBox("Devi scegliere un nome per il file da salvare.", MsgBoxStyle.Information, "Asc Client - Salvataggio File")
            ElseIf TextBox2.Text = "" Then
                MsgBox("Devi scrivere del testo per il messaggio da salvare.", MsgBoxStyle.Information, "Asc Client - Salvataggio File")
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

    Public Sub OpenSelectedFile()

        Dim StringaDaConfrontare As String = ""
        Try
            If ListBox1.SelectedItems.Count > 0 Then
                If ListBox1.SelectedItem.ToString.Contains("- ") Then
                    SelectedItem = ListBox1.SelectedItem.ToString
                Else
                    SelectedItem = "- " & ListBox1.SelectedItem.ToString
                End If
                If ListBox1.SelectedItem.ToString.Contains("SVC") Then
                    leggi = System.IO.File.OpenText(Application.StartupPath & "\Messaggi Ricevuti\" & ListBox1.SelectedItem.ToString & ".txt")
                    letto = leggi.ReadToEnd
                    leggi.Close()
                    TextBox4.Text = letto
                    If AutomaticOpen = True Then
                        AutomaticOpen = False
                    Else
                        Dim FileInfo As New FileInfo(Application.StartupPath & "\Messaggi Ricevuti\" & ListBox1.SelectedItem.ToString & ".txt")
                        If FileInfo.Name.ToString.Contains("- ") Then
                        Else
                            If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\- " & ListBox1.SelectedItem.ToString & ".txt") Then
                                For r = 1 To 9999
                                    If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\- " & ListBox1.SelectedItem.ToString & "(" & r & ")" & ".txt") Then
                                    Else
                                        Rename(Application.StartupPath & "\Messaggi Ricevuti\" & ListBox1.SelectedItem.ToString & ".txt", Application.StartupPath & "\Messaggi Ricevuti\- " & ListBox1.SelectedItem.ToString & "(" & r & ")" & ".txt")
                                        RefreshListView()
                                    End If
                                Next
                            Else
                                Rename(Application.StartupPath & "\Messaggi Ricevuti\" & ListBox1.SelectedItem.ToString & ".txt", Application.StartupPath & "\Messaggi Ricevuti\- " & ListBox1.SelectedItem.ToString & ".txt")
                                RefreshListView()
                            End If
                        End If
                    End If
                Else
                    Dim MySplitter() As String = Split(FilesNameContainer, "|")
                    leggi = System.IO.File.OpenText(MySplitter(ListBox1.SelectedIndex).ToString)
                    letto = leggi.ReadToEnd
                    leggi.Close()
                    Dim Splitter() As String = Split(letto, vbCrLf)
                    StringaDaConfrontare = Splitter(0).ToString & "  " & Splitter(1).ToString
                    Dim FileInfo As New FileInfo(MySplitter(ListBox1.SelectedIndex).ToString)
                    If FileInfo.Name.ToString.Contains("- ") Then
                        If "- " & StringaDaConfrontare = ListBox1.SelectedItem.ToString Then
                            TextBox4.Text = letto
                            Exit Sub
                        End If
                    Else
                        If StringaDaConfrontare = ListBox1.SelectedItem.ToString Then
                            TextBox4.Text = letto
                            If AutomaticOpen = True Then
                                AutomaticOpen = False
                            Else
                                Dim MySplitter2() As String = Split(MySplitter(ListBox1.SelectedIndex).ToString, "\")
                                Dim MySplitter3() As String = Split(MySplitter2(MySplitter2.Count - 1).ToString, ".txt")
                                If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\- " & MySplitter3(0).ToString & ".txt") Then
                                    For r = 1 To 9999
                                        If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\- " & MySplitter3(0).ToString & "(" & r & ")" & ".txt") Then
                                        Else
                                            Rename(MySplitter(ListBox1.SelectedIndex).ToString, Application.StartupPath & "\Messaggi Ricevuti\- " & MySplitter3(0).ToString & "(" & r & ")" & ".txt")
                                            RefreshListView()
                                            Exit Sub
                                        End If
                                    Next
                                Else
                                    Rename(MySplitter(ListBox1.SelectedIndex).ToString, Application.StartupPath & "\Messaggi Ricevuti\- " & MySplitter2(MySplitter2.Count - 1).ToString)
                                    RefreshListView()
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged

        If RadioButton1.Checked = True Then
            AutomaticOpen = True
            RefreshLstViewWithSelection()
        End If

    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged

        If RadioButton2.Checked = True Then
            AutomaticOpen = True
            RefreshLstViewWithSelection()
        End If

    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged

        If RadioButton3.Checked = True Then
            AutomaticOpen = True
            RefreshLstViewWithSelection()
        End If

    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged

        If RadioButton4.Checked = True Then
            AutomaticOpen = True
            RefreshLstViewWithSelection()
        End If

    End Sub

    Private Sub RadioButton5_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton5.CheckedChanged

        If RadioButton5.Checked = True Then
            AutomaticOpen = True
            RefreshLstViewWithSelection()
        End If

    End Sub

    Public Sub RefreshLstViewWithSelection()

        Try
            FilesNameContainer = ""
            If RadioButton1.Checked = True Then
                Try
                    Dim StringToBeAdded As String
                    If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti") Then
                        ListBox1.Items.Clear()
                        Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Messaggi Ricevuti")
                        For r = 0 To UBound(FilesCollection)
                            If FilesCollection(r).ToString.Contains("SVC") Then
                            Else
                                leggi = System.IO.File.OpenText(FilesCollection(r).ToString)
                                letto = leggi.ReadToEnd
                                leggi.Close()
                                Dim FileInfo As New FileInfo(FilesCollection(r).ToString)
                                Dim Splitter() As String = Split(letto, vbCrLf)
                                StringToBeAdded = Splitter(0).ToString & "  " & Splitter(1).ToString
                                If StringToBeAdded <> "" Then
                                    Dim Splitter2() As String = Split(StringToBeAdded, " ")
                                    If Splitter2(0).ToString = "Z" Then
                                        If FileInfo.Name.ToString.Contains("- ") Then
                                            ListBox1.Items.Add("- " & StringToBeAdded)
                                            FilesNameContainer += FilesCollection(r).ToString & "|"
                                        Else
                                            ListBox1.Items.Add(StringToBeAdded)
                                            FilesNameContainer += FilesCollection(r).ToString & "|"
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                Catch ex As Exception

                End Try
            ElseIf RadioButton2.Checked = True Then
                Try
                    Dim StringToBeAdded As String
                    If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti") Then
                        ListBox1.Items.Clear()
                        Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Messaggi Ricevuti")
                        For r = 0 To UBound(FilesCollection)
                            If FilesCollection(r).ToString.Contains("SVC") Then
                            Else
                                leggi = System.IO.File.OpenText(FilesCollection(r).ToString)
                                letto = leggi.ReadToEnd
                                leggi.Close()
                                Dim FileInfo As New FileInfo(FilesCollection(r).ToString)
                                Dim Splitter() As String = Split(letto, vbCrLf)
                                StringToBeAdded = Splitter(0).ToString & "  " & Splitter(1).ToString
                                If StringToBeAdded <> "" Then
                                    Dim Splitter2() As String = Split(StringToBeAdded, " ")
                                    If Splitter2(0).ToString = "O" Then
                                        If FileInfo.Name.ToString.Contains("- ") Then
                                            ListBox1.Items.Add("- " & StringToBeAdded)
                                            FilesNameContainer += FilesCollection(r).ToString & "|"
                                        Else
                                            ListBox1.Items.Add(StringToBeAdded)
                                            FilesNameContainer += FilesCollection(r).ToString & "|"
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                Catch ex As Exception

                End Try
            ElseIf RadioButton3.Checked = True Then
                Try
                    Dim StringToBeAdded As String
                    If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti") Then
                        ListBox1.Items.Clear()
                        Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Messaggi Ricevuti")
                        For r = 0 To UBound(FilesCollection)
                            If FilesCollection(r).ToString.Contains("SVC") Then
                            Else
                                leggi = System.IO.File.OpenText(FilesCollection(r).ToString)
                                letto = leggi.ReadToEnd
                                leggi.Close()
                                Dim FileInfo As New FileInfo(FilesCollection(r).ToString)
                                Dim Splitter() As String = Split(letto, vbCrLf)
                                StringToBeAdded = Splitter(0).ToString & "  " & Splitter(1).ToString
                                If StringToBeAdded <> "" Then
                                    Dim Splitter2() As String = Split(StringToBeAdded, " ")
                                    If Splitter2(0).ToString = "P" Then
                                        If FileInfo.Name.ToString.Contains("- ") Then
                                            ListBox1.Items.Add("- " & StringToBeAdded)
                                            FilesNameContainer += FilesCollection(r).ToString & "|"
                                        Else
                                            ListBox1.Items.Add(StringToBeAdded)
                                            FilesNameContainer += FilesCollection(r).ToString & "|"
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                Catch ex As Exception

                End Try
            ElseIf RadioButton4.Checked = True Then
                Try
                    Dim StringToBeAdded As String
                    If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti") Then
                        ListBox1.Items.Clear()
                        Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Messaggi Ricevuti")
                        For r = 0 To UBound(FilesCollection)
                            If FilesCollection(r).ToString.Contains("SVC") Then
                            Else
                                leggi = System.IO.File.OpenText(FilesCollection(r).ToString)
                                letto = leggi.ReadToEnd
                                leggi.Close()
                                Dim FileInfo As New FileInfo(FilesCollection(r).ToString)
                                Dim Splitter() As String = Split(letto, vbCrLf)
                                StringToBeAdded = Splitter(0).ToString & "  " & Splitter(1).ToString
                                If StringToBeAdded <> "" Then
                                    Dim Splitter2() As String = Split(StringToBeAdded, " ")
                                    If Splitter2(0).ToString = "R" Then
                                        If FileInfo.Name.ToString.Contains("- ") Then
                                            ListBox1.Items.Add("- " & StringToBeAdded)
                                            FilesNameContainer += FilesCollection(r).ToString & "|"
                                        Else
                                            ListBox1.Items.Add(StringToBeAdded)
                                            FilesNameContainer += FilesCollection(r).ToString & "|"
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                Catch ex As Exception

                End Try
            ElseIf RadioButton5.Checked = True Then
                If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti") Then
                    ListBox1.Items.Clear()
                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Messaggi Ricevuti")
                    For r = 0 To UBound(FilesCollection)
                        If FilesCollection(r).ToString.Contains("SVC") Then
                            Dim FileInfo As New FileInfo(FilesCollection(r).ToString)
                            Dim Splitter() As String = Split(FileInfo.Name.ToString, ".txt")
                            ListBox1.Items.Add(Splitter(0).ToString)
                            FilesNameContainer += FilesCollection(r).ToString & "|"
                        End If
                    Next
                End If
            Else
                RefreshListView()
            End If
            Dim ItemsCounter As Integer = 0
            For r = 0 To ListBox1.Items.Count - 1
                If ListBox1.Items(r).ToString = SelectedItem Then
                    ItemsCounter += 1
                End If
            Next
            If ItemsCounter < 2 Then
                For r = 0 To ListBox1.Items.Count - 1
                    If ListBox1.Items(r).ToString = SelectedItem Then
                        ListBox1.SetSelected(r, True)
                        Exit For
                    End If
                Next
            End If
            RefreshLabelsNumber()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub RefreshLabelsNumber()

        Try
            Dim StringToBeAdded As String = ""
            Dim TotalZreceived As Integer = 0
            Dim TotalOreceived As Integer = 0
            Dim TotalPreceived As Integer = 0
            Dim TotalRreceived As Integer = 0
            Dim TotalDreceived As Integer = 0
            Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Messaggi Ricevuti")
            For r = 0 To UBound(FilesCollection)
                If FilesCollection(r).ToString.Contains("SVC") Then
                    TotalDreceived += 1
                Else
                    leggi = System.IO.File.OpenText(FilesCollection(r).ToString)
                    letto = leggi.ReadToEnd
                    leggi.Close()
                    Dim Splitter() As String = Split(letto, vbCrLf)
                    StringToBeAdded = Splitter(0).ToString & "  " & Splitter(1).ToString
                    If StringToBeAdded <> "" Then
                        If StringToBeAdded(0).ToString.Contains("Z") Then
                            TotalZreceived += 1
                        ElseIf StringToBeAdded(0).ToString.Contains("O") Then
                            TotalOreceived += 1
                        ElseIf StringToBeAdded(0).ToString.Contains("P") Then
                            TotalPreceived += 1
                        ElseIf StringToBeAdded(0).ToString.Contains("R") Then
                            TotalRreceived += 1
                        End If
                    End If
                End If
            Next
            RadioButton1.Text = "Z : " & TotalZreceived.ToString
            RadioButton2.Text = "O : " & TotalOreceived.ToString
            RadioButton3.Text = "P : " & TotalPreceived.ToString
            RadioButton4.Text = "R : " & TotalRreceived.ToString
            RadioButton5.Text = "D : " & TotalDreceived.ToString
            AutomaticallyOpen()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub AutomaticallyOpen()

        Try
            If AutomaticOpen = True Then
                If ListBox1.Items.Count > 0 Then
                    For r = 0 To ListBox1.Items.Count - 1
                        ListBox1.SetSelected(r, False)
                    Next
                    For r = 0 To ListBox1.Items.Count - 1
                        If ListBox1.Items(r).ToString.Contains("- ") Then
                            'Il file è già stato visualizzato.
                        Else
                            ListBox1.SetSelected(r, True)
                            Exit For
                        End If
                    Next
                    OpenSelectedFile()
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click

        If ListBox1.SelectedItems.Count > 1 Then
            DecodaMultiplo()
        Else
            Decoda()
        End If

    End Sub

    Public Sub Decoda()

        Try
            If ListBox1.SelectedItems.Count > 0 Then
                If ListBox1.SelectedItem.ToString.Contains("SVC") Then
                    System.IO.File.Delete(Application.StartupPath & "\Messaggi Ricevuti\" & ListBox1.SelectedItem.ToString & ".txt")
                Else
                    Dim MySplitter() As String = Split(FilesNameContainer, "|")
                    System.IO.File.Delete(MySplitter(ListBox1.SelectedIndex).ToString)
                    TextBox4.Clear()
                    TextBox2.Clear()
                    RefreshListView()
                    Exit Sub
                End If
            Else
                MsgBox("Deve essere selezionato un messaggio per poter proseguire.", MsgBoxStyle.Information, "Asc Client - Decoda")
            End If
            TextBox4.Clear()
            TextBox2.Clear()
            RefreshListView()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub DecodaMultiplo()

        Try
            Dim ListboxSelectedIndices As ListBox.SelectedIndexCollection = ListBox1.SelectedIndices
            For r = 0 To ListboxSelectedIndices.Count - 1
                If ListBox1.Items(ListboxSelectedIndices(r)).ToString.Contains("SVC") Then
                    System.IO.File.Delete(Application.StartupPath & "\Messaggi Ricevuti\" & ListBox1.SelectedItem.ToString & ".txt")
                Else
                    Dim MySplitter() As String = Split(FilesNameContainer, "|")
                    System.IO.File.Delete(MySplitter(ListboxSelectedIndices(r)).ToString)
                End If
            Next
            TextBox4.Clear()
            TextBox2.Clear()
            RefreshListView()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub Archiviazione()

        Try
            Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Messaggi Ricevuti")
            For r = 0 To UBound(FilesCollection)
                Dim FileInfo As New FileInfo(FilesCollection(r).ToString)
                If FileInfo.Name.ToString.Contains("- ") Then

                Else
                    If FileInfo.Name.ToString.Contains("SVC") Then
                        leggi = System.IO.File.OpenText(Application.StartupPath & "\Messaggi Ricevuti\" & FileInfo.Name.ToString)
                        letto = leggi.ReadToEnd
                        leggi.Close()
                        Dim JulianDateDaUtilizzare As String = ""
                        Dim MySplitter() As String = Split(letto, vbCrLf)
                        JulianDateDaUtilizzare = MySplitter(0).ToString

                        If Val(JulianDateDaUtilizzare) > 0 Then
                            Dim DefinitiveString As String = ""
                            For x = 1 To UBound(MySplitter)
                                DefinitiveString += MySplitter(x).ToString & vbCrLf
                            Next
                            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & FileInfo.Name.ToString)
                            scrivi.Write(DefinitiveString)
                            scrivi.Close()
                            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & JulianDateDaUtilizzare & "\SVC") Then
                            Else
                                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & JulianDateDaUtilizzare & "\SVC")
                            End If
                            If System.IO.File.Exists(Application.StartupPath & "\Messaggi\" & JulianDateDaUtilizzare & "\SVC\" & FileInfo.Name.ToString) Then
                                'Do texts comparison
                                Dim text1 As String = ""
                                Dim text2 As String = ""
                                Dim reader As System.IO.StreamReader = System.IO.File.OpenText(Application.StartupPath & "\Messaggi\" & JulianDateDaUtilizzare & "\SVC\" & FileInfo.Name.ToString)
                                text1 = reader.ReadToEnd
                                reader.Close()
                                reader = System.IO.File.OpenText(Application.StartupPath & "\Messaggi Ricevuti\" & FileInfo.Name.ToString)
                                text2 = reader.ReadToEnd
                                reader.Close()
                                If text1 = text2 Then
                                Else
                                    For c = 1 To 9999
                                        If System.IO.File.Exists(Application.StartupPath & "\Messaggi\" & JulianDateDaUtilizzare & "\SVC\(" & c.ToString & ")" & FileInfo.Name.ToString) Then
                                        Else
                                            System.IO.File.Copy(Application.StartupPath & "\Messaggi Ricevuti\" & FileInfo.Name.ToString, _
                                Application.StartupPath & "\Messaggi\" & JulianDateDaUtilizzare & "\SVC\(" & c.ToString & ")" & FileInfo.Name.ToString)
                                            Exit For
                                        End If
                                    Next
                                End If
                            Else
                                System.IO.File.Copy(Application.StartupPath & "\Messaggi Ricevuti\" & FileInfo.Name.ToString, _
                                Application.StartupPath & "\Messaggi\" & JulianDateDaUtilizzare & "\SVC\" & FileInfo.Name.ToString)
                            End If
                        End If

                    Else

                        leggi = System.IO.File.OpenText(Application.StartupPath & "\Messaggi Ricevuti\" & FileInfo.Name.ToString)
                        letto = leggi.ReadToEnd
                        leggi.Close()
                        Dim Splitter2() As String = Split(letto, vbCrLf)
                        Dim DataArchiviazione As String = FileInfo.Name(5).ToString & FileInfo.Name(6).ToString & FileInfo.Name(7).ToString
                        If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & DataArchiviazione) Then
                        Else
                            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & DataArchiviazione)
                        End If
                        If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & DataArchiviazione & "\Partenze") Then
                        Else
                            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & DataArchiviazione & "\Partenze")
                        End If
                        If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & DataArchiviazione & "\Arrivi") Then
                        Else
                            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & DataArchiviazione & "\Arrivi")
                        End If
                        If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi\" & DataArchiviazione & "\Intelligence") Then
                        Else
                            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi\" & DataArchiviazione & "\Intelligence")
                        End If
                        If FileInfo.Name(0).ToString = "I" Then
                            If System.IO.File.Exists(Application.StartupPath & "\Messaggi\" & DataArchiviazione & "\Intelligence\" & FileInfo.Name.ToString) Then
                                'Do nothing
                            Else
                                System.IO.File.Copy(Application.StartupPath & "\Messaggi Ricevuti\" & FileInfo.Name.ToString, _
                                    Application.StartupPath & "\Messaggi\" & DataArchiviazione & "\Intelligence\" & FileInfo.Name.ToString)
                            End If
                        Else
                            If Splitter2(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                                Or Splitter2(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                                Or Splitter2(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                                If System.IO.File.Exists(Application.StartupPath & "\Messaggi\" & DataArchiviazione & "\Partenze\" & FileInfo.Name.ToString) Then
                                    'Do nothing
                                Else
                                    System.IO.File.Copy(Application.StartupPath & "\Messaggi Ricevuti\" & FileInfo.Name.ToString, _
                                        Application.StartupPath & "\Messaggi\" & DataArchiviazione & "\Partenze\" & FileInfo.Name.ToString)
                                End If
                            Else
                                If System.IO.File.Exists(Application.StartupPath & "\Messaggi\" & DataArchiviazione & "\Arrivi\" & FileInfo.Name.ToString) Then
                                    'Do nothing
                                Else
                                    System.IO.File.Copy(Application.StartupPath & "\Messaggi Ricevuti\" & FileInfo.Name.ToString, _
                                        Application.StartupPath & "\Messaggi\" & DataArchiviazione & "\Arrivi\" & FileInfo.Name.ToString)
                                End If
                            End If
                        End If

                    End If
                End If
            Next
        Catch ex As Exception
        End Try

    End Sub

    Private Sub MenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem12.Click

        Cerca.Show()

    End Sub

    Private Sub ListBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            OpenSelectedFile()
        End If

    End Sub

    Private Sub ListBox1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBox1.MouseDoubleClick

        If e.Button = Windows.Forms.MouseButtons.Left Then
            OpenSelectedFile()
        End If

    End Sub

    Private Sub MenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem14.Click

        My.Settings.x = 35
        NomeUnita.Show()

    End Sub

    Private Sub MenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem16.Click

        IPAddress.Show()

    End Sub

    Public Sub SaveNewSize()

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\Memorized") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Memorized")
            End If
            If Me.WindowState = FormWindowState.Maximized Then
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\LastSize.txt")
                scrivi.Write("MAX")
                scrivi.Close()
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\LastSize.txt")
                scrivi.Write(Me.Width & "|" & Me.Height & "|")
                scrivi.Close()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        If Me.WindowState = FormWindowState.Maximized Then
            SaveNewSize()
        End If

    End Sub

    Private Sub Form1_ResizeBegin(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeBegin

        Resizing = True

    End Sub

    Private Sub Form1_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd

        Resizing = False
        ResizeForm()
        SaveNewSize()

    End Sub

    Private Sub Form1_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged

        If Resizing = False Then
            ResizeForm()
        End If

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        Try
            TextBox4.Clear()
            TextBox2.Clear()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick

        Try
            If client.client.Connected = False Then
                If RectangleShape1.BackColor = Color.Green Then
                    RectangleShape1.BackColor = Color.Red
                    RectangleShape1.BorderColor = Color.Red
                    MenuItem18.Text = "Connetti"
                    Timering = 0
                    Timer4.Start()
                End If
            End If
        Catch ex As Exception
            If RectangleShape1.BackColor = Color.Green Then
                RectangleShape1.BackColor = Color.Red
                RectangleShape1.BorderColor = Color.Red
                MenuItem18.Text = "Connetti"
                Timering = 0
                Timer4.Start()
            End If
        End Try

        Try
            If MessaggioRicevuto.Contains("AscExchangeChkClientConnection") Then
                If MessaggioRicevuto.Contains(My.Computer.Name.ToString) Then
                    MessaggioRicevuto = ""
                    client.Send(Requests.DataFile, "ClientIsConnected=True")
                Else
                    MessaggioRicevuto = ""
                End If
            ElseIf MessaggioRicevuto = "ChkingConnectedClientsList" Then
                MessaggioRicevuto = ""
                client.Send(Requests.StringMessage, "NewClientConnected|" & My.Computer.Name.ToString & "|")
            ElseIf MessaggioRicevuto = "ChiusuraAscExchangeInCorso" Then
                MessaggioRicevuto = ""
                Try
                    MenuItem18.Text = "Connetti"
                Catch ex As Exception

                End Try
                Try
                    client.Send(RequestTags.Disconnect)
                    client.client.GetStream.Close()
                    client.client.Close()
                    client.client = Nothing
                Catch ex As Exception

                End Try
                If Timer2.Enabled = True Then
                    Timer2.Stop()
                End If
                If Timer1.Enabled = True Then
                    Timer1.Stop()
                End If
                Try
                    If client.client.Connected = False Then
                        If RectangleShape1.BackColor = Color.Green Then
                            RectangleShape1.BackColor = Color.Red
                            RectangleShape1.BorderColor = Color.Red
                            MenuItem18.Text = "Connetti"
                            Timering = 0
                            Timer4.Start()
                        End If
                    End If
                Catch ex As Exception

                End Try
            Else
                MessaggioRicevuto = ""
            End If
        Catch ex As Exception
            If Timer2.Enabled = False Then
                Timer2.Start()
            End If
            Exit Sub
        End Try

    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\temp") Then
                Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\temp")
                If FilesCollection.Count > 0 Then
                    Smista()
                End If
            End If
        Catch ex As Exception
            Exit Sub
        End Try

    End Sub

    Public Sub Smista()

        Try
            If client.client.Available = 0 Then
                Try
                    If Timer3.Enabled = True Then
                        Timer3.Stop()
                    End If
                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\temp")
                    For z = 0 To UBound(FilesCollection)
                        Dim lettore As System.IO.StreamReader
                        Dim letto As String
                        lettore = System.IO.File.OpenText(FilesCollection(z).ToString)
                        letto = lettore.ReadToEnd
                        lettore.Close()
                        Dim definitivestring As String = ""
                        Dim mysplitter() As String = Split(letto, vbCrLf)
                        For x = 1 To UBound(mysplitter)
                            definitivestring += mysplitter(x).ToString & vbCrLf
                        Next
                        If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti") Then
                            If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt") Then
                                For r = 1 To 9999
                                    If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & "(" & r.ToString & ").txt") Then
                                    Else
                                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & "(" & r.ToString & ").txt")
                                        scrivi.Write(definitivestring)
                                        scrivi.Close()
                                        System.IO.File.Delete(FilesCollection(z).ToString)
                                        RefreshListView()
                                        Exit For
                                    End If
                                Next
                            Else
                                Try
                                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt")
                                    scrivi.Write(definitivestring)
                                    scrivi.Close()
                                    leggi = System.IO.File.OpenText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt")
                                    letto = leggi.ReadToEnd
                                    leggi.Close()
                                    Dim Splitter() As String = Split(letto, vbCrLf)
                                    Dim MyString As String = Splitter(0).ToString
                                    If MyString(0) = "Z" Then
                                        RadioButton1.Checked = True
                                    End If
                                    System.IO.File.Delete(FilesCollection(z).ToString)
                                    RefreshListView()
                                Catch ex As Exception

                                End Try
                            End If

                        Else

                            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Ricevuti")
                            If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt") Then
                                For r = 1 To 9999
                                    If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & "(" & r.ToString & ").txt") Then
                                    Else
                                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & "(" & r.ToString & ").txt")
                                        scrivi.Write(definitivestring)
                                        scrivi.Close()
                                        System.IO.File.Delete(FilesCollection(z).ToString)
                                        RefreshListView()
                                        Exit For
                                    End If
                                Next
                            Else
                                Try
                                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt")
                                    scrivi.Write(definitivestring)
                                    scrivi.Close()
                                    leggi = System.IO.File.OpenText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt")
                                    letto = leggi.ReadToEnd
                                    leggi.Close()
                                    Dim Splitter() As String = Split(letto, vbCrLf)
                                    Dim MyString As String = Splitter(0).ToString
                                    If MyString(0) = "Z" Then
                                        RadioButton1.Checked = True
                                    End If
                                    System.IO.File.Delete(FilesCollection(z).ToString)
                                    RefreshListView()
                                Catch ex As Exception

                                End Try
                            End If
                        End If
                        My.Settings.FileName = mysplitter(0).ToString
                        Archiviazione()
                    Next
                    If Timer3.Enabled = False Then
                        Timer3.Start()
                    End If
                Catch ex As Exception
                    If Timer3.Enabled = False Then
                        Timer3.Start()
                    End If
                End Try
            End If
        Catch ex As Exception
            Try
                If Timer3.Enabled = True Then
                    Timer3.Stop()
                End If
                Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\temp")
                For z = 0 To UBound(FilesCollection)
                    Dim lettore As System.IO.StreamReader
                    Dim letto As String
                    lettore = System.IO.File.OpenText(FilesCollection(z).ToString)
                    letto = lettore.ReadToEnd
                    lettore.Close()
                    Dim definitivestring As String = ""
                    Dim mysplitter() As String = Split(letto, vbCrLf)
                    For x = 1 To UBound(mysplitter)
                        definitivestring += mysplitter(x).ToString & vbCrLf
                    Next
                    If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti") Then
                        If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt") Then
                            For r = 1 To 9999
                                If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & "(" & r.ToString & ").txt") Then
                                Else
                                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & "(" & r.ToString & ").txt")
                                    scrivi.Write(definitivestring)
                                    scrivi.Close()
                                    System.IO.File.Delete(FilesCollection(z).ToString)
                                    RefreshListView()
                                    Exit For
                                End If
                            Next
                        Else
                            Try
                                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt")
                                scrivi.Write(definitivestring)
                                scrivi.Close()
                                leggi = System.IO.File.OpenText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt")
                                letto = leggi.ReadToEnd
                                leggi.Close()
                                Dim Splitter() As String = Split(letto, vbCrLf)
                                Dim MyString As String = Splitter(0).ToString
                                If MyString(0) = "Z" Then
                                    RadioButton1.Checked = True
                                End If
                                System.IO.File.Delete(FilesCollection(z).ToString)
                                RefreshListView()
                            Catch ex2 As Exception

                            End Try
                        End If

                    Else

                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Ricevuti")
                        If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt") Then
                            For r = 1 To 9999
                                If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & "(" & r.ToString & ").txt") Then
                                Else
                                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & "(" & r.ToString & ").txt")
                                    scrivi.Write(definitivestring)
                                    scrivi.Close()
                                    System.IO.File.Delete(FilesCollection(z).ToString)
                                    RefreshListView()
                                    Exit For
                                End If
                            Next
                        Else
                            Try
                                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt")
                                scrivi.Write(definitivestring)
                                scrivi.Close()
                                leggi = System.IO.File.OpenText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt")
                                letto = leggi.ReadToEnd
                                leggi.Close()
                                Dim Splitter() As String = Split(letto, vbCrLf)
                                Dim MyString As String = Splitter(0).ToString
                                If MyString(0) = "Z" Then
                                    RadioButton1.Checked = True
                                End If
                                System.IO.File.Delete(FilesCollection(z).ToString)
                                RefreshListView()
                            Catch ex2 As Exception

                            End Try
                        End If
                    End If
                    My.Settings.FileName = mysplitter(0).ToString
                    Archiviazione()
                Next
                If Timer3.Enabled = False Then
                    Timer3.Start()
                End If
            Catch ex1 As Exception
                If Timer3.Enabled = False Then
                    Timer3.Start()
                End If
            End Try
        End Try

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Try
            If MenuItem18.Text = "Disconnetti" Then
                If client.isConnected = True Then
                    Try
                        client.Send(Requests.StringMessage, "NewClientConnected|" & My.Computer.Name.ToString & "|")
                    Catch ex As Exception
                        Exit Sub
                    End Try
                End If
            End If
        Catch ex As Exception
            Exit Sub
        End Try

    End Sub

    Private Sub MenuItem18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem18.Click

        ConnectClient()

    End Sub

    Dim Timering As Integer = 0
    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick

        Try
            Timering += 1
            If Timering < 79 Then
                If MenuItem18.Text = "Disconnetti" Then
                    If RectangleShape1.Width < 93 Then
                        RectangleShape1.Width += 1
                    Else
                        RectangleShape1.Width = 0
                    End If
                ElseIf MenuItem18.Text = "Connetti" Then
                    If RectangleShape1.Width < 79 Then
                        RectangleShape1.Width += 1
                    Else
                        RectangleShape1.Width = 0
                    End If
                End If
            Else
                If MenuItem18.Text = "Disconnetti" Then
                    RectangleShape1.Width = 93
                ElseIf MenuItem18.Text = "Connetti" Then
                    RectangleShape1.Width = 79
                End If
                Timering = 0
                Timer4.Stop()
            End If
        Catch ex As Exception

        End Try

    End Sub
End Class
