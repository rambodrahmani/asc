Imports System.Net, System.Net.Sockets, System.IO, System.IO.Directory
Imports Asc.tcpConnection

Public Class Main

    Dim IsFirstChkLastBT As Integer = 0
    Dim IsFisrChkLastBT As Integer = 0
    Dim IsFirstChkBtIsRegular As Integer = 0
    Dim IsFirstControlloPOC As Integer = 0
    Dim IsFirstSelectBTasLastLine As Integer = 0
    Dim IsFirstChkBTcounter As Integer = 0
    Dim IsFirstChkFollowingSICline As Integer = 0
    Dim IsFirstChkGDOIsRegular As Integer = 0
    Dim IsFirstChkFmIsRegular As Integer = 0
    Dim FinalSectionNr As Integer = 0
    Dim OpenedFile As Integer = 0
    Dim OpenedFileNr As Integer = 0
    Dim FinalSectionPresent As Integer = 0
    Dim ListView1LastSelectedItem As String = "NONSELEZIONATO"
    Dim RecreateTxtBox3 As String = ""
    '*** Dichiarazione keybd_event inizio ***
    Declare Sub keybd_event Lib "user32" _
        (ByVal bVk As Byte, ByVal bScan As Byte, _
         ByVal dwFlags As Integer, ByVal dwExtraInfo As Integer)
    '*** Dichiarazione keybd_event fine ***
    Dim InvioAnnullato As Integer = 0
    Dim TxtBox3LenCounter As Integer = 0
    Dim ManualPrint As Integer = 0
    Dim IsPCNAvailable As String
    Dim TxtBox3Text As String
    Dim PCNSplitter() As String
    Dim contatoreDecret As Integer = 0
    Dim stoppami As Integer = 0
    Dim numeroriga As Integer = 0
    Dim RigheTextBox() As String
    Dim NumFoglioStampa As Integer = 0
    Dim MSGIDdef As String
    Dim MSGID() As String
    Dim numaparola As String
    Dim rigaemail As String
    Dim dastampare As String
    Dim riga() As String
    Dim primadec As Integer = 0
    Dim infoufficio() As String
    Dim contatto() As String
    Dim ora As String = Date.Now.Hour & ":" & Date.Now.Minute & ":" & Date.Now.Second
    Dim datagiorno As String
    Dim chkAllegato As String
    Dim scrivi As System.IO.StreamWriter
    Dim leggi As System.IO.StreamReader
    Dim letto As String
    Dim int1 As Integer
    Dim int2 As Integer
    Dim objOutlookMsg As Object
    Dim objOutlook As Object
    Private stringToPrint As String
    Dim FileName As String = ""
    Dim PCNCompleteEmailText As String = ""
    Dim PCNPartedEmailText As String = ""
    Dim MeOldWidth As Double = 0
    Dim MeOldHeihgt As Double = 0
    Dim Resizing As Boolean = False
    Dim Timering As Integer = 0
    Dim AspettaAncora As Integer = 0
    Dim NumeroUltimoUfficio As Integer = 0
    Dim ListaClients As String = ""
    Dim AspettaIlProcessoDiInvioCode As Integer = 0
    Dim CollezioneElementiCheckati As String
    Dim sezionedecretazione As String = ""

#Region "DichiarazioneClient"
    Public WithEvents client As tcpConnection
    Dim MessaggioRicevuto As String = ""
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

    Public Sub Smista()

        Try
            If client.client.Available = 0 Then
                If Timer4.Enabled = True Then
                    Timer4.Stop()
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
                    Try
                        If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti") Then
                            If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt") Then
                                For r = 1 To 9999
                                    If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & "(" & r & ")" & ".txt") Then
                                    Else
                                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & "(" & r & ")" & ".txt")
                                        scrivi.Write(definitivestring)
                                        scrivi.Close()
                                        System.IO.File.Delete(FilesCollection(z).ToString)
                                        Exit For
                                    End If
                                Next
                            Else
                                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt")
                                scrivi.Write(definitivestring)
                                scrivi.Close()
                                System.IO.File.Delete(FilesCollection(z).ToString)
                            End If
                        Else
                            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Ricevuti")
                            If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt") Then
                                For r = 1 To 9999
                                    If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & "(" & r & ")" & ".txt") Then
                                    Else
                                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & "(" & r & ")" & ".txt")
                                        scrivi.Write(definitivestring)
                                        scrivi.Close()
                                        System.IO.File.Delete(FilesCollection(z).ToString)
                                        Exit For
                                    End If
                                Next
                            Else
                                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt")
                                scrivi.Write(definitivestring)
                                scrivi.Close()
                                System.IO.File.Delete(FilesCollection(z).ToString)
                            End If
                        End If
                    Catch ex As Exception
                        If Timer4.Enabled = False Then
                            Timer4.Start()
                        End If
                        MsgBox("Errore durante il salvataggio di un messaggio ricevuto da un client, messaggio di errore rilevato: " & ex.Message, MsgBoxStyle.Critical, "Asc Server - Errore Salvataggio")
                    End Try
                    RefreshListBox()
                Next
                If Timer4.Enabled = False Then
                    Timer4.Start()
                End If
            End If
        Catch ex As Exception
            Try
                If Timer4.Enabled = True Then
                    Timer4.Stop()
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
                    Try
                        If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti") Then
                            If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt") Then
                                For r = 1 To 9999
                                    If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & "(" & r & ")" & ".txt") Then
                                    Else
                                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & "(" & r & ")" & ".txt")
                                        scrivi.Write(definitivestring)
                                        scrivi.Close()
                                        System.IO.File.Delete(FilesCollection(z).ToString)
                                        Exit For
                                    End If
                                Next
                            Else
                                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt")
                                scrivi.Write(definitivestring)
                                scrivi.Close()
                                System.IO.File.Delete(FilesCollection(z).ToString)
                            End If
                        Else
                            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Ricevuti")
                            If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt") Then
                                For r = 1 To 9999
                                    If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & "(" & r & ")" & ".txt") Then
                                    Else
                                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & "(" & r & ")" & ".txt")
                                        scrivi.Write(definitivestring)
                                        scrivi.Close()
                                        System.IO.File.Delete(FilesCollection(z).ToString)
                                        Exit For
                                    End If
                                Next
                            Else
                                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(0).ToString & ".txt")
                                scrivi.Write(definitivestring)
                                scrivi.Close()
                                System.IO.File.Delete(FilesCollection(z).ToString)
                            End If
                        End If
                    Catch ex1 As Exception
                        If Timer4.Enabled = False Then
                            Timer4.Start()
                        End If
                        MsgBox("Errore durante il salvataggio di un messaggio ricevuto da un client, messaggio di errore rilevato: " & ex.Message, MsgBoxStyle.Critical, "Asc Server - Errore Salvataggio")
                    End Try
                    RefreshListBox()
                Next
                If Timer4.Enabled = False Then
                    Timer4.Start()
                End If
            Catch ex1 As Exception
                If Timer4.Enabled = False Then
                    Timer4.Start()
                End If
            End Try
        End Try

    End Sub

    Public Sub LoadContacts()

        Try
            StatusBar1.Text = "Stato: Caricamento lista contatti iniziato"
            CheckedListBox1.Items.Clear()
            leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\listauffici.txt")
            letto = leggi.ReadToEnd
            leggi.Close()
            contatto = Split(letto, vbCrLf)
            If contatto.Count > 0 Then
                For r = 0 To UBound(contatto)
                    If contatto(r) <> "" Then
                        leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                        letto = leggi.ReadToEnd
                        leggi.Close()
                        infoufficio = Split(letto, "#")
                        If infoufficio.Count > 3 Then
                            CheckedListBox1.Items.Add(contatto(r).ToUpper)
                        Else
                            CheckedListBox1.Items.Add(contatto(r))
                        End If
                    End If
                Next
            Else
                CheckedListBox1.Items.Clear()
            End If
            StatusBar1.Text = "Stato: Pronto per l'utilizzo."
        Catch ex As Exception

        End Try

    End Sub

    Public Sub TrovaMSGID()

        Try
            StatusBar1.Text = "Stato Invio: Controllo riga MSGID"
            MSGIDdef = ""
            MSGID = Split(TextBox3.Text, vbCrLf)
            For r = 0 To MSGID.Count - 1
                If MSGID(r).Contains("MSGID") Or MSGID(r).Contains("OGGETTO") Or MSGID(r).Contains("SUBJ") _
                    Or MSGID(r).Contains("SUBJECT") Or MSGID(r).Contains("ARG") Then
                    MSGIDdef = MSGID(r)
                End If
            Next
        Catch ex As Exception

        End Try

        If MSGIDdef <> "" Then
            ChkProgressiveNumber()
        Else
            Dim AlertMessage = MsgBox("Non è stato trovato alcun campo MSGID, si desidera proseguire comunque?", MsgBoxStyle.YesNo, "Asc Server - MSGID Mancante")
            If AlertMessage = 6 Then
                ChkProgressiveNumber()
            End If
        End If

    End Sub

    Public Sub CreateTempFiles()

        Try
            Dim UserDesktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
            If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                If System.IO.Directory.Exists(UserDesktopPath & "\Messaggi in Partenza") Then
                    scrivi = System.IO.File.CreateText(UserDesktopPath & "\Messaggi in Partenza\" & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & ".txt")
                    scrivi.Write(TextBox3.Text)
                    scrivi.Close()
                Else
                    System.IO.Directory.CreateDirectory(UserDesktopPath & "\Messaggi in Partenza")
                    scrivi = System.IO.File.CreateText(UserDesktopPath & "\Messaggi in Partenza\" & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & ".txt")
                    scrivi.Write(TextBox3.Text)
                    scrivi.Close()
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ChkProgressiveNumber()

        Try
            StatusBar1.Text = "Stato Invio: Controllo numero progressivo"
            Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
            If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                Dim filescollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno)
                If filescollection.Count > 1 Then
                    If filescollection.Count - 1 > My.Settings.NumEmail2 Then
                        MsgBox("Rilevato errore nella numerazione progressiva in partenza. Invio annullato.", MsgBoxStyle.Information, "Asc Server - Errore Numerazione")
                        StatusBar1.Text = "Invio Annullato. Errore numerazione progressiva in pratenza rilevato"
                        InvioAnnullato = 1
                        Exit Sub
                    End If
                    If filescollection.Count - 1 < My.Settings.NumEmail2 Then
                        MsgBox("Rilevato errore nella numerazione progressiva in partenza. Invio annullato.", MsgBoxStyle.Information, "Asc Server - Errore Numerazione")
                        StatusBar1.Text = "Invio Annullato. Errore numerazione progressiva in pratenza rilevato"
                        InvioAnnullato = 1
                        Exit Sub
                    End If

                    For r = 1 To My.Settings.NumEmail2
                        If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\emailinviate" & r.ToString("D4") & "P.txt") Then

                        Else
                            MsgBox("Rilevato errore nella numerazione progressiva in partenza. Invio annullato.", MsgBoxStyle.Information, "Asc Server - Errore Numerazione")
                            StatusBar1.Text = "Invio Annullato. Errore numerazione progressiva in pratenza rilevato"
                            InvioAnnullato = 1
                            Exit Sub
                        End If
                        If r = My.Settings.NumEmail2 Then
                            My.Settings.NumEmail2 += 1
                            FileName = "T" & Date.Now.Year.ToString & My.Settings.JlDate.ToString("D3") & My.Settings.NumEmail2.ToString("D4")
                            CreateTempFiles()
                        End If
                    Next
                Else
                    My.Settings.NumEmail2 += 1
                    FileName = "T" & Date.Now.Year.ToString & My.Settings.JlDate.ToString("D3") & My.Settings.NumEmail2.ToString("D4")
                    CreateTempFiles()
                End If

            Else

                Dim filescollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno)
                If filescollection.Count > 1 Then
                    If filescollection.Count - 1 > My.Settings.NumEmail Then
                        MsgBox("Rilevato errore nella numerazione progressiva in arrivo. Invio annullato.", MsgBoxStyle.Information, "Asc Server - Errore Numerazione")
                        StatusBar1.Text = "Invio Annullato. Errore numerazione progressiva in arrivo rilevato"
                        InvioAnnullato = 1
                        Exit Sub
                    End If
                    If filescollection.Count - 1 < My.Settings.NumEmail Then
                        MsgBox("Rilevato errore nella numerazione progressiva in arrivo. Invio annullato.", MsgBoxStyle.Information, "Asc Server - Errore Numerazione")
                        StatusBar1.Text = "Invio Annullato. Errore numerazione progressiva in arrivo rilevato"
                        InvioAnnullato = 1
                        Exit Sub
                    End If

                    For r = 1 To My.Settings.NumEmail
                        If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\emailinviate" & r.ToString("D4") & "A.txt") Then

                        Else
                            MsgBox("Rilevato errore nella numerazione progressiva in arrivo. Invio annullato.", MsgBoxStyle.Information, "Asc Server - Errore Numerazione")
                            StatusBar1.Text = "Invio Annullato. Errore numerazione progressiva in arrivo rilevato"
                            InvioAnnullato = 1
                            Exit Sub
                        End If
                        If r = My.Settings.NumEmail Then
                            My.Settings.NumEmail += 1
                            FileName = "R" & Date.Now.Year.ToString & My.Settings.JlDate.ToString("D3") & My.Settings.NumEmail.ToString("D4")
                        End If
                    Next
                Else
                    My.Settings.NumEmail += 1
                    FileName = "R" & Date.Now.Year.ToString & My.Settings.JlDate.ToString("D3") & My.Settings.NumEmail.ToString("D4")
                End If

            End If

            LoadList()

        Catch ex As Exception

        End Try

    End Sub

    Public Sub LoadList()

        Try
            StatusBar1.Text = "Stato Invio: Caricamento lista destinatari"
            Try
                contatoreDecret = 0
                TextBox3.Text += vbCrLf & vbCrLf & "Email inviata a: "
                sezionedecretazione = vbCrLf & vbCrLf & "Email inviata a: "
                For z = 0 To CheckedListBox1.Items.Count - 1
                    CheckedListBox1.SelectedIndex = z
                    If CheckedListBox1.GetItemChecked(z) = True Then
                        If CheckedListBox1.Items(z).ToString.Contains("[") Then
                            Dim InstantSplitter() As String = Split(CheckedListBox1.Items(z).ToString, "  [")
                            TextBox3.Text += InstantSplitter(0).ToString & ", "
                            sezionedecretazione += InstantSplitter(0).ToString & ", "
                            contatoreDecret += 1
                            If contatoreDecret = 4 Then
                                TextBox3.Text += vbCrLf
                                sezionedecretazione += vbCrLf
                                contatoreDecret = 0
                            End If
                        Else
                            TextBox3.Text += CheckedListBox1.SelectedItem.ToString & ", "
                            sezionedecretazione += CheckedListBox1.SelectedItem.ToString & ", "
                            contatoreDecret += 1
                            If contatoreDecret = 4 Then
                                TextBox3.Text += vbCrLf
                                sezionedecretazione += vbCrLf
                                contatoreDecret = 0
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
            End Try
            Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
            If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                TextBox3.Text += vbCrLf & "Nr. Progressivo Email: " & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3")
                sezionedecretazione += vbCrLf & "Nr. Progressivo Email: " & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3")
            Else
                TextBox3.Text += vbCrLf & "Nr. Progressivo Email: " & My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3")
                sezionedecretazione += vbCrLf & "Nr. Progressivo Email: " & My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3")
            End If
            TextBox3.Text += vbCrLf & "Il: " & Date.Now.Day.ToString & "." & Date.Now.Month.ToString & "." & Date.Now.Year.ToString & "  Alle: " & ora
            sezionedecretazione += vbCrLf & "Il: " & Date.Now.Day.ToString & "." & Date.Now.Month.ToString & "." & Date.Now.Year.ToString & "  Alle: " & ora
            TextBox3.Text += vbCrLf & "CHIAVE: " & FileName
            sezionedecretazione += vbCrLf & "CHIAVE: " & FileName
        Catch ex As Exception

        End Try

        Try
            ListBox2.Items.Clear()
            If CheckedListBox1.Items.Count > 0 Then
                CheckedListBox1.SelectedIndex = 0

                Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
                If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                    Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                    Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                    riga = Split(TextBox3.Text, vbCrLf)
                    If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt") Then
                        scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                        scrivi.Write(My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & "#")
                        scrivi.Close()
                    Else
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                        scrivi.Write(My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & "#")
                        scrivi.Close()
                    End If

                    If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt") Then
                        scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt")
                        scrivi.Write(TextBox3.Text)
                        scrivi.Close()
                    Else
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt")
                        scrivi.Write(TextBox3.Text)
                        scrivi.Close()
                    End If

                    For r = 0 To CheckedListBox1.Items.Count - 1
                        CheckedListBox1.SelectedIndex = r
                        If CheckedListBox1.GetItemChecked(r) = True Then
                            If CheckedListBox1.Items(r).ToString.Contains("[") Then
                                Dim InstantSplitter() As String = Split(CheckedListBox1.Items(r).ToString, "  [")
                                scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                                scrivi.Write(InstantSplitter(0).ToString & "/")
                                scrivi.Close()
                            Else
                                scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                                scrivi.Write(CheckedListBox1.SelectedItem.ToString & "/")
                                scrivi.Close()
                            End If

                            leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                            letto = leggi.ReadToEnd
                            leggi.Close()
                            infoufficio = Split(letto, "#")
                            For z = 0 To UBound(infoufficio) - 2
                                If infoufficio(z + 1) <> "" Then
                                    ListBox2.Items.Add(infoufficio(z + 1))
                                End If
                            Next
                        End If
                    Next

                    StatusBar1.Text = "Stato Invio: Creazione dati database"

                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                    scrivi.Write("#" & MSGIDdef & "#" & datagiorno & " " & ora & "#" & chkAllegato & "#" & riga(0) & "#" & riga(1) & "#" & vbCrLf)
                    scrivi.Close()

                    If My.Settings.TipoDiDistribuzione = "Manuale" Then
                        For z = 0 To CheckedListBox1.Items.Count - 1
                            CheckedListBox1.SelectedIndex = z
                            If CheckedListBox1.GetItemChecked(z) = True Then
                                If CheckedListBox1.Items(z).ToString <> "" Then
                                    If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "P" & My.Settings.JlDate.ToString("D3")) Then
                                        If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "P" & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt") Then
                                        Else
                                            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "P" & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt")
                                            scrivi.Write(TextBox3.Text)
                                            scrivi.Close()
                                        End If
                                    Else
                                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "P" & My.Settings.JlDate.ToString("D3"))
                                        If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "P" & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt") Then
                                        Else
                                            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "P" & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt")
                                            scrivi.Write(TextBox3.Text)
                                            scrivi.Close()
                                        End If
                                    End If
                                End If
                            End If
                        Next

                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\TestoEmail.txt")
                        scrivi.Write(TextBox3.Text)
                        scrivi.Close()

                        MsgBox("Invio Completato, Email inviata a " & ListBox2.Items.Count & " indirizzi.", MsgBoxStyle.Information, "Asc Server - Invio Completato")
                        StatusBar1.Text = "Stato Invio: Invio completato"
                        SaveProgressiveNumbers()

                        OpenFileDialog1.FileName = ""
                        chkAllegato = "NO"

                        Dim MySplitter2() As String = Split(TextBox3.Text, vbCrLf)
                        If MySplitter2(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                            Or MySplitter2(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                            Or MySplitter2(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                            Try
                                System.IO.File.Delete(Application.StartupPath & "\" & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & ".txt")
                            Catch ex As Exception

                            End Try
                        Else
                            Try
                                System.IO.File.Delete(Application.StartupPath & "\" & My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3") & ".txt")
                            Catch ex As Exception

                            End Try
                        End If
                        If Timer2.Enabled = False Then
                            Timer2.Start()
                        End If
                        If My.Settings.SavePrintSett = 1 Then
                            Stampa()
                        Else
                            My.Settings.x = 9
                            PrintSett.Show()
                        End If
                        Exit Sub
                    End If

                Else 'Codice nel caso non contenga le due stringhe indicate

                    riga = Split(TextBox3.Text, vbCrLf)
                    If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt") Then
                        scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                        scrivi.Write(My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3") & "#")
                        scrivi.Close()
                    Else
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                        scrivi.Write(My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3") & "#")
                        scrivi.Close()
                    End If

                    If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt") Then
                        scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt")
                        scrivi.Write(TextBox3.Text)
                        scrivi.Close()
                    Else
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt")
                        scrivi.Write(TextBox3.Text)
                        scrivi.Close()
                    End If

                    For r = 0 To CheckedListBox1.Items.Count - 1
                        CheckedListBox1.SelectedIndex = r
                        If CheckedListBox1.GetItemChecked(r) = True Then
                            If CheckedListBox1.Items(r).ToString.Contains("[") Then
                                Dim InstantSplitter() As String = Split(CheckedListBox1.Items(r).ToString, "  [")
                                scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                                scrivi.Write(InstantSplitter(0).ToString & "/")
                                scrivi.Close()
                            Else
                                scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                                scrivi.Write(CheckedListBox1.SelectedItem.ToString & "/")
                                scrivi.Close()
                            End If

                            leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                            letto = leggi.ReadToEnd
                            leggi.Close()
                            infoufficio = Split(letto, "#")
                            For z = 0 To UBound(infoufficio) - 2
                                If infoufficio(z + 1) <> "" Then
                                    ListBox2.Items.Add(infoufficio(z + 1))
                                End If
                            Next
                        End If
                    Next

                    StatusBar1.Text = "Stato Invio: Creazione dati database"

                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                    scrivi.Write("#" & MSGIDdef & "#" & datagiorno & " " & ora & "#" & chkAllegato & "#" & riga(0) & "#" & riga(1) & "#" & IsPCNAvailable & "#" & vbCrLf)
                    scrivi.Close()

                    If My.Settings.TipoDiDistribuzione = "Manuale" Then
                        For z = 0 To CheckedListBox1.Items.Count - 1
                            CheckedListBox1.SelectedIndex = z
                            If CheckedListBox1.GetItemChecked(z) = True Then
                                If CheckedListBox1.Items(z).ToString <> "" Then
                                    If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "A" & My.Settings.JlDate.ToString("D3")) Then
                                        If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "A" & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt") Then
                                        Else
                                            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "A" & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt")
                                            scrivi.Write(TextBox3.Text)
                                            scrivi.Close()
                                        End If
                                    Else
                                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "A" & My.Settings.JlDate.ToString("D3"))
                                        If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "A" & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt") Then
                                        Else
                                            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "A" & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt")
                                            scrivi.Write(TextBox3.Text)
                                            scrivi.Close()
                                        End If
                                    End If
                                End If
                            End If
                        Next

                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\TestoEmail.txt")
                        scrivi.Write(TextBox3.Text)
                        scrivi.Close()

                        MsgBox("Invio Completato, Email inviata a " & ListBox2.Items.Count & " indirizzi.", MsgBoxStyle.Information, "Asc Server - Invio Completato")
                        StatusBar1.Text = "Stato Invio: Invio completato"
                        SaveProgressiveNumbers()

                        OpenFileDialog1.FileName = ""
                        chkAllegato = "NO"

                        Dim MySplitter2() As String = Split(TextBox3.Text, vbCrLf)
                        If MySplitter2(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                            Or MySplitter2(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                            Or MySplitter2(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                            Try
                                System.IO.File.Delete(Application.StartupPath & "\" & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & ".txt")
                            Catch ex As Exception

                            End Try
                        Else
                            Try
                                System.IO.File.Delete(Application.StartupPath & "\" & My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3") & ".txt")
                            Catch ex As Exception

                            End Try
                        End If
                        If Timer2.Enabled = False Then
                            Timer2.Start()
                        End If
                        If My.Settings.SavePrintSett = 1 Then
                            Stampa()
                        Else
                            My.Settings.x = 9
                            PrintSett.Show()
                        End If
                        Exit Sub
                    End If

                End If

                For r = 0 To ListBox2.Items.Count - 1
                    ListBox2.SetSelected(r, True)
                    If ListBox2.SelectedItem = "" Then
                        ListBox2.Items.RemoveAt(r)
                    End If
                Next

            End If
        Catch ex As Exception

        End Try

        Try
            StatusBar1.Text = "Stato Invio: Controllo allegati"
            Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
            If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\" & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & ".txt")
                scrivi.Write(TextBox3.Text)
                scrivi.Close()
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\" & My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3") & ".txt")
                scrivi.Write(TextBox3.Text)
                scrivi.Close()
            End If

            scrivi = System.IO.File.CreateText(Application.StartupPath & "\TestoEmail.txt")
            scrivi.Write(TextBox3.Text)
            scrivi.Close()
        Catch ex As Exception

        End Try

        If ListBox2.Items.Count > 0 Then
            TrattaDistribution()
        Else
            MsgBox("L'ufficio che avete selezionato sembra non contenere alcun contatto. Invio annullato.", MsgBoxStyle.Information, "Asc Server - Errore Invio")
            Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
            If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                My.Settings.NumEmail2 -= 1
            Else
                My.Settings.NumEmail -= 1
            End If
            Exit Sub
        End If

    End Sub

    Public Sub AddExer()

        Try
            StatusBar1.Text = "Stato Invio: Aggiunta EXER"
            Dim mysplitter() As String
            mysplitter = Split(TextBox3.Text, vbCrLf)
            TextBox3.Clear()
            For r = 0 To UBound(mysplitter)
                If r <> 1 Then
                    TextBox3.Text += mysplitter(r) & vbCrLf
                Else
                    TextBox3.Text += mysplitter(r) & "                EXER" & vbCrLf
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub

    Dim indirizzononclass As Boolean = False
    Dim listaindirizzinonclass As String = ""
    Public Sub TrattaDistribution()

        If InvioAnnullato = 1 Then
            Exit Sub
        End If

        Try
            StatusBar1.Text = "Stato Invio: Invio iniziato, smistamento"
            ListBox2.SetSelected(0, True)
            int2 = ListBox2.Items.Count

            If System.IO.File.Exists(Application.StartupPath & "\Uffici\indirizziselezionati.txt") Then
                Dim leggiora As System.IO.StreamReader = System.IO.File.OpenText(Application.StartupPath & "\Uffici\indirizziselezionati.txt")
                listaindirizzinonclass = leggiora.ReadToEnd
                leggiora.Close()
            End If

            For r = 0 To int2 - 1
                ListBox2.SetSelected(r, True)
                If listaindirizzinonclass.Contains(ListBox2.SelectedItem.ToString) Then
                    indirizzononclass = True
                Else
                    indirizzononclass = False
                End If
                If ListBox2.SelectedItem.ToString.Contains("@") Then
                    SendEmailToEmailAddress()
                Else
                    SendEmailToMachineAddress()
                End If

                If r = int2 - 1 Then
                    MsgBox("Invio Completato, Email inviata a " & ListBox2.Items.Count & " indirizzi.", MsgBoxStyle.Information, "Asc Server - Invio Completato")
                    StatusBar1.Text = "Stato Invio: Invio completato"
                    SaveProgressiveNumbers()

                    OpenFileDialog1.FileName = ""
                    chkAllegato = "NO"

                    Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
                    If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                        Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                        Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                        Try
                            System.IO.File.Delete(Application.StartupPath & "\" & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & ".txt")
                        Catch ex As Exception

                        End Try
                    Else
                        Try
                            System.IO.File.Delete(Application.StartupPath & "\" & My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3") & ".txt")
                        Catch ex As Exception

                        End Try
                    End If
                    If Timer2.Enabled = False Then
                        Timer2.Start()
                    End If
                    If My.Settings.SavePrintSett = 1 Then
                        Stampa()
                    Else
                        My.Settings.x = 9
                        PrintSett.Show()
                    End If
                End If

            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Sub SendEmailToEmailAddress()

        Try
            If TextBox3.Text.Contains("R I S E R V A T I S S I M O") Or TextBox3.Text.Contains("S E G R E T O") Or TextBox3.Text.Contains("S E G R E T I S S I M O") _
            Or TextBox3.Text.Contains("C O N F I D E N T I A L") Or TextBox3.Text.Contains("S E C R E T") Or TextBox3.Text.Contains("T O P S E C R E T") _
            Or TextBox3.Text.Contains("N A T O C O N F I D E N T I A L") Or TextBox3.Text.Contains("N A T O S E C R E T") Or TextBox3.Text.Contains("N A T O T O P S E C R E T") _
            Or TextBox3.Text.Contains("R I S E R V A T O") Or TextBox3.Text.Contains("R E S T R I C T E D") Or TextBox3.Text.Contains("N A T O R E S T R I C T E D") Then
                If indirizzononclass = True Then
                    Dim MySplitter2() As String = Split(TextBox3.Text, vbCrLf)
                    Dim TestoMessaggio As String = ""
                    For x = 0 To UBound(MySplitter2)
                        If MySplitter2(x).ToString <> "BT" Then
                            TestoMessaggio += MySplitter2(x).ToString & vbCrLf
                        Else
                            TestoMessaggio += MySplitter2(x).ToString & vbCrLf
                            TestoMessaggio += "-------Trattasi di messaggio classificato, recarsi in Segreteria Comando per la visualizzazione-------" & vbCrLf & "BT" & vbCrLf & sezionedecretazione
                            Exit For
                        End If
                    Next

                    Dim objOutlook As Microsoft.Office.Interop.Outlook.Application
                    Dim objEmail As Microsoft.Office.Interop.Outlook.MailItem
                    objOutlook = CType(CreateObject("Outlook.Application"), Microsoft.Office.Interop.Outlook.Application)
                    objEmail = objOutlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
                    With objEmail
                        .Subject = TextBox1.Text
                        .To = ListBox2.SelectedItem.ToString
                        .Body = TestoMessaggio
                        Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
                        If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                            Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                            Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                            .Attachments.Add(Application.StartupPath & "\" & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & ".txt")
                        Else
                            .Attachments.Add(Application.StartupPath & "\" & My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3") & ".txt")
                        End If
                        If OpenFileDialog1.FileName <> "" Then
                            .Attachments.Add(OpenFileDialog1.FileName)
                        End If
                        .Send()
                        indirizzononclass = False
                    End With
                Else
                    Dim objOutlook As Microsoft.Office.Interop.Outlook.Application
                    Dim objEmail As Microsoft.Office.Interop.Outlook.MailItem
                    objOutlook = CType(CreateObject("Outlook.Application"), Microsoft.Office.Interop.Outlook.Application)
                    objEmail = objOutlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
                    With objEmail
                        .Subject = TextBox1.Text
                        .To = ListBox2.SelectedItem.ToString
                        .Body = TextBox3.Text
                        Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
                        If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                            Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                            Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                            .Attachments.Add(Application.StartupPath & "\" & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & ".txt")
                        Else
                            .Attachments.Add(Application.StartupPath & "\" & My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3") & ".txt")
                        End If
                        If OpenFileDialog1.FileName <> "" Then
                            .Attachments.Add(OpenFileDialog1.FileName)
                        End If
                        .Send()
                    End With
                End If
            Else
                Dim objOutlook As Microsoft.Office.Interop.Outlook.Application
                Dim objEmail As Microsoft.Office.Interop.Outlook.MailItem
                objOutlook = CType(CreateObject("Outlook.Application"), Microsoft.Office.Interop.Outlook.Application)
                objEmail = objOutlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
                With objEmail
                    .Subject = TextBox1.Text
                    .To = ListBox2.SelectedItem.ToString
                    .Body = TextBox3.Text
                    Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
                    If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                        Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                        Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                        .Attachments.Add(Application.StartupPath & "\" & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & ".txt")
                    Else
                        .Attachments.Add(Application.StartupPath & "\" & My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3") & ".txt")
                    End If
                    If OpenFileDialog1.FileName <> "" Then
                        .Attachments.Add(OpenFileDialog1.FileName)
                    End If
                    .Send()
                End With
            End If
        Catch ex As Exception
            MsgBox("Errore invio: " & ex.Message, MsgBoxStyle.Critical, "Asc Server - Errore Invio")
        End Try

    End Sub

    Public Sub SendEmailToMachineAddress()

        Try
            Dim scrittore As System.IO.StreamWriter
            If TextBox3.Text.Contains("R I S E R V A T I S S I M O") Or TextBox3.Text.Contains("S E G R E T O") Or TextBox3.Text.Contains("S E G R E T I S S I M O") _
            Or TextBox3.Text.Contains("C O N F I D E N T I A L") Or TextBox3.Text.Contains("S E C R E T") Or TextBox3.Text.Contains("T O P S E C R E T") _
            Or TextBox3.Text.Contains("N A T O C O N F I D E N T I A L") Or TextBox3.Text.Contains("N A T O S E C R E T") Or TextBox3.Text.Contains("N A T O T O P S E C R E T") _
            Or TextBox3.Text.Contains("R I S E R V A T O") Or TextBox3.Text.Contains("R E S T R I C T E D") Or TextBox3.Text.Contains("N A T O R E S T R I C T E D") Then
                If indirizzononclass = True Then
                    Dim MySplitter2() As String = Split(TextBox3.Text, vbCrLf)
                    Dim TestoMessaggio As String = ""
                    For x = 0 To UBound(MySplitter2)
                        If MySplitter2(x).ToString <> "BT" Then
                            TestoMessaggio += MySplitter2(x).ToString & vbCrLf
                        Else
                            TestoMessaggio += MySplitter2(x).ToString & vbCrLf
                            TestoMessaggio += "-------Trattasi di messaggio classificato, recarsi in Segreteria Comando per la visualizzazione-------" & vbCrLf & "BT" & vbCrLf & sezionedecretazione
                            Exit For
                        End If
                    Next
                    If System.IO.Directory.Exists(Application.StartupPath & "\temppartenze") Then
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\temppartenze")
                    End If
                    If System.IO.File.Exists(Application.StartupPath & "\temppartenze\temppartenza.txt") Then
                        For r = 1 To 9999
                            If System.IO.File.Exists(Application.StartupPath & "\temppartenze\temppartenza(" & r.ToString & ").txt") Then
                            Else
                                scrittore = System.IO.File.CreateText(Application.StartupPath & "\temppartenze\temppartenza(" & r.ToString & ").txt")
                                scrittore.Write("MessaggioDalServer" & vbCrLf)
                                scrittore.Write(ListBox2.SelectedItem.ToString & vbCrLf)
                                scrittore.Write(FileName & vbCrLf)
                                scrittore.Write(TestoMessaggio)
                                scrittore.Close()
                                Exit For
                            End If
                        Next
                    Else
                        scrittore = System.IO.File.CreateText(Application.StartupPath & "\temppartenze\temppartenza.txt")
                        scrittore.Write("MessaggioDalServer" & vbCrLf)
                        scrittore.Write(ListBox2.SelectedItem.ToString & vbCrLf)
                        scrittore.Write(FileName & vbCrLf)
                        scrittore.Write(TestoMessaggio)
                        scrittore.Close()
                    End If
                    indirizzononclass = False
                Else
                    If System.IO.Directory.Exists(Application.StartupPath & "\temppartenze") Then
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\temppartenze")
                    End If
                    If System.IO.File.Exists(Application.StartupPath & "\temppartenze\temppartenza.txt") Then
                        For r = 1 To 9999
                            If System.IO.File.Exists(Application.StartupPath & "\temppartenze\temppartenza(" & r.ToString & ").txt") Then
                            Else
                                scrittore = System.IO.File.CreateText(Application.StartupPath & "\temppartenze\temppartenza(" & r.ToString & ").txt")
                                scrittore.Write("MessaggioDalServer" & vbCrLf)
                                scrittore.Write(ListBox2.SelectedItem.ToString & vbCrLf)
                                scrittore.Write(FileName & vbCrLf)
                                scrittore.Write(TextBox3.Text)
                                scrittore.Close()
                                Exit For
                            End If
                        Next
                    Else
                        scrittore = System.IO.File.CreateText(Application.StartupPath & "\temppartenze\temppartenza.txt")
                        scrittore.Write("MessaggioDalServer" & vbCrLf)
                        scrittore.Write(ListBox2.SelectedItem.ToString & vbCrLf)
                        scrittore.Write(FileName & vbCrLf)
                        scrittore.Write(TextBox3.Text)
                        scrittore.Close()
                    End If
                End If
            Else
                If System.IO.Directory.Exists(Application.StartupPath & "\temppartenze") Then
                Else
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\temppartenze")
                End If
                If System.IO.File.Exists(Application.StartupPath & "\temppartenze\temppartenza.txt") Then
                    For r = 1 To 9999
                        If System.IO.File.Exists(Application.StartupPath & "\temppartenze\temppartenza(" & r.ToString & ").txt") Then
                        Else
                            scrittore = System.IO.File.CreateText(Application.StartupPath & "\temppartenze\temppartenza(" & r.ToString & ").txt")
                            scrittore.Write("MessaggioDalServer" & vbCrLf)
                            scrittore.Write(ListBox2.SelectedItem.ToString & vbCrLf)
                            scrittore.Write(FileName & vbCrLf)
                            scrittore.Write(TextBox3.Text)
                            scrittore.Close()
                            Exit For
                        End If
                    Next
                Else
                    scrittore = System.IO.File.CreateText(Application.StartupPath & "\temppartenze\temppartenza.txt")
                    scrittore.Write("MessaggioDalServer" & vbCrLf)
                    scrittore.Write(ListBox2.SelectedItem.ToString & vbCrLf)
                    scrittore.Write(FileName & vbCrLf)
                    scrittore.Write(TextBox3.Text)
                    scrittore.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox("Errore invio email, messaggio ex: " & ex.Message, MsgBoxStyle.Critical, "Asc Server - Errore Invio")
        End Try

    End Sub

    Public Sub ChkServerFilesUsed()

        Try
            StatusBar1.Text = "Stato Invio: Controllo utilizzo file ricevuto"
            Dim lstview1selectedindex As Integer
            Dim PositiveRow As Integer = 0
            Dim nomedasplittare As String = ""
            If ListView1LastSelectedItem <> "NONSELEZIONATO" Then
                lstview1selectedindex = Val(ListView1LastSelectedItem)
                nomedasplittare = ListView1.Items(lstview1selectedindex).Text.ToString & ".txt"
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Messaggi Ricevuti\" & nomedasplittare)
                letto = leggi.ReadToEnd
                leggi.Close()
                Dim TxtFileSplitter() As String = Split(letto, vbCrLf)
                For r = 0 To UBound(TxtFileSplitter)
                    If TextBox3.Text.Contains(TxtFileSplitter(r).ToString) Then
                        PositiveRow += 1
                    End If
                Next
            End If

            Dim TextBox3RowSplitter() As String = Split(TextBox3.Text, vbCrLf)
            If PositiveRow > TextBox3RowSplitter.Count - 3 Then
                ListView1.Items(lstview1selectedindex).ForeColor = Color.Green
                If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\UsedIndicies.txt") Then
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Messaggi Ricevuti\UsedIndicies.txt")
                    scrivi.Write(nomedasplittare & "#")
                    scrivi.Close()
                Else
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\UsedIndicies.txt")
                    scrivi.Write(nomedasplittare & "#")
                    scrivi.Close()
                End If
            End If
        Catch ex As Exception

        End Try

        TrovaMSGID()

    End Sub

    Public Sub ChkCharsErrorsPresence()

        Try
            StatusBar1.Text = "Stato Invio: Controllo ortografico"
            Dim SelectionPointStart As Integer = 0
            Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
            For r = 0 To UBound(MySplitter)
                If MySplitter(r).ToString.Length > 69 Then
                    MsgBox("Riga " & (r + 1).ToString & " non conforme allo standard di compilazione, prego correggere.", MsgBoxStyle.Information, "Asc Server - Errore Compilazione")
                    TextBox3.SelectionStart = TextBox3.Text.IndexOf(MySplitter(r).ToString)
                    TextBox3.SelectionLength = MySplitter(r).ToString.Length
                    TextBox3.HideSelection = False
                    Exit Sub
                End If
            Next

            ChkServerFilesUsed()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub RemoveProbableTemps()

        Try
            StatusBar1.Text = "Stato Invio: Rimozione files temporanei"
            For r = 0 To NumFoglioStampa
                System.IO.File.Delete(Application.StartupPath & "\fogliostampa" & r & ".txt")
                NumFoglioStampa = 0
                numeroriga = 0
                stoppami = 0
            Next
        Catch ex As Exception

        End Try

        Try
            System.IO.File.Delete(Application.StartupPath & "\" & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & ".txt")
            System.IO.File.Delete(Application.StartupPath & "\" & My.Settings.NumEmail.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & ".txt")
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ChkLastBT()

        Try
            If IsFirstChkLastBT = 0 Then
                IsFirstChkLastBT = 1
                If TextBox3.Lines(TextBox3.Lines.Count - 2).ToString = "BT" Then
                    ChkCharsErrorsPresence()
                Else
                    MsgBox("Errore battitura rilevato, il messaggio deve terminare con la stringa " & """" & "BT" & """" & ". Invio annullato.", MsgBoxStyle.Information, "Asc Server- Errore Battitura BT")
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
                If TextBox3.Text.Contains("PDC:") Or TextBox3.Text.Contains("POC:") Or TextBox3.Text.Contains("P.D.C.:") Or TextBox3.Text.Contains("P.O.C.:") Then
                    ChkLastBT()
                Else
                    Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
                    If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                        Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                        Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
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
                StatusBar1.Text = "Stato Invio: Controllo classifica di riservatezza"
                Dim MyTxtBoxSpliiter() As String = Split(TextBox3.Text, vbCrLf)
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
                                                MsgBox("Non è stata rilevata alcuna riga SIC, impossibile continuare.", MsgBoxStyle.Critical, "Asc Server - Sic Mancante")
                                                Exit Sub
                                            End If
                                        ElseIf MyTxtBoxSpliiter(r + 1).ToString.Contains("UNCLASSIFIED") Or _
                                        MyTxtBoxSpliiter(r + 1).ToString.Contains("UNCLASS") Or _
                                        MyTxtBoxSpliiter(r + 1).ToString.Contains("NATO UNCLASS") Or _
                                        MyTxtBoxSpliiter(r + 1).ToString.Contains("UNCLAS") Or _
                                        MyTxtBoxSpliiter(r + 1).ToString.Contains("NATO UNCLAS") Or _
                                        MyTxtBoxSpliiter(r + 1).ToString.Contains("NATO UNCLASSIFIED") Or _
                                        MyTxtBoxSpliiter(r + 1).ToString.Contains("NON CLASSIFICATO") Then
                                            If MyTxtBoxSpliiter(r + 2).ToString.Contains("SIC") Then
                                                ControlloPOC()
                                                Exit Sub
                                            Else
                                                MsgBox("Non è stata rilevata alcuna riga SIC, impossibile continuare.", MsgBoxStyle.Critical, "Asc Server - Sic Mancante")
                                                Exit Sub
                                            End If
                                        Else
                                            MsgBox("Classifica di segretezza mancante. Impossibile continuare.", MsgBoxStyle.Critical, "Asc Server - Classifica Segretezza Mancante")
                                            Exit Sub
                                        End If
                                    ElseIf BTcounter = 1 Then
                                        If ClassificaturaPresente = 1 Then
                                            If MyTxtBoxSpliiter(r - 1).ToString.Contains("DISTRUGGERE") Or _
                                                MyTxtBoxSpliiter(r - 1).ToString.Contains("DESTROY") Or _
                                                MyTxtBoxSpliiter(r - 1).ToString.Contains("DISTRUZIONE") Then
                                                ControlloPOC()
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
                ControlloPOC()
            End If
        Catch ex As Exception

        End Try

    End Sub
    Public Sub SelectBTasLastLine()

        Try
            If IsFirstSelectBTasLastLine = 0 Then
                IsFirstSelectBTasLastLine = 1
                StatusBar1.Text = "Stato Invio: Controllo ultimo BT"
                Dim BtCounter As Integer = 0
                Dim Textbox3Splitter() As String = Split(TextBox3.Text, vbCrLf)
                TextBox3.Clear()
                For r = 0 To UBound(Textbox3Splitter)
                    If Textbox3Splitter(r).ToString.Contains("BT") Then
                        Dim MySplitter() As String = Split(Textbox3Splitter(r).ToString & " ", " ")
                        For z = 0 To UBound(MySplitter)
                            If MySplitter(z).ToString = "" Or MySplitter(z).ToString = " " Then
                                If z = MySplitter.Count - 1 Then
                                    If BtCounter = 0 Then
                                        TextBox3.Text += Textbox3Splitter(r).ToString & vbCrLf
                                        BtCounter += 1
                                    Else
                                        TextBox3.Text += Textbox3Splitter(r).ToString & vbCrLf
                                        ChkBtIsRegular()
                                        Exit Sub
                                    End If
                                End If
                            Else
                                If MySplitter(z).ToString <> "BT" Then
                                    TextBox3.Text += Textbox3Splitter(r).ToString & vbCrLf
                                    Exit For
                                End If
                            End If
                        Next
                    Else
                        TextBox3.Text += Textbox3Splitter(r).ToString & vbCrLf
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
                StatusBar1.Text = "Stato Invio: Controllo BT"
                Dim BtCounter As Integer = 0
                Dim Textbox3Splitter() As String = Split(TextBox3.Text, vbCrLf)
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
                    MsgBox("Conteggio BT non conforme, numero BT presenti non conforme.", MsgBoxStyle.Information, "Asc Server - BT Non Conforme")
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ChkFollowingSICline()

        Try
            If IsFirstChkFollowingSICline = 0 Then
                IsFirstChkFollowingSICline = 1
                StatusBar1.Text = "Stato Invio: Controllo riga SIC"
                Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
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
                Dim MySplitter2() As String = Split(TextBox3.Text, vbCrLf)
                If MySplitter2(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                    Or MySplitter2(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                    Or MySplitter2(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                    Dim mysplitter() As String = Split(TextBox3.Text, vbCrLf)
                    Dim mysplittertostring As String = mysplitter(0).ToString
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
                                Dim AlertMessage = MsgBox("Data GDO non conforme: giorno superiore al corrente, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Server- Data GDO Errata")
                                If AlertMessage = 6 Then
                                    If Date.Now.Month = 1 Then
                                        If MyGDOsplitter(3).ToString = "GEN" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "GEN" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 2 Then
                                        If MyGDOsplitter(3).ToString = "FEB" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "FEB" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 3 Then
                                        If MyGDOsplitter(3).ToString = "MAR" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "MAR" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 4 Then
                                        If MyGDOsplitter(3).ToString = "APR" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "APR" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 5 Then
                                        If MyGDOsplitter(3).ToString = "MAG" Or MyGDOsplitter(3).ToString = "MAY" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "MAG" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 6 Then
                                        If MyGDOsplitter(3).ToString = "GIU" Or MyGDOsplitter(3).ToString = "JUN" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "GIU" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 7 Then
                                        If MyGDOsplitter(3).ToString = "LUG" Or MyGDOsplitter(3).ToString = "JUL" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "LUG" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 8 Then
                                        If MyGDOsplitter(3).ToString = "AGO" Or MyGDOsplitter(3).ToString = "AUG" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "AGO" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 9 Then
                                        If MyGDOsplitter(3).ToString = "SET" Or MyGDOsplitter(3).ToString = "SEP" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "SET" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 10 Then
                                        If MyGDOsplitter(3).ToString = "OTT" Or MyGDOsplitter(3).ToString = "OCT" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "OTT" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 11 Then
                                        If MyGDOsplitter(3).ToString = "NOV" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "NOV" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 12 Then
                                        If MyGDOsplitter(3).ToString = "DIC" Or MyGDOsplitter(3).ToString = "DEC" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "DIC" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
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
                                Dim AlertMessage = MsgBox("Data GDO non conforme: giorno inferiore al corrente, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Server- Data GDO Errata")
                                If AlertMessage = 6 Then
                                    If Date.Now.Month = 1 Then
                                        If MyGDOsplitter(3).ToString = "GEN" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "GEN" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 2 Then
                                        If MyGDOsplitter(3).ToString = "FEB" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "FEB" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 3 Then
                                        If MyGDOsplitter(3).ToString = "MAR" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "MAR" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 4 Then
                                        If MyGDOsplitter(3).ToString = "APR" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "APR" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 5 Then
                                        If MyGDOsplitter(3).ToString = "MAG" Or MyGDOsplitter(3).ToString = "MAY" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "MAG" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 6 Then
                                        If MyGDOsplitter(3).ToString = "GIU" Or MyGDOsplitter(3).ToString = "JUN" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "GIU" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 7 Then
                                        If MyGDOsplitter(3).ToString = "LUG" Or MyGDOsplitter(3).ToString = "JUL" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "LUG" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 8 Then
                                        If MyGDOsplitter(3).ToString = "AGO" Or MyGDOsplitter(3).ToString = "AUG" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "AGO" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 9 Then
                                        If MyGDOsplitter(3).ToString = "SET" Or MyGDOsplitter(3).ToString = "SEP" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "SET" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 10 Then
                                        If MyGDOsplitter(3).ToString = "OTT" Or MyGDOsplitter(3).ToString = "OCT" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "OTT" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 11 Then
                                        If MyGDOsplitter(3).ToString = "NOV" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "NOV" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 12 Then
                                        If MyGDOsplitter(3).ToString = "DIC" Or MyGDOsplitter(3).ToString = "DEC" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "DIC" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
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
                                        Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "GEN" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                        If AlertMessage1 = 6 Then
                                            ChkFollowingSICline()
                                        Else
                                            Exit Sub
                                        End If
                                    End If
                                ElseIf Date.Now.Month = 2 Then
                                    If MyGDOsplitter(3).ToString = "FEB" Then
                                        ChkFollowingSICline()
                                    Else
                                        Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "FEB" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                        If AlertMessage1 = 6 Then
                                            ChkFollowingSICline()
                                        Else
                                            Exit Sub
                                        End If
                                    End If
                                ElseIf Date.Now.Month = 3 Then
                                    If MyGDOsplitter(3).ToString = "MAR" Then
                                        ChkFollowingSICline()
                                    Else
                                        Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "MAR" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                        If AlertMessage1 = 6 Then
                                            ChkFollowingSICline()
                                        Else
                                            Exit Sub
                                        End If
                                    End If
                                ElseIf Date.Now.Month = 4 Then
                                    If MyGDOsplitter(3).ToString = "APR" Then
                                        ChkFollowingSICline()
                                    Else
                                        Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "APR" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                        If AlertMessage1 = 6 Then
                                            ChkFollowingSICline()
                                        Else
                                            Exit Sub
                                        End If
                                    End If
                                ElseIf Date.Now.Month = 5 Then
                                    If MyGDOsplitter(3).ToString = "MAG" Or MyGDOsplitter(3).ToString = "MAY" Then
                                        ChkFollowingSICline()
                                    Else
                                        Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "MAG" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                        If AlertMessage1 = 6 Then
                                            ChkFollowingSICline()
                                        Else
                                            Exit Sub
                                        End If
                                    End If
                                ElseIf Date.Now.Month = 6 Then
                                    If MyGDOsplitter(3).ToString = "GIU" Or MyGDOsplitter(3).ToString = "JUN" Then
                                        ChkFollowingSICline()
                                    Else
                                        Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "GIU" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                        If AlertMessage1 = 6 Then
                                            ChkFollowingSICline()
                                        Else
                                            Exit Sub
                                        End If
                                    End If
                                ElseIf Date.Now.Month = 7 Then
                                    If MyGDOsplitter(3).ToString = "LUG" Or MyGDOsplitter(3).ToString = "JUL" Then
                                        ChkFollowingSICline()
                                    Else
                                        Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "LUG" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                        If AlertMessage1 = 6 Then
                                            ChkFollowingSICline()
                                        Else
                                            Exit Sub
                                        End If
                                    End If
                                ElseIf Date.Now.Month = 8 Then
                                    If MyGDOsplitter(3).ToString = "AGO" Or MyGDOsplitter(3).ToString = "AUG" Then
                                        ChkFollowingSICline()
                                    Else
                                        Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "AGO" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                        If AlertMessage1 = 6 Then
                                            ChkFollowingSICline()
                                        Else
                                            Exit Sub
                                        End If
                                    End If
                                ElseIf Date.Now.Month = 9 Then
                                    If MyGDOsplitter(3).ToString = "SET" Or MyGDOsplitter(3).ToString = "SEP" Then
                                        ChkFollowingSICline()
                                    Else
                                        Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "SET" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                        If AlertMessage1 = 6 Then
                                            ChkFollowingSICline()
                                        Else
                                            Exit Sub
                                        End If
                                    End If
                                ElseIf Date.Now.Month = 10 Then
                                    If MyGDOsplitter(3).ToString = "OTT" Or MyGDOsplitter(3).ToString = "OCT" Then
                                        ChkFollowingSICline()
                                    Else
                                        Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "OTT" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                        If AlertMessage1 = 6 Then
                                            ChkFollowingSICline()
                                        Else
                                            Exit Sub
                                        End If
                                    End If
                                ElseIf Date.Now.Month = 11 Then
                                    If MyGDOsplitter(3).ToString = "NOV" Then
                                        ChkFollowingSICline()
                                    Else
                                        Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "NOV" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                        If AlertMessage1 = 6 Then
                                            ChkFollowingSICline()
                                        Else
                                            Exit Sub
                                        End If
                                    End If
                                ElseIf Date.Now.Month = 12 Then
                                    If MyGDOsplitter(3).ToString = "DIC" Or MyGDOsplitter(3).ToString = "DEC" Then
                                        ChkFollowingSICline()
                                    Else
                                        Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "DIC" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                        If AlertMessage1 = 6 Then
                                            ChkFollowingSICline()
                                        Else
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            MsgBox("GDO non conforme. Invio annullato.", MsgBoxStyle.Information, "Asc Server- GDO Errato")
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
                                    Dim AlertMessage = MsgBox("Data GDO non conforme: giorno superiore al corrente, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Server- Data GDO Errata")
                                    If AlertMessage = 6 Then
                                        If Date.Now.Month = 1 Then
                                            If MyGDOsplitter(2).ToString = "GEN" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "GEN" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 2 Then
                                            If MyGDOsplitter(2).ToString = "FEB" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "FEB" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 3 Then
                                            If MyGDOsplitter(2).ToString = "MAR" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "MAR" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 4 Then
                                            If MyGDOsplitter(2).ToString = "APR" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "APR" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 5 Then
                                            If MyGDOsplitter(2).ToString = "MAG" Or MyGDOsplitter(2).ToString = "MAY" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "MAG" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 6 Then
                                            If MyGDOsplitter(2).ToString = "GIU" Or MyGDOsplitter(2).ToString = "JUN" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "GIU" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 7 Then
                                            If MyGDOsplitter(2).ToString = "LUG" Or MyGDOsplitter(2).ToString = "JUL" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "LUG" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 8 Then
                                            If MyGDOsplitter(2).ToString = "AGO" Or MyGDOsplitter(2).ToString = "AUG" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "AGO" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 9 Then
                                            If MyGDOsplitter(2).ToString = "SET" Or MyGDOsplitter(2).ToString = "SEP" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "SET" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 10 Then
                                            If MyGDOsplitter(2).ToString = "OTT" Or MyGDOsplitter(2).ToString = "OCT" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "OTT" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 11 Then
                                            If MyGDOsplitter(2).ToString = "NOV" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "NOV" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 12 Then
                                            If MyGDOsplitter(2).ToString = "DIC" Or MyGDOsplitter(2).ToString = "DEC" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "DIC" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
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
                                    Dim AlertMessage = MsgBox("Data GDO non conforme: giorno inferiore al corrente, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Server- Data GDO Errata")
                                    If AlertMessage = 6 Then
                                        If Date.Now.Month = 1 Then
                                            If MyGDOsplitter(2).ToString = "GEN" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "GEN" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 2 Then
                                            If MyGDOsplitter(2).ToString = "FEB" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "FEB" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 3 Then
                                            If MyGDOsplitter(2).ToString = "MAR" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "MAR" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 4 Then
                                            If MyGDOsplitter(2).ToString = "APR" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "APR" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 5 Then
                                            If MyGDOsplitter(2).ToString = "MAG" Or MyGDOsplitter(2).ToString = "MAY" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "MAG" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 6 Then
                                            If MyGDOsplitter(2).ToString = "GIU" Or MyGDOsplitter(2).ToString = "JUN" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "GIU" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 7 Then
                                            If MyGDOsplitter(2).ToString = "LUG" Or MyGDOsplitter(2).ToString = "JUL" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "LUG" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 8 Then
                                            If MyGDOsplitter(2).ToString = "AGO" Or MyGDOsplitter(2).ToString = "AUG" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "AGO" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 9 Then
                                            If MyGDOsplitter(2).ToString = "SET" Or MyGDOsplitter(2).ToString = "SEP" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "SET" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 10 Then
                                            If MyGDOsplitter(2).ToString = "OTT" Or MyGDOsplitter(2).ToString = "OCT" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "OTT" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 11 Then
                                            If MyGDOsplitter(2).ToString = "NOV" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "NOV" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf Date.Now.Month = 12 Then
                                            If MyGDOsplitter(2).ToString = "DIC" Or MyGDOsplitter(2).ToString = "DEC" Then
                                                ChkFollowingSICline()
                                            Else
                                                Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "DIC" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                                If AlertMessage1 = 6 Then
                                                    ChkFollowingSICline()
                                                Else
                                                    Exit Sub
                                                End If
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
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "GEN" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 2 Then
                                        If MyGDOsplitter(2).ToString = "FEB" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "FEB" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 3 Then
                                        If MyGDOsplitter(2).ToString = "MAR" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "MAR" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 4 Then
                                        If MyGDOsplitter(2).ToString = "APR" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "APR" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 5 Then
                                        If MyGDOsplitter(2).ToString = "MAG" Or MyGDOsplitter(2).ToString = "MAY" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "MAG" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 6 Then
                                        If MyGDOsplitter(2).ToString = "GIU" Or MyGDOsplitter(2).ToString = "JUN" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "GIU" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 7 Then
                                        If MyGDOsplitter(2).ToString = "LUG" Or MyGDOsplitter(2).ToString = "JUL" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "LUG" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 8 Then
                                        If MyGDOsplitter(2).ToString = "AGO" Or MyGDOsplitter(2).ToString = "AUG" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "AGO" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 9 Then
                                        If MyGDOsplitter(2).ToString = "SET" Or MyGDOsplitter(2).ToString = "SEP" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "SET" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 10 Then
                                        If MyGDOsplitter(2).ToString = "OTT" Or MyGDOsplitter(2).ToString = "OCT" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "OTT" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 11 Then
                                        If MyGDOsplitter(2).ToString = "NOV" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "NOV" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf Date.Now.Month = 12 Then
                                        If MyGDOsplitter(2).ToString = "DIC" Or MyGDOsplitter(2).ToString = "DEC" Then
                                            ChkFollowingSICline()
                                        Else
                                            Dim AlertMessage1 = MsgBox("Mese del GDO non conforme con quello corrente, il mese corretto è " & """" & "DIC" & """" & ". Si desidera continuare comunque?.", MsgBoxStyle.YesNo, "Asc Server- Mese GDO Non Conforme")
                                            If AlertMessage1 = 6 Then
                                                ChkFollowingSICline()
                                            Else
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                MsgBox("GDO non conforme. Invio annullato.", MsgBoxStyle.Information, "Asc Server- GDO Errato")
                                Exit Sub
                            End If
                        Else
                            MsgBox("Qualifica di precedenza GDO errata. Invio annullato.", MsgBoxStyle.Information, "Asc Server- Qualifica Errata")
                            Exit Sub
                        End If
                    End If
                Else
                    ChkFollowingSICline()
                End If
            End If
        Catch ex As Exception
            MsgBox("GDO non conforme. Invio annullato. Errore segnalato: " & ex.Message.ToArray, MsgBoxStyle.Information, "Asc Server- GDO Errato")
            Exit Sub
        End Try

    End Sub

    Public Sub ChkFmIsRegular()

        Try
            If IsFirstChkFmIsRegular = 0 Then
                IsFirstChkFmIsRegular = 1
                StatusBar1.Text = "Stato Invio: Controllo riga FM"
                Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
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
                    Dim RigaTrovata As String = MySplitter(1).ToString
                    Dim NomeUnitaConfronto As String = My.Settings.NomeUnita
                    Dim Counter As Integer = 0
                    If RigaTrovata.Contains("FM ITS ") Then
                        For r = 0 To RigaTrovata.Length - 8
                            If r < RigaTrovata.Length And r < NomeUnitaConfronto.Length Then
                                If RigaTrovata(r + 7) = NomeUnitaConfronto(r) Then
                                    Counter += 1
                                End If
                            End If
                        Next
                    ElseIf RigaTrovata.Contains("FM NAVE ") Then
                        For r = 0 To RigaTrovata.Length - 9
                            If r < RigaTrovata.Length And r < NomeUnitaConfronto.Length Then
                                If RigaTrovata(r + 8) = NomeUnitaConfronto(r) Then
                                    Counter += 1
                                End If
                            End If
                        Next
                    ElseIf RigaTrovata.Contains("FM SMG ") Then
                        For r = 0 To RigaTrovata.Length - 8
                            If r < RigaTrovata.Length And r < NomeUnitaConfronto.Length Then
                                If RigaTrovata(r + 7) = NomeUnitaConfronto(r) Then
                                    Counter += 1
                                End If
                            End If
                        Next
                    End If
                    If Counter >= 4 Then
                        Dim AlertMessage = MsgBox("Durante il controllo ortografico è stata rilevata una possibilità di errore nella riga mittente, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Server - Possibile Errore")
                        If AlertMessage = 6 Then
                            ChkGDOIsRegular()
                        End If
                    Else
                        ChkGDOIsRegular()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            If My.Settings.TipoDiDistribuzione = "Automatica" Then
                If MenuItem42.Text = "Connetti" Then
                    If CheckedListBox1.SelectedItems.Count > 0 Then
                        For r = 0 To CheckedListBox1.Items.Count - 1
                            CheckedListBox1.SelectedIndex = r
                            If CheckedListBox1.GetItemChecked(r) = True Then
                                leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                                letto = leggi.ReadToEnd
                                leggi.Close()
                                infoufficio = Split(letto, "#")
                                For z = 0 To UBound(infoufficio) - 2
                                    If infoufficio(z + 1) <> "" Then
                                        ListBox2.Items.Add(infoufficio(z + 1))
                                    End If
                                Next
                            End If
                        Next

                        int2 = ListBox2.Items.Count
                        For r = 0 To int2 - 1
                            ListBox2.SetSelected(r, True)
                            If ListBox2.SelectedItem.ToString.Contains("@") Then
                                'Next
                            Else
                                MsgBox("Negli uffici selezionati è stato rilevato l'indirizzo di una macchina [Indirizzo Rilevato Non Idoneo: " & """" & ListBox2.SelectedItem.ToString & """" & "], poichè Asc Server non risulta connesso ad Asc Exchange è impossibile continuare. Invio Annullato.", MsgBoxStyle.Critical, "Asc Server - Errore Invio")
                                ListBox2.Items.Clear()
                                Exit Sub
                            End If
                        Next

                        ListBox2.Items.Clear()

                        If CheckBox1.Checked = True Then
                            Tratta()
                        Else
                            If TextBox3.Text.Contains("R I S E R V A T I S S I M O") Or TextBox3.Text.Contains("S E G R E T O") Or TextBox3.Text.Contains("S E G R E T I S S I M O") _
                                Or TextBox3.Text.Contains("C O N F I D E N T I A L") Or TextBox3.Text.Contains("S E C R E T") Or TextBox3.Text.Contains("T O P S E C R E T") _
                                Or TextBox3.Text.Contains("N A T O C O N F I D E N T I A L") Or TextBox3.Text.Contains("N A T O S E C R E T") Or TextBox3.Text.Contains("N A T O T O P S E C R E T") _
                                Then
                                TrattaPCN()
                            Else
                                Tratta()
                            End If
                        End If
                    Else
                        MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                    End If
                Else
                    If CheckBox1.Checked = True Then
                        Tratta()
                    Else
                        If TextBox3.Text.Contains("R I S E R V A T I S S I M O") Or TextBox3.Text.Contains("S E G R E T O") Or TextBox3.Text.Contains("S E G R E T I S S I M O") _
                            Or TextBox3.Text.Contains("C O N F I D E N T I A L") Or TextBox3.Text.Contains("S E C R E T") Or TextBox3.Text.Contains("T O P S E C R E T") _
                            Or TextBox3.Text.Contains("N A T O C O N F I D E N T I A L") Or TextBox3.Text.Contains("N A T O S E C R E T") Or TextBox3.Text.Contains("N A T O T O P S E C R E T") _
                            Then
                            TrattaPCN()
                        Else
                            Tratta()
                        End If
                    End If
                End If
            ElseIf My.Settings.TipoDiDistribuzione = "Manuale" Then
                If CheckBox1.Checked = True Then
                    Tratta()
                Else
                    If TextBox3.Text.Contains("R I S E R V A T I S S I M O") Or TextBox3.Text.Contains("S E G R E T O") Or TextBox3.Text.Contains("S E G R E T I S S I M O") _
                        Or TextBox3.Text.Contains("C O N F I D E N T I A L") Or TextBox3.Text.Contains("S E C R E T") Or TextBox3.Text.Contains("T O P S E C R E T") _
                        Or TextBox3.Text.Contains("N A T O C O N F I D E N T I A L") Or TextBox3.Text.Contains("N A T O S E C R E T") Or TextBox3.Text.Contains("N A T O T O P S E C R E T") _
                        Then
                        TrattaPCN()
                    Else
                        Tratta()
                    End If
                End If
            Else
                MsgBox("Non è ancora stata selezionata alcun tipo di distribuzione: Selezionare un tipo di distribuzione dal menu " & """" & "Distribuzione" & """" & " per poter continaure.", MsgBoxStyle.Critical, "Asc Server - Tipo di Distribuzione")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub Tratta()

        Try
            Try
                FileName = ""
                IsFirstChkLastBT = 0
                IsFisrChkLastBT = 0
                IsFirstChkBtIsRegular = 0
                IsFirstSelectBTasLastLine = 0
                IsFirstChkBTcounter = 0
                IsFirstChkFollowingSICline = 0
                IsFirstChkGDOIsRegular = 0
                IsFirstChkFmIsRegular = 0
                IsFirstControlloPOC = 0
                StatusBar1.Text = "Stato Invio: Azzeramento parametri"
            Catch ex As Exception

            End Try

            If TextBox3.Text <> "" Then
                If OpenedFile = 0 Then
                    RemoveProbableTemps()
                    Try
                        InvioAnnullato = 0

                        If CheckedListBox1.CheckedItems.Count < 1 Then
                            MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                            Exit Sub
                        End If


                        If CheckBox1.Checked = True Then
                            AddExer()
                        End If

                        ChkFmIsRegular()
                    Catch ex As Exception

                    End Try
                Else
                    If FinalSectionPresent = 1 Then
                        RemoveProbableTemps()
                        Try
                            InvioAnnullato = 0

                            If CheckedListBox1.CheckedItems.Count < 1 Then
                                MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                                Exit Sub
                            End If

                            If CheckBox1.Checked = True Then
                                AddExer()
                            End If

                            ChkFmIsRegular()
                        Catch ex As Exception

                        End Try
                    Else
                        Dim AlertMessage = MsgBox("L'inserimento di tutte le sezioni non è stato completato. Manca la parte finale, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Server - Sezione Finale Mancante")
                        If AlertMessage = 6 Then
                            RemoveProbableTemps()
                            Try
                                InvioAnnullato = 0

                                If CheckedListBox1.CheckedItems.Count < 1 Then
                                    MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                                    Exit Sub
                                End If

                                If CheckBox1.Checked = True Then
                                    AddExer()
                                End If

                                ChkFmIsRegular()
                            Catch ex As Exception

                            End Try
                        End If
                    End If
                End If
            Else
                MsgBox("Impossibile inviare il messaggio: testo del messaggio mancante. Invio Annullato.", MsgBoxStyle.Information, "Asc Server - Testo Mancante")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Errore durante il Tratta, ex message: " & ex.Message.ToString, MsgBoxStyle.Critical, "Asc Server Server - Errore")
        End Try

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Try
            If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                chkAllegato = "SI"
            Else
                chkAllegato = "NO"
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        If TextBox3.Text <> "" Then
            ManualPrint = 1
            Try
                If My.Settings.SavePrintSett = 1 Then
                    Stampa()
                Else
                    My.Settings.x = 9
                    PrintSett.Show()
                End If
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        e.Graphics.DrawString(dastampare, TextBox3.Font, Brushes.Black, New System.Drawing.RectangleF(My.Settings.Top, My.Settings.Left, My.Settings.Width, My.Settings.Height))
        e.HasMorePages = False

    End Sub

    Public Sub Stampa()

        Try
            StatusBar1.Text = "Stato: Inizio Stampa"
            RigheTextBox = Split(TextBox3.Text, vbCrLf)
            If System.IO.File.Exists(Application.StartupPath & "\TestoEmail.txt") Then
                leggi = System.IO.File.OpenText(Application.StartupPath & "\TestoEmail.txt")
                letto = leggi.ReadToEnd
                leggi.Close()
            End If
        Catch ex As Exception

        End Try

        If letto = TextBox3.Text Then
            Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
            If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
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
            Else
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
            End If
        Else
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
        End If

        IniziaStampa()

    End Sub
    Public Sub IniziaStampa()

        StatusBar1.Text = "Stato: Stampa in corso"
        If CheckBox2.Checked = True Then
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
                    leggi = System.IO.File.OpenText(Application.StartupPath & "\fogliostampa" & r & ".txt")
                    letto = leggi.ReadToEnd
                    leggi.Close()
                    dastampare = letto
                    PrintDocument1.Print()
                Next
            Catch ex As Exception

            End Try
        Else
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
        End If

        Try
            For r = 0 To NumFoglioStampa
                System.IO.File.Delete(Application.StartupPath & "\fogliostampa" & r & ".txt")
                NumFoglioStampa = 0
                numeroriga = 0
                stoppami = 0
                FinalSectionPresent = 0
                OpenedFile = 0
            Next
        Catch ex As Exception

        End Try

        If ManualPrint = 0 Then
            CleanAll()
        Else
            ManualPrint = 0
        End If
        CheckBox1.Checked = False
        CheckBox2.Checked = False

    End Sub

    Public Sub CleanAll()

        Try
            primadec = 0
            TextBox3.Clear()
            TextBox1.Clear()
            FinalSectionPresent = 0
            OpenedFile = 0
            CheckBox1.Checked = False
            CheckBox2.Checked = False
            For r = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(r, False)
            Next
            CheckedListBox1.SelectedIndex = 0
            StatusBar1.Text = "Stato: Processo invio terminato con successo"
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        CleanAll()

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Try
            datagiorno = My.Settings.DataOggi
            ora = Date.Now.Hour & ":" & Date.Now.Minute & ":" & Date.Now.Second
            Label6.Text = datagiorno & "   " & ora
        Catch ex As Exception

        End Try

        Try
            If ora = "0:0:0" Then
                SendDateUpdateRequest()
            End If
            If MenuItem40.Checked = True Then
                If ora = My.Settings.OraAggiornamento Then
                    SendDateUpdateRequest()
                End If
            End If
        Catch ex As Exception

        End Try

        Try
            If client.client.Connected = False Then
                If RectangleShape1.BackColor = Color.Green Then
                    RectangleShape1.BackColor = Color.Red
                    RectangleShape1.BorderColor = Color.Red
                    MenuItem42.Text = "Connetti"
                    Timering = 0
                    Timer5.Start()
                    Label8.Text = "LISTA UFFICI  [Connettersi ad Exchange]"
                    If Timer6.Enabled = True Then
                        Timer6.Stop()
                    End If
                End If
            End If
        Catch ex As Exception
            If RectangleShape1.BackColor = Color.Green Then
                RectangleShape1.BackColor = Color.Red
                RectangleShape1.BorderColor = Color.Red
                MenuItem42.Text = "Connetti"
                Timering = 0
                Timer5.Start()
                Label8.Text = "LISTA UFFICI  [Connettersi ad Exchange]"
                If Timer6.Enabled = True Then
                    Timer6.Stop()
                End If
            End If
        End Try

    End Sub

    Public Sub SendDateUpdateRequest()

        Timer1.Stop()

        My.Settings.SelectedDate = ""
        My.Settings.SelectedDate2 = ""

        My.Settings.DataOggi = Date.Now.Day.ToString & "." & Date.Now.Month.ToString & "." & Date.Now.Year.ToString

        If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEARRIVI") Then
        Else
            My.Settings.JlDate = Date.Now.DayOfYear
            My.Settings.DataIeri = My.Settings.DataOggi
            datagiorno = My.Settings.DataOggi
        End If

        Dim MostraAlert As Boolean = False
        If My.Settings.DataIeri <> My.Settings.DataOggi Then
            Dim Msgbox1 = MsgBox("Volete creare la cartella per il nuovo giorno?", vbYesNo, "Asc Server - Aggiornamento Data")
            If Msgbox1 = 6 Then
                My.Settings.NumEmail = 0
                My.Settings.NumEmail2 = 0
                My.Settings.NumEmailIntelligence = 0
                Label4.ForeColor = Color.Black
                Label6.ForeColor = Color.Black
                My.Settings.DataIeri = My.Settings.DataOggi
                datagiorno = My.Settings.DataOggi
                My.Settings.JlDate = Date.Now.DayOfYear
                MostraAlert = True
                Timer1.Start()
            Else
                My.Settings.DataOggi = My.Settings.DataIeri
                Label4.ForeColor = Color.Red
                Label6.ForeColor = Color.Red
                datagiorno = My.Settings.DataOggi
                Label6.Text = datagiorno
                Label4.Text = My.Settings.JlDate.ToString("D3")
                Timer1.Start()
            End If
        Else
            Timer1.Start()
        End If

        datagiorno = My.Settings.DataOggi

        Label4.Text = My.Settings.JlDate.ToString("D3")

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\Email") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email")
            End If
            'DATABASE ARRIVI
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEARRIVI") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\DATABASEARRIVI")
            End If
            'DATABASE PARTENZE
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\DATABASEPARTENZE")
            End If
            'DATABASE INTELLIGENCE
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\INTELLIGENCE") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\INTELLIGENCE")
            End If
        Catch ex As Exception

        End Try

        Try
            'DATABASE ARRIVI
            If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\dateusate.txt") Then
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Email\DATABASEARRIVI\dateusate.txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                If letto.Contains(datagiorno) Then
                Else
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEARRIVI\dateusate.txt")
                    scrivi.WriteLine(datagiorno)
                    scrivi.Close()
                End If
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEARRIVI\dateusate.txt")
                scrivi.WriteLine(datagiorno)
                scrivi.Close()
            End If
            'DATABASE PARTENZE
            If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\dateusate.txt") Then
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Email\DATABASEPARTENZE\dateusate.txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                If letto.Contains(datagiorno) Then
                Else
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEPARTENZE\dateusate.txt")
                    scrivi.WriteLine(datagiorno)
                    scrivi.Close()
                End If
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEPARTENZE\dateusate.txt")
                scrivi.WriteLine(datagiorno)
                scrivi.Close()
            End If
            'DATABASE INTELLIGENCE
            If System.IO.File.Exists(Application.StartupPath & "\Email\INTELLIGENCE\dateusate.txt") Then
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Email\INTELLIGENCE\dateusate.txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                If letto.Contains(datagiorno) Then
                Else
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\INTELLIGENCE\dateusate.txt")
                    scrivi.WriteLine(datagiorno)
                    scrivi.Close()
                End If
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\INTELLIGENCE\dateusate.txt")
                scrivi.WriteLine(datagiorno)
                scrivi.Close()
            End If
        Catch ex As Exception

        End Try

        Try
            'DATABASE ARRIVI
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno) Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno)
                My.Settings.NumEmail = 0
            End If
            'DATABASE PARTENZE
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno) Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno)
                My.Settings.NumEmail2 = 0
            End If
            'DATABASE INTELLIGENCE
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno) Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno)
                My.Settings.NumEmailIntelligence = 0
            End If
        Catch ex As Exception

        End Try

        Try
            'DATABASE ARRIVI
            If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt") Then
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                scrivi.Close()
            End If
            'DATABASE PARTENZE
            If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt") Then
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                scrivi.Close()
            End If
            'DATABASE INTELLIGENCE
            If System.IO.File.Exists(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt") Then
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                scrivi.Close()
            End If
        Catch ex As Exception

        End Try

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\Uffici") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Uffici")
                My.Settings.Ufficio = 0
            End If

            If System.IO.File.Exists(Application.StartupPath & "\Uffici\listauffici.txt") Then
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Uffici\listauffici.txt")
                scrivi.Close()
            End If
        Catch ex As Exception

        End Try

        SaveOthersProgressive()

        If MostraAlert = True Then
            Dim AlertMessage = MsgBox("Per poter continuare ad utilizzare il software bisogna prima effettuare la chiusura dei messaggi. Si desidera effettuare la chiusura adesso?", MsgBoxStyle.YesNo, "Asc Server - Chiusure")
            If AlertMessage = 6 Then
                If My.Settings.DataUltimaChiusura <> "" Then
                    If My.Settings.DataUltimaChiusura <> My.Settings.DataOggi Then
                        Chiusure.Show()
                    Else
                        MsgBox("La chiusura dei messaggi risulta già essere stata effettuata.", MsgBoxStyle.Information, "Asc Server - Chiusure")
                    End If
                Else
                    Chiusure.Show()
                End If
            End If
        End If

    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            TextBox3.Focus()
        End If

    End Sub

    Private Sub VisualizzaGliUfficiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VisualizzaGliUfficiToolStripMenuItem.Click

        Uffici.Show()

    End Sub

    Private Sub SelezionaTuttiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelezionaTuttiToolStripMenuItem.Click

        Try
            For r = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(r, True)
            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Sub DeselezionaTuttiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeselezionaTuttiToolStripMenuItem.Click

        Try
            For r = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(r, False)
            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Main_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        SaveProgressiveNumbers()

        SaveOthersProgressive()

        Try
            If Timer2.Enabled = True Then
                Dim AlertMessage2 = MsgBox("Il processo di invio messaggi ad asc exchange è attualmente in atto, si desidera procedere alla chiusura di Asc Server? [SCONSIGLIATO]", MsgBoxStyle.YesNo, "Asc Server - Chiusura")
                If AlertMessage2 = 6 Then
                    'Chiusura
                Else
                    e.Cancel = True
                    Exit Sub
                End If
            End If

            If client.client.Connected = True Then
                If client.client.Available > 0 Then
                    Dim AlertMessage = MsgBox("Il client sta ricevendo files da Asc Exchange, se viene chiuso in questo momento, i files andranno persi per sempre, si desidera continuare? [Sconsigliato]", MsgBoxStyle.YesNo, "Asc Exchange - Chiusura")
                    If AlertMessage = 6 Then
                        Try
                            client.Send(Requests.StringMessage, "ServerDisconnected")
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
                            Timer3.Stop()
                            Timer4.Stop()
                            Timer1.Stop()
                        Catch ex As Exception

                        End Try

                        Try
                            If System.IO.Directory.Exists(Application.StartupPath & "\Memorized") Then
                            Else
                                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Memorized")
                            End If
                            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\DataChiusura.txt")
                            scrivi.Write(My.Settings.DataOggi)
                            scrivi.Close()
                            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\JDChiusura.txt")
                            scrivi.Write(My.Settings.JlDate)
                            scrivi.Close()
                            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoArriviChiusura.txt")
                            scrivi.Write(My.Settings.NumEmail)
                            scrivi.Close()
                            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoPartenzeChiusura.txt")
                            scrivi.Write(My.Settings.NumEmail2)
                            scrivi.Close()
                            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoIntellChiusura.txt")
                            scrivi.Write(My.Settings.NumEmailIntelligence)
                            scrivi.Close()
                            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoUfficiChiusura.txt")
                            scrivi.Write(My.Settings.Ufficio)
                            scrivi.Close()
                        Catch ex As Exception

                        End Try

                        Try
                            TerSelectionDate.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            Accesso.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            Cerca.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            ContactsAdd.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            ContactsRemove.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            DATABASE_ARRIVI.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            DATABASE_PARTENZE.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            MessageText.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            Numerazione.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            PrintSett.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            Uffici.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            UfficiAdd.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            RimuoviUffici.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            ConnectedClientsList.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            CodeUffici.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            DateSelection.Close()
                        Catch ex As Exception
                        End Try

                        Try
                            Chiusure.Close()
                        Catch ex As Exception
                        End Try
                    Else
                        e.Cancel = True
                    End If
                Else
                    Try
                        client.Send(Requests.StringMessage, "ServerDisconnected")
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
                        Timer3.Stop()
                        Timer4.Stop()
                        Timer1.Stop()
                    Catch ex As Exception

                    End Try

                    Try
                        If System.IO.Directory.Exists(Application.StartupPath & "\Memorized") Then
                        Else
                            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Memorized")
                        End If
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\DataChiusura.txt")
                        scrivi.Write(My.Settings.DataOggi)
                        scrivi.Close()
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\JDChiusura.txt")
                        scrivi.Write(My.Settings.JlDate)
                        scrivi.Close()
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoArriviChiusura.txt")
                        scrivi.Write(My.Settings.NumEmail)
                        scrivi.Close()
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoPartenzeChiusura.txt")
                        scrivi.Write(My.Settings.NumEmail2)
                        scrivi.Close()
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoIntellChiusura.txt")
                        scrivi.Write(My.Settings.NumEmailIntelligence)
                        scrivi.Close()
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoUfficiChiusura.txt")
                        scrivi.Write(My.Settings.Ufficio)
                        scrivi.Close()
                    Catch ex As Exception

                    End Try

                    Try
                        TerSelectionDate.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        Accesso.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        Cerca.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        ContactsAdd.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        ContactsRemove.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        DATABASE_ARRIVI.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        DATABASE_PARTENZE.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        MessageText.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        Numerazione.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        PrintSett.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        Uffici.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        UfficiAdd.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        RimuoviUffici.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        ConnectedClientsList.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        CodeUffici.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        DateSelection.Close()
                    Catch ex As Exception
                    End Try

                    Try
                        Chiusure.Close()
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
                    Timer3.Stop()
                    Timer4.Stop()
                    Timer1.Stop()
                Catch ex As Exception

                End Try

                Try
                    If System.IO.Directory.Exists(Application.StartupPath & "\Memorized") Then
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\Memorized")
                    End If
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\DataChiusura.txt")
                    scrivi.Write(My.Settings.DataOggi)
                    scrivi.Close()
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\JDChiusura.txt")
                    scrivi.Write(My.Settings.JlDate)
                    scrivi.Close()
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoArriviChiusura.txt")
                    scrivi.Write(My.Settings.NumEmail)
                    scrivi.Close()
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoPartenzeChiusura.txt")
                    scrivi.Write(My.Settings.NumEmail2)
                    scrivi.Close()
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoIntellChiusura.txt")
                    scrivi.Write(My.Settings.NumEmailIntelligence)
                    scrivi.Close()
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoUfficiChiusura.txt")
                    scrivi.Write(My.Settings.Ufficio)
                    scrivi.Close()
                Catch ex As Exception

                End Try

                Try
                    TerSelectionDate.Close()
                Catch ex As Exception
                End Try

                Try
                    Accesso.Close()
                Catch ex As Exception
                End Try

                Try
                    Cerca.Close()
                Catch ex As Exception
                End Try

                Try
                    ContactsAdd.Close()
                Catch ex As Exception
                End Try

                Try
                    ContactsRemove.Close()
                Catch ex As Exception
                End Try

                Try
                    DATABASE_ARRIVI.Close()
                Catch ex As Exception
                End Try

                Try
                    DATABASE_PARTENZE.Close()
                Catch ex As Exception
                End Try

                Try
                    MessageText.Close()
                Catch ex As Exception
                End Try

                Try
                    Numerazione.Close()
                Catch ex As Exception
                End Try

                Try
                    PrintSett.Close()
                Catch ex As Exception
                End Try

                Try
                    Uffici.Close()
                Catch ex As Exception
                End Try

                Try
                    UfficiAdd.Close()
                Catch ex As Exception
                End Try

                Try
                    RimuoviUffici.Close()
                Catch ex As Exception
                End Try

                Try
                    ConnectedClientsList.Close()
                Catch ex As Exception
                End Try

                Try
                    CodeUffici.Close()
                Catch ex As Exception
                End Try

                Try
                    DateSelection.Close()
                Catch ex As Exception
                End Try

                Try
                    Chiusure.Close()
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
                Timer3.Stop()
                Timer4.Stop()
                Timer1.Stop()
            Catch ex As Exception

            End Try

            Try
                If System.IO.Directory.Exists(Application.StartupPath & "\Memorized") Then
                Else
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Memorized")
                End If
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\DataChiusura.txt")
                scrivi.Write(My.Settings.DataOggi)
                scrivi.Close()
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\JDChiusura.txt")
                scrivi.Write(My.Settings.JlDate)
                scrivi.Close()
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoArriviChiusura.txt")
                scrivi.Write(My.Settings.NumEmail)
                scrivi.Close()
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoPartenzeChiusura.txt")
                scrivi.Write(My.Settings.NumEmail2)
                scrivi.Close()
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoIntellChiusura.txt")
                scrivi.Write(My.Settings.NumEmailIntelligence)
                scrivi.Close()
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoUfficiChiusura.txt")
                scrivi.Write(My.Settings.Ufficio)
                scrivi.Close()
            Catch ex As Exception

            End Try

            Try
                TerSelectionDate.Close()
            Catch ex As Exception
            End Try

            Try
                Accesso.Close()
            Catch ex As Exception
            End Try

            Try
                Cerca.Close()
            Catch ex As Exception
            End Try

            Try
                ContactsAdd.Close()
            Catch ex As Exception
            End Try

            Try
                ContactsRemove.Close()
            Catch ex As Exception
            End Try

            Try
                DATABASE_ARRIVI.Close()
            Catch ex As Exception
            End Try

            Try
                DATABASE_PARTENZE.Close()
            Catch ex As Exception
            End Try

            Try
                MessageText.Close()
            Catch ex As Exception
            End Try

            Try
                Numerazione.Close()
            Catch ex As Exception
            End Try

            Try
                PrintSett.Close()
            Catch ex As Exception
            End Try

            Try
                Uffici.Close()
            Catch ex As Exception
            End Try

            Try
                UfficiAdd.Close()
            Catch ex As Exception
            End Try

            Try
                RimuoviUffici.Close()
            Catch ex As Exception
            End Try

            Try
                ConnectedClientsList.Close()
            Catch ex As Exception
            End Try

            Try
                CodeUffici.Close()
            Catch ex As Exception
            End Try

            Try
                DateSelection.Close()
            Catch ex As Exception
            End Try

            Try
                Chiusure.Close()
            Catch ex As Exception
            End Try
        End Try

    End Sub

    Private Sub EsciToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EsciToolStripMenuItem.Click

        Me.Close()

    End Sub

    Private Sub AzzeraCampiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AzzeraCampiToolStripMenuItem.Click

        CleanAll()

    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click

        Me.Close()

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click

        Cerca.Show()

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        Try
            OpenFileDialog2.Filter = "File Testo|*.txt"
            OpenFileDialog2.FileName = ""
            If OpenFileDialog2.ShowDialog() = Windows.Forms.DialogResult.OK Then
                leggi = System.IO.File.OpenText(OpenFileDialog2.FileName)
                letto = leggi.ReadToEnd
                leggi.Close()
                Dim Btcounter As Integer = 0
                '*** Transform String to Integer ***
                OpenedFileNr = 0
                Dim NumberToString As String = ""
                Dim MySplitter() As String = Split(letto, vbCrLf)
                Dim MySecondSplitter() As String
                For r = 0 To UBound(MySplitter)
                    If MySplitter(r).ToString.Contains("PARTE FINALE DI ") Or MySplitter(r).ToString.Contains("FINAL SECTION OF ") Then
                        MySecondSplitter = Split(MySplitter(r).ToString, " ")
                        If MySecondSplitter(0).ToString = "FINAL" And MySecondSplitter(1).ToString = "SECTION" And MySecondSplitter(2).ToString = "OF" Then
                            If MySecondSplitter(3).ToString = "1" Or MySecondSplitter(3).ToString = "ONE" Then
                                FinalSectionNr = 1
                            ElseIf MySecondSplitter(3).ToString = "2" Or MySecondSplitter(3).ToString = "TWO" Then
                                FinalSectionNr = 2
                            ElseIf MySecondSplitter(3).ToString = "3" Or MySecondSplitter(3).ToString = "THREE" Then
                                FinalSectionNr = 3
                            ElseIf MySecondSplitter(3).ToString = "4" Or MySecondSplitter(3).ToString = "FOUR" Then
                                FinalSectionNr = 4
                            ElseIf MySecondSplitter(3).ToString = "5" Or MySecondSplitter(3).ToString = "FIVE" Then
                                FinalSectionNr = 5
                            ElseIf MySecondSplitter(3).ToString = "6" Or MySecondSplitter(3).ToString = "SIX" Then
                                FinalSectionNr = 6
                            ElseIf MySecondSplitter(3).ToString = "7" Or MySecondSplitter(3).ToString = "SEVEN" Then
                                FinalSectionNr = 7
                            ElseIf MySecondSplitter(3).ToString = "8" Or MySecondSplitter(3).ToString = "EIGHT" Then
                                FinalSectionNr = 8
                            ElseIf MySecondSplitter(3).ToString = "9" Or MySecondSplitter(3).ToString = "NINE" Then
                                FinalSectionNr = 9
                            ElseIf MySecondSplitter(3).ToString = "10" Or MySecondSplitter(3).ToString = "TEN" Then
                                FinalSectionNr = 10
                            ElseIf MySecondSplitter(3).ToString = "11" Or MySecondSplitter(3).ToString = "ELEVEN" Then
                                FinalSectionNr = 11
                            ElseIf MySecondSplitter(3).ToString = "12" Or MySecondSplitter(3).ToString = "TWELVE" Then
                                FinalSectionNr = 12
                            ElseIf MySecondSplitter(3).ToString = "13" Or MySecondSplitter(3).ToString = "THIRTEEN" Then
                                FinalSectionNr = 13
                            ElseIf MySecondSplitter(3).ToString = "14" Or MySecondSplitter(3).ToString = "FOURTEEN" Then
                                FinalSectionNr = 14
                            ElseIf MySecondSplitter(3).ToString = "15" Or MySecondSplitter(3).ToString = "FIFTEEN" Then
                                FinalSectionNr = 15
                            ElseIf MySecondSplitter(3).ToString = "16" Or MySecondSplitter(3).ToString = "SIXTEEN" Then
                                FinalSectionNr = 16
                            ElseIf MySecondSplitter(3).ToString = "17" Or MySecondSplitter(3).ToString = "SEVENTEEN" Then
                                FinalSectionNr = 17
                            ElseIf MySecondSplitter(3).ToString = "18" Or MySecondSplitter(3).ToString = "EIGHTEEN" Then
                                FinalSectionNr = 18
                            ElseIf MySecondSplitter(3).ToString = "19" Or MySecondSplitter(3).ToString = "NINETEEN" Then
                                FinalSectionNr = 19
                            ElseIf MySecondSplitter(3).ToString = "20" Or MySecondSplitter(3).ToString = "TWENTY" Then
                                FinalSectionNr = 20
                            ElseIf MySecondSplitter(3).ToString = "21" Or MySecondSplitter(3).ToString = "TWENTYONE" Then
                                FinalSectionNr = 21
                            ElseIf MySecondSplitter(3).ToString = "22" Or MySecondSplitter(3).ToString = "TWENTYTWO" Then
                                FinalSectionNr = 22
                            ElseIf MySecondSplitter(3).ToString = "23" Or MySecondSplitter(3).ToString = "TWENTYTHREE" Then
                                FinalSectionNr = 23
                            ElseIf MySecondSplitter(3).ToString = "24" Or MySecondSplitter(3).ToString = "TWENTYFOUR" Then
                                FinalSectionNr = 24
                            ElseIf MySecondSplitter(3).ToString = "25" Or MySecondSplitter(3).ToString = "TWENTYFIVE" Then
                                FinalSectionNr = 25
                            ElseIf MySecondSplitter(3).ToString = "26" Or MySecondSplitter(3).ToString = "TWENTYSIX" Then
                                FinalSectionNr = 26
                            ElseIf MySecondSplitter(3).ToString = "27" Or MySecondSplitter(3).ToString = "TWENTYSEVEN" Then
                                FinalSectionNr = 27
                            ElseIf MySecondSplitter(3).ToString = "28" Or MySecondSplitter(3).ToString = "TWENTYEIGHT" Then
                                FinalSectionNr = 28
                            ElseIf MySecondSplitter(3).ToString = "29" Or MySecondSplitter(3).ToString = "TWENTYNINE" Then
                                FinalSectionNr = 29
                            ElseIf MySecondSplitter(3).ToString = "30" Or MySecondSplitter(3).ToString = "THIRTY" Then
                                FinalSectionNr = 30
                            ElseIf MySecondSplitter(3).ToString = "31" Or MySecondSplitter(3).ToString = "THIRTYONE" Then
                                FinalSectionNr = 31
                            ElseIf MySecondSplitter(3).ToString = "32" Or MySecondSplitter(3).ToString = "THIRTYTWO" Then
                                FinalSectionNr = 32
                            ElseIf MySecondSplitter(3).ToString = "33" Or MySecondSplitter(3).ToString = "THIRTYTHREE" Then
                                FinalSectionNr = 33
                            ElseIf MySecondSplitter(3).ToString = "34" Or MySecondSplitter(3).ToString = "THIRTYFOUR" Then
                                FinalSectionNr = 34
                            ElseIf MySecondSplitter(3).ToString = "35" Or MySecondSplitter(3).ToString = "THIRTYFIVE" Then
                                FinalSectionNr = 35
                            ElseIf MySecondSplitter(3).ToString = "36" Or MySecondSplitter(3).ToString = "THIRTYSIX" Then
                                FinalSectionNr = 36
                            ElseIf MySecondSplitter(3).ToString = "37" Or MySecondSplitter(3).ToString = "THIRTYSEVEN" Then
                                FinalSectionNr = 37
                            ElseIf MySecondSplitter(3).ToString = "38" Or MySecondSplitter(3).ToString = "THIRTYEIGHT" Then
                                FinalSectionNr = 38
                            ElseIf MySecondSplitter(3).ToString = "39" Or MySecondSplitter(3).ToString = "THIRTYNINE" Then
                                FinalSectionNr = 39
                            ElseIf MySecondSplitter(3).ToString = "40" Or MySecondSplitter(3).ToString = "FORTY" Then
                                FinalSectionNr = 40
                            ElseIf MySecondSplitter(3).ToString = "41" Or MySecondSplitter(3).ToString = "FORTYONE" Then
                                FinalSectionNr = 41
                            ElseIf MySecondSplitter(3).ToString = "42" Or MySecondSplitter(3).ToString = "FORTYTWO" Then
                                FinalSectionNr = 42
                            ElseIf MySecondSplitter(3).ToString = "43" Or MySecondSplitter(3).ToString = "FORTYTHREE" Then
                                FinalSectionNr = 43
                            ElseIf MySecondSplitter(3).ToString = "44" Or MySecondSplitter(3).ToString = "FORTYFOUR" Then
                                FinalSectionNr = 44
                            ElseIf MySecondSplitter(3).ToString = "45" Or MySecondSplitter(3).ToString = "FORTYFIVE" Then
                                FinalSectionNr = 45
                            ElseIf MySecondSplitter(3).ToString = "46" Or MySecondSplitter(3).ToString = "FORTYSIX" Then
                                FinalSectionNr = 46
                            ElseIf MySecondSplitter(3).ToString = "47" Or MySecondSplitter(3).ToString = "FORTYSEVEN" Then
                                FinalSectionNr = 47
                            ElseIf MySecondSplitter(3).ToString = "48" Or MySecondSplitter(3).ToString = "FORTYEIGHT" Then
                                FinalSectionNr = 48
                            ElseIf MySecondSplitter(3).ToString = "49" Or MySecondSplitter(3).ToString = "FORTYNINE" Then
                                FinalSectionNr = 49
                            ElseIf MySecondSplitter(3).ToString = "50" Or MySecondSplitter(3).ToString = "FIFTY" Then
                                FinalSectionNr = 50
                            ElseIf MySecondSplitter(3).ToString = "51" Or MySecondSplitter(3).ToString = "FIFTYONE" Then
                                FinalSectionNr = 51
                            ElseIf MySecondSplitter(3).ToString = "52" Or MySecondSplitter(3).ToString = "FIFTYTWO" Then
                                FinalSectionNr = 52
                            ElseIf MySecondSplitter(3).ToString = "53" Or MySecondSplitter(3).ToString = "FIFTYTHREE" Then
                                FinalSectionNr = 53
                            ElseIf MySecondSplitter(3).ToString = "54" Or MySecondSplitter(3).ToString = "FIFTYFOUR" Then
                                FinalSectionNr = 54
                            ElseIf MySecondSplitter(3).ToString = "55" Or MySecondSplitter(3).ToString = "FIFTYFIVE" Then
                                FinalSectionNr = 55
                            ElseIf MySecondSplitter(3).ToString = "56" Or MySecondSplitter(3).ToString = "FIFTYSIX" Then
                                FinalSectionNr = 56
                            ElseIf MySecondSplitter(3).ToString = "57" Or MySecondSplitter(3).ToString = "FIFTYSEVEN" Then
                                FinalSectionNr = 57
                            ElseIf MySecondSplitter(3).ToString = "58" Or MySecondSplitter(3).ToString = "FIFTYEIGHT" Then
                                FinalSectionNr = 58
                            ElseIf MySecondSplitter(3).ToString = "59" Or MySecondSplitter(3).ToString = "FIFTYNINE" Then
                                FinalSectionNr = 59
                            ElseIf MySecondSplitter(3).ToString = "60" Or MySecondSplitter(3).ToString = "SIXTY" Then
                                FinalSectionNr = 60
                            ElseIf MySecondSplitter(3).ToString = "61" Or MySecondSplitter(3).ToString = "SIXTYONE" Then
                                FinalSectionNr = 61
                            ElseIf MySecondSplitter(3).ToString = "62" Or MySecondSplitter(3).ToString = "SIXTYTWO" Then
                                FinalSectionNr = 62
                            ElseIf MySecondSplitter(3).ToString = "63" Or MySecondSplitter(3).ToString = "SIXTYTHREE" Then
                                FinalSectionNr = 63
                            ElseIf MySecondSplitter(3).ToString = "64" Or MySecondSplitter(3).ToString = "SIXTYFOUR" Then
                                FinalSectionNr = 64
                            ElseIf MySecondSplitter(3).ToString = "65" Or MySecondSplitter(3).ToString = "SIXTYFIVE" Then
                                FinalSectionNr = 65
                            ElseIf MySecondSplitter(3).ToString = "66" Or MySecondSplitter(3).ToString = "SIXTYSIX" Then
                                FinalSectionNr = 66
                            ElseIf MySecondSplitter(3).ToString = "67" Or MySecondSplitter(3).ToString = "SIXTYSEVEN" Then
                                FinalSectionNr = 67
                            ElseIf MySecondSplitter(3).ToString = "68" Or MySecondSplitter(3).ToString = "SIXTYEIGHT" Then
                                FinalSectionNr = 68
                            ElseIf MySecondSplitter(3).ToString = "69" Or MySecondSplitter(3).ToString = "SIXTYNINE" Then
                                FinalSectionNr = 69
                            ElseIf MySecondSplitter(3).ToString = "70" Or MySecondSplitter(3).ToString = "SEVENTY" Then
                                FinalSectionNr = 70
                            ElseIf MySecondSplitter(3).ToString = "71" Or MySecondSplitter(3).ToString = "SEVENTYONE" Then
                                FinalSectionNr = 71
                            ElseIf MySecondSplitter(3).ToString = "72" Or MySecondSplitter(3).ToString = "SEVENTYTWO" Then
                                FinalSectionNr = 72
                            ElseIf MySecondSplitter(3).ToString = "73" Or MySecondSplitter(3).ToString = "SEVENTYTHREE" Then
                                FinalSectionNr = 73
                            ElseIf MySecondSplitter(3).ToString = "74" Or MySecondSplitter(3).ToString = "SEVENTYFOUR" Then
                                FinalSectionNr = 74
                            ElseIf MySecondSplitter(3).ToString = "75" Or MySecondSplitter(3).ToString = "SEVENTYFIVE" Then
                                FinalSectionNr = 75
                            End If
                            Exit For
                        Else
                            If MySecondSplitter(0).ToString = "PARTE" And MySecondSplitter(1).ToString = "FINALE" And MySecondSplitter(2).ToString = "DI" Then
                                If MySecondSplitter(3).ToString = "1" Then
                                    FinalSectionNr = 1
                                ElseIf MySecondSplitter(3).ToString = "2" Then
                                    FinalSectionNr = 2
                                ElseIf MySecondSplitter(3).ToString = "3" Then
                                    FinalSectionNr = 3
                                ElseIf MySecondSplitter(3).ToString = "4" Then
                                    FinalSectionNr = 4
                                ElseIf MySecondSplitter(3).ToString = "5" Then
                                    FinalSectionNr = 5
                                ElseIf MySecondSplitter(3).ToString = "6" Then
                                    FinalSectionNr = 6
                                ElseIf MySecondSplitter(3).ToString = "7" Then
                                    FinalSectionNr = 7
                                ElseIf MySecondSplitter(3).ToString = "8" Then
                                    FinalSectionNr = 8
                                ElseIf MySecondSplitter(3).ToString = "9" Then
                                    FinalSectionNr = 9
                                ElseIf MySecondSplitter(3).ToString = "10" Then
                                    FinalSectionNr = 10
                                ElseIf MySecondSplitter(3).ToString = "11" Then
                                    FinalSectionNr = 11
                                ElseIf MySecondSplitter(3).ToString = "12" Then
                                    FinalSectionNr = 12
                                ElseIf MySecondSplitter(3).ToString = "13" Then
                                    FinalSectionNr = 13
                                ElseIf MySecondSplitter(3).ToString = "14" Then
                                    FinalSectionNr = 14
                                ElseIf MySecondSplitter(3).ToString = "15" Then
                                    FinalSectionNr = 15
                                ElseIf MySecondSplitter(3).ToString = "16" Then
                                    FinalSectionNr = 16
                                ElseIf MySecondSplitter(3).ToString = "17" Then
                                    FinalSectionNr = 17
                                ElseIf MySecondSplitter(3).ToString = "18" Then
                                    FinalSectionNr = 18
                                ElseIf MySecondSplitter(3).ToString = "19" Then
                                    FinalSectionNr = 19
                                ElseIf MySecondSplitter(3).ToString = "20" Then
                                    FinalSectionNr = 20
                                ElseIf MySecondSplitter(3).ToString = "21" Then
                                    FinalSectionNr = 21
                                ElseIf MySecondSplitter(3).ToString = "22" Then
                                    FinalSectionNr = 22
                                ElseIf MySecondSplitter(3).ToString = "23" Then
                                    FinalSectionNr = 23
                                ElseIf MySecondSplitter(3).ToString = "24" Then
                                    FinalSectionNr = 24
                                ElseIf MySecondSplitter(3).ToString = "25" Then
                                    FinalSectionNr = 25
                                ElseIf MySecondSplitter(3).ToString = "26" Then
                                    FinalSectionNr = 26
                                ElseIf MySecondSplitter(3).ToString = "27" Then
                                    FinalSectionNr = 27
                                ElseIf MySecondSplitter(3).ToString = "28" Then
                                    FinalSectionNr = 28
                                ElseIf MySecondSplitter(3).ToString = "29" Then
                                    FinalSectionNr = 29
                                ElseIf MySecondSplitter(3).ToString = "30" Then
                                    FinalSectionNr = 30
                                ElseIf MySecondSplitter(3).ToString = "31" Then
                                    FinalSectionNr = 31
                                ElseIf MySecondSplitter(3).ToString = "32" Then
                                    FinalSectionNr = 32
                                ElseIf MySecondSplitter(3).ToString = "33" Then
                                    FinalSectionNr = 33
                                ElseIf MySecondSplitter(3).ToString = "34" Then
                                    FinalSectionNr = 34
                                ElseIf MySecondSplitter(3).ToString = "35" Then
                                    FinalSectionNr = 35
                                ElseIf MySecondSplitter(3).ToString = "36" Then
                                    FinalSectionNr = 36
                                ElseIf MySecondSplitter(3).ToString = "37" Then
                                    FinalSectionNr = 37
                                ElseIf MySecondSplitter(3).ToString = "38" Then
                                    FinalSectionNr = 38
                                ElseIf MySecondSplitter(3).ToString = "39" Then
                                    FinalSectionNr = 39
                                ElseIf MySecondSplitter(3).ToString = "40" Then
                                    FinalSectionNr = 40
                                ElseIf MySecondSplitter(3).ToString = "41" Then
                                    FinalSectionNr = 41
                                ElseIf MySecondSplitter(3).ToString = "42" Then
                                    FinalSectionNr = 42
                                ElseIf MySecondSplitter(3).ToString = "43" Then
                                    FinalSectionNr = 43
                                ElseIf MySecondSplitter(3).ToString = "44" Then
                                    FinalSectionNr = 44
                                ElseIf MySecondSplitter(3).ToString = "45" Then
                                    FinalSectionNr = 45
                                ElseIf MySecondSplitter(3).ToString = "46" Then
                                    FinalSectionNr = 46
                                ElseIf MySecondSplitter(3).ToString = "47" Then
                                    FinalSectionNr = 47
                                ElseIf MySecondSplitter(3).ToString = "48" Then
                                    FinalSectionNr = 48
                                ElseIf MySecondSplitter(3).ToString = "49" Then
                                    FinalSectionNr = 49
                                ElseIf MySecondSplitter(3).ToString = "50" Then
                                    FinalSectionNr = 50
                                ElseIf MySecondSplitter(3).ToString = "51" Then
                                    FinalSectionNr = 51
                                ElseIf MySecondSplitter(3).ToString = "52" Then
                                    FinalSectionNr = 52
                                ElseIf MySecondSplitter(3).ToString = "53" Then
                                    FinalSectionNr = 53
                                ElseIf MySecondSplitter(3).ToString = "54" Then
                                    FinalSectionNr = 54
                                ElseIf MySecondSplitter(3).ToString = "55" Then
                                    FinalSectionNr = 55
                                ElseIf MySecondSplitter(3).ToString = "56" Then
                                    FinalSectionNr = 56
                                ElseIf MySecondSplitter(3).ToString = "57" Then
                                    FinalSectionNr = 57
                                ElseIf MySecondSplitter(3).ToString = "58" Then
                                    FinalSectionNr = 58
                                ElseIf MySecondSplitter(3).ToString = "59" Then
                                    FinalSectionNr = 59
                                ElseIf MySecondSplitter(3).ToString = "60" Then
                                    FinalSectionNr = 60
                                ElseIf MySecondSplitter(3).ToString = "61" Then
                                    FinalSectionNr = 61
                                ElseIf MySecondSplitter(3).ToString = "62" Then
                                    FinalSectionNr = 62
                                ElseIf MySecondSplitter(3).ToString = "63" Then
                                    FinalSectionNr = 63
                                ElseIf MySecondSplitter(3).ToString = "64" Then
                                    FinalSectionNr = 64
                                ElseIf MySecondSplitter(3).ToString = "65" Then
                                    FinalSectionNr = 65
                                ElseIf MySecondSplitter(3).ToString = "66" Then
                                    FinalSectionNr = 66
                                ElseIf MySecondSplitter(3).ToString = "67" Then
                                    FinalSectionNr = 67
                                ElseIf MySecondSplitter(3).ToString = "68" Then
                                    FinalSectionNr = 68
                                ElseIf MySecondSplitter(3).ToString = "69" Then
                                    FinalSectionNr = 69
                                ElseIf MySecondSplitter(3).ToString = "70" Then
                                    FinalSectionNr = 70
                                ElseIf MySecondSplitter(3).ToString = "71" Then
                                    FinalSectionNr = 71
                                ElseIf MySecondSplitter(3).ToString = "72" Then
                                    FinalSectionNr = 72
                                ElseIf MySecondSplitter(3).ToString = "73" Then
                                    FinalSectionNr = 73
                                ElseIf MySecondSplitter(3).ToString = "74" Then
                                    FinalSectionNr = 74
                                ElseIf MySecondSplitter(3).ToString = "75" Then
                                    FinalSectionNr = 75
                                End If
                                Exit For
                            End If
                        End If
                        Exit For
                    End If
                    MySecondSplitter = Split(MySplitter(r).ToString, " ")
                    If MySecondSplitter.Count = 4 Then
                        If MySecondSplitter(0).ToString = "PARTE" And MySecondSplitter(2).ToString = "DI" Then
                            If MySecondSplitter(1).ToString = "1" Then
                                OpenedFileNr = 1
                            ElseIf MySecondSplitter(1).ToString = "2" Then
                                OpenedFileNr = 2
                            ElseIf MySecondSplitter(1).ToString = "3" Then
                                OpenedFileNr = 3
                            ElseIf MySecondSplitter(1).ToString = "4" Then
                                OpenedFileNr = 4
                            ElseIf MySecondSplitter(1).ToString = "5" Then
                                OpenedFileNr = 5
                            ElseIf MySecondSplitter(1).ToString = "6" Then
                                OpenedFileNr = 6
                            ElseIf MySecondSplitter(1).ToString = "7" Then
                                OpenedFileNr = 7
                            ElseIf MySecondSplitter(1).ToString = "8" Then
                                OpenedFileNr = 8
                            ElseIf MySecondSplitter(1).ToString = "9" Then
                                OpenedFileNr = 9
                            ElseIf MySecondSplitter(1).ToString = "10" Then
                                OpenedFileNr = 10
                            ElseIf MySecondSplitter(1).ToString = "11" Then
                                OpenedFileNr = 11
                            ElseIf MySecondSplitter(1).ToString = "12" Then
                                OpenedFileNr = 12
                            ElseIf MySecondSplitter(1).ToString = "13" Then
                                OpenedFileNr = 13
                            ElseIf MySecondSplitter(1).ToString = "14" Then
                                OpenedFileNr = 14
                            ElseIf MySecondSplitter(1).ToString = "15" Then
                                OpenedFileNr = 15
                            ElseIf MySecondSplitter(1).ToString = "16" Then
                                OpenedFileNr = 16
                            ElseIf MySecondSplitter(1).ToString = "17" Then
                                OpenedFileNr = 17
                            ElseIf MySecondSplitter(1).ToString = "18" Then
                                OpenedFileNr = 18
                            ElseIf MySecondSplitter(1).ToString = "19" Then
                                OpenedFileNr = 19
                            ElseIf MySecondSplitter(1).ToString = "20" Then
                                OpenedFileNr = 20
                            ElseIf MySecondSplitter(1).ToString = "21" Then
                                OpenedFileNr = 21
                            ElseIf MySecondSplitter(1).ToString = "22" Then
                                OpenedFileNr = 22
                            ElseIf MySecondSplitter(1).ToString = "23" Then
                                OpenedFileNr = 23
                            ElseIf MySecondSplitter(1).ToString = "24" Then
                                OpenedFileNr = 24
                            ElseIf MySecondSplitter(1).ToString = "25" Then
                                OpenedFileNr = 25
                            ElseIf MySecondSplitter(1).ToString = "26" Then
                                OpenedFileNr = 26
                            ElseIf MySecondSplitter(1).ToString = "27" Then
                                OpenedFileNr = 27
                            ElseIf MySecondSplitter(1).ToString = "28" Then
                                OpenedFileNr = 28
                            ElseIf MySecondSplitter(1).ToString = "29" Then
                                OpenedFileNr = 29
                            ElseIf MySecondSplitter(1).ToString = "30" Then
                                OpenedFileNr = 30
                            ElseIf MySecondSplitter(1).ToString = "31" Then
                                OpenedFileNr = 31
                            ElseIf MySecondSplitter(1).ToString = "32" Then
                                OpenedFileNr = 32
                            ElseIf MySecondSplitter(1).ToString = "33" Then
                                OpenedFileNr = 33
                            ElseIf MySecondSplitter(1).ToString = "34" Then
                                OpenedFileNr = 34
                            ElseIf MySecondSplitter(1).ToString = "35" Then
                                OpenedFileNr = 35
                            ElseIf MySecondSplitter(1).ToString = "36" Then
                                OpenedFileNr = 36
                            ElseIf MySecondSplitter(1).ToString = "37" Then
                                OpenedFileNr = 37
                            ElseIf MySecondSplitter(1).ToString = "38" Then
                                OpenedFileNr = 38
                            ElseIf MySecondSplitter(1).ToString = "39" Then
                                OpenedFileNr = 39
                            ElseIf MySecondSplitter(1).ToString = "40" Then
                                OpenedFileNr = 40
                            ElseIf MySecondSplitter(1).ToString = "41" Then
                                OpenedFileNr = 41
                            ElseIf MySecondSplitter(1).ToString = "42" Then
                                OpenedFileNr = 42
                            ElseIf MySecondSplitter(1).ToString = "43" Then
                                OpenedFileNr = 43
                            ElseIf MySecondSplitter(1).ToString = "44" Then
                                OpenedFileNr = 44
                            ElseIf MySecondSplitter(1).ToString = "45" Then
                                OpenedFileNr = 45
                            ElseIf MySecondSplitter(1).ToString = "46" Then
                                OpenedFileNr = 46
                            ElseIf MySecondSplitter(1).ToString = "47" Then
                                OpenedFileNr = 47
                            ElseIf MySecondSplitter(1).ToString = "48" Then
                                OpenedFileNr = 48
                            ElseIf MySecondSplitter(1).ToString = "49" Then
                                OpenedFileNr = 49
                            ElseIf MySecondSplitter(1).ToString = "50" Then
                                OpenedFileNr = 50
                            ElseIf MySecondSplitter(1).ToString = "51" Then
                                OpenedFileNr = 51
                            ElseIf MySecondSplitter(1).ToString = "52" Then
                                OpenedFileNr = 52
                            ElseIf MySecondSplitter(1).ToString = "53" Then
                                OpenedFileNr = 53
                            ElseIf MySecondSplitter(1).ToString = "54" Then
                                OpenedFileNr = 54
                            ElseIf MySecondSplitter(1).ToString = "55" Then
                                OpenedFileNr = 55
                            ElseIf MySecondSplitter(1).ToString = "56" Then
                                OpenedFileNr = 56
                            ElseIf MySecondSplitter(1).ToString = "57" Then
                                OpenedFileNr = 57
                            ElseIf MySecondSplitter(1).ToString = "58" Then
                                OpenedFileNr = 58
                            ElseIf MySecondSplitter(1).ToString = "59" Then
                                OpenedFileNr = 59
                            ElseIf MySecondSplitter(1).ToString = "60" Then
                                OpenedFileNr = 60
                            ElseIf MySecondSplitter(1).ToString = "61" Then
                                OpenedFileNr = 61
                            ElseIf MySecondSplitter(1).ToString = "62" Then
                                OpenedFileNr = 62
                            ElseIf MySecondSplitter(1).ToString = "63" Then
                                OpenedFileNr = 63
                            ElseIf MySecondSplitter(1).ToString = "64" Then
                                OpenedFileNr = 64
                            ElseIf MySecondSplitter(1).ToString = "65" Then
                                OpenedFileNr = 65
                            ElseIf MySecondSplitter(1).ToString = "66" Then
                                OpenedFileNr = 66
                            ElseIf MySecondSplitter(1).ToString = "67" Then
                                OpenedFileNr = 67
                            ElseIf MySecondSplitter(1).ToString = "68" Then
                                OpenedFileNr = 68
                            ElseIf MySecondSplitter(1).ToString = "69" Then
                                OpenedFileNr = 69
                            ElseIf MySecondSplitter(1).ToString = "70" Then
                                OpenedFileNr = 70
                            ElseIf MySecondSplitter(1).ToString = "71" Then
                                OpenedFileNr = 71
                            ElseIf MySecondSplitter(1).ToString = "72" Then
                                OpenedFileNr = 72
                            ElseIf MySecondSplitter(1).ToString = "73" Then
                                OpenedFileNr = 73
                            ElseIf MySecondSplitter(1).ToString = "74" Then
                                OpenedFileNr = 74
                            ElseIf MySecondSplitter(1).ToString = "75" Then
                                OpenedFileNr = 75
                            End If
                            Exit For
                        Else
                            If MySecondSplitter(0).ToString = "SECTION" And MySecondSplitter(2).ToString = "OF" Then
                                If MySecondSplitter(1).ToString = "1" Or MySecondSplitter(1).ToString = "ONE" Then
                                    OpenedFileNr = 1
                                    NumberToString = "ONE"
                                ElseIf MySecondSplitter(1).ToString = "2" Or MySecondSplitter(1).ToString = "TWO" Then
                                    OpenedFileNr = 2
                                    NumberToString = "TWO"
                                ElseIf MySecondSplitter(1).ToString = "3" Or MySecondSplitter(1).ToString = "THREE" Then
                                    OpenedFileNr = 3
                                    NumberToString = "THREE"
                                ElseIf MySecondSplitter(1).ToString = "4" Or MySecondSplitter(1).ToString = "FOUR" Then
                                    OpenedFileNr = 4
                                    NumberToString = "FOUR"
                                ElseIf MySecondSplitter(1).ToString = "5" Or MySecondSplitter(1).ToString = "FIVE" Then
                                    OpenedFileNr = 5
                                    NumberToString = "FIVE"
                                ElseIf MySecondSplitter(1).ToString = "6" Or MySecondSplitter(1).ToString = "SIX" Then
                                    OpenedFileNr = 6
                                    NumberToString = "SIX"
                                ElseIf MySecondSplitter(1).ToString = "7" Or MySecondSplitter(1).ToString = "SEVEN" Then
                                    OpenedFileNr = 7
                                    NumberToString = "SEVEN"
                                ElseIf MySecondSplitter(1).ToString = "8" Or MySecondSplitter(1).ToString = "EIGHT" Then
                                    OpenedFileNr = 8
                                    NumberToString = "EIGHT"
                                ElseIf MySecondSplitter(1).ToString = "9" Or MySecondSplitter(1).ToString = "NINE" Then
                                    OpenedFileNr = 9
                                    NumberToString = "NINE"
                                ElseIf MySecondSplitter(1).ToString = "10" Or MySecondSplitter(1).ToString = "TEN" Then
                                    OpenedFileNr = 10
                                    NumberToString = "TEN"
                                ElseIf MySecondSplitter(1).ToString = "11" Or MySecondSplitter(1).ToString = "ELEVEN" Then
                                    OpenedFileNr = 11
                                    NumberToString = "ELEVEN"
                                ElseIf MySecondSplitter(1).ToString = "12" Or MySecondSplitter(1).ToString = "TWELVE" Then
                                    OpenedFileNr = 12
                                    NumberToString = "TWELVE"
                                ElseIf MySecondSplitter(1).ToString = "13" Or MySecondSplitter(1).ToString = "THIRTEEN" Then
                                    OpenedFileNr = 13
                                    NumberToString = "THIRTEEN"
                                ElseIf MySecondSplitter(1).ToString = "14" Or MySecondSplitter(1).ToString = "FOURTEEN" Then
                                    OpenedFileNr = 14
                                    NumberToString = "FOURTEEN"
                                ElseIf MySecondSplitter(1).ToString = "15" Or MySecondSplitter(1).ToString = "FIFTEEN" Then
                                    OpenedFileNr = 15
                                    NumberToString = "FIFTEEN"
                                ElseIf MySecondSplitter(1).ToString = "16" Or MySecondSplitter(1).ToString = "SIXTEEN" Then
                                    OpenedFileNr = 16
                                    NumberToString = "SIXTEEN"
                                ElseIf MySecondSplitter(1).ToString = "17" Or MySecondSplitter(1).ToString = "SEVENTEEN" Then
                                    OpenedFileNr = 17
                                    NumberToString = "SEVENTEEN"
                                ElseIf MySecondSplitter(1).ToString = "18" Or MySecondSplitter(1).ToString = "EIGHTEEN" Then
                                    OpenedFileNr = 18
                                    NumberToString = "EIGHTEEN"
                                ElseIf MySecondSplitter(1).ToString = "19" Or MySecondSplitter(1).ToString = "NINETEEN" Then
                                    OpenedFileNr = 19
                                    NumberToString = "NINETEEN"
                                ElseIf MySecondSplitter(1).ToString = "20" Or MySecondSplitter(1).ToString = "TWENTY" Then
                                    OpenedFileNr = 20
                                    NumberToString = "TWENTY"
                                ElseIf MySecondSplitter(1).ToString = "21" Or MySecondSplitter(1).ToString = "TWENTYONE" Then
                                    OpenedFileNr = 21
                                    NumberToString = "TWENTYONE"
                                ElseIf MySecondSplitter(1).ToString = "22" Or MySecondSplitter(1).ToString = "TWENTYTWO" Then
                                    OpenedFileNr = 22
                                    NumberToString = "TWENTYTWO"
                                ElseIf MySecondSplitter(1).ToString = "23" Or MySecondSplitter(1).ToString = "TWENTYTHREE" Then
                                    OpenedFileNr = 23
                                    NumberToString = "TWENTYTHREE"
                                ElseIf MySecondSplitter(1).ToString = "24" Or MySecondSplitter(1).ToString = "TWENTYFOUR" Then
                                    OpenedFileNr = 24
                                    NumberToString = "TWENTYFOUR"
                                ElseIf MySecondSplitter(1).ToString = "25" Or MySecondSplitter(1).ToString = "TWENTYFIVE" Then
                                    OpenedFileNr = 25
                                    NumberToString = "TWENTYFIVE"
                                ElseIf MySecondSplitter(1).ToString = "26" Or MySecondSplitter(1).ToString = "TWENTYSIX" Then
                                    OpenedFileNr = 26
                                    NumberToString = "TWENTYSIX"
                                ElseIf MySecondSplitter(1).ToString = "27" Or MySecondSplitter(1).ToString = "TWENTYSEVEN" Then
                                    OpenedFileNr = 27
                                    NumberToString = "TWENTYSEVEN"
                                ElseIf MySecondSplitter(1).ToString = "28" Or MySecondSplitter(1).ToString = "TWENTYEIGHT" Then
                                    OpenedFileNr = 28
                                    NumberToString = "TWENTYEIGHT"
                                ElseIf MySecondSplitter(1).ToString = "29" Or MySecondSplitter(1).ToString = "TWENTYNINE" Then
                                    OpenedFileNr = 29
                                    NumberToString = "TWENTYNINE"
                                ElseIf MySecondSplitter(1).ToString = "30" Or MySecondSplitter(1).ToString = "THIRTY" Then
                                    OpenedFileNr = 30
                                    NumberToString = "THIRTY"
                                ElseIf MySecondSplitter(1).ToString = "31" Or MySecondSplitter(1).ToString = "THIRTYONE" Then
                                    OpenedFileNr = 31
                                    NumberToString = "THIRTYONE"
                                ElseIf MySecondSplitter(1).ToString = "32" Or MySecondSplitter(1).ToString = "THIRTYTWO" Then
                                    OpenedFileNr = 32
                                    NumberToString = "THIRTYTWO"
                                ElseIf MySecondSplitter(1).ToString = "33" Or MySecondSplitter(1).ToString = "THIRTYTHREE" Then
                                    OpenedFileNr = 33
                                    NumberToString = "THIRTYTHREE"
                                ElseIf MySecondSplitter(1).ToString = "34" Or MySecondSplitter(1).ToString = "THIRTYFOUR" Then
                                    OpenedFileNr = 34
                                    NumberToString = "THIRTYFOUR"
                                ElseIf MySecondSplitter(1).ToString = "35" Or MySecondSplitter(1).ToString = "THIRTYFIVE" Then
                                    OpenedFileNr = 35
                                    NumberToString = "THIRTYFIVE"
                                ElseIf MySecondSplitter(1).ToString = "36" Or MySecondSplitter(1).ToString = "THIRTYSIX" Then
                                    OpenedFileNr = 36
                                    NumberToString = "THIRTYSIX"
                                ElseIf MySecondSplitter(1).ToString = "37" Or MySecondSplitter(1).ToString = "THIRTYSEVEN" Then
                                    OpenedFileNr = 37
                                    NumberToString = "THIRTYSEVEN"
                                ElseIf MySecondSplitter(1).ToString = "38" Or MySecondSplitter(1).ToString = "THIRTYEIGHT" Then
                                    OpenedFileNr = 38
                                    NumberToString = "THIRTYEIGHT"
                                ElseIf MySecondSplitter(1).ToString = "39" Or MySecondSplitter(1).ToString = "THIRTYNINE" Then
                                    OpenedFileNr = 39
                                    NumberToString = "THIRTYNINE"
                                ElseIf MySecondSplitter(1).ToString = "40" Or MySecondSplitter(1).ToString = "FORTY" Then
                                    OpenedFileNr = 40
                                    NumberToString = "FORTY"
                                ElseIf MySecondSplitter(1).ToString = "41" Or MySecondSplitter(1).ToString = "FORTYONE" Then
                                    OpenedFileNr = 41
                                    NumberToString = "FORTYONE"
                                ElseIf MySecondSplitter(1).ToString = "42" Or MySecondSplitter(1).ToString = "FORTYTWO" Then
                                    OpenedFileNr = 42
                                    NumberToString = "FORTYTWO"
                                ElseIf MySecondSplitter(1).ToString = "43" Or MySecondSplitter(1).ToString = "FORTYTHREE" Then
                                    OpenedFileNr = 43
                                    NumberToString = "FORTYTHREE"
                                ElseIf MySecondSplitter(1).ToString = "44" Or MySecondSplitter(1).ToString = "FORTYFOUR" Then
                                    OpenedFileNr = 44
                                    NumberToString = "FORTYFOUR"
                                ElseIf MySecondSplitter(1).ToString = "45" Or MySecondSplitter(1).ToString = "FORTYFIVE" Then
                                    OpenedFileNr = 45
                                    NumberToString = "FORTYFIVE"
                                ElseIf MySecondSplitter(1).ToString = "46" Or MySecondSplitter(1).ToString = "FORTYSIX" Then
                                    OpenedFileNr = 46
                                    NumberToString = "FORTYSIX"
                                ElseIf MySecondSplitter(1).ToString = "47" Or MySecondSplitter(1).ToString = "FORTYSEVEN" Then
                                    OpenedFileNr = 47
                                    NumberToString = "FORTYSEVEN"
                                ElseIf MySecondSplitter(1).ToString = "48" Or MySecondSplitter(1).ToString = "FORTYEIGHT" Then
                                    OpenedFileNr = 48
                                    NumberToString = "FORTYEIGHT"
                                ElseIf MySecondSplitter(1).ToString = "49" Or MySecondSplitter(1).ToString = "FORTYNINE" Then
                                    OpenedFileNr = 49
                                    NumberToString = "FORTYNINE"
                                ElseIf MySecondSplitter(1).ToString = "50" Or MySecondSplitter(1).ToString = "FIFTY" Then
                                    OpenedFileNr = 50
                                    NumberToString = "FIFTY"
                                ElseIf MySecondSplitter(1).ToString = "51" Or MySecondSplitter(1).ToString = "FIFTYONE" Then
                                    OpenedFileNr = 51
                                    NumberToString = "FIFTYONE"
                                ElseIf MySecondSplitter(1).ToString = "52" Or MySecondSplitter(1).ToString = "FIFTYTWO" Then
                                    OpenedFileNr = 52
                                    NumberToString = "FIFTYTWO"
                                ElseIf MySecondSplitter(1).ToString = "53" Or MySecondSplitter(1).ToString = "FIFTYTHREE" Then
                                    OpenedFileNr = 53
                                    NumberToString = "FIFTYTHREE"
                                ElseIf MySecondSplitter(1).ToString = "54" Or MySecondSplitter(1).ToString = "FIFTYFOUR" Then
                                    OpenedFileNr = 54
                                    NumberToString = "FIFTYFOUR"
                                ElseIf MySecondSplitter(1).ToString = "55" Or MySecondSplitter(1).ToString = "FIFTYFIVE" Then
                                    OpenedFileNr = 55
                                    NumberToString = "FIFTYFIVE"
                                ElseIf MySecondSplitter(1).ToString = "56" Or MySecondSplitter(1).ToString = "FIFTYSIX" Then
                                    OpenedFileNr = 56
                                    NumberToString = "FIFTYSIX"
                                ElseIf MySecondSplitter(1).ToString = "57" Or MySecondSplitter(1).ToString = "FIFTYSEVEN" Then
                                    OpenedFileNr = 57
                                    NumberToString = "FIFTYSEVEN"
                                ElseIf MySecondSplitter(1).ToString = "58" Or MySecondSplitter(1).ToString = "FIFTYEIGHT" Then
                                    OpenedFileNr = 58
                                    NumberToString = "FIFTYEIGHT"
                                ElseIf MySecondSplitter(1).ToString = "59" Or MySecondSplitter(1).ToString = "FIFTYNINE" Then
                                    OpenedFileNr = 59
                                    NumberToString = "FIFTYNINE"
                                ElseIf MySecondSplitter(1).ToString = "60" Or MySecondSplitter(1).ToString = "SIXTY" Then
                                    OpenedFileNr = 60
                                    NumberToString = "SIXTY"
                                ElseIf MySecondSplitter(1).ToString = "61" Or MySecondSplitter(1).ToString = "SIXTYONE" Then
                                    OpenedFileNr = 61
                                    NumberToString = "SIXTYONE"
                                ElseIf MySecondSplitter(1).ToString = "62" Or MySecondSplitter(1).ToString = "SIXTYTWO" Then
                                    OpenedFileNr = 62
                                    NumberToString = "SIXTYTWO"
                                ElseIf MySecondSplitter(1).ToString = "63" Or MySecondSplitter(1).ToString = "SIXTYTHREE" Then
                                    OpenedFileNr = 63
                                    NumberToString = "SIXTYTHREE"
                                ElseIf MySecondSplitter(1).ToString = "64" Or MySecondSplitter(1).ToString = "SIXTYFOUR" Then
                                    OpenedFileNr = 64
                                    NumberToString = "SIXTYFOUR"
                                ElseIf MySecondSplitter(1).ToString = "65" Or MySecondSplitter(1).ToString = "SIXTYFIVE" Then
                                    OpenedFileNr = 65
                                    NumberToString = "SIXTYFIVE"
                                ElseIf MySecondSplitter(1).ToString = "66" Or MySecondSplitter(1).ToString = "SIXTYSIX" Then
                                    OpenedFileNr = 66
                                    NumberToString = "SIXTYSIX"
                                ElseIf MySecondSplitter(1).ToString = "67" Or MySecondSplitter(1).ToString = "SIXTYSEVEN" Then
                                    OpenedFileNr = 67
                                    NumberToString = "SIXTYSEVEN"
                                ElseIf MySecondSplitter(1).ToString = "68" Or MySecondSplitter(1).ToString = "SIXTYEIGHT" Then
                                    OpenedFileNr = 68
                                    NumberToString = "SIXTYEIGHT"
                                ElseIf MySecondSplitter(1).ToString = "69" Or MySecondSplitter(1).ToString = "SIXTYNINE" Then
                                    OpenedFileNr = 69
                                    NumberToString = "SIXTYNINE"
                                ElseIf MySecondSplitter(1).ToString = "70" Or MySecondSplitter(1).ToString = "SEVENTY" Then
                                    OpenedFileNr = 70
                                    NumberToString = "SEVENTY"
                                ElseIf MySecondSplitter(1).ToString = "71" Or MySecondSplitter(1).ToString = "SEVENTYONE" Then
                                    OpenedFileNr = 71
                                    NumberToString = "SEVENTYONE"
                                ElseIf MySecondSplitter(1).ToString = "72" Or MySecondSplitter(1).ToString = "SEVENTYTWO" Then
                                    OpenedFileNr = 72
                                    NumberToString = "SEVENTYTWO"
                                ElseIf MySecondSplitter(1).ToString = "73" Or MySecondSplitter(1).ToString = "SEVENTYTHREE" Then
                                    OpenedFileNr = 73
                                    NumberToString = "SEVENTYTHREE"
                                ElseIf MySecondSplitter(1).ToString = "74" Or MySecondSplitter(1).ToString = "SEVENTYFOUR" Then
                                    OpenedFileNr = 74
                                    NumberToString = "SEVENTYFOUR"
                                ElseIf MySecondSplitter(1).ToString = "75" Or MySecondSplitter(1).ToString = "SEVENTYFIVE" Then
                                    OpenedFileNr = 75
                                    NumberToString = "SEVENTYFIVE"
                                End If
                                Exit For
                            End If
                        End If
                    Else
                        'nuova parte
                        If r = MySplitter.Count - 1 Then
                            Try
                                leggi = System.IO.File.OpenText(OpenFileDialog2.FileName)
                                letto = leggi.ReadToEnd
                                leggi.Close()
                                TextBox3.Text = letto
                                Exit Sub
                            Catch ex As Exception

                            End Try
                        End If
                        'fine nuova parte
                    End If
                Next
            '*** Transform String to Integer End ***

            '*** Check GDO Compatibility ***
            If OpenedFile <> 0 Then
                Dim MyThirdSplitter() As String = Split(letto, vbCrLf)
                Dim GDOline As String = MyThirdSplitter(0).ToString
                MyThirdSplitter = Split(TextBox3.Text, vbCrLf)
                If MyThirdSplitter(0).ToString <> GDOline Then
                    MsgBox("Sezione selezionata errata: GDO non corrispondente. Impossibile continuare.", MsgBoxStyle.Information, "Asc Server - Sezione Errata")
                    Exit Sub
                End If
            End If
            '*** Check GDO Compatibility End***

            If OpenedFileNr = OpenedFile + 1 Then
                If OpenedFile = 0 Then
                    Dim MySplitter2() As String = Split(letto, vbCrLf)
                    For r = 0 To UBound(MySplitter2)
                        If MySplitter2(r).ToString <> "BT" Then
                            TextBox3.Text += MySplitter2(r).ToString & vbCrLf
                        Else
                            If Btcounter = 0 Then
                                TextBox3.Text += MySplitter2(r).ToString & vbCrLf
                                Btcounter = 1
                            Else
                                OpenedFile += 1
                                Exit Sub
                            End If
                        End If
                    Next
                Else
                    Dim MySplitter2() As String = Split(letto, vbCrLf)
                    For r = 0 To UBound(MySplitter2)
                        If r <> MySplitter2.Count - 1 Then
                            If MySplitter2(r).ToString.Contains("SECTION " & NumberToString & " OF") Or MySplitter2(r).ToString.Contains("SECTION " & OpenedFileNr & " OF") _
                                Or MySplitter2(r).ToString.Contains("PARTE " & OpenedFileNr & " DI") Then
                                For z = r To UBound(MySplitter2)
                                    If MySplitter2(z).ToString <> "BT" Then
                                        TextBox3.Text += MySplitter2(z).ToString & vbCrLf
                                    Else
                                        OpenedFile += 1
                                        Exit Sub
                                    End If
                                Next
                            End If
                        End If
                    Next
                End If
            Else
                If FinalSectionNr = OpenedFile + 1 Then
                    Dim MySplitter2() As String = Split(letto, vbCrLf)
                    For r = 0 To UBound(MySplitter2)
                        If MySplitter2(r).ToString.Contains("FINAL SECTION OF") Or MySplitter(r).ToString.Contains("PARTE FINALE DI") Then
                            For z = r To UBound(MySplitter2)
                                TextBox3.Text += MySplitter2(z).ToString & vbCrLf
                            Next
                        End If
                    Next
                    OpenedFile = 0
                    FinalSectionPresent = 1
                    Exit Sub
                Else
                    MsgBox("Sezione selezionata errata: errore numerazione. Impossibile continuare.", MsgBoxStyle.Information, "Asc Server - Sezione Errata")
                End If
            End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem8.Click

        Uffici.Show()

    End Sub

    Private Sub MenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem13.Click

        MsgBox("Asc Server è un software realizzato da Rambod Rahmani. Questo programma è tutelato dalle norme internazionali sul Copyright: La riproduzione o distribuzione di esso, o parte di esso, è perseguibile penalmente fino alla massima pena in accordo alla viggente legislazione Italiana", MsgBoxStyle.Information, "Asc Server - V 2.2")

    End Sub

    Private Sub MenuItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem15.Click

        Process.Start("http://www.rambodrahmani.altervista.org")
        MsgBox("Per qualsiasi informazione rivolgersi a " & """" & "rambodrahmani@yahoo.it" & """", MsgBoxStyle.Information, "Asc Server - Info")

    End Sub

    Private Sub MenuItem19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem19.Click

        PrintSett.Show()

    End Sub

    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            If My.Settings.TipoDiDistribuzione <> "" Then
                If My.Settings.TipoDiDistribuzione = "Automatica" Then
                    MenuItem45.Checked = True
                    MenuItem47.Checked = False
                ElseIf My.Settings.TipoDiDistribuzione = "Manuale" Then
                    MenuItem47.Checked = True
                    MenuItem45.Checked = False
                End If
            Else
                MsgBox("Non è ancora stato scelto che tipo di distribuzione utilizzare. Si ricorda di farlo il prima possibile.", MsgBoxStyle.Information, "Asc Server - Tipo di Distribuzione")
            End If
            My.Settings.NumeroUfficioUltimo = 0
            If My.Settings.OraAggiornamento <> "" Then
                MenuItem40.Text = My.Settings.OraAggiornamento
            End If
            If My.Settings.AggiornamentoAutoSelect = True Then
                MenuItem40.Checked = True
            End If
            If My.Settings.IPAddress = "" Then
                MsgBox("Non è ancora stato inserito alcun indirizzo IP, per un corretto funzionamento di Asc Server si consiglia di inserirlo adesso.", MsgBoxStyle.Information, "Asc Server - Indirizzo IP Mancante")
                IPAddress.Show()
                IPAddress.TopMost = True
            End If
        Catch ex As Exception

        End Try

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
            LoadAscServer()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ResizeForm()

        Try
            If Me.WindowState <> FormWindowState.Minimized Then
                If MeOldWidth <> Me.Width Or MeOldHeihgt <> Me.Height Then
                    MeOldWidth = Me.Width
                    MeOldHeihgt = Me.Height

                    'x : 1024 = 36.42578125 : 100
                    Dim LarghezzaTxtBox1 As Double = 36.42578125
                    TextBox1.Width = ((Me.Width * LarghezzaTxtBox1) / 100)

                    Button2.Left = TextBox1.Right + 7
                    Button4.Left = Button2.Right + 7
                    Button3.Left = Button4.Left
                    Button3.Top = Button4.Top + 55
                    Button5.Left = Button3.Left - 88
                    Button5.Top = Button3.Top
                    Button11.Left = Button5.Left - 100
                    Button11.Top = Button5.Top

                    CheckBox2.Left = Button11.Left - 236
                    CheckBox2.Top = Button11.Top
                    CheckBox1.Left = Button11.Left - 236
                    CheckBox1.Top = CheckBox2.Top + 23

                    PictureBox1.Left = CheckBox2.Left - 90

                    'x : 1024 = 63.96484375 : 100
                    Dim LarghezzaTxtBox3 As Double = 63.96484375
                    TextBox3.Width = ((Me.Width * LarghezzaTxtBox3) / 100)
                    TextBox3.Height = Me.Height - 283

                    While (TextBox3.Right <> Button4.Right)
                        If TextBox3.Right > Button4.Right Then
                            TextBox3.Width -= 1
                        ElseIf TextBox3.Right < Button4.Right Then
                            TextBox3.Width += 1
                        End If
                    End While

                    Button1.Left = TextBox3.Right - 79
                    Button1.Top = TextBox3.Bottom + 8
                    Button9.Left = Button1.Left - 128
                    Button9.Top = Button1.Top
                    Button6.Left = Button9.Left - 87
                    Button6.Top = Button1.Top

                    Label5.Top = TextBox3.Bottom + 8
                    Label6.Top = TextBox3.Bottom + 8
                    Label7.Top = Label5.Bottom + 5
                    Label4.Top = Label6.Bottom + 5

                    GroupBox1.Left = TextBox3.Right + 34
                    GroupBox1.Width = (Me.Width - (TextBox3.Width + 75))
                    'x : 771 = 57.06874189364462 : 100
                    Dim AltezzaGrpBx1 As Double = 57.068741893644621
                    GroupBox1.Height = ((Me.Height * AltezzaGrpBx1) / 100)

                    CheckedListBox1.Left = Label1.Right + 3
                    CheckedListBox1.Top = Label1.Top
                    CheckedListBox1.Width = GroupBox1.Width - 35
                    CheckedListBox1.Height = GroupBox1.Height - 60
                    RectangleShape2.Width = CheckedListBox1.Width
                    Label8.Left = RectangleShape2.Left + 5

                    ListBox2.Top = CheckedListBox1.Top + 10
                    ListBox2.Left = CheckedListBox1.Left + 10
                    While (ListBox2.Bottom > CheckedListBox1.Bottom)
                        ListBox2.Height -= 2
                    End While
                    While (ListBox2.Right > CheckedListBox1.Right)
                        ListBox2.Width -= 2
                    End While

                    GroupBox2.Left = TextBox3.Right + 34
                    GroupBox2.Top = GroupBox1.Bottom + 6
                    GroupBox2.Width = (Me.Width - (TextBox3.Width + 75))
                    'x : 763 = 27.39187418086501 : 100
                    Dim AltezzaGrpBx2 As Double = 27.39187418086501
                    GroupBox2.Height = ((Me.Height * AltezzaGrpBx2) / 100)

                    ListView1.Width = GroupBox2.Width - 12
                    ListView1.Height = GroupBox2.Height - 62
                    ListView1.Columns(0).Width = ListView1.Width - 4

                    Button10.Left = ListView1.Left
                    Button10.Top = ListView1.Bottom + 6
                    Button8.Left = ListView1.Right - 79
                    Button8.Top = ListView1.Bottom + 6
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub LoadAscServer()

        Timer1.Start()

        IsPCNAvailable = "NO"
        chkAllegato = "NO"
        My.Settings.CloseDTBNow = 0
        My.Settings.SelectedDate = ""
        My.Settings.SelectedDate2 = ""
        My.Settings.Intelligence22Ter = ""

        My.Settings.DataOggi = Date.Now.Day.ToString & "." & Date.Now.Month.ToString & "." & Date.Now.Year.ToString

        If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEARRIVI") Then
        Else
            My.Settings.JlDate = Date.Now.DayOfYear
            My.Settings.DataIeri = My.Settings.DataOggi
            datagiorno = My.Settings.DataOggi
        End If

        If My.Settings.DataIeri <> My.Settings.DataOggi Then
            Dim Msgbox1 = MsgBox("Benvenuti in Asc, Volete creare la cartella per il nuovo giorno?", vbYesNo, "Asc Server - Benvenuti in Asc")
            If Msgbox1 = 6 Then
                If Date.Now.DayOfYear - My.Settings.JlDate > 1 Then
                    My.Settings.NumEmail = 0
                    My.Settings.NumEmail2 = 0
                    My.Settings.NumEmailIntelligence = 0
                    Dim Splitter() As String = Split(My.Settings.DataIeri, ".")
                    Dim GiorniMese As Integer = Date.DaysInMonth(Splitter(2).ToString, Splitter(1).ToString)

                    If GiorniMese = 31 Then
                        If Splitter(0).ToString <> "31" Then
                            My.Settings.DataOggi = (Val(Splitter(0).ToString) + 1).ToString & "." & Splitter(1).ToString & "." & Splitter(2).ToString
                            My.Settings.DataIeri = (Val(Splitter(0).ToString) + 1).ToString & "." & Splitter(1).ToString & "." & Splitter(2).ToString
                            datagiorno = My.Settings.DataOggi
                            My.Settings.JlDate = My.Settings.JlDate + 1
                            Label4.ForeColor = Color.Red
                            Label6.ForeColor = Color.Red
                            Label6.Text = datagiorno
                            Label4.Text = My.Settings.JlDate.ToString("D3")
                        Else
                            If Splitter(1).ToString <> "12" Then
                                My.Settings.DataOggi = "1." & (Val(Splitter(1).ToString) + 1).ToString & "." & Splitter(2).ToString
                                My.Settings.DataIeri = "1." & (Val(Splitter(1).ToString) + 1).ToString & "." & Splitter(2).ToString
                                datagiorno = My.Settings.DataOggi
                                My.Settings.JlDate = My.Settings.JlDate + 1
                                Label4.ForeColor = Color.Red
                                Label6.ForeColor = Color.Red
                                Label6.Text = datagiorno
                                Label4.Text = My.Settings.JlDate.ToString("D3")
                            Else
                                My.Settings.DataOggi = "1." & "1." & (Val(Splitter(2).ToString) + 1).ToString
                                My.Settings.DataIeri = "1." & "1." & (Val(Splitter(2).ToString) + 1).ToString
                                datagiorno = My.Settings.DataOggi
                                My.Settings.JlDate = My.Settings.JlDate + 1
                                Label4.ForeColor = Color.Red
                                Label6.ForeColor = Color.Red
                                Label6.Text = datagiorno
                                Label4.Text = My.Settings.JlDate.ToString("D3")

                            End If
                        End If
                    ElseIf GiorniMese = 30 Then
                        If Splitter(0).ToString <> "30" Then
                            My.Settings.DataOggi = (Val(Splitter(0).ToString) + 1).ToString & "." & Splitter(1).ToString & "." & Splitter(2).ToString
                            My.Settings.DataIeri = (Val(Splitter(0).ToString) + 1).ToString & "." & Splitter(1).ToString & "." & Splitter(2).ToString
                            datagiorno = My.Settings.DataOggi
                            My.Settings.JlDate = My.Settings.JlDate + 1
                            Label4.ForeColor = Color.Red
                            Label6.ForeColor = Color.Red
                            Label6.Text = datagiorno
                            Label4.Text = My.Settings.JlDate.ToString("D3")
                        Else
                            If Splitter(1).ToString <> "12" Then
                                My.Settings.DataOggi = "1." & (Val(Splitter(1).ToString) + 1).ToString & "." & Splitter(2).ToString
                                My.Settings.DataIeri = "1." & (Val(Splitter(1).ToString) + 1).ToString & "." & Splitter(2).ToString
                                datagiorno = My.Settings.DataOggi
                                My.Settings.JlDate = My.Settings.JlDate + 1
                                Label4.ForeColor = Color.Red
                                Label6.ForeColor = Color.Red
                                Label6.Text = datagiorno
                                Label4.Text = My.Settings.JlDate.ToString("D3")
                            Else
                                My.Settings.DataOggi = "1." & "1." & (Val(Splitter(2).ToString) + 1).ToString
                                My.Settings.DataIeri = "1." & "1." & (Val(Splitter(2).ToString) + 1).ToString
                                datagiorno = My.Settings.DataOggi
                                My.Settings.JlDate = My.Settings.JlDate + 1
                                Label4.ForeColor = Color.Red
                                Label6.ForeColor = Color.Red
                                Label6.Text = datagiorno
                                Label4.Text = My.Settings.JlDate.ToString("D3")
                            End If
                        End If
                    ElseIf GiorniMese = 29 Then
                        If Splitter(0).ToString <> "29" Then
                            My.Settings.DataOggi = (Val(Splitter(0).ToString) + 1).ToString & "." & Splitter(1).ToString & "." & Splitter(2).ToString
                            My.Settings.DataIeri = (Val(Splitter(0).ToString) + 1).ToString & "." & Splitter(1).ToString & "." & Splitter(2).ToString
                            datagiorno = My.Settings.DataOggi
                            My.Settings.JlDate = My.Settings.JlDate + 1
                            Label4.ForeColor = Color.Red
                            Label6.ForeColor = Color.Red
                            Label6.Text = datagiorno
                            Label4.Text = My.Settings.JlDate.ToString("D3")
                        Else
                            If Splitter(1).ToString <> "12" Then
                                My.Settings.DataOggi = "1." & (Val(Splitter(1).ToString) + 1).ToString & "." & Splitter(2).ToString
                                My.Settings.DataIeri = "1." & (Val(Splitter(1).ToString) + 1).ToString & "." & Splitter(2).ToString
                                datagiorno = My.Settings.DataOggi
                                My.Settings.JlDate = My.Settings.JlDate + 1
                                Label4.ForeColor = Color.Red
                                Label6.ForeColor = Color.Red
                                Label6.Text = datagiorno
                                Label4.Text = My.Settings.JlDate.ToString("D3")
                            Else
                                My.Settings.DataOggi = "1." & "1." & (Val(Splitter(2).ToString) + 1).ToString
                                My.Settings.DataIeri = "1." & "1." & (Val(Splitter(2).ToString) + 1).ToString
                                datagiorno = My.Settings.DataOggi
                                My.Settings.JlDate = My.Settings.JlDate + 1
                                Label4.ForeColor = Color.Red
                                Label6.ForeColor = Color.Red
                                Label6.Text = datagiorno
                                Label4.Text = My.Settings.JlDate.ToString("D3")
                            End If
                        End If
                    End If

                Else
                    My.Settings.NumEmail = 0
                    My.Settings.NumEmail2 = 0
                    My.Settings.NumEmailIntelligence = 0
                    My.Settings.DataIeri = My.Settings.DataOggi
                    datagiorno = My.Settings.DataOggi
                    My.Settings.JlDate = Date.Now.DayOfYear
                End If
            Else
                My.Settings.DataOggi = My.Settings.DataIeri
                Label4.ForeColor = Color.Red
                Label6.ForeColor = Color.Red
                datagiorno = My.Settings.DataOggi
                Label6.Text = datagiorno
                Label4.Text = My.Settings.JlDate.ToString("D3")
            End If
        End If

        datagiorno = My.Settings.DataOggi

        Label4.Text = My.Settings.JlDate.ToString("D3")

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\Email") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email")
            End If
            'DATABASE ARRIVI
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEARRIVI") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\DATABASEARRIVI")
            End If
            'DATABASE PARTENZE
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\DATABASEPARTENZE")
            End If
            'DATABASE INTELLIGENCE
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\INTELLIGENCE") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\INTELLIGENCE")
            End If
        Catch ex As Exception

        End Try

        Try
            'DATABASE ARRIVI
            If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\dateusate.txt") Then
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Email\DATABASEARRIVI\dateusate.txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                If letto.Contains(datagiorno) Then
                Else
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEARRIVI\dateusate.txt")
                    scrivi.WriteLine(datagiorno)
                    scrivi.Close()
                End If
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEARRIVI\dateusate.txt")
                scrivi.WriteLine(datagiorno)
                scrivi.Close()
            End If
            'DATABASE PARTENZE
            If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\dateusate.txt") Then
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Email\DATABASEPARTENZE\dateusate.txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                If letto.Contains(datagiorno) Then
                Else
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEPARTENZE\dateusate.txt")
                    scrivi.WriteLine(datagiorno)
                    scrivi.Close()
                End If
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEPARTENZE\dateusate.txt")
                scrivi.WriteLine(datagiorno)
                scrivi.Close()
            End If
            'DATABASE INTELLIGENCE
            If System.IO.File.Exists(Application.StartupPath & "\Email\INTELLIGENCE\dateusate.txt") Then
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Email\INTELLIGENCE\dateusate.txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                If letto.Contains(datagiorno) Then
                Else
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\INTELLIGENCE\dateusate.txt")
                    scrivi.WriteLine(datagiorno)
                    scrivi.Close()
                End If
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\INTELLIGENCE\dateusate.txt")
                scrivi.WriteLine(datagiorno)
                scrivi.Close()
            End If
        Catch ex As Exception

        End Try

        Try
            'DATABASE ARRIVI
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno) Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno)
                My.Settings.NumEmail = 0
            End If
            'DATABASE PARTENZE
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno) Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno)
                My.Settings.NumEmail2 = 0
            End If
            'DATABASE INTELLIGENCE
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno) Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno)
                My.Settings.NumEmailIntelligence = 0
            End If
        Catch ex As Exception

        End Try

        Try
            'DATABASE ARRIVI
            If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt") Then
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                scrivi.Close()
            End If
            'DATABASE PARTENZE
            If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt") Then
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                scrivi.Close()
            End If
            'DATABASE INTELLIGENCE
            If System.IO.File.Exists(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt") Then
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                scrivi.Close()
            End If
        Catch ex As Exception

        End Try

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\Uffici") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Uffici")
                My.Settings.Ufficio = 0
            End If

            If System.IO.File.Exists(Application.StartupPath & "\Uffici\listauffici.txt") Then
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Uffici\listauffici.txt")
                scrivi.Close()
            End If
        Catch ex As Exception

        End Try

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\temp") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\temp")
            End If
        Catch ex As Exception

        End Try

        LoadContacts()

        If My.Settings.MenuItem32 = True Then
            MenuItem32.Checked = True
        Else
            MenuItem32.Checked = False
        End If
        If My.Settings.MenuItem33 = True Then
            MenuItem33.Checked = True
        Else
            MenuItem33.Checked = False
        End If

        RefreshListBox()

        SaveProgressiveNumbers()

        SaveOthersProgressive()

        Timer4.Start()

    End Sub

    Public Sub ConnectClient()

        If MenuItem42.Text = "Connetti" Then
            If My.Settings.IPAddress <> "" Then
                Try
                    client = New tcpConnection(My.Settings.IPAddress, CInt(My.Settings.PortaConnessione))
                    MenuItem42.Text = "Disconnetti"
                    client.Send(Requests.StringMessage, "ServerConnected")
                    If Timer3.Enabled = False Then
                        Timer3.Start()
                    End If
                    If Timer2.Enabled = False Then
                        Timer2.Start()
                    End If
                    If Timer6.Enabled = False Then
                        Timer6.Start()
                    End If
                    Try
                        RectangleShape1.Width = 93
                        RectangleShape1.BackColor = Color.Green
                        RectangleShape1.BorderColor = Color.Green
                    Catch ex As Exception

                    End Try
                    Timer5.Start()
                    Label8.Text = "LISTA UFFICI"
                Catch ex As Exception
                    MenuItem42.Text = "Connetti"
                    MsgBox("Impossibile connetersi al server. Errore Rilevato: " & ex.Message, MsgBoxStyle.Critical, "Asc Server - Errore Connessione")
                End Try
            Else
                MsgBox("Non è stato possibile tentare la connessione poichè non è stato ancora inserito alcun Indirizzo IP. Impostare L'indirizzo IP per poter continuare.", MsgBoxStyle.Information, "Asc Client - Indirizzo IP Mancante")
                IPAddress.Show()
            End If
        ElseIf MenuItem42.Text = "Disconnetti" Then
            Try
                client.Send(Requests.StringMessage, "ServerDisconnected")
            Catch ex As Exception

            End Try
            Try
                MenuItem42.Text = "Connetti"
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
                If Timer3.Enabled = True Then
                    Timer3.Stop()
                End If
                If Timer2.Enabled = True Then
                    Timer2.Stop()
                End If
                If Timer6.Enabled = True Then
                    Timer6.Stop()
                End If
            Catch ex As Exception

            End Try
            Try
                RectangleShape1.Width = 79
                RectangleShape1.BackColor = Color.Red
                RectangleShape1.BorderColor = Color.Red
            Catch ex As Exception

            End Try
            Timer5.Start()
            Label8.Text = "LISTA UFFICI  [Connettersi ad Exchange]"
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        TerSelectionDate.Show()

    End Sub

    Public Sub LoadPrivateList()

        For r = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SelectedIndex = r
            If CheckedListBox1.GetItemChecked(r) = True Then
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                infoufficio = Split(letto, "#")
                For z = 0 To UBound(infoufficio) - 2
                    If infoufficio(z + 1) <> "" Then
                        ListBox2.Items.Add(infoufficio(z + 1))
                    End If
                Next
            End If
        Next

    End Sub

    '*** Inizio invio email privata ***
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click

        Try
            If MenuItem42.Text = "Connetti" Then
                If CheckedListBox1.SelectedItems.Count > 0 Then
                    For r = 0 To CheckedListBox1.Items.Count - 1
                        CheckedListBox1.SelectedIndex = r
                        If CheckedListBox1.GetItemChecked(r) = True Then
                            leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                            letto = leggi.ReadToEnd
                            leggi.Close()
                            infoufficio = Split(letto, "#")
                            For z = 0 To UBound(infoufficio) - 2
                                If infoufficio(z + 1) <> "" Then
                                    ListBox2.Items.Add(infoufficio(z + 1))
                                End If
                            Next
                        End If
                    Next

                    int2 = ListBox2.Items.Count
                    For r = 0 To int2 - 1
                        ListBox2.SetSelected(r, True)
                        If ListBox2.SelectedItem.ToString.Contains("@") Then
                            'Next
                        Else
                            MsgBox("Negli uffici selezionati è stato rilevato l'indirizzo di una macchina [Indirizzo Rilevato Non Idoneo: " & """" & ListBox2.SelectedItem.ToString & """" & "], poichè Asc Server non risulta connesso ad Asc Exchange è impossibile continuare. Invio Annullato.", MsgBoxStyle.Critical, "Asc Server - Errore Invio")
                            Exit Sub
                        End If
                    Next

                    If TextBox3.Text <> "" Then
                        If OpenedFile = 0 Then
                            ListBox2.Items.Clear()
                            PrivateEmailDistribution()
                        Else
                            If FinalSectionPresent = 1 Then
                                ListBox2.Items.Clear()
                                PrivateEmailDistribution()
                            Else
                                Dim AlertMessage = MsgBox("L'inserimento di tutte le sezioni non è stato completato. Manca la parte finale, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Server - Sezione Finale Mancante")
                                If AlertMessage = 6 Then
                                    ListBox2.Items.Clear()
                                    PrivateEmailDistribution()
                                End If
                            End If
                        End If
                    Else
                        MsgBox("Impossibile inviare il messaggio: testo del messaggio mancante. Invio Annullato.", MsgBoxStyle.Information, "Asc Server - Testo Mancante")
                    End If
                Else
                    MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                End If
            Else
                If TextBox3.Text <> "" Then
                    If OpenedFile = 0 Then
                        ListBox2.Items.Clear()
                        PrivateEmailDistribution()
                    Else
                        If FinalSectionPresent = 1 Then
                            ListBox2.Items.Clear()
                            PrivateEmailDistribution()
                        Else
                            Dim AlertMessage = MsgBox("L'inserimento di tutte le sezioni non è stato completato. Manca la parte finale, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Server - Sezione Finale Mancante")
                            If AlertMessage = 6 Then
                                ListBox2.Items.Clear()
                                PrivateEmailDistribution()
                            End If
                        End If
                    End If
                Else
                    MsgBox("Impossibile inviare il messaggio: testo del messaggio mancante. Invio Annullato.", MsgBoxStyle.Information, "Asc Server - Testo Mancante")
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub PrivateEmailDistribution()

        FileName = "SVC"

        If CheckedListBox1.CheckedItems.Count < 1 Then
            MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
            Exit Sub
        End If

        LoadPrivateList()
        StatusBar1.Text = "Stato Invio: Invio email privata in corso"

        Try
            ListBox2.SetSelected(0, True)
            int2 = ListBox2.Items.Count
            For r = 0 To int2 - 1
                ListBox2.SetSelected(r, True)
                If ListBox2.SelectedItem.ToString.Contains("@") Then
                    SendPrivateEmailToEmailAddress()
                Else
                    SendPrivateEmailToMachineAddress()
                End If

                If r = int2 - 1 Then
                    StatusBar1.Text = "Stato Invio: Invio completato con successo"
                    MsgBox("Invio Completato, Email inviata a " & ListBox2.Items.Count & " indirizzi.", MsgBoxStyle.Information, "Asc Server - Invio Completato")
                    OpenFileDialog1.FileName = ""
                    chkAllegato = "NO"
                    If Timer2.Enabled = False Then
                        Timer2.Start()
                    End If
                    CleanAll()
                End If
            Next
        Catch ex As Exception
            MsgBox("Errore invio email privata, messaggio ex: " & ex.Message, MsgBoxStyle.Critical, "Errore Invio")
        End Try

    End Sub

    Public Sub SendPrivateEmailToEmailAddress()

        Try
            Dim objOutlook As Microsoft.Office.Interop.Outlook.Application
            Dim objEmail As Microsoft.Office.Interop.Outlook.MailItem
            objOutlook = CType(CreateObject("Outlook.Application"), Microsoft.Office.Interop.Outlook.Application)
            objEmail = objOutlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
            With objEmail
                .Subject = TextBox1.Text
                .To = ListBox2.SelectedItem.ToString
                .Body = TextBox3.Text
                If OpenFileDialog1.FileName <> "" Then
                    .Attachments.Add(OpenFileDialog1.FileName)
                End If
                .Send()
            End With
        Catch ex As Exception
            MsgBox("Errore invio email privata, messaggio ex: " & ex.Message, MsgBoxStyle.Critical, "Asc Server - Errore Invio")
        End Try

    End Sub

    Public Sub SendPrivateEmailToMachineAddress()

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\temppartenze") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\temppartenze")
            End If
            Dim scrittore As System.IO.StreamWriter
            If System.IO.File.Exists(Application.StartupPath & "\temppartenze\temppartenza.txt") Then
                For r = 1 To 9999
                    If System.IO.File.Exists(Application.StartupPath & "\temppartenze\temppartenza(" & r.ToString & ").txt") Then
                    Else
                        scrittore = System.IO.File.CreateText(Application.StartupPath & "\temppartenze\temppartenza(" & r.ToString & ").txt")
                        scrittore.Write("MessaggioDalServer" & vbCrLf)
                        scrittore.Write(ListBox2.SelectedItem.ToString & vbCrLf)
                        scrittore.Write(FileName & vbCrLf)
                        scrittore.Write(My.Settings.JlDate.ToString("D3") & vbCrLf)
                        scrittore.Write(TextBox3.Text)
                        scrittore.Close()
                        Exit For
                    End If
                Next
            Else
                scrittore = System.IO.File.CreateText(Application.StartupPath & "\temppartenze\temppartenza.txt")
                scrittore.Write("MessaggioDalServer" & vbCrLf)
                scrittore.Write(ListBox2.SelectedItem.ToString & vbCrLf)
                scrittore.Write(FileName & vbCrLf)
                scrittore.Write(My.Settings.JlDate.ToString("D3") & vbCrLf)
                scrittore.Write(TextBox3.Text)
                scrittore.Close()
            End If
        Catch ex As Exception
            MsgBox("Errore invio email privata, messaggio ex: " & ex.Message, MsgBoxStyle.Critical, "Asc Server - Errore Invio")
        End Try

    End Sub
    '*** Fine invio email privata ***

    '**** INIZIO EMAIL PCN ***
    Public Sub TrattaPCN()

        Try
            PCNCompleteEmailText = ""
            PCNPartedEmailText = ""

            If TextBox3.Text <> "" Then
                If OpenedFile = 0 Then
                    If CheckedListBox1.CheckedItems.Count < 1 Then
                        MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                        Exit Sub
                    End If

                    StatusBar1.Text = "Stato Invio: Invio email PCN iniziato"

                    IsPCNAvailable = "SI"

                    Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
                    If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                        Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                        Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                        My.Settings.NumEmail2 += 1
                        FileName = "T" & Date.Now.Year.ToString & My.Settings.JlDate.ToString("D3") & My.Settings.NumEmail2.ToString("D4")
                        CreateTempFiles()
                    Else
                        My.Settings.NumEmail += 1
                        FileName = "R" & Date.Now.Year.ToString & My.Settings.JlDate.ToString("D3") & My.Settings.NumEmail.ToString("D4")
                    End If

                    ListBox2.Items.Clear()

                    Try
                        MSGIDdef = ""
                        MSGID = Split(TextBox3.Text, vbCrLf)
                        For r = 0 To MSGID.Count - 1
                            If MSGID(r).Contains("MSGID") Or MSGID(r).Contains("OGGETTO") Or MSGID(r).Contains("SUBJ") _
                                Or MSGID(r).Contains("SUBJECT") Or MSGID(r).Contains("ARG") Or MSGID(r).Contains("OPER") Then
                                MSGIDdef = MSGID(r)
                            End If
                        Next
                    Catch ex As Exception

                    End Try

                    If MSGIDdef <> "" Then
                        TaglioBt()
                    Else
                        Dim AlertMessage = MsgBox("Non è stato trovato alcun campo MSGID, si desidera proseguire comunque?", MsgBoxStyle.YesNo, "Asc Server - MSGID Mancante")
                        If AlertMessage = 6 Then
                            TaglioBt()
                        End If
                    End If
                Else
                    If FinalSectionPresent = 1 Then
                        If CheckedListBox1.CheckedItems.Count < 1 Then
                            MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                            Exit Sub
                        End If

                        StatusBar1.Text = "Stato Invio: Invio email PCN iniziato"

                        IsPCNAvailable = "SI"

                        Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
                        If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                            Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                            Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                            My.Settings.NumEmail2 += 1
                            FileName = "T" & Date.Now.Year.ToString & My.Settings.JlDate.ToString("D3") & My.Settings.NumEmail2.ToString("D4")
                            CreateTempFiles()
                        Else
                            My.Settings.NumEmail += 1
                            FileName = "R" & Date.Now.Year.ToString & My.Settings.JlDate.ToString("D3") & My.Settings.NumEmail.ToString("D4")
                        End If

                        ListBox2.Items.Clear()

                        Try
                            MSGIDdef = ""
                            MSGID = Split(TextBox3.Text, vbCrLf)
                            For r = 0 To MSGID.Count - 1
                                If MSGID(r).Contains("MSGID") Or MSGID(r).Contains("OGGETTO") Or MSGID(r).Contains("SUBJ") _
                                    Or MSGID(r).Contains("SUBJECT") Or MSGID(r).Contains("ARG") Or MSGID(r).Contains("OPER") Then
                                    MSGIDdef = MSGID(r)
                                End If
                            Next
                        Catch ex As Exception

                        End Try

                        If MSGIDdef <> "" Then
                            TaglioBt()
                        Else
                            Dim AlertMessage = MsgBox("Non è stato trovato alcun campo MSGID, si desidera proseguire comunque?", MsgBoxStyle.YesNo, "Asc Server - MSGID Mancante")
                            If AlertMessage = 6 Then
                                TaglioBt()
                            End If
                        End If
                    Else
                        Dim AlertMessage2 = MsgBox("L'inserimento di tutte le sezioni non è stato completato. Manca la parte finale, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Server - Sezione Finale Mancante")
                        If AlertMessage2 = 6 Then
                            If CheckedListBox1.CheckedItems.Count < 1 Then
                                MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                                Exit Sub
                            End If

                            StatusBar1.Text = "Stato Invio: Invio email PCN iniziato"

                            IsPCNAvailable = "SI"

                            Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
                            If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                                Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                                Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                                My.Settings.NumEmail2 += 1
                                FileName = "T" & Date.Now.Year.ToString & My.Settings.JlDate.ToString("D3") & My.Settings.NumEmail2.ToString("D4")
                                CreateTempFiles()
                            Else
                                My.Settings.NumEmail += 1
                                FileName = "R" & Date.Now.Year.ToString & My.Settings.JlDate.ToString("D3") & My.Settings.NumEmail.ToString("D4")
                            End If

                            ListBox2.Items.Clear()

                            Try
                                MSGIDdef = ""
                                MSGID = Split(TextBox3.Text, vbCrLf)
                                For r = 0 To MSGID.Count - 1
                                    If MSGID(r).Contains("MSGID") Or MSGID(r).Contains("OGGETTO") Or MSGID(r).Contains("SUBJ") _
                                        Or MSGID(r).Contains("SUBJECT") Or MSGID(r).Contains("ARG") Or MSGID(r).Contains("OPER") Then
                                        MSGIDdef = MSGID(r)
                                    End If
                                Next
                            Catch ex As Exception

                            End Try

                            If MSGIDdef <> "" Then
                                TaglioBt()
                            Else
                                Dim AlertMessage = MsgBox("Non è stato trovato alcun campo MSGID, si desidera proseguire comunque?", MsgBoxStyle.YesNo, "Asc Server - MSGID Mancante")
                                If AlertMessage = 6 Then
                                    TaglioBt()
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                MsgBox("Impossibile inviare il messaggio: testo del messaggio mancante. Invio Annullato.", MsgBoxStyle.Information, "Asc Server - Testo Mancante")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Errore durante il Tratta PCN, ex message: " & ex.Message.ToString, MsgBoxStyle.Critical, "Asc Server - Errore")
        End Try

    End Sub

    Public Sub TaglioBt()

        Try
            Dim lstview1selectedindex As Integer
            Dim PositiveRow As Integer = 0
            Dim nomedasplittare As String = ""
            If ListView1LastSelectedItem <> "NONSELEZIONATO" Then
                lstview1selectedindex = Val(ListView1LastSelectedItem)
                nomedasplittare = ListView1.Items(lstview1selectedindex).Text.ToString & ".txt"
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Messaggi Ricevuti\" & nomedasplittare)
                letto = leggi.ReadToEnd
                leggi.Close()
                Dim TxtFileSplitter() As String = Split(letto, vbCrLf)
                For r = 0 To UBound(TxtFileSplitter)
                    If TextBox3.Text.Contains(TxtFileSplitter(r).ToString) Then
                        PositiveRow += 1
                    End If
                Next
            End If

            Dim TextBox3RowSplitter() As String = Split(TextBox3.Text, vbCrLf)
            If PositiveRow > TextBox3RowSplitter.Count - 3 Then
                ListView1.Items(lstview1selectedindex).ForeColor = Color.Green
                If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\UsedIndicies.txt") Then
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Messaggi Ricevuti\UsedIndicies.txt")
                    scrivi.Write(nomedasplittare & "#")
                    scrivi.Close()
                Else
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\UsedIndicies.txt")
                    scrivi.Write(nomedasplittare & "#")
                    scrivi.Close()
                End If
            End If
        Catch ex As Exception

        End Try

        Dim MySplitter3() As String = Split(TextBox3.Text, vbCrLf)
        If MySplitter3(1).ToString = "FM ITS " & My.Settings.NomeUnita _
            Or MySplitter3(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
            Or MySplitter3(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
            Try
                contatoreDecret = 0
                TextBox3.Text += vbCrLf & vbCrLf & "Email inviata a: "
                For z = 0 To CheckedListBox1.Items.Count - 1
                    CheckedListBox1.SelectedIndex = z
                    If CheckedListBox1.GetItemChecked(z) = True Then
                        If CheckedListBox1.Items(z).ToString.Contains("[") Then
                            Dim InstantSplitter() As String = Split(CheckedListBox1.Items(z).ToString, "  [")
                            TextBox3.Text += InstantSplitter(0).ToString & ", "
                            contatoreDecret += 1
                            If contatoreDecret = 4 Then
                                TextBox3.Text += vbCrLf
                                contatoreDecret = 0
                            End If
                        Else
                            TextBox3.Text += CheckedListBox1.SelectedItem.ToString & ", "
                            contatoreDecret += 1
                            If contatoreDecret = 4 Then
                                TextBox3.Text += vbCrLf
                                contatoreDecret = 0
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
            End Try
            TextBox3.Text += vbCrLf & "Nr. Progressivo Email: " & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3")
            TextBox3.Text += vbCrLf & "Il: " & Date.Now.Day.ToString & "." & Date.Now.Month.ToString & "." & Date.Now.Year.ToString & "  Alle: " & ora
            TextBox3.Text += vbCrLf & "CHIAVE: " & FileName
        Else
            Try
                contatoreDecret = 0
                TextBox3.Text += vbCrLf & vbCrLf & "Email inviata a: "
                For z = 0 To CheckedListBox1.Items.Count - 1
                    CheckedListBox1.SelectedIndex = z
                    If CheckedListBox1.GetItemChecked(z) = True Then
                        If CheckedListBox1.Items(z).ToString.Contains("[") Then
                            Dim InstantSplitter() As String = Split(CheckedListBox1.Items(z).ToString, "  [")
                            TextBox3.Text += InstantSplitter(0).ToString & ", "
                            contatoreDecret += 1
                            If contatoreDecret = 4 Then
                                TextBox3.Text += vbCrLf
                                contatoreDecret = 0
                            End If
                        Else
                            TextBox3.Text += CheckedListBox1.SelectedItem.ToString & ", "
                            contatoreDecret += 1
                            If contatoreDecret = 4 Then
                                TextBox3.Text += vbCrLf
                                contatoreDecret = 0
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
            End Try
            TextBox3.Text += vbCrLf & "Nr. Progressivo Email: " & My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3")
            TextBox3.Text += vbCrLf & "Il: " & Date.Now.Day.ToString & "." & Date.Now.Month.ToString & "." & Date.Now.Year.ToString & "  Alle: " & ora
            TextBox3.Text += vbCrLf & "CHIAVE: " & FileName
        End If

        Try
            Dim MySplitter2() As String = Split(TextBox3.Text, vbCrLf)
            If MySplitter2(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                Or MySplitter2(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                Or MySplitter2(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Alta Classifica") Then
                    If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Alta Classifica\" & My.Settings.DataOggi) Then
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Alta Classifica\" & My.Settings.DataOggi & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt")
                        scrivi.Write(TextBox3.Text)
                        scrivi.Close()
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Alta Classifica\" & My.Settings.DataOggi)
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Alta Classifica\" & My.Settings.DataOggi & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt")
                        scrivi.Write(TextBox3.Text)
                        scrivi.Close()
                    End If
                Else
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Alta Classifica")
                    If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Alta Classifica\" & My.Settings.DataOggi) Then
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Alta Classifica\" & My.Settings.DataOggi & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt")
                        scrivi.Write(TextBox3.Text)
                        scrivi.Close()
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Alta Classifica\" & My.Settings.DataOggi)
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Alta Classifica\" & My.Settings.DataOggi & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt")
                        scrivi.Write(TextBox3.Text)
                        scrivi.Close()
                    End If
                End If
            Else
                If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Alta Classifica") Then
                    If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Alta Classifica\" & My.Settings.DataOggi) Then
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Alta Classifica\" & My.Settings.DataOggi & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt")
                        scrivi.Write(TextBox3.Text)
                        scrivi.Close()
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Alta Classifica\" & My.Settings.DataOggi)
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Alta Classifica\" & My.Settings.DataOggi & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt")
                        scrivi.Write(TextBox3.Text)
                        scrivi.Close()
                    End If
                Else
                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Alta Classifica")
                    If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Alta Classifica\" & My.Settings.DataOggi) Then
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Alta Classifica\" & My.Settings.DataOggi & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt")
                        scrivi.Write(TextBox3.Text)
                        scrivi.Close()
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Alta Classifica\" & My.Settings.DataOggi)
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Alta Classifica\" & My.Settings.DataOggi & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt")
                        scrivi.Write(TextBox3.Text)
                        scrivi.Close()
                    End If
                End If
            End If
        Catch ex As Exception

        End Try

        PCNCompleteEmailText = TextBox3.Text

        Try
            Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
            TextBox3.Clear()
            For r = 0 To UBound(MySplitter)
                If MySplitter(r).ToString = "BT" Then
                    TextBox3.Text += MySplitter(r).ToString & vbCrLf
                    For z = r + 1 To UBound(MySplitter)
                        If z = r + 6 Then
                            TextBox3.Text += MySplitter(z).ToString & vbCrLf
                            TextBox3.Text += "=================TESTO FIGURATIVO=================" & vbCrLf & "BT" & vbCrLf
                            PCNPartedEmailText = TextBox3.Text
                            LoadPCNEmail()
                            Exit Sub
                        Else
                            TextBox3.Text += MySplitter(z).ToString & vbCrLf
                        End If
                    Next
                Else
                    TextBox3.Text += MySplitter(r).ToString & vbCrLf
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Sub LoadPCNEmail()

        Dim MySplitter() As String = Split(TextBox3.Text, vbCrLf)
        If MySplitter(1).ToString = "FM ITS " & My.Settings.NomeUnita _
            Or MySplitter(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
            Or MySplitter(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
            Try
                contatoreDecret = 0
                TextBox3.Text += vbCrLf & vbCrLf & "Email inviata a: "
                For z = 0 To CheckedListBox1.Items.Count - 1
                    CheckedListBox1.SelectedIndex = z
                    If CheckedListBox1.GetItemChecked(z) = True Then
                        If CheckedListBox1.Items(z).ToString.Contains("[") Then
                            Dim InstantSplitter() As String = Split(CheckedListBox1.Items(z).ToString, "  [")
                            TextBox3.Text += InstantSplitter(0).ToString & ", "
                            contatoreDecret += 1
                            If contatoreDecret = 4 Then
                                TextBox3.Text += vbCrLf
                                contatoreDecret = 0
                            End If
                        Else
                            TextBox3.Text += CheckedListBox1.SelectedItem.ToString & ", "
                            contatoreDecret += 1
                            If contatoreDecret = 4 Then
                                TextBox3.Text += vbCrLf
                                contatoreDecret = 0
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
            End Try
            TextBox3.Text += vbCrLf & "Nr. Progressivo Email: " & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3")
            TextBox3.Text += vbCrLf & "Il: " & Date.Now.Day.ToString & "." & Date.Now.Month.ToString & "." & Date.Now.Year.ToString & "  Alle: " & ora
            TextBox3.Text += vbCrLf & "CHIAVE: " & FileName
        Else
            Try
                contatoreDecret = 0
                TextBox3.Text += vbCrLf & vbCrLf & "Email inviata a: "
                For z = 0 To CheckedListBox1.Items.Count - 1
                    CheckedListBox1.SelectedIndex = z
                    If CheckedListBox1.GetItemChecked(z) = True Then
                        If CheckedListBox1.Items(z).ToString.Contains("[") Then
                            Dim InstantSplitter() As String = Split(CheckedListBox1.Items(z).ToString, "  [")
                            TextBox3.Text += InstantSplitter(0).ToString & ", "
                            contatoreDecret += 1
                            If contatoreDecret = 4 Then
                                TextBox3.Text += vbCrLf
                                contatoreDecret = 0
                            End If
                        Else
                            TextBox3.Text += CheckedListBox1.SelectedItem.ToString & ", "
                            contatoreDecret += 1
                            If contatoreDecret = 4 Then
                                TextBox3.Text += vbCrLf
                                contatoreDecret = 0
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
            End Try
            TextBox3.Text += vbCrLf & "Nr. Progressivo Email: " & My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3")
            TextBox3.Text += vbCrLf & "Il: " & Date.Now.Day.ToString & "." & Date.Now.Month.ToString & "." & Date.Now.Year.ToString & "  Alle: " & ora
            TextBox3.Text += vbCrLf & "CHIAVE: " & FileName
        End If

        StatusBar1.Text = "Stato Invio: Caricamento lista destinatari"
        ListBox2.Items.Clear()

        If CheckedListBox1.Items.Count > 0 Then
            CheckedListBox1.SelectedIndex = 0
            riga = Split(TextBox3.Text, vbCrLf)

            Dim MySplitter2() As String = Split(TextBox3.Text, vbCrLf)
            If MySplitter2(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                Or MySplitter2(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                Or MySplitter2(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                riga = Split(TextBox3.Text, vbCrLf)
                If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt") Then
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                    scrivi.Write(My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & "#")
                    scrivi.Close()
                Else
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                    scrivi.Write(My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & "#")
                    scrivi.Close()
                End If

                If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt") Then
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt")
                    scrivi.Write(TextBox3.Text)
                    scrivi.Close()
                Else
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt")
                    scrivi.Write(TextBox3.Text)
                    scrivi.Close()
                End If

                For r = 0 To CheckedListBox1.Items.Count - 1
                    CheckedListBox1.SelectedIndex = r
                    If CheckedListBox1.GetItemChecked(r) = True Then
                        If CheckedListBox1.Items(r).ToString.Contains("[") Then
                            Dim InstantSplitter() As String = Split(CheckedListBox1.Items(r).ToString, "  [")
                            scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                            scrivi.Write(InstantSplitter(0).ToString & "/")
                            scrivi.Close()
                        Else
                            scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                            scrivi.Write(CheckedListBox1.SelectedItem.ToString & "/")
                            scrivi.Close()
                        End If

                        leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                        letto = leggi.ReadToEnd
                        leggi.Close()
                        infoufficio = Split(letto, "#")
                        For z = 0 To UBound(infoufficio) - 2
                            If infoufficio(z + 1) <> "" Then
                                ListBox2.Items.Add(infoufficio(z + 1))
                            End If
                        Next
                    End If
                Next

                scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                scrivi.Write("#" & MSGIDdef & "#" & datagiorno & " " & ora & "#" & chkAllegato & "#" & riga(0) & "#" & riga(1) & "#" & vbCrLf)
                scrivi.Close()

                If My.Settings.TipoDiDistribuzione = "Manuale" Then
                    For z = 0 To CheckedListBox1.Items.Count - 1
                        CheckedListBox1.SelectedIndex = z
                        If CheckedListBox1.GetItemChecked(z) = True Then
                            If CheckedListBox1.Items(z).ToString <> "" Then
                                If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "P" & My.Settings.JlDate.ToString("D3")) Then
                                    If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "P" & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt") Then
                                    Else
                                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "P" & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt")
                                        scrivi.Write(TextBox3.Text)
                                        scrivi.Close()
                                    End If
                                Else
                                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "P" & My.Settings.JlDate.ToString("D3"))
                                    If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "P" & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt") Then
                                    Else
                                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "P" & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmail2.ToString("D4") & "P.txt")
                                        scrivi.Write(TextBox3.Text)
                                        scrivi.Close()
                                    End If
                                End If
                            End If
                        End If
                    Next

                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\TestoEmail.txt")
                    scrivi.Write(TextBox3.Text)
                    scrivi.Close()

                    StatusBar1.Text = "Invio Emails completato con successo."
                    MsgBox("Invio Completato, Email inviata a " & ListBox2.Items.Count & " indirizzi.", MsgBoxStyle.Information, "Asc Server - Invio Completato")

                    If MenuItem33.Checked = True Then
                        StampaPCNCompleta()
                    Else
                        If MenuItem32.Checked = True Then
                            StampaPCNParted()
                        End If
                    End If

                    SaveProgressiveNumbers()

                    OpenFileDialog1.FileName = ""
                    chkAllegato = "NO"
                    IsPCNAvailable = "NO"

                    Dim MySplitter3() As String = Split(TextBox3.Text, vbCrLf)
                    If MySplitter3(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                        Or MySplitter3(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                        Or MySplitter3(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                        Try
                            System.IO.File.Delete(Application.StartupPath & "\" & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & ".txt")
                        Catch ex As Exception

                        End Try
                    Else
                        Try
                            System.IO.File.Delete(Application.StartupPath & "\" & My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3") & ".txt")
                        Catch ex As Exception

                        End Try
                    End If
                    If Timer2.Enabled = False Then
                        Timer2.Start()
                    End If
                    CleanAll()
                    Exit Sub
                End If

            Else

                riga = Split(TextBox3.Text, vbCrLf)
                If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt") Then
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                    scrivi.Write(My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3") & "#")
                    scrivi.Close()
                Else
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                    scrivi.Write(My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3") & "#")
                    scrivi.Close()
                End If

                If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt") Then
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt")
                    scrivi.Write(TextBox3.Text)
                    scrivi.Close()
                Else
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt")
                    scrivi.Write(TextBox3.Text)
                    scrivi.Close()
                End If

                For r = 0 To CheckedListBox1.Items.Count - 1
                    CheckedListBox1.SelectedIndex = r
                    If CheckedListBox1.GetItemChecked(r) = True Then
                        If CheckedListBox1.Items(r).ToString.Contains("[") Then
                            Dim InstantSplitter() As String = Split(CheckedListBox1.Items(r).ToString, "  [")
                            scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                            scrivi.Write(InstantSplitter(0).ToString & "/")
                            scrivi.Close()
                        Else
                            scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                            scrivi.Write(CheckedListBox1.SelectedItem.ToString & "/")
                            scrivi.Close()
                        End If

                        leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                        letto = leggi.ReadToEnd
                        leggi.Close()
                        infoufficio = Split(letto, "#")
                        For z = 0 To UBound(infoufficio) - 2
                            If infoufficio(z + 1) <> "" Then
                                ListBox2.Items.Add(infoufficio(z + 1))
                            End If
                        Next
                    End If
                Next

                scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                scrivi.Write("#" & MSGIDdef & "#" & datagiorno & " " & ora & "#" & chkAllegato & "#" & riga(0) & "#" & riga(1) & "#" & IsPCNAvailable & "#" & vbCrLf)
                scrivi.Close()

                If My.Settings.TipoDiDistribuzione = "Manuale" Then
                    For z = 0 To CheckedListBox1.Items.Count - 1
                        CheckedListBox1.SelectedIndex = z
                        If CheckedListBox1.GetItemChecked(z) = True Then
                            If CheckedListBox1.Items(z).ToString <> "" Then
                                If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "A" & My.Settings.JlDate.ToString("D3")) Then
                                    If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "A" & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt") Then
                                    Else
                                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "A" & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt")
                                        scrivi.Write(TextBox3.Text)
                                        scrivi.Close()
                                    End If
                                Else
                                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "A" & My.Settings.JlDate.ToString("D3"))
                                    If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "A" & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt") Then
                                    Else
                                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & "A" & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmail.ToString("D4") & "A.txt")
                                        scrivi.Write(TextBox3.Text)
                                        scrivi.Close()
                                    End If
                                End If
                            End If
                        End If
                    Next

                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\TestoEmail.txt")
                    scrivi.Write(TextBox3.Text)
                    scrivi.Close()

                    StatusBar1.Text = "Invio Emails completato con successo."
                    MsgBox("Invio Completato, Email inviata a " & ListBox2.Items.Count & " indirizzi.", MsgBoxStyle.Information, "Asc Server - Invio Completato")

                    If MenuItem33.Checked = True Then
                        StampaPCNCompleta()
                    Else
                        If MenuItem32.Checked = True Then
                            StampaPCNParted()
                        End If
                    End If

                    SaveProgressiveNumbers()

                    OpenFileDialog1.FileName = ""
                    chkAllegato = "NO"
                    IsPCNAvailable = "NO"

                    Dim MySplitter3() As String = Split(TextBox3.Text, vbCrLf)
                    If MySplitter3(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                        Or MySplitter3(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                        Or MySplitter3(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                        Try
                            System.IO.File.Delete(Application.StartupPath & "\" & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & ".txt")
                        Catch ex As Exception

                        End Try
                    Else
                        Try
                            System.IO.File.Delete(Application.StartupPath & "\" & My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3") & ".txt")
                        Catch ex As Exception

                        End Try
                    End If
                    If Timer2.Enabled = False Then
                        Timer2.Start()
                    End If
                    CleanAll()
                    Exit Sub
                End If
            End If
        End If

        Try
            StatusBar1.Text = "Stato Invio: Controllo allegati"
            scrivi = System.IO.File.CreateText(Application.StartupPath & "\" & My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3") & ".txt")
            scrivi.Write(TextBox3.Text)
            scrivi.Close()
        Catch ex As Exception

        End Try

        TrattaPCNDistribution()

    End Sub

    Public Sub TrattaPCNDistribution()

        Try
            ListBox2.SetSelected(0, True)
            int2 = ListBox2.Items.Count
            StatusBar1.Text = "Stato Invio: Invio iniziato, smistamento"
            Dim TestoDiRiferimento As String = TextBox3.Text
            For r = 0 To int2 - 1
                ListBox2.SetSelected(r, True)
                If ListBox2.SelectedItem.ToString.Contains("@") Then
                    PCNSendEmailToEmailAddress()
                Else
                    PCNSendEmailToMachineAddress()
                End If

                If r = int2 - 1 Then
                    StatusBar1.Text = "Invio Emails completato con successo."
                    MsgBox("Invio Completato, Email inviata a " & ListBox2.Items.Count & " indirizzi.", MsgBoxStyle.Information, "Asc Server - Invio Completato")

                    If MenuItem33.Checked = True Then
                        StampaPCNCompleta()
                    Else
                        If MenuItem32.Checked = True Then
                            StampaPCNParted()
                        End If
                    End If
                    SaveProgressiveNumbers()

                    OpenFileDialog1.FileName = ""
                    chkAllegato = "NO"
                    IsPCNAvailable = "NO"
                    Dim MySplitter3() As String = Split(TestoDiRiferimento, vbCrLf)
                    If MySplitter3(1).ToString = "FM ITS " & My.Settings.NomeUnita _
                        Or MySplitter3(1).ToString = "FM NAVE " & My.Settings.NomeUnita _
                        Or MySplitter3(1).ToString = "FM SMG " & My.Settings.NomeUnita Then
                        Try
                            System.IO.File.Delete(Application.StartupPath & "\" & My.Settings.NumEmail2.ToString("D4") & "P - " & My.Settings.JlDate.ToString("D3") & ".txt")
                        Catch ex As Exception

                        End Try
                    Else
                        Try
                            System.IO.File.Delete(Application.StartupPath & "\" & My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3") & ".txt")
                        Catch ex As Exception

                        End Try
                    End If
                    If Timer2.Enabled = False Then
                        Timer2.Start()
                    End If
                    CleanAll()
                End If
            Next
        Catch ex As Exception
            MsgBox("Errore invio email PCN, messaggio ex: " & ex.Message, MsgBoxStyle.Critical, "Errore Invio")
        End Try

    End Sub

    Public Sub PCNSendEmailToEmailAddress()

        Try
            Dim objOutlook As Microsoft.Office.Interop.Outlook.Application
            Dim objEmail As Microsoft.Office.Interop.Outlook.MailItem
            objOutlook = CType(CreateObject("Outlook.Application"), Microsoft.Office.Interop.Outlook.Application)
            objEmail = objOutlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
            With objEmail
                .Subject = TextBox1.Text
                .To = ListBox2.SelectedItem.ToString
                .Body = TextBox3.Text
                .Attachments.Add(Application.StartupPath & "\" & My.Settings.NumEmail.ToString("D4") & "A - " & My.Settings.JlDate.ToString("D3") & ".txt")
                If OpenFileDialog1.FileName <> "" Then
                    .Attachments.Add(OpenFileDialog1.FileName)
                End If
                .Send()
            End With
        Catch ex As Exception
            MsgBox("Errore invio email PCN, messaggio ex: " & ex.Message, MsgBoxStyle.Critical, "Asc Server - Errore Invio")
        End Try

    End Sub

    Public Sub PCNSendEmailToMachineAddress()

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\temppartenze") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\temppartenze")
            End If
            Dim scrittore As System.IO.StreamWriter
            If System.IO.File.Exists(Application.StartupPath & "\temppartenze\temppartenza.txt") Then
                For r = 1 To 9999
                    If System.IO.File.Exists(Application.StartupPath & "\temppartenze\temppartenza(" & r.ToString & ").txt") Then
                    Else
                        scrittore = System.IO.File.CreateText(Application.StartupPath & "\temppartenze\temppartenza(" & r.ToString & ").txt")
                        scrittore.Write("MessaggioDalServer" & vbCrLf)
                        scrittore.Write(ListBox2.SelectedItem.ToString & vbCrLf)
                        scrittore.Write(FileName & vbCrLf)
                        scrittore.Write(TextBox3.Text)
                        scrittore.Close()
                        Exit For
                    End If
                Next
            Else
                scrittore = System.IO.File.CreateText(Application.StartupPath & "\temppartenze\temppartenza.txt")
                scrittore.Write("MessaggioDalServer" & vbCrLf)
                scrittore.Write(ListBox2.SelectedItem.ToString & vbCrLf)
                scrittore.Write(FileName & vbCrLf)
                scrittore.Write(TextBox3.Text)
                scrittore.Close()
            End If
        Catch ex As Exception
            MsgBox("Errore invio email PCN, messaggio ex: " & ex.Message, MsgBoxStyle.Critical, "Asc Server - Errore Invio")
        End Try

    End Sub

    Public Sub StampaPCNCompleta()

        Try
            RigheTextBox = Split(PCNCompleteEmailText, vbCrLf)
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

        IniziaStampaPCN()

    End Sub
    Public Sub IniziaStampaPCN()

        If CheckBox2.Checked = True Then
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
                    leggi = System.IO.File.OpenText(Application.StartupPath & "\fogliostampa" & r & ".txt")
                    letto = leggi.ReadToEnd
                    leggi.Close()
                    dastampare = letto
                    PrintDocument1.Print()
                Next
            Catch ex As Exception

            End Try
        Else
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
        End If

        Try
            For r = 0 To NumFoglioStampa
                System.IO.File.Delete(Application.StartupPath & "\fogliostampa" & r & ".txt")
                NumFoglioStampa = 0
                numeroriga = 0
                stoppami = 0
                FinalSectionPresent = 0
                OpenedFile = 0
            Next
        Catch ex As Exception

        End Try

        If ManualPrint = 0 Then
            If MenuItem32.Checked = True Then
                StampaPCNParted()
            Else
                CleanAll()
            End If
        Else
            ManualPrint = 0
        End If

    End Sub

    Public Sub StampaPCNParted()

        Try
            RigheTextBox = Split(PCNPartedEmailText, vbCrLf)
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

        IniziaStampaPCNParted()

    End Sub
    Public Sub IniziaStampaPCNParted()

        If CheckBox2.Checked = True Then
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
                    leggi = System.IO.File.OpenText(Application.StartupPath & "\fogliostampa" & r & ".txt")
                    letto = leggi.ReadToEnd
                    leggi.Close()
                    dastampare = letto
                    PrintDocument1.Print()
                Next
            Catch ex As Exception

            End Try
        Else
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
        End If

        Try
            For r = 0 To NumFoglioStampa
                System.IO.File.Delete(Application.StartupPath & "\fogliostampa" & r & ".txt")
                NumFoglioStampa = 0
                numeroriga = 0
                stoppami = 0
                FinalSectionPresent = 0
                OpenedFile = 0
            Next
        Catch ex As Exception

        End Try

        If ManualPrint = 0 Then
            CleanAll()
        Else
            ManualPrint = 0
        End If
        CheckBox1.Checked = False
        CheckBox2.Checked = False

    End Sub
    '**** FINE PCN EMAIL CODING ***

    Private Sub MenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem16.Click

        SaveFileDialog1.Filter = "Text File | *.txt"
        If SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            scrivi = System.IO.File.CreateText(SaveFileDialog1.FileName)
            scrivi.Write(TextBox3.Text)
            scrivi.Close()
        End If

    End Sub

    Private Sub MenuItem17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem17.Click

        Numerazione.Show()

    End Sub

    Private Sub TextBox3_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.GotFocus

        If My.Computer.Keyboard.CapsLock = False Then
            keybd_event(Windows.Forms.Keys.Capital, 0, 0, 0)
        End If

    End Sub

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown

        If My.Computer.Keyboard.CapsLock = False Then
            keybd_event(Windows.Forms.Keys.CapsLock, 0, 0, 0)
        End If

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click

        If MenuItem42.Text = "Connetti" Then
            MsgBox("Asc Server non risulta connesso ad Asc Exchange, impossibile continuare. Connettersi ad Asc Exchange e riprovare.", MsgBoxStyle.Information, "Asc Server - Connessione ad Asc Exchange")
        Else
            Try
                ConnectedClientsList.ListBox1.Items.Clear()
                client.Send(Requests.StringMessage, "ServerRequestsConnectedClientList")
            Catch ex As Exception

            End Try

            Try
                ConnectedClientsList.Close()
            Catch ex As Exception

            End Try
            Try
                ConnectedClientsList.Show()
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click

        If ListView1.SelectedItems.Count > 0 Then
            Try
                Dim lstviewselecteditem As ListViewItem
                For Each lstviewselecteditem In ListView1.SelectedItems
                    ListView1.Items.RemoveAt(lstviewselecteditem.Index)
                    System.IO.File.Delete(Application.StartupPath & "\Messaggi Ricevuti\" & lstviewselecteditem.Text.ToString & ".txt")
                Next
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\UsedIndicies.txt")
                scrivi.Close()
                For r = 0 To ListView1.Items.Count - 1
                    If ListView1.Items(r).ForeColor = Color.Green Then
                        scrivi = System.IO.File.AppendText(Application.StartupPath & "\Messaggi Ricevuti\UsedIndicies.txt")
                        scrivi.Write(ListView1.Items(r).Text.ToString & ".txt#")
                        scrivi.Close()
                    End If
                Next
            Catch ex As Exception
                MsgBox("Errore eliminazione messaggio, ex message:" & ex.Message, MsgBoxStyle.Critical, "Asc Server - Errore")
            End Try
        Else
            MsgBox("Non è stato selezionato alcun messaggio da eliminare, impossibile continuare.", MsgBoxStyle.Information, "Asc Server - Eliminazione Messaggio")
        End If

        RefreshListBox()

    End Sub

    Public Sub RefreshListBox()

        Try
            ListView1.Items.Clear()
            Dim MyPrimarySplitter() As String

            Try
                If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\UsedIndicies.txt") Then
                    leggi = System.IO.File.OpenText(Application.StartupPath & "\Messaggi Ricevuti\UsedIndicies.txt")
                    letto = leggi.ReadToEnd
                    leggi.Close()
                    MyPrimarySplitter = Split(letto, "#")
                End If
            Catch ex As Exception

            End Try

            Try
                Dim FilesCollection() As String
                FilesCollection = System.IO.Directory.GetFiles(Application.StartupPath & "\Messaggi Ricevuti")
                If FilesCollection.Count > 0 Then
                    For r = 0 To UBound(FilesCollection)
                        Dim FileInfo As New FileInfo(FilesCollection(r).ToString)
                        If FileInfo.Name.ToString <> "" And FileInfo.Name.ToString <> "UsedIndicies.txt" Then
                            Dim Splitter() As String = Split(FileInfo.Name.ToString, ".txt")
                            ListView1.Items.Add(Splitter(0).ToString)
                        End If
                    Next
                End If
            Catch ex As Exception

            End Try

            For y = 0 To ListView1.Items.Count - 1
                ListView1.Items(y).Selected = True
                Dim lstviewselecteditem As ListViewItem
                For Each lstviewselecteditem In ListView1.SelectedItems
                    Dim nomedasplittare As String = lstviewselecteditem.ToString & "{"
                    Dim MyLstViewItemSplitter() As String = Split(nomedasplittare.ToString, "{")
                    Dim MySecondSplitter() As String = Split(MyLstViewItemSplitter(1), "}")
                    If MyPrimarySplitter.Count > 0 Then
                        For p = 0 To UBound(MyPrimarySplitter)
                            If MySecondSplitter(0) & ".txt" = MyPrimarySplitter(p).ToString Then
                                lstviewselecteditem.ForeColor = Color.Green
                            End If
                        Next
                    End If
                Next
            Next
        Catch ex As Exception
        End Try

        Try
            Dim lstbox1items As Integer = ListView1.Items.Count
            ListView1.Columns(0).Text = lstbox1items & " messaggi in partenza"
        Catch ex As Exception

        End Try

    End Sub

    Public Sub OpenSelectedFile()

        Try
            Dim lstviewselecteditem As ListViewItem
            For Each lstviewselecteditem In ListView1.SelectedItems
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Messaggi Ricevuti\" & lstviewselecteditem.Text.ToString & ".txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                TextBox3.Text = letto
                ListView1LastSelectedItem = lstviewselecteditem.Index.ToString
            Next
        Catch ex As Exception
            MsgBox("Errore durante l'apertura di un messaggio ricevuto dai clients, errore rilevato: " & ex.Message.ToString, MsgBoxStyle.Critical, "Asc Server - Errore")
        End Try

    End Sub

    Private Sub ListView1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListView1.KeyDown

        If e.KeyCode = Keys.Enter Then
            OpenSelectedFile()
        End If

    End Sub

    Private Sub ListView1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListView1.MouseDoubleClick

        If e.Button = Windows.Forms.MouseButtons.Left Then
            OpenSelectedFile()
        End If

    End Sub

    '*** STARTING - INTELLIGENCE - CODING ***
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

        Try
            If My.Settings.TipoDiDistribuzione = "Automatica" Then
                If MenuItem42.Text = "Connetti" Then
                    If CheckedListBox1.SelectedItems.Count > 0 Then
                        For r = 0 To CheckedListBox1.Items.Count - 1
                            CheckedListBox1.SelectedIndex = r
                            If CheckedListBox1.GetItemChecked(r) = True Then
                                leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                                letto = leggi.ReadToEnd
                                leggi.Close()
                                infoufficio = Split(letto, "#")
                                For z = 0 To UBound(infoufficio) - 2
                                    If infoufficio(z + 1) <> "" Then
                                        ListBox2.Items.Add(infoufficio(z + 1))
                                    End If
                                Next
                            End If
                        Next

                        int2 = ListBox2.Items.Count
                        For r = 0 To int2 - 1
                            ListBox2.SetSelected(r, True)
                            If ListBox2.SelectedItem.ToString.Contains("@") Then
                                'Next
                            Else
                                MsgBox("Negli uffici selezionati è stato rilevato l'indirizzo di una macchina [Indirizzo Rilevato Non Idoneo: " & """" & ListBox2.SelectedItem.ToString & """" & "], poichè Asc Server non risulta connesso ad Asc Exchange è impossibile continuare. Invio Annullato.", MsgBoxStyle.Critical, "Asc Server - Errore Invio")
                                Exit Sub
                            End If
                        Next

                        ListBox2.Items.Clear()

                        If TextBox3.Text <> "" Then
                            If OpenedFile = 0 Then
                                Try
                                    If CheckedListBox1.CheckedItems.Count < 1 Then
                                        MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                                        Exit Sub
                                    End If

                                    StatusBar1.Text = "Stato Invio: Invio messaggio INTELLIGENCE iniziato"

                                    IntelligenceFindMSGID()
                                Catch ex As Exception

                                End Try
                            Else
                                If FinalSectionPresent = 1 Then
                                    Try
                                        If CheckedListBox1.CheckedItems.Count < 1 Then
                                            MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                                            Exit Sub
                                        End If

                                        StatusBar1.Text = "Stato Invio: Invio messaggio INTELLIGENCE iniziato"

                                        IntelligenceFindMSGID()
                                    Catch ex As Exception

                                    End Try
                                Else
                                    Dim AlertMessage = MsgBox("L'inserimento di tutte le sezioni non è stato completato. Manca la parte finale, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Server - Sezione Finale Mancante")
                                    If AlertMessage = 6 Then
                                        Try
                                            If CheckedListBox1.CheckedItems.Count < 1 Then
                                                MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                                                Exit Sub
                                            End If

                                            StatusBar1.Text = "Stato Invio: Invio messaggio INTELLIGENCE iniziato"

                                            IntelligenceFindMSGID()
                                        Catch ex As Exception

                                        End Try
                                    End If
                                End If
                            End If
                        Else
                            MsgBox("Impossibile inviare il messaggio: testo del messaggio mancante. Invio Annullato.", MsgBoxStyle.Information, "Asc Server - Testo Mancante")
                        End If
                    Else
                        MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                    End If
                Else
                    If TextBox3.Text <> "" Then
                        If OpenedFile = 0 Then
                            Try
                                If CheckedListBox1.CheckedItems.Count < 1 Then
                                    MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                                    Exit Sub
                                End If

                                StatusBar1.Text = "Stato Invio: Invio messaggio INTELLIGENCE iniziato"

                                IntelligenceFindMSGID()
                            Catch ex As Exception

                            End Try
                        Else
                            If FinalSectionPresent = 1 Then
                                Try
                                    If CheckedListBox1.CheckedItems.Count < 1 Then
                                        MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                                        Exit Sub
                                    End If

                                    StatusBar1.Text = "Stato Invio: Invio messaggio INTELLIGENCE iniziato"

                                    IntelligenceFindMSGID()
                                Catch ex As Exception

                                End Try
                            Else
                                Dim AlertMessage = MsgBox("L'inserimento di tutte le sezioni non è stato completato. Manca la parte finale, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Server - Sezione Finale Mancante")
                                If AlertMessage = 6 Then
                                    Try
                                        If CheckedListBox1.CheckedItems.Count < 1 Then
                                            MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                                            Exit Sub
                                        End If

                                        StatusBar1.Text = "Stato Invio: Invio messaggio INTELLIGENCE iniziato"

                                        IntelligenceFindMSGID()
                                    Catch ex As Exception

                                    End Try
                                End If
                            End If
                        End If
                    Else
                        MsgBox("Impossibile inviare il messaggio: testo del messaggio mancante. Invio Annullato.", MsgBoxStyle.Information, "Asc Server - Testo Mancante")
                    End If
                End If
            ElseIf My.Settings.TipoDiDistribuzione = "Manuale" Then
                If TextBox3.Text <> "" Then
                    If OpenedFile = 0 Then
                        Try
                            If CheckedListBox1.CheckedItems.Count < 1 Then
                                MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                                Exit Sub
                            End If

                            StatusBar1.Text = "Stato Invio: Invio messaggio INTELLIGENCE iniziato"

                            IntelligenceFindMSGID()
                        Catch ex As Exception

                        End Try
                    Else
                        If FinalSectionPresent = 1 Then
                            Try
                                If CheckedListBox1.CheckedItems.Count < 1 Then
                                    MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                                    Exit Sub
                                End If

                                StatusBar1.Text = "Stato Invio: Invio messaggio INTELLIGENCE iniziato"

                                IntelligenceFindMSGID()
                            Catch ex As Exception

                            End Try
                        Else
                            Dim AlertMessage = MsgBox("L'inserimento di tutte le sezioni non è stato completato. Manca la parte finale, si desidera continuare comunque?", MsgBoxStyle.YesNo, "Asc Server - Sezione Finale Mancante")
                            If AlertMessage = 6 Then
                                Try
                                    If CheckedListBox1.CheckedItems.Count < 1 Then
                                        MsgBox("Non è stato selezionato alcun ufficio, selezionare almeno un ufficio per poter continuare.", MsgBoxStyle.Information, "Asc Server - Uffici non Selezionati")
                                        Exit Sub
                                    End If

                                    StatusBar1.Text = "Stato Invio: Invio messaggio INTELLIGENCE iniziato"

                                    IntelligenceFindMSGID()
                                Catch ex As Exception

                                End Try
                            End If
                        End If
                    End If
                Else
                    MsgBox("Impossibile inviare il messaggio: testo del messaggio mancante. Invio Annullato.", MsgBoxStyle.Information, "Asc Server - Testo Mancante")
                End If
            Else
                MsgBox("Non è ancora stata selezionata alcun tipo di distribuzione: Selezionare un tipo di distribuzione dal menu " & """" & "Distribuzione" & """" & " per poter continaure.", MsgBoxStyle.Critical, "Asc Server - Tipo di Distribuzione")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub IntelligenceFindMSGID()

        Try
            MSGIDdef = ""
            MSGID = Split(TextBox3.Text, vbCrLf)
            For r = 0 To MSGID.Count - 1
                If MSGID(r).Contains("MSGID") Or MSGID(r).Contains("OGGETTO") Or MSGID(r).Contains("SUBJ") _
                    Or MSGID(r).Contains("SUBJECT") Or MSGID(r).Contains("ARG") Or MSGID(r).Contains("OPER") Then
                    MSGIDdef = MSGID(r)
                End If
            Next
        Catch ex As Exception

        End Try

        If MSGIDdef <> "" Then
            IntelligenceChkProgressiveNumber()
        Else
            Dim AlertMessage = MsgBox("Non è stato trovato alcun campo MSGID, si desidera proseguire comunque?", MsgBoxStyle.YesNo, "Asc Server - MSGID Mancante")
            If AlertMessage = 6 Then
                IntelligenceChkProgressiveNumber()
            End If
        End If

    End Sub

    Public Sub IntelligenceChkProgressiveNumber()

        Dim filescollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno)
        If filescollection.Count > 1 Then

            If filescollection.Count - 1 > My.Settings.NumEmailIntelligence Then
                MsgBox("Rilevato errore nella numerazione progressiva in Intelligence. Invio annullato.", MsgBoxStyle.Information, "Asc Server - Errore Numerazione")
                StatusBar1.Text = "Invio Annullato. Errore numerazione progressiva in Intelligence rilevato"
                InvioAnnullato = 1
                Exit Sub
            End If

            If filescollection.Count - 1 < My.Settings.NumEmailIntelligence Then
                MsgBox("Rilevato errore nella numerazione progressiva in Intelligence. Invio annullato.", MsgBoxStyle.Information, "Asc Server - Errore Numerazione")
                StatusBar1.Text = "Invio Annullato. Errore numerazione progressiva in Intelligence rilevato"
                InvioAnnullato = 1
                Exit Sub
            End If

            For r = 1 To My.Settings.NumEmailIntelligence
                If System.IO.File.Exists(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\emailinviate" & r.ToString("D4") & "INTEL.txt") Then

                Else
                    MsgBox("Rilevato errore nella numerazione progressiva in Intelligence. Invio annullato.", MsgBoxStyle.Information, "Asc Server - Errore Numerazione")
                    StatusBar1.Text = "Invio Annullato. Errore numerazione progressiva in Intelligence rilevato"
                    InvioAnnullato = 1
                    Exit Sub
                End If
                If r = My.Settings.NumEmailIntelligence Then
                    My.Settings.NumEmailIntelligence += 1
                    FileName = "I" & Date.Now.Year.ToString & My.Settings.JlDate.ToString("D3") & My.Settings.NumEmailIntelligence.ToString("D4")
                End If
            Next
        Else
            My.Settings.NumEmailIntelligence += 1
            FileName = "I" & Date.Now.Year.ToString & My.Settings.JlDate.ToString("D3") & My.Settings.NumEmailIntelligence.ToString("D4")
        End If

        LoadIntelligenceList()

    End Sub

    Public Sub LoadIntelligenceList()

        Try
            Try
                contatoreDecret = 0
                TextBox3.Text += vbCrLf & vbCrLf & "Email inviata a: "
                For z = 0 To CheckedListBox1.Items.Count - 1
                    CheckedListBox1.SelectedIndex = z
                    If CheckedListBox1.GetItemChecked(z) = True Then
                        If CheckedListBox1.Items(z).ToString.Contains("[") Then
                            Dim InstantSplitter() As String = Split(CheckedListBox1.Items(z).ToString, "  [")
                            TextBox3.Text += InstantSplitter(0).ToString & ", "
                            contatoreDecret += 1
                            If contatoreDecret = 4 Then
                                TextBox3.Text += vbCrLf
                                contatoreDecret = 0
                            End If
                        Else
                            TextBox3.Text += CheckedListBox1.SelectedItem.ToString & ", "
                            contatoreDecret += 1
                            If contatoreDecret = 4 Then
                                TextBox3.Text += vbCrLf
                                contatoreDecret = 0
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
            End Try
            TextBox3.Text += vbCrLf & "Nr. Progressivo Email: " & My.Settings.NumEmailIntelligence.ToString("D4") & "INTEL - " & My.Settings.JlDate.ToString("D3")
            TextBox3.Text += vbCrLf & "Il: " & Date.Now.Day.ToString & "." & Date.Now.Month.ToString & "." & Date.Now.Year.ToString & "  Alle: " & ora
            TextBox3.Text += vbCrLf & "CHIAVE: " & FileName
        Catch ex As Exception

        End Try

        Try
            StatusBar1.Text = "Stato Invio: Caricamento lista destinatari"
            ListBox2.Items.Clear()
            If CheckedListBox1.Items.Count > 0 Then
                CheckedListBox1.SelectedIndex = 0

                riga = Split(TextBox3.Text, vbCrLf)
                If System.IO.File.Exists(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt") Then
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                    scrivi.Write(My.Settings.NumEmailIntelligence.ToString("D4") & "INTEL - " & My.Settings.JlDate.ToString("D3") & "#")
                    scrivi.Close()
                Else
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                    scrivi.Write(My.Settings.NumEmailIntelligence.ToString("D4") & "INTEL - " & My.Settings.JlDate.ToString("D3") & "#")
                    scrivi.Close()
                End If

                If System.IO.File.Exists(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\emailinviate" & My.Settings.NumEmailIntelligence.ToString("D4") & "INTEL.txt") Then
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\emailinviate" & My.Settings.NumEmailIntelligence.ToString("D4") & "INTEL.txt")
                    scrivi.Write(TextBox3.Text)
                    scrivi.Close()
                Else
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\emailinviate" & My.Settings.NumEmailIntelligence.ToString("D4") & "INTEL.txt")
                    scrivi.Write(TextBox3.Text)
                    scrivi.Close()
                End If

                For r = 0 To CheckedListBox1.Items.Count - 1
                    CheckedListBox1.SelectedIndex = r
                    If CheckedListBox1.GetItemChecked(r) = True Then
                        If CheckedListBox1.Items(r).ToString.Contains("[") Then
                            Dim InstantSplitter() As String = Split(CheckedListBox1.Items(r).ToString, "  [")
                            scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                            scrivi.Write(InstantSplitter(0).ToString & "/")
                            scrivi.Close()
                        Else
                            scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                            scrivi.Write(CheckedListBox1.SelectedItem.ToString & "/")
                            scrivi.Close()
                        End If
                        
                        leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                        letto = leggi.ReadToEnd
                        leggi.Close()
                        infoufficio = Split(letto, "#")
                        For z = 0 To UBound(infoufficio) - 2
                            If infoufficio(z + 1) <> "" Then
                                ListBox2.Items.Add(infoufficio(z + 1))
                            End If
                        Next
                    End If
                Next

                scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                scrivi.Write("#" & MSGIDdef & "#" & datagiorno & " " & ora & "#" & chkAllegato & "#" & riga(0) & "#" & riga(1) & "#" & vbCrLf)
                scrivi.Close()

                If My.Settings.TipoDiDistribuzione = "Manuale" Then
                    For z = 0 To CheckedListBox1.Items.Count - 1
                        CheckedListBox1.SelectedIndex = z
                        If CheckedListBox1.GetItemChecked(z) = True Then
                            If CheckedListBox1.Items(z).ToString <> "" Then
                                If System.IO.Directory.Exists(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & My.Settings.JlDate.ToString("D3")) Then
                                    If System.IO.File.Exists(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmailIntelligence.ToString("D4") & "INTEL.txt") Then
                                    Else
                                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmailIntelligence.ToString("D4") & "INTEL.txt")
                                        scrivi.Write(TextBox3.Text)
                                        scrivi.Close()
                                    End If
                                Else
                                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & My.Settings.JlDate.ToString("D3"))
                                    If System.IO.File.Exists(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmailIntelligence.ToString("D4") & "INTEL.txt") Then
                                    Else
                                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\" & CheckedListBox1.Items(z).ToString & My.Settings.JlDate.ToString("D3") & "\emailinviate" & My.Settings.NumEmailIntelligence.ToString("D4") & "INTEL.txt")
                                        scrivi.Write(TextBox3.Text)
                                        scrivi.Close()
                                    End If
                                End If
                            End If
                        End If
                    Next

                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\TestoEmail.txt")
                    scrivi.Write(TextBox3.Text)
                    scrivi.Close()

                    StatusBar1.Text = "Invio Emails completato con successo."
                    MsgBox("Invio Completato, Email inviata a " & ListBox2.Items.Count & " indirizzi.", MsgBoxStyle.Information, "Asc Server - Invio Completato")
                    SaveProgressiveNumbers()

                    OpenFileDialog1.FileName = ""
                    chkAllegato = "NO"

                    Try
                        System.IO.File.Delete(Application.StartupPath & "\" & My.Settings.NumEmailIntelligence.ToString("D4") & "INTEL - " & My.Settings.JlDate.ToString("D3") & ".txt")
                    Catch ex As Exception

                    End Try
                    If Timer2.Enabled = False Then
                        Timer2.Start()
                    End If
                    CleanAll()
                    Exit Sub
                End If

                For r = 0 To ListBox2.Items.Count - 1
                    ListBox2.SetSelected(r, True)
                    If ListBox2.SelectedItem = "" Then
                        ListBox2.Items.RemoveAt(r)
                    End If
                Next
            End If

        Catch ex As Exception

        End Try

        Try
            StatusBar1.Text = "Stato Invio: Controllo allegati"
            scrivi = System.IO.File.CreateText(Application.StartupPath & "\" & My.Settings.NumEmailIntelligence.ToString("D4") & "INTEL - " & My.Settings.JlDate.ToString("D3") & ".txt")
            scrivi.Write(TextBox3.Text)
            scrivi.Close()

            scrivi = System.IO.File.CreateText(Application.StartupPath & "\TestoEmail.txt")
            scrivi.Write(TextBox3.Text)
            scrivi.Close()
        Catch ex As Exception

        End Try

        TrattaIntelligenceDistribution()

    End Sub

    Public Sub TrattaIntelligenceDistribution()

        If InvioAnnullato = 1 Then
            Exit Sub
        End If

        Try
            ListBox2.SetSelected(0, True)
            int2 = ListBox2.Items.Count
            StatusBar1.Text = "Stato Invio: Invio iniziato, smistamento"
            For r = 0 To int2 - 1
                ListBox2.SetSelected(r, True)
                If ListBox2.SelectedItem.ToString.Contains("@") Then
                    IntelligenceSendEmailToEmailAddress()
                Else
                    IntelligenceSendEmailToMachineAddress()
                End If

                If r = int2 - 1 Then
                    StatusBar1.Text = "Invio Emails completato con successo."
                    MsgBox("Invio Completato, Email inviata a " & ListBox2.Items.Count & " indirizzi.", MsgBoxStyle.Information, "Asc Server - Invio Completato")
                    SaveProgressiveNumbers()

                    OpenFileDialog1.FileName = ""
                    chkAllegato = "NO"

                    Try
                        System.IO.File.Delete(Application.StartupPath & "\" & My.Settings.NumEmailIntelligence.ToString("D4") & "INTEL - " & My.Settings.JlDate.ToString("D3") & ".txt")
                    Catch ex As Exception

                    End Try
                    If Timer2.Enabled = False Then
                        Timer2.Start()
                    End If
                    CleanAll()
                End If

            Next
        Catch ex As Exception
            MsgBox("Errore invio email intelligence, messaggio ex: " & ex.Message, MsgBoxStyle.Critical, "Errore Invio")
        End Try

    End Sub

    Public Sub IntelligenceSendEmailToEmailAddress()

        Try
            Dim objOutlook As Microsoft.Office.Interop.Outlook.Application
            Dim objEmail As Microsoft.Office.Interop.Outlook.MailItem
            objOutlook = CType(CreateObject("Outlook.Application"), Microsoft.Office.Interop.Outlook.Application)
            objEmail = objOutlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
            With objEmail
                .Subject = TextBox1.Text
                .To = ListBox2.SelectedItem.ToString
                .Body = TextBox3.Text
                .Attachments.Add(Application.StartupPath & "\" & My.Settings.NumEmailIntelligence.ToString("D4") & "INTEL - " & My.Settings.JlDate.ToString("D3") & ".txt")
                If OpenFileDialog1.FileName <> "" Then
                    .Attachments.Add(OpenFileDialog1.FileName)
                End If
                .Send()
            End With
        Catch ex As Exception
            MsgBox("Errore invio email intelligence, messaggio ex: " & ex.Message, MsgBoxStyle.Critical, "Asc Server - Errore Invio")
        End Try

    End Sub

    Public Sub IntelligenceSendEmailToMachineAddress()

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\temppartenze") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\temppartenze")
            End If
            Dim scrittore As System.IO.StreamWriter
            If System.IO.File.Exists(Application.StartupPath & "\temppartenze\temppartenza.txt") Then
                For r = 1 To 9999
                    If System.IO.File.Exists(Application.StartupPath & "\temppartenze\temppartenza(" & r.ToString & ").txt") Then
                    Else
                        scrittore = System.IO.File.CreateText(Application.StartupPath & "\temppartenze\temppartenza(" & r.ToString & ").txt")
                        scrittore.Write("MessaggioDalServer" & vbCrLf)
                        scrittore.Write(ListBox2.SelectedItem.ToString & vbCrLf)
                        scrittore.Write(FileName & vbCrLf)
                        scrittore.Write(TextBox3.Text)
                        scrittore.Close()
                        Exit For
                    End If
                Next
            Else
                scrittore = System.IO.File.CreateText(Application.StartupPath & "\temppartenze\temppartenza.txt")
                scrittore.Write("MessaggioDalServer" & vbCrLf)
                scrittore.Write(ListBox2.SelectedItem.ToString & vbCrLf)
                scrittore.Write(FileName & vbCrLf)
                scrittore.Write(TextBox3.Text)
                scrittore.Close()
            End If
        Catch ex As Exception
            MsgBox("Errore invio email intelligence, messaggio ex: " & ex.Message, MsgBoxStyle.Critical, "Asc Server - Errore Invio")
        End Try

    End Sub
    '*** ENDS - INTELLIGENCE - CODING ***

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click

        INTEL22TerSelectionDate.Show()

    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

        If TextBox3.Text = "" Then
            OpenedFile = 0
            FinalSectionPresent = 0
        End If

    End Sub

    Private Sub MenuItem23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem23.Click

        TextBox3.SelectAll()

    End Sub

    Private Sub MenuItem26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem26.Click

        TextBox3.Cut()

    End Sub

    Private Sub MenuItem24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem24.Click

        TextBox3.Copy()

    End Sub

    Private Sub MenuItem25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem25.Click

        TextBox3.Paste()

    End Sub

    Private Sub MenuItem29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem29.Click

        If MenuItem42.Text = "Disconnetti" Then
            Try
                client.Send(Requests.StringMessage, "ServerRequestsCode")
            Catch ex As Exception

            End Try
            CodeUffici.Show()
        Else
            MsgBox("Asc Server non risulta connesso ad Asc Echange, impossibile continuare. Connettersi e riprovare.", MsgBoxStyle.Information, "Asc Server - Visualizza Code")
        End If

    End Sub

    Public Sub SaveProgressiveNumbers()

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\Memorized") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Memorized")
            End If
            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoArriviChiusura.txt")
            scrivi.Write(My.Settings.NumEmail)
            scrivi.Close()
            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoPartenzeChiusura.txt")
            scrivi.Write(My.Settings.NumEmail2)
            scrivi.Close()
            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoIntellChiusura.txt")
            scrivi.Write(My.Settings.NumEmailIntelligence)
            scrivi.Close()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub SaveOthersProgressive()

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\Memorized") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Memorized")
            End If
            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\DataChiusura.txt")
            scrivi.Write(My.Settings.DataOggi)
            scrivi.Close()
            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\JDChiusura.txt")
            scrivi.Write(My.Settings.JlDate)
            scrivi.Close()
            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Memorized\NrProgressivoUfficiChiusura.txt")
            scrivi.Write(My.Settings.Ufficio)
            scrivi.Close()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MenuItem30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem30.Click

        Try
            My.Settings.x = 8
            Accesso.Show()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MenuItem32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem32.Click

        If MenuItem32.Checked = True Then
            MenuItem32.Checked = False
            My.Settings.MenuItem32 = False
        Else
            MenuItem32.Checked = True
            My.Settings.MenuItem32 = True
        End If

    End Sub

    Private Sub MenuItem33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem33.Click

        If MenuItem33.Checked = True Then
            MenuItem33.Checked = False
            My.Settings.MenuItem33 = False
        Else
            MenuItem33.Checked = True
            My.Settings.MenuItem33 = True
        End If

    End Sub

    Private Sub Main_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        If Me.WindowState = FormWindowState.Maximized Then
            SaveNewSize()
        End If

    End Sub

    Private Sub Main_ResizeBegin(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeBegin

        Resizing = True

    End Sub

    Private Sub Main_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd

        Resizing = False
        ResizeForm()
        SaveNewSize()

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

    Private Sub Main_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged

        If Resizing = False Then
            ResizeForm()
        End If

    End Sub

    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem9.Click

        DATABASE_ARRIVI.Show()

    End Sub

    Private Sub MenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem10.Click

        DATABASE_PARTENZE.Show()

    End Sub

    Private Sub MenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem11.Click

        DATABASE_INTELLIGENCE.Show()

    End Sub

    Public Sub LoadExchangeSettingsForm()

        Try
            ExchangeSettings.Label5.Visible = False
            ExchangeSettings.Button1.Visible = True
            ExchangeSettings.GroupBox1.Visible = True
            ExchangeSettings.GroupBox2.Visible = True
            ExchangeSettings.NumericUpDown1.Value = (My.Settings.SettingsFrequenzaControllo / 1000)
            ExchangeSettings.NumericUpDown2.Value = (My.Settings.SettingsFrequenzaInvio / 1000)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MenuItem21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem21.Click

        Data.Show()

    End Sub

    Private Sub MenuItem36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem36.Click

        Accesso.Show()
        My.Settings.x = 5

    End Sub

    Public Sub RequestSettings()

        Try
            If MenuItem42.Text = "Disconnetti" Then
                client.Send(Requests.StringMessage, "ServerRequestsSettings")
                ExchangeSettings.Show()
            Else
                MsgBox("Impossibile ricevere le impostazioni da Asc Exchange, Asc Server non risulta connesso ad Asc Exchange. Connettersi e riprova.", MsgBoxStyle.Critical, "Asc Server - Errore Connessione")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MenuItem37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem37.Click

        IPAddress.Show()

    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick

        Try
            If MessaggioRicevuto.Contains("ConnectedClientListAnswearing") Then
                Try
                    Dim MySplitter() As String = Split(MessaggioRicevuto, "|")
                    MessaggioRicevuto = ""
                    For r = 1 To UBound(MySplitter)
                        ConnectedClientsList.ListBox1.Items.Add(MySplitter(r).ToString)
                    Next
                Catch ex As Exception

                End Try
            ElseIf MessaggioRicevuto.Contains("InvioSettingsDallExchange") Then
                Try
                    Dim MySplitter() As String = Split(MessaggioRicevuto, "|")
                    MessaggioRicevuto = ""
                    My.Settings.SettingsFrequenzaControllo = MySplitter(1).ToString
                    My.Settings.SettingsFrequenzaInvio = MySplitter(2).ToString
                    LoadExchangeSettingsForm()
                Catch ex As Exception

                End Try
            ElseIf MessaggioRicevuto.Contains("AscExchangeCode") Then
                Try
                    Dim MySplitter() As String = Split(MessaggioRicevuto, "|")
                    MessaggioRicevuto = ""
                    If MySplitter(1).ToString.Contains("Nulla") Then
                        MsgBox("Non sono presenti code da visualizzare.", MsgBoxStyle.Information, "Asc Server - Code Uffici")
                        CodeUffici.Close()
                    Else
                        CodeUffici.ListView1.Items.Clear()
                        CodeUffici.ListView1.Enabled = True
                        CodeUffici.MenuItem1.Enabled = True
                        CodeUffici.MenuItem3.Enabled = True
                        CodeUffici.Label1.Visible = False
                        For r = 1 To UBound(MySplitter)
                            If MySplitter(r).ToString <> "" Then
                                CodeUffici.ListView1.Items.Add(MySplitter(r).ToString)
                            End If
                        Next
                    End If
                Catch ex As Exception

                End Try
            ElseIf MessaggioRicevuto.Contains("AscExchangeMessaggiCartella") Then
                Try
                    Dim MySplitter() As String = Split(MessaggioRicevuto, "|")
                    MessaggioRicevuto = ""
                    CodeUffici.Label1.Visible = False
                    If MySplitter(1).ToString.Contains("Nulla") Then
                        MsgBox("Non sono presenti messaggi da visualizzare.", MsgBoxStyle.Information, "Asc Server - Code Uffici")
                        CodeUffici.ListView1.Items.Add("Nulla")
                    Else
                        Dim MySecondSplitter() As String
                        For r = 1 To UBound(MySplitter)
                            MySecondSplitter = Split(MySplitter(r).ToString, ".txt")
                            If MySplitter(r).ToString <> "" Then
                                CodeUffici.ListView1.Items.Add(MySecondSplitter(0).ToString)
                            End If
                        Next
                    End If
                Catch ex As Exception

                End Try
            ElseIf MessaggioRicevuto = "ChkServerIsConnected" Then
                MessaggioRicevuto = ""
                client.Send(Requests.DataFile, "ServerIsConnected=True")
            ElseIf MessaggioRicevuto.Contains("L'UfficioPerCuiSiAvevaFattoRichiestaEConnesso") Then
                MessaggioRicevuto = ""
                Try
                    CollezioneElementiCheckati = ""
                    For l = 0 To CheckedListBox1.Items.Count - 1
                        If CheckedListBox1.Items(l).ToString <> "" Then
                            If CheckedListBox1.GetItemChecked(l) = True Then
                                CollezioneElementiCheckati += CheckedListBox1.Items(l).ToString & vbCrLf
                            End If
                        End If
                    Next
                Catch ex As Exception

                End Try
                Dim ListaElementi As String = ""
                For r = 0 To CheckedListBox1.Items.Count - 1
                    If CheckedListBox1.Items(r).ToString <> "" Then
                        ListaElementi += CheckedListBox1.Items(r).ToString & "|"
                    End If
                Next
                CheckedListBox1.Items.Clear()
                Dim InstantSplitter() As String = Split(ListaElementi, "|")
                For x = 0 To UBound(InstantSplitter)
                    If x <> NumeroUltimoUfficio - 1 Then
                        If InstantSplitter(x).ToString <> "" Then
                            CheckedListBox1.Items.Add(InstantSplitter(x).ToString)
                        End If
                    Else
                        If InstantSplitter(x).ToString.Contains("[ON-LINE]") Then
                            Dim SecondInstantSplitter() As String = Split(InstantSplitter(x).ToString, "  [ON-LINE]")
                            CheckedListBox1.Items.Add(SecondInstantSplitter(0).ToString & "  [ON-LINE]")
                        ElseIf InstantSplitter(x).ToString.Contains("[OFF-LINE]") Then
                            Dim SecondInstantSplitter() As String = Split(InstantSplitter(x).ToString, "  [OFF-LINE]")
                            CheckedListBox1.Items.Add(SecondInstantSplitter(0).ToString & "  [ON-LINE]")
                        Else
                            CheckedListBox1.Items.Add(InstantSplitter(x).ToString & "  [ON-LINE]")
                        End If
                    End If
                Next
                RistabilisciIndiciSelezionati()
                AspettaAncora = 0
            ElseIf MessaggioRicevuto.Contains("L'UfficioPerCuiSiAvevaFattoRichiestaNonEConnesso") Then
                MessaggioRicevuto = ""
                AspettaAncora = 11
            ElseIf MessaggioRicevuto.Contains("RispostaDalProcessoDiOttenimentoCode") Then
                Dim InstantSplitter() As String = Split(MessaggioRicevuto, "§")
                MessaggioRicevuto = ""
                MessaggioFinaleRicevuto = ""
                For r = 1 To UBound(InstantSplitter)
                    If InstantSplitter(r).ToString <> "" Then
                        MessaggioFinaleRicevuto += InstantSplitter(r).ToString & "§"
                    End If
                Next
                CalcolaCode()
            ElseIf MessaggioRicevuto = "ChiusuraAscExchangeInCorso" Then
                MessaggioRicevuto = ""
                Try
                    MenuItem42.Text = "Connetti"
                Catch ex As Exception

                End Try
                Try
                    client.Send(RequestTags.Disconnect)
                    client.client.GetStream.Close()
                    client.client.Close()
                    client.client = Nothing
                Catch ex As Exception

                End Try
                If Timer3.Enabled = True Then
                    Timer3.Stop()
                End If
                Try
                    If client.client.Connected = False Then
                        If RectangleShape1.BackColor = Color.Green Then
                            RectangleShape1.BackColor = Color.Red
                            RectangleShape1.BorderColor = Color.Red
                            MenuItem42.Text = "Connetti"
                            Timering = 0
                            Timer5.Start()
                        End If
                    End If
                Catch ex As Exception

                End Try
                Label8.Text = "LISTA UFFICI  [Connettersi ad Exchange]"
                If Timer6.Enabled = True Then
                    Timer6.Stop()
                End If
            Else
                MessaggioRicevuto = ""
            End If
        Catch ex As Exception
            If Timer3.Enabled = False Then
                Timer3.Start()
            End If
            Exit Sub
        End Try

    End Sub

    Dim MessaggioFinaleRicevuto As String = ""
    Dim ClientDaCercare As String = ""
    Dim UfficioDaIncrementare As String = ""
    Dim NumeroDaIncrementare As Integer = 0
    Public Sub CalcolaCode()

        Try
            Timer7.Stop()
            Try
                CollezioneElementiCheckati = ""
                For l = 0 To CheckedListBox1.Items.Count - 1
                    If CheckedListBox1.Items(l).ToString <> "" Then
                        If CheckedListBox1.GetItemChecked(l) = True Then
                            CollezioneElementiCheckati += CheckedListBox1.Items(l).ToString & vbCrLf
                        End If
                    End If
                Next
            Catch ex As Exception

            End Try
            Dim MySplitter() As String = Split(MessaggioFinaleRicevuto, "§")
            For r = 0 To UBound(MySplitter)
                Dim MySplitter2() As String = Split(MySplitter(r).ToString, "#")
                If MySplitter2(0).ToString <> "" Then
                    ClientDaCercare = MySplitter2(0).ToString
                    For z = 0 To My.Settings.Ufficio - 1
                        leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & z.ToString & ".txt")
                        letto = leggi.ReadToEnd
                        leggi.Close()
                        If letto.Contains(ClientDaCercare) Then
                            'trovato l'ufficio
                            Dim MySplitter3() As String = Split(letto, "#")
                            UfficioDaIncrementare = MySplitter3(0).ToString
                            'contiamo il numero di incremento
                            If MySplitter2(1).ToString <> "" Then
                                NumeroDaIncrementare = 0
                                Dim MySplitter4() As String = Split(MySplitter2(1).ToString, "|")
                                For t = 0 To UBound(MySplitter4)
                                    If MySplitter4(t).ToString <> "" Then
                                        NumeroDaIncrementare += 1
                                    End If
                                Next
                            End If
                            'fine conteggio
                            Dim ListaElementi As String = ""
                            For f = 0 To CheckedListBox1.Items.Count - 1
                                If CheckedListBox1.Items(f).ToString <> "" Then
                                    ListaElementi += CheckedListBox1.Items(f).ToString & "|"
                                End If
                            Next
                            CheckedListBox1.Items.Clear()
                            Dim InstantSplitter() As String = Split(ListaElementi, "|")
                            For x = 0 To UBound(InstantSplitter)
                                If InstantSplitter(x).ToString <> "" Then
                                    If InstantSplitter(x).ToString.ToLower.Contains(UfficioDaIncrementare.ToLower) Then
                                        CheckedListBox1.Items.Add(InstantSplitter(x).ToString & "  [" & NumeroDaIncrementare.ToString("D4") & "]")
                                    Else
                                        CheckedListBox1.Items.Add(InstantSplitter(x).ToString)
                                    End If
                                End If
                            Next
                            Exit For
                        End If
                    Next
                End If
            Next
            RistabilisciIndiciSelezionati()
            Label8.Text = "LISTA UFFICI"
            Timer7.Start()
            AspettaIlProcessoDiInvioCode = -10
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick

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

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\temppartenze") Then
                Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\temppartenze")
                If FilesCollection.Count > 0 Then
                    For r = 0 To UBound(FilesCollection)
                        client.SendFile(Requests.DataFile, FilesCollection(r).ToString)
                        System.IO.File.Delete(FilesCollection(r).ToString)
                        Exit For
                    Next
                Else
                    Timer2.Stop()
                End If
            Else
                Timer2.Stop()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MenuItem38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem38.Click

        Timer1.Start()

        IsPCNAvailable = "NO"
        chkAllegato = "NO"
        My.Settings.CloseDTBNow = 0
        My.Settings.SelectedDate = ""
        My.Settings.SelectedDate2 = ""
        My.Settings.Intelligence22Ter = ""
        My.Settings.DataOggi = Date.Now.Day.ToString & "." & Date.Now.Month.ToString & "." & Date.Now.Year.ToString

        If Date.Now.DayOfYear - My.Settings.JlDate > 1 Then
            My.Settings.NumEmail = 0
            My.Settings.NumEmail2 = 0
            My.Settings.NumEmailIntelligence = 0
            Dim Splitter() As String = Split(My.Settings.DataIeri, ".")
            Dim GiorniMese As Integer = Date.DaysInMonth(Splitter(2).ToString, Splitter(1).ToString)

            If GiorniMese = 31 Then
                If Splitter(0).ToString <> "31" Then
                    My.Settings.DataOggi = (Val(Splitter(0).ToString) + 1).ToString & "." & Splitter(1).ToString & "." & Splitter(2).ToString
                    My.Settings.DataIeri = (Val(Splitter(0).ToString) + 1).ToString & "." & Splitter(1).ToString & "." & Splitter(2).ToString
                    datagiorno = My.Settings.DataOggi
                    My.Settings.JlDate = My.Settings.JlDate + 1
                    Label4.ForeColor = Color.Red
                    Label6.ForeColor = Color.Red
                    Label6.Text = datagiorno
                    Label4.Text = My.Settings.JlDate.ToString("D3")
                Else
                    If Splitter(1).ToString <> "12" Then
                        My.Settings.DataOggi = "1." & (Val(Splitter(1).ToString) + 1).ToString & "." & Splitter(2).ToString
                        My.Settings.DataIeri = "1." & (Val(Splitter(1).ToString) + 1).ToString & "." & Splitter(2).ToString
                        datagiorno = My.Settings.DataOggi
                        My.Settings.JlDate = My.Settings.JlDate + 1
                        Label4.ForeColor = Color.Red
                        Label6.ForeColor = Color.Red
                        Label6.Text = datagiorno
                        Label4.Text = My.Settings.JlDate.ToString("D3")
                    Else
                        My.Settings.DataOggi = "1." & "1." & (Val(Splitter(2).ToString) + 1).ToString
                        My.Settings.DataIeri = "1." & "1." & (Val(Splitter(2).ToString) + 1).ToString
                        datagiorno = My.Settings.DataOggi
                        My.Settings.JlDate = My.Settings.JlDate + 1
                        Label4.ForeColor = Color.Red
                        Label6.ForeColor = Color.Red
                        Label6.Text = datagiorno
                        Label4.Text = My.Settings.JlDate.ToString("D3")
                    End If
                End If
            ElseIf GiorniMese = 30 Then
                If Splitter(0).ToString <> "30" Then
                    My.Settings.DataOggi = (Val(Splitter(0).ToString) + 1).ToString & "." & Splitter(1).ToString & "." & Splitter(2).ToString
                    My.Settings.DataIeri = (Val(Splitter(0).ToString) + 1).ToString & "." & Splitter(1).ToString & "." & Splitter(2).ToString
                    datagiorno = My.Settings.DataOggi
                    My.Settings.JlDate = My.Settings.JlDate + 1
                    Label4.ForeColor = Color.Red
                    Label6.ForeColor = Color.Red
                    Label6.Text = datagiorno
                    Label4.Text = My.Settings.JlDate.ToString("D3")
                Else
                    If Splitter(1).ToString <> "12" Then
                        My.Settings.DataOggi = "1." & (Val(Splitter(1).ToString) + 1).ToString & "." & Splitter(2).ToString
                        My.Settings.DataIeri = "1." & (Val(Splitter(1).ToString) + 1).ToString & "." & Splitter(2).ToString
                        datagiorno = My.Settings.DataOggi
                        My.Settings.JlDate = My.Settings.JlDate + 1
                        Label4.ForeColor = Color.Red
                        Label6.ForeColor = Color.Red
                        Label6.Text = datagiorno
                        Label4.Text = My.Settings.JlDate.ToString("D3")
                    Else
                        My.Settings.DataOggi = "1." & "1." & (Val(Splitter(2).ToString) + 1).ToString
                        My.Settings.DataIeri = "1." & "1." & (Val(Splitter(2).ToString) + 1).ToString
                        datagiorno = My.Settings.DataOggi
                        My.Settings.JlDate = My.Settings.JlDate + 1
                        Label4.ForeColor = Color.Red
                        Label6.ForeColor = Color.Red
                        Label6.Text = datagiorno
                        Label4.Text = My.Settings.JlDate.ToString("D3")
                    End If
                End If
            ElseIf GiorniMese = 29 Then
                If Splitter(0).ToString <> "29" Then
                    My.Settings.DataOggi = (Val(Splitter(0).ToString) + 1).ToString & "." & Splitter(1).ToString & "." & Splitter(2).ToString
                    My.Settings.DataIeri = (Val(Splitter(0).ToString) + 1).ToString & "." & Splitter(1).ToString & "." & Splitter(2).ToString
                    datagiorno = My.Settings.DataOggi
                    My.Settings.JlDate = My.Settings.JlDate + 1
                    Label4.ForeColor = Color.Red
                    Label6.ForeColor = Color.Red
                    Label6.Text = datagiorno
                    Label4.Text = My.Settings.JlDate.ToString("D3")
                Else
                    If Splitter(1).ToString <> "12" Then
                        My.Settings.DataOggi = "1." & (Val(Splitter(1).ToString) + 1).ToString & "." & Splitter(2).ToString
                        My.Settings.DataIeri = "1." & (Val(Splitter(1).ToString) + 1).ToString & "." & Splitter(2).ToString
                        datagiorno = My.Settings.DataOggi
                        My.Settings.JlDate = My.Settings.JlDate + 1
                        Label4.ForeColor = Color.Red
                        Label6.ForeColor = Color.Red
                        Label6.Text = datagiorno
                        Label4.Text = My.Settings.JlDate.ToString("D3")
                    Else
                        My.Settings.DataOggi = "1." & "1." & (Val(Splitter(2).ToString) + 1).ToString
                        My.Settings.DataIeri = "1." & "1." & (Val(Splitter(2).ToString) + 1).ToString
                        datagiorno = My.Settings.DataOggi
                        My.Settings.JlDate = My.Settings.JlDate + 1
                        Label4.ForeColor = Color.Red
                        Label6.ForeColor = Color.Red
                        Label6.Text = datagiorno
                        Label4.Text = My.Settings.JlDate.ToString("D3")
                    End If
                End If
            End If

        Else
            Label4.ForeColor = Color.Black
            Label6.ForeColor = Color.Black
            My.Settings.NumEmail = 0
            My.Settings.NumEmail2 = 0
            My.Settings.NumEmailIntelligence = 0
            My.Settings.DataIeri = My.Settings.DataOggi
            datagiorno = My.Settings.DataOggi
            My.Settings.JlDate = Date.Now.DayOfYear
        End If

        datagiorno = My.Settings.DataOggi

        Label4.Text = My.Settings.JlDate.ToString("D3")

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\Email") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email")
            End If
            'DATABASE ARRIVI
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEARRIVI") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\DATABASEARRIVI")
            End If
            'DATABASE PARTENZE
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\DATABASEPARTENZE")
            End If
            'DATABASE INTELLIGENCE
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\INTELLIGENCE") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\INTELLIGENCE")
            End If
        Catch ex As Exception

        End Try

        Try
            'DATABASE ARRIVI
            If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\dateusate.txt") Then
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Email\DATABASEARRIVI\dateusate.txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                If letto.Contains(datagiorno) Then
                Else
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEARRIVI\dateusate.txt")
                    scrivi.WriteLine(datagiorno)
                    scrivi.Close()
                End If
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEARRIVI\dateusate.txt")
                scrivi.WriteLine(datagiorno)
                scrivi.Close()
            End If
            'DATABASE PARTENZE
            If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\dateusate.txt") Then
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Email\DATABASEPARTENZE\dateusate.txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                If letto.Contains(datagiorno) Then
                Else
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\DATABASEPARTENZE\dateusate.txt")
                    scrivi.WriteLine(datagiorno)
                    scrivi.Close()
                End If
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEPARTENZE\dateusate.txt")
                scrivi.WriteLine(datagiorno)
                scrivi.Close()
            End If
            'DATABASE INTELLIGENCE
            If System.IO.File.Exists(Application.StartupPath & "\Email\INTELLIGENCE\dateusate.txt") Then
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Email\INTELLIGENCE\dateusate.txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                If letto.Contains(datagiorno) Then
                Else
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Email\INTELLIGENCE\dateusate.txt")
                    scrivi.WriteLine(datagiorno)
                    scrivi.Close()
                End If
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\INTELLIGENCE\dateusate.txt")
                scrivi.WriteLine(datagiorno)
                scrivi.Close()
            End If
        Catch ex As Exception

        End Try

        Try
            'DATABASE ARRIVI
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno) Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno)
                My.Settings.NumEmail = 0
            End If
            'DATABASE PARTENZE
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno) Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno)
                My.Settings.NumEmail2 = 0
            End If
            'DATABASE INTELLIGENCE
            If System.IO.Directory.Exists(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno) Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno)
                My.Settings.NumEmailIntelligence = 0
            End If
        Catch ex As Exception

        End Try

        Try
            'DATABASE ARRIVI
            If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt") Then
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEARRIVI\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                scrivi.Close()
            End If
            'DATABASE PARTENZE
            If System.IO.File.Exists(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt") Then
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\DATABASEPARTENZE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                scrivi.Close()
            End If
            'DATABASE INTELLIGENCE
            If System.IO.File.Exists(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt") Then
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Email\INTELLIGENCE\" & datagiorno & "\datiemailinviate" & datagiorno & ".txt")
                scrivi.Close()
            End If
        Catch ex As Exception

        End Try

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\Uffici") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Uffici")
                My.Settings.Ufficio = 0
            End If

            If System.IO.File.Exists(Application.StartupPath & "\Uffici\listauffici.txt") Then
            Else
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Uffici\listauffici.txt")
                scrivi.Close()
            End If
        Catch ex As Exception

        End Try

        Try
            MsgBox("Data aggiornata al: " & My.Settings.DataOggi, MsgBoxStyle.Information, "Asc Server - Aggiornamento data")

            SaveOthersProgressive()

            Dim AlertMessage = MsgBox("Per poter continuare ad utilizzare il software bisogna prima effettuare la chiusura dei messaggi. Si desidera effettuare la chiusura adesso?", MsgBoxStyle.YesNo, "Asc Server - Chiusure")
            If AlertMessage = 6 Then
                If My.Settings.DataUltimaChiusura <> "" Then
                    If My.Settings.DataUltimaChiusura <> My.Settings.DataOggi Then
                        Chiusure.Show()
                    Else
                        MsgBox("La chiusura dei messaggi risulta già essere stata effettuata.", MsgBoxStyle.Information, "Asc Server - Chiusure")
                    End If
                Else
                    Chiusure.Show()
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MenuItem41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem41.Click

        AggiornaOra.Show()

    End Sub

    Private Sub MenuItem40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem40.Click

        Try
            If MenuItem40.Text <> "[h:m:s] [Impostare Ora]" Then
                If MenuItem40.Checked = False Then
                    MenuItem40.Checked = True
                    MsgBox("E' stato selezionato l'aggiornamento automatico della data per le ore " & My.Settings.OraAggiornamento & ".", MsgBoxStyle.Information, "Asc Server - Ora Aggiornamento Automatico")
                    My.Settings.AggiornamentoAutoSelect = True
                Else
                    MsgBox("E' stato deselezionato l'aggiornamento automatico della data.", MsgBoxStyle.Information, "Asc Server - Ora Aggiornamento")
                    MenuItem40.Checked = False
                    My.Settings.AggiornamentoAutoSelect = False
                End If
            Else
                MsgBox("Prima di poter utilizzare l'aggiornamento automatico della data bisogna impostare l'ora.", MsgBoxStyle.Information, "Asc Server - Scegliere Ora")
                AggiornaOra.Show()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MenuItem42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem42.Click

        ConnectClient()

    End Sub

    Private Sub Timer5_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer5.Tick

        Try
            Timering += 1
            If Timering < 79 Then
                If MenuItem42.Text = "Disconnetti" Then
                    If RectangleShape1.Width < 93 Then
                        RectangleShape1.Width += 1
                    Else
                        RectangleShape1.Width = 0
                    End If
                ElseIf MenuItem42.Text = "Connetti" Then
                    If RectangleShape1.Width < 79 Then
                        RectangleShape1.Width += 1
                    Else
                        RectangleShape1.Width = 0
                    End If
                End If
            Else
                If MenuItem42.Text = "Disconnetti" Then
                    RectangleShape1.Width = 93
                ElseIf MenuItem42.Text = "Connetti" Then
                    RectangleShape1.Width = 79
                End If
                Timering = 0
                Timer5.Stop()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Timer6_Tick(sender As System.Object, e As System.EventArgs) Handles Timer6.Tick

        Try
            If AspettaAncora = 0 Then
                AspettaAncora += 1
                InviaRichiesta()
            Else
                AspettaAncora += 1
                If AspettaAncora > 10 Then
                    Try
                        CollezioneElementiCheckati = ""
                        For l = 0 To CheckedListBox1.Items.Count - 1
                            If CheckedListBox1.Items(l).ToString <> "" Then
                                If CheckedListBox1.GetItemChecked(l) = True Then
                                    CollezioneElementiCheckati += CheckedListBox1.Items(l).ToString & vbCrLf
                                End If
                            End If
                        Next
                    Catch ex As Exception

                    End Try
                    'Nessuna risposta ricevuta, segna come offline
                    Dim ListaElementi As String = ""
                    For r = 0 To CheckedListBox1.Items.Count - 1
                        If CheckedListBox1.Items(r).ToString <> "" Then
                            ListaElementi += CheckedListBox1.Items(r).ToString & "|"
                        End If
                    Next
                    CheckedListBox1.Items.Clear()
                    Dim InstantSplitter() As String = Split(ListaElementi, "|")
                    For x = 0 To UBound(InstantSplitter)
                        If x <> NumeroUltimoUfficio - 1 Then
                            If InstantSplitter(x).ToString <> "" Then
                                CheckedListBox1.Items.Add(InstantSplitter(x).ToString)
                            End If
                        Else
                            If InstantSplitter(x).ToString <> "" Then
                                If InstantSplitter(x).ToString.Contains("[ON-LINE]") Then
                                    Dim SecondInstantSplitter() As String = Split(InstantSplitter(x).ToString, "  [ON-LINE]")
                                    CheckedListBox1.Items.Add(SecondInstantSplitter(0).ToString & "  [OFF-LINE]")
                                ElseIf InstantSplitter(x).ToString.Contains("[OFF-LINE]") Then
                                    Dim SecondInstantSplitter() As String = Split(InstantSplitter(x).ToString, "  [OFF-LINE]")
                                    CheckedListBox1.Items.Add(SecondInstantSplitter(0).ToString & "  [OFF-LINE]")
                                Else
                                    CheckedListBox1.Items.Add(InstantSplitter(x).ToString & "  [OFF-LINE]")
                                End If
                            End If
                        End If
                    Next
                    AspettaAncora = 0
                End If
            End If
            RistabilisciIndiciSelezionati()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub InviaRichiesta()

        Try
            NumeroUltimoUfficio = My.Settings.NumeroUfficioUltimo
            If NumeroUltimoUfficio = 0 Then
                My.Settings.NumeroUfficioUltimo = NumeroUltimoUfficio
                For r = 0 To CheckedListBox1.Items.Count - 1
                    leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                    letto = leggi.ReadToEnd
                    leggi.Close()
                    Dim InstantSplitter() As String = Split(letto, "#")
                    ListaClients = ""
                    For x = 1 To UBound(InstantSplitter)
                        If InstantSplitter(x).ToString <> "" Then
                            If InstantSplitter(x).ToString.Contains("@") Then
                            Else
                                If InstantSplitter(x).ToString <> "" Then
                                    ListaClients += InstantSplitter(x).ToString & "|"
                                End If
                            End If
                        End If
                    Next
                    If ListaClients = "" Then
                        NumeroUltimoUfficio += 1
                        My.Settings.NumeroUfficioUltimo = NumeroUltimoUfficio
                        AspettaAncora = 0
                        Exit Sub
                    Else
                        Dim SecondInstantSplitter() As String = Split(ListaClients, "|")
                        If SecondInstantSplitter.Count > 0 Then
                            For y = 0 To UBound(SecondInstantSplitter)
                                If SecondInstantSplitter(y).ToString.Length > 2 Then
                                    client.Send(2, "ControlloConnettivitàPerL'ufficio|" & ListaClients)
                                    NumeroUltimoUfficio += 1
                                    My.Settings.NumeroUfficioUltimo = NumeroUltimoUfficio
                                    Exit Sub
                                End If
                            Next
                        Else
                            NumeroUltimoUfficio += 1
                            My.Settings.NumeroUfficioUltimo = NumeroUltimoUfficio
                            AspettaAncora = 0
                            Exit Sub
                        End If
                    End If
                Next
            Else
                If NumeroUltimoUfficio < CheckedListBox1.Items.Count Then
                    My.Settings.NumeroUfficioUltimo = NumeroUltimoUfficio
                    For r = NumeroUltimoUfficio To CheckedListBox1.Items.Count - 1
                        leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                        letto = leggi.ReadToEnd
                        leggi.Close()
                        Dim InstantSplitter() As String = Split(letto, "#")
                        ListaClients = ""
                        For x = 1 To UBound(InstantSplitter)
                            If InstantSplitter(x).ToString.Contains("@") Then
                            Else
                                If InstantSplitter(x).ToString <> "" Then
                                    If InstantSplitter(x).ToString <> "" Then
                                        ListaClients += InstantSplitter(x).ToString & "|"
                                    End If
                                End If
                            End If
                        Next
                        If ListaClients = "" Then
                            NumeroUltimoUfficio += 1
                            My.Settings.NumeroUfficioUltimo = NumeroUltimoUfficio
                            AspettaAncora = 0
                            Exit Sub
                        Else
                            Dim SecondInstantSplitter() As String = Split(ListaClients, "|")
                            If SecondInstantSplitter.Count > 0 Then
                                For y = 0 To UBound(SecondInstantSplitter)
                                    If SecondInstantSplitter(y).ToString.Length > 2 Then
                                        client.Send(2, "ControlloConnettivitàPerL'ufficio|" & ListaClients)
                                        NumeroUltimoUfficio += 1
                                        My.Settings.NumeroUfficioUltimo = NumeroUltimoUfficio
                                        Exit Sub
                                    End If
                                Next
                            Else
                                NumeroUltimoUfficio += 1
                                My.Settings.NumeroUfficioUltimo = NumeroUltimoUfficio
                                AspettaAncora = 0
                                Exit Sub
                            End If
                        End If
                    Next
                Else
                    NumeroUltimoUfficio = 0
                    My.Settings.NumeroUfficioUltimo = NumeroUltimoUfficio
                    AspettaAncora = 0
                    Timer6.Stop()
                    OttieniMessaggiInCoda()
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub OttieniMessaggiInCoda()

        Try
            client.Send(2, "IlServerRichiedeL'inizioDelProcessoDiInvioCode")
            Timer7.Start()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Timer7_Tick(sender As System.Object, e As System.EventArgs) Handles Timer7.Tick

        Try
            AspettaIlProcessoDiInvioCode += 1
            If AspettaIlProcessoDiInvioCode = 20 Then
                AspettaAncora = 0
                Label8.Text = "LISTA UFFICI"
                Timer6.Start()
                AspettaIlProcessoDiInvioCode = 0
                Timer7.Stop()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub RistabilisciIndiciSelezionati()

        Try
            Dim MySplitter() As String = Split(CollezioneElementiCheckati, vbCrLf)
            For r = 0 To UBound(MySplitter)
                If MySplitter(r).ToString <> "" Then
                    If MySplitter(r).ToString.Contains("[") Then
                        Dim MySecondSplitter() As String = Split(MySplitter(r).ToString, "[")
                        For x = 0 To CheckedListBox1.Items.Count - 1
                            If CheckedListBox1.Items(x).ToString.Contains(MySecondSplitter(0).ToString) Then
                                CheckedListBox1.SetItemChecked(x, True)
                            End If
                        Next
                    Else
                        For x = 0 To CheckedListBox1.Items.Count - 1
                            If CheckedListBox1.Items(x).ToString.Contains(MySplitter(r).ToString) Then
                                CheckedListBox1.SetItemChecked(x, True)
                            End If
                        Next
                    End If
                End If
            Next
            CollezioneElementiCheckati = ""
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MenuItem45_Click(sender As System.Object, e As System.EventArgs) Handles MenuItem45.Click

        Try
            My.Settings.TipoDiDistribuzione = "Automatica"
            MenuItem45.Checked = True
            MenuItem47.Checked = False
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MenuItem47_Click(sender As System.Object, e As System.EventArgs) Handles MenuItem47.Click

        Try
            My.Settings.TipoDiDistribuzione = "Manuale"
            MenuItem45.Checked = False
            MenuItem47.Checked = True
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MenuItem49_Click(sender As System.Object, e As System.EventArgs) Handles MenuItem49.Click

        SelezioneIndirizzi.Show()

    End Sub

    Private Sub MenuItem51_Click(sender As System.Object, e As System.EventArgs) Handles MenuItem51.Click

        Try
            If My.Settings.DataUltimaChiusura <> "" Then
                If My.Settings.DataUltimaChiusura <> My.Settings.DataOggi Then
                    Chiusure.Show()
                Else
                    MsgBox("La chiusura dei messaggi risulta già essere stata effettuata.", MsgBoxStyle.Information, "Asc Server - Chiusure")
                End If
            Else
                Chiusure.Show()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Timer8_Tick(sender As System.Object, e As System.EventArgs) Handles Timer8.Tick

        SaveProgressiveNumbers()

        SaveOthersProgressive()

    End Sub
End Class
