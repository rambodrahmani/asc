Imports System.Net, System.Net.Sockets, System.IO, System.IO.Directory

Public Class Form1

    Dim scrivi As System.IO.StreamWriter
    Dim PresendingClientList As String = ""
    Dim leggi As System.IO.StreamReader
    Dim letto As String = ""
    Dim FileName As String = ""
    Dim AspettaServer As Integer = 0
    Dim SelectedFolder As String = ""
    Dim CartelleDaSaltare As String = ""
    Dim SelectedClient As String = ""
    Dim ErroreRilevato As String = ""
    Dim MessaggioRicevuto As String = ""
    Dim RicezioneInCorso As Boolean = False
    Dim InvioMessaggiInCorso As Boolean = False
    Dim AcceleraTimer As Boolean = False

#Region "Dichiarazione Server"
    'This is the number of bytes to transfer in one block
    Const PACKET_SIZE As Short = 16384
    Private WithEvents listener As clsListener
    'These are message tags designed for this particular app
    Enum Requests As Byte
        DataFile = 1
        StringMessage = 2
    End Enum
#End Region

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Dim AlertMessage = MsgBox("Si sta chiudendo Asc Exchange, il che renderà Asc Server e Asc Client NON operativi, si è sicuri di voler procedere?", MsgBoxStyle.YesNo, "Asc Exchange - Chiusura")
        If AlertMessage = 6 Then
            Try
                listener.Broadcast(Requests.StringMessage, "ChiusuraAscExchangeInCorso")
            Catch ex As Exception

            End Try
        Else
            e.Cancel = True
        End If

    End Sub
    
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If My.Settings.NomeUtente = "" Or My.Settings.PasswordUtente = "" Then
            MsgBox("Il nome utente e la password non sono ancora stati impostati. Impostarli per poter procedere.", MsgBoxStyle.Information, "Asc Exchange - Utenza")
            Utenza.Show()
            Exit Sub
        End If

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Ricevuti Dal Server")
            End If

            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti Dai Clients") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Ricevuti Dai Clients")
            End If

            If System.IO.Directory.Exists(Application.StartupPath & "\temp") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\temp")
            End If
        Catch ex As Exception

        End Try

        Label3.Text = "Disconnesso"
        LoadServer()

        Try
            Timer1.Start()
            Timer2.Start()
            Timer6.Start()
        Catch ex As Exception

        End Try

        GetLocalIP()
        SetSettings()

    End Sub

    Public Sub GetLocalIP()

        Try
            Dim IPList As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName)
            For Each IPaddress In IPList.AddressList
                If (IPaddress.AddressFamily = Sockets.AddressFamily.InterNetwork) Then
                    Me.Text += " [INDIRIZZO IP MACCHINA: " & IPaddress.ToString & "]"
                    Label5.Text = IPaddress.ToString
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Sub LoadForm()

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Ricevuti Dal Server")
            End If

            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti Dai Clients") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Ricevuti Dai Clients")
            End If

            If System.IO.Directory.Exists(Application.StartupPath & "\temp") Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\temp")
            End If
        Catch ex As Exception

        End Try

        Label3.Text = "Disconnesso"
        LoadServer()

        Try
            Timer1.Start()
            Timer2.Start()
            Timer6.Start()
        Catch ex As Exception

        End Try

        GetLocalIP()
        SetSettings()

    End Sub

#Region "Eventi Server"

    Public Sub LoadServer()

        Try
            listener = New clsListener(3818, PACKET_SIZE)
            Label2.Text = "Server Connesso."
        Catch ex As Exception
            Label2.Text = "Server Disconnesso."
        End Try

    End Sub

    Private Sub listener_ConnectionRequest(ByVal requestor As System.Net.Sockets.TcpClient, ByRef AllowConnection As Boolean) Handles listener.ConnectionRequest

        'Qui viene determinato se accettare le connesioni o meno
        Debug.Print("Connection Request")
        AllowConnection = True

    End Sub

#End Region

#Region "Eventi Ricezione Dati"

    Public Sub listener_DataReceived(ByVal Sender As tcpConnection, ByVal msgtag As Byte, ByVal mStream As System.IO.MemoryStream) Handles listener.DataReceived

        Try
            Select Case msgtag
                Case Requests.DataFile
                    'file data, save to a local file
                    RicezioneInCorso = True
                    If System.IO.Directory.Exists(Application.StartupPath & "\temp") Then
                        If System.IO.File.Exists(Application.StartupPath & "\temp\temp.txt") Then
                            For r = 1 To 9999
                                If System.IO.File.Exists(Application.StartupPath & "\temp\temp(" & r.ToString & ").txt") Then
                                Else
                                    SaveFile(Application.StartupPath & "\temp\temp(" & r.ToString & ").txt", mStream)
                                    Exit Sub
                                End If
                            Next
                        Else
                            SaveFile(Application.StartupPath & "\temp\temp.txt", mStream)
                        End If
                    Else
                        System.IO.Directory.CreateDirectory(Application.StartupPath & "\temp")
                        If System.IO.File.Exists(Application.StartupPath & "\temp\temp.txt") Then
                            For r = 1 To 9999
                                If System.IO.File.Exists(Application.StartupPath & "\temp\temp(" & r.ToString & ").txt") Then
                                Else
                                    SaveFile(Application.StartupPath & "\temp\temp(" & r.ToString & ").txt", mStream)
                                    Exit Sub
                                End If
                            Next
                        Else
                            SaveFile(Application.StartupPath & "\temp\temp.txt", mStream)
                        End If
                    End If
            End Select
        Catch ex As Exception

        End Try

    End Sub

    Public Sub Smista()

        Try

            If RicezioneInCorso = False Then
                If Timer2.Enabled = True Then
                    Timer2.Stop()
                End If
                Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\temp")
                For y = 0 To UBound(FilesCollection)
                    Dim lettore As System.IO.StreamReader
                    Dim letto As String = ""
                    lettore = System.IO.File.OpenText(FilesCollection(y).ToString)
                    letto = lettore.ReadToEnd
                    lettore.Close()
                    Dim MySplitter() As String = Split(letto, vbCrLf)
                    If MySplitter(0).ToString = "MessaggioDalServer" Then
                        Try
                            Dim DefinitiveString As String = ""
                            For z = 2 To UBound(MySplitter)
                                DefinitiveString += MySplitter(z).ToString & vbCrLf
                            Next
                            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server") Then
                                If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString) Then
                                    If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString & "\" & MySplitter(2).ToString & ".txt") Then
                                        For r = 1 To 9999
                                            If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString & "\" & MySplitter(2).ToString & "(" & r.ToString & ").txt") Then
                                            Else
                                                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString & "\" & MySplitter(2).ToString & "(" & r.ToString & ").txt")
                                                scrivi.Write(DefinitiveString)
                                                scrivi.Close()
                                                System.IO.File.Delete(FilesCollection(y).ToString)
                                                Timer2.Start()
                                                Exit Sub
                                            End If
                                        Next
                                    Else
                                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString & "\" & MySplitter(2).ToString & ".txt")
                                        scrivi.Write(DefinitiveString)
                                        scrivi.Close()
                                        System.IO.File.Delete(FilesCollection(y).ToString)
                                        Timer2.Start()
                                        Exit Sub
                                    End If
                                Else
                                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString)
                                    If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString & "\" & MySplitter(2).ToString & ".txt") Then
                                        For r = 1 To 9999
                                            If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString & "\" & MySplitter(2).ToString & "(" & r.ToString & ").txt") Then
                                            Else
                                                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString & "\" & MySplitter(2).ToString & "(" & r.ToString & ").txt")
                                                scrivi.Write(DefinitiveString)
                                                scrivi.Close()
                                                System.IO.File.Delete(FilesCollection(y).ToString)
                                                Timer2.Start()
                                                Exit Sub
                                            End If
                                        Next
                                    Else
                                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString & "\" & MySplitter(2).ToString & ".txt")
                                        scrivi.Write(DefinitiveString)
                                        scrivi.Close()
                                        System.IO.File.Delete(FilesCollection(y).ToString)
                                        Timer2.Start()
                                        Exit Sub
                                    End If
                                End If
                            Else
                                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Ricevuti Dal Server")
                                If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString) Then
                                    If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString & "\" & MySplitter(2).ToString & ".txt") Then
                                        For r = 1 To 9999
                                            If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString & "\" & MySplitter(2).ToString & "(" & r.ToString & ").txt") Then
                                            Else
                                                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString & "\" & MySplitter(2).ToString & "(" & r.ToString & ").txt")
                                                scrivi.Write(DefinitiveString)
                                                scrivi.Close()
                                                System.IO.File.Delete(FilesCollection(y).ToString)
                                                Timer2.Start()
                                                Exit Sub
                                            End If
                                        Next
                                    Else
                                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString & "\" & MySplitter(2).ToString & ".txt")
                                        scrivi.Write(DefinitiveString)
                                        scrivi.Close()
                                        System.IO.File.Delete(FilesCollection(y).ToString)
                                        Timer2.Start()
                                        Exit Sub
                                    End If
                                Else
                                    System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString)
                                    If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString & "\" & MySplitter(2).ToString & ".txt") Then
                                        For r = 1 To 9999
                                            If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString & "\" & MySplitter(2).ToString & "(" & r.ToString & ").txt") Then
                                            Else
                                                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString & "\" & MySplitter(2).ToString & "(" & r.ToString & ").txt")
                                                scrivi.Write(DefinitiveString)
                                                scrivi.Close()
                                                System.IO.File.Delete(FilesCollection(y).ToString)
                                                Timer2.Start()
                                                Exit Sub
                                            End If
                                        Next
                                    Else
                                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & MySplitter(1).ToString & "\" & MySplitter(2).ToString & ".txt")
                                        scrivi.Write(DefinitiveString)
                                        scrivi.Close()
                                        System.IO.File.Delete(FilesCollection(y).ToString)
                                        Timer2.Start()
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            Timer2.Start()
                        End Try
                    ElseIf MySplitter(0).ToString = "MessaggioDalClient" Then
                        Try
                            Dim DefinitiveString As String = ""
                            For r = 1 To UBound(MySplitter)
                                DefinitiveString += MySplitter(r).ToString & vbCrLf
                            Next
                            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti Dai Clients") Then
                                If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti Dai Clients\" & MySplitter(1).ToString & ".txt") Then
                                    For r = 1 To 9999
                                        If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti Dai Clients\" & MySplitter(1).ToString & "(" & r & ")" & ".txt") Then
                                        Else
                                            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti Dai Clients\" & MySplitter(1).ToString & "(" & r & ")" & ".txt")
                                            scrivi.Write(DefinitiveString)
                                            scrivi.Close()
                                            System.IO.File.Delete(FilesCollection(y).ToString)
                                            Timer2.Start()
                                            Exit Sub
                                        End If
                                    Next
                                Else
                                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti Dai Clients\" & MySplitter(1).ToString & ".txt")
                                    scrivi.Write(DefinitiveString)
                                    scrivi.Close()
                                    System.IO.File.Delete(FilesCollection(y).ToString)
                                    Timer2.Start()
                                    Exit Sub
                                End If
                            Else
                                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Messaggi Ricevuti Dai Clients")
                                If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti Dai Clients\" & MySplitter(1).ToString & ".txt") Then
                                    For r = 1 To 9999
                                        If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti Dai Clients\" & MySplitter(1).ToString & "(" & r & ")" & ".txt") Then
                                        Else
                                            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti Dai Clients\" & MySplitter(1).ToString & "(" & r & ")" & ".txt")
                                            scrivi.Write(DefinitiveString)
                                            scrivi.Close()
                                            System.IO.File.Delete(FilesCollection(y).ToString)
                                            Timer2.Start()
                                            Exit Sub
                                        End If
                                    Next
                                Else
                                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti Dai Clients\" & MySplitter(1).ToString & ".txt")
                                    scrivi.Write(DefinitiveString)
                                    scrivi.Close()
                                    System.IO.File.Delete(FilesCollection(y).ToString)
                                    Timer2.Start()
                                    Exit Sub
                                End If
                            End If
                        Catch ex As Exception
                            Timer2.Start()
                        End Try
                    End If
                Next
            End If
        Catch ex As Exception
            If Timer2.Enabled = False Then
                Timer2.Start()
            End If
            Exit Sub
        End Try

    End Sub

    Private Sub SaveFile(ByVal FilePath As String, ByVal mstream As System.IO.MemoryStream)

        Try
            'save file to path specified
            Dim FS As New FileStream(FilePath, IO.FileMode.Create, IO.FileAccess.Write)
            mstream.WriteTo(FS)
            mstream.Flush()
            mstream.Close()
            FS.Close()
            RicezioneInCorso = False
        Catch ex As Exception

        End Try

    End Sub

    Private Sub listener_MessageReceived(ByVal sender As tcpConnection, ByVal message As Byte) Handles listener.MessageReceived


    End Sub

    Private Sub listener_StringReceived(ByVal Sender As tcpConnection, ByVal msgTag As Byte, ByVal message As String) Handles listener.StringReceived

        Try
            Select Case msgTag
                Case Requests.DataFile
                    If message = "ServerIsConnected=True" Then
                        If Timer4.Enabled = True Then
                            Timer4.Stop()
                        End If
                        AspettaServer = 0
                        Dim filescollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Messaggi Ricevuti Dai Clients")
                        For r = 0 To UBound(filescollection)
                            Sender.SendFile(msgTag, filescollection(r).ToString)
                            System.IO.File.Delete(filescollection(r).ToString)
                            If r = filescollection.Count - 1 Then
                                InvioMessaggiInCorso = False
                                Exit Sub
                            End If
                            SendMessagesToServer()
                            Exit Sub
                        Next
                    ElseIf message = "ClientIsConnected=True" Then
                        If Timer5.Enabled = True Then
                            Timer5.Stop()
                        End If
                        AspettaServer = 0
                        Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & SelectedClient)
                        For r = 0 To UBound(FilesCollection)
                            Sender.SendFile(msgTag, FilesCollection(r).ToString)
                            System.IO.File.Delete(FilesCollection(r).ToString)
                            AcceleraTimer = True
                            InvioMessaggiInCorso = False
                            Exit Sub
                        Next
                    End If
                Case Requests.StringMessage
                    MessaggioRicevuto = message
            End Select
        Catch ex As Exception

        End Try

    End Sub

#End Region

    Public Sub DeleteFolder()

        Try
            System.IO.Directory.Delete(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & SelectedFolder, True)
        Catch ex As Exception

        End Try

    End Sub

    Public Sub DeleteMessage()

        Try
            System.IO.File.Delete(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & SelectedFolder)
        Catch ex As Exception
        End Try

    End Sub

    Public Sub SendMessagesForFolder()

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & SelectedFolder) Then
                Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Messaggi Ricevuti Dal Server\" & SelectedFolder)
                If FilesCollection.Count > 0 Then
                    Dim definitivestring As String = ""
                    For r = 0 To UBound(FilesCollection)
                        Dim FileInfo As New FileInfo(FilesCollection(r).ToString)
                        definitivestring += FileInfo.Name & "|"
                    Next
                    listener.Broadcast(Requests.StringMessage, "AscExchangeMessaggiCartella|" & definitivestring)
                Else
                    listener.Broadcast(Requests.StringMessage, "AscExchangeMessaggiCartella|Nulla|")
                End If
            Else
                listener.Broadcast(Requests.StringMessage, "AscExchangeMessaggiCartella|Nulla|")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub SendAscExchangeCode()

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server") Then
                Dim MainFolderSubFoldersCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\Messaggi Ricevuti Dal Server")
                Dim FoldersContainer As String = ""
                If MainFolderSubFoldersCollection.Count > 0 Then
                    For r = 0 To UBound(MainFolderSubFoldersCollection)
                        Dim DirectoryInfo As New DirectoryInfo(MainFolderSubFoldersCollection(r).ToString)
                        FoldersContainer += DirectoryInfo.Name.ToString & "|"
                    Next
                Else
                    listener.Broadcast(Requests.StringMessage, "AscExchangeCode|Nulla")
                End If
                listener.Broadcast(Requests.StringMessage, "AscExchangeCode|" & FoldersContainer)
            Else
                listener.Broadcast(Requests.StringMessage, "AscExchangeCode|Nulla")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub SendSettings()

        Try
            listener.Broadcast(Requests.StringMessage, "InvioSettingsDallExchange" & "|" & My.Settings.FrequenzaControllo & "|" & My.Settings.FrequenzaInvio & "|")
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        If Me.WindowState = FormWindowState.Minimized Then
            Try
                Me.Hide()
                NotifyIcon1.Visible = True
                NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
                NotifyIcon1.BalloonTipText = "Asc Exchange è stato minimizzato a icona, ma è comunque attivo!"
                NotifyIcon1.BalloonTipTitle = "Asc Exchange è Minimzzato"
                NotifyIcon1.ShowBalloonTip(6)
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick

        If e.Button = Windows.Forms.MouseButtons.Left Then
            Try
                Me.Show()
                Me.TopMost = True
                Me.WindowState = FormWindowState.Normal
                NotifyIcon1.Visible = False
                Me.TopMost = False
            Catch ex As Exception

            End Try
        End If

    End Sub

    Public Sub SetSettings()

        If Timer3.Enabled = True Then
            Timer3.Stop()
        End If
        If My.Settings.FrequenzaControllo <> 0 Then
            Try
                Timer3.Interval = My.Settings.FrequenzaControllo
            Catch ex As Exception

            End Try
        Else
            Try
                Timer3.Interval = "15000"
            Catch ex As Exception

            End Try
        End If
        If Timer3.Enabled = False Then
            Timer3.Start()
        End If

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        Me.Close()

    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click

        Me.WindowState = FormWindowState.Minimized

    End Sub

    Private Sub MenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem8.Click

        My.Settings.x = 2
        Accesso.Show()

    End Sub

    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem9.Click

        My.Settings.x = 1
        Accesso.Show()

    End Sub

    Public Sub SendConnectedClientsList()

        Try
            Dim ConnectedClientsList As String = ""
            For r = 0 To ListBox1.Items.Count - 1
                ConnectedClientsList += ListBox1.Items(r).ToString & "|"
            Next
            listener.Broadcast(Requests.StringMessage, "ConnectedClientListAnswearing|" & ConnectedClientsList)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Try
            If MessaggioRicevuto = "ServerConnected" Then
                Label3.Text = "Connesso"
                MessaggioRicevuto = ""
            ElseIf MessaggioRicevuto = "ServerDisconnected" Then
                Label3.Text = "Disconnesso"
                MessaggioRicevuto = ""
            ElseIf MessaggioRicevuto.Contains("NewClientConnected") Then
                Dim MySplitter() As String = Split(MessaggioRicevuto, "|")
                MessaggioRicevuto = ""
                If ListBox1.Items.Count > 0 Then
                    For r = 0 To ListBox1.Items.Count - 1
                        If ListBox1.Items(r).ToString = MySplitter(1).ToString Then
                            Exit For
                        Else
                            If r = ListBox1.Items.Count - 1 Then
                                ListBox1.Items.Add(MySplitter(1).ToString)
                            End If
                        End If
                    Next
                Else
                    ListBox1.Items.Add(MySplitter(1).ToString)
                End If
            ElseIf MessaggioRicevuto.Contains("ClientDisconnected") Then
                Dim MySplitter() As String = Split(MessaggioRicevuto, "|")
                MessaggioRicevuto = ""
                ListBox1.Items.Remove(MySplitter(1).ToString)
            ElseIf MessaggioRicevuto = "ServerRequestsConnectedClientList" Then
                MessaggioRicevuto = ""
                SendConnectedClientsList()
            ElseIf MessaggioRicevuto = "ServerRequestsSettings" Then
                MessaggioRicevuto = ""
                SendSettings()
            ElseIf MessaggioRicevuto.Contains("NewSettingsFromServerForExchange") Then
                Try
                    Dim MySplitter() As String = Split(MessaggioRicevuto, "|")
                    MessaggioRicevuto = ""
                    My.Settings.FrequenzaControllo = (Val(MySplitter(1).ToString) * 1000)
                    My.Settings.FrequenzaInvio = (Val(MySplitter(2).ToString) * 1000)
                    SetSettings()
                Catch ex As Exception

                End Try
            ElseIf MessaggioRicevuto = "ServerRequestsCode" Then
                MessaggioRicevuto = ""
                SendAscExchangeCode()
            ElseIf MessaggioRicevuto.Contains("MessagesRequestFor") Then
                Try
                    Dim MySplitter() As String = Split(MessaggioRicevuto, "|")
                    MessaggioRicevuto = ""
                    SelectedFolder = MySplitter(1).ToString
                    SendMessagesForFolder()
                Catch ex As Exception

                End Try
            ElseIf MessaggioRicevuto.Contains("DeleteMessageRequestfor") Then
                Try
                    Dim MySplitter() As String = Split(MessaggioRicevuto, "|")
                    MessaggioRicevuto = ""
                    SelectedFolder = MySplitter(1).ToString
                    DeleteMessage()
                Catch ex As Exception

                End Try
            ElseIf MessaggioRicevuto.Contains("DeleteFolderRequestfor") Then
                Try
                    Dim MySplitter() As String = Split(MessaggioRicevuto, "|")
                    MessaggioRicevuto = ""
                    SelectedFolder = MySplitter(1).ToString
                    DeleteFolder()
                Catch ex As Exception

                End Try
            ElseIf MessaggioRicevuto.Contains("ControlloConnettivitàPerL'ufficio") Then
                Try
                    Try
                        Dim MySplitter() As String = Split(MessaggioRicevuto, "|")
                        MessaggioRicevuto = ""
                        Dim ListaClients As String = ""
                        For r = 0 To ListBox1.Items.Count - 1
                            For x = 0 To UBound(MySplitter)
                                If ListBox1.Items(r).ToString = MySplitter(x).ToString Then
                                    listener.Broadcast(Requests.StringMessage, "L'UfficioPerCuiSiAvevaFattoRichiestaEConnesso")
                                    Exit Sub
                                End If
                            Next
                        Next
                        listener.Broadcast(Requests.StringMessage, "L'UfficioPerCuiSiAvevaFattoRichiestaNonEConnesso")
                    Catch ex As Exception

                    End Try
                Catch ex As Exception

                End Try
            ElseIf MessaggioRicevuto = "IlServerRichiedeL'inizioDelProcessoDiInvioCode" Then
                MessaggioRicevuto = ""
                ProcessoOttenimentoCode()
            Else
                MessaggioRicevuto = ""
            End If
        Catch ex As Exception
            If Timer1.Enabled = False Then
                Timer1.Start()
            End If
            Exit Sub
        End Try

    End Sub

    Public Sub ProcessoOttenimentoCode()

        Try
            Dim MessaggioFinaleDaComunicare As String = ""
            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server") Then
                Dim MainFolderSubFoldersCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\Messaggi Ricevuti Dal Server")
                For r = 0 To UBound(MainFolderSubFoldersCollection)
                    Dim FolderName As New DirectoryInfo(MainFolderSubFoldersCollection(r).ToString)
                    MessaggioFinaleDaComunicare += FolderName.Name.ToString & "#"
                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(MainFolderSubFoldersCollection(r).ToString)
                    For x = 0 To UBound(FilesCollection)
                        Dim FileInfo As New FileInfo(FilesCollection(x).ToString)
                        MessaggioFinaleDaComunicare += FileInfo.Name.ToString & "|"
                    Next
                    MessaggioFinaleDaComunicare += "#§"
                Next
                listener.Broadcast(Requests.StringMessage, "RispostaDalProcessoDiOttenimentoCode§" & MessaggioFinaleDaComunicare)
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick

        Try
            If System.IO.Directory.Exists(Application.StartupPath & "\temp") Then
                Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\temp")
                If FilesCollection.Count > 0 Then
                    Smista()
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub SendMessagesToServer()

        Try
            If Timer3.Enabled = True Then
                Timer3.Stop()
            End If
            InvioMessaggiInCorso = True
            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti Dai Clients") Then
                Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\Messaggi Ricevuti Dai Clients")
                If FilesCollection.Count > 0 Then
                    ChkServerIsConnected()
                Else
                    SendMessagesToClients()
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ChkServerIsConnected()

        Try
            Timer4.Start()
            listener.Broadcast(Requests.StringMessage, "ChkServerIsConnected")
        Catch ex As Exception

        End Try

    End Sub

    Public Sub SendMessagesToClients()

        Try
            If Timer4.Enabled = True Then
                Timer4.Stop()
            End If
            If Timer5.Enabled = True Then
                Timer5.Stop()
            End If
            If System.IO.Directory.Exists(Application.StartupPath & "\Messaggi Ricevuti Dal Server") Then
                Dim DirectoriesCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\Messaggi Ricevuti Dal Server")
                If DirectoriesCollection.Count > 0 Then
                    ChkSelectedClientIsConncted()
                Else
                    CartelleDaSaltare = ""
                    InvioMessaggiInCorso = False
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ChkSelectedClientIsConncted()

        Dim DirectoriesCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\Messaggi Ricevuti Dal Server")
        If DirectoriesCollection.Count > 0 Then
            For r = 0 To UBound(DirectoriesCollection)
                Dim FilesCollection() As String = System.IO.Directory.GetFiles(DirectoriesCollection(r).ToString)
                If FilesCollection.Count > 0 Then
                    Dim DirecotyInfo As New DirectoryInfo(DirectoriesCollection(r).ToString)
                    If CartelleDaSaltare.Contains(DirecotyInfo.Name.ToString) Then
                        If r = DirectoriesCollection.Count - 1 Then
                            CartelleDaSaltare = ""
                            InvioMessaggiInCorso = False
                        End If
                    Else
                        SelectedClient = DirecotyInfo.Name.ToString
                        listener.Broadcast(Requests.StringMessage, "AscExchangeChkClientConnection|" & SelectedClient)
                        Timer5.Start()
                        Exit For
                    End If
                Else
                    System.IO.Directory.Delete(DirectoriesCollection(r).ToString)
                    If r = DirectoriesCollection.Count - 1 Then
                        CartelleDaSaltare = ""
                        InvioMessaggiInCorso = False
                    End If
                End If
            Next
        Else
            SendMessagesToServer()
        End If

    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick

        If AspettaServer = 7 Then
            AspettaServer = 0
            SendMessagesToClients()
        Else
            AspettaServer += 1
        End If

    End Sub

    Private Sub Timer5_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer5.Tick

        If AspettaServer = 7 Then
            CartelleDaSaltare += SelectedClient & vbCrLf
            AspettaServer = 0
            SendMessagesToClients()
        Else
            AspettaServer += 1
        End If

    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick

        SendMessagesToServer()

    End Sub

    Private Sub Timer6_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer6.Tick

        Try
            If InvioMessaggiInCorso = False Then
                If Timer3.Enabled = False Then
                    If AcceleraTimer = True Then
                        AcceleraTimer = False
                        Timer3.Interval = "3000"
                        Timer3.Start()
                    Else
                        Timer3.Interval = My.Settings.FrequenzaControllo
                        Timer3.Start()
                    End If
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub
End Class