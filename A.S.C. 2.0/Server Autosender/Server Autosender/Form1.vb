Imports System.Net, System.Net.Sockets, System.IO, System.IO.Directory

Public Class Form1

    Dim PresendingClientList As String = ""
    Dim leggi As System.IO.StreamReader
    Dim letto As String = ""
    Dim FileName As String = ""
    Dim Listener As TcpListener
    Dim Client As TcpClient
    Dim ClientList As New List(Of ChatClient)
    Dim sReader As StreamReader
    Dim cClient As ChatClient
    Delegate Sub _xUpdate(ByVal Str As String, ByVal Relay As Boolean)
    Dim ServerIsActive As Integer = 0
    Dim ServerIsActive1 As Integer = 0
    Dim sWriter As StreamWriter
    Dim NickPrefix As Integer = New Random().Next(1111, 9999)
    Dim scrivi As System.IO.StreamWriter

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        If PresendingClientList <> "" Then
            ChkFilesToBeSent()
        Else
            RequestConnectedClientList()
        End If

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Dim ActiveProcessesCounter As Integer = 0
            Dim ActiveProcessesList() As Process = Diagnostics.Process.GetProcesses
            Dim SingleProcess As Process
            For Each SingleProcess In ActiveProcessesList
                If SingleProcess.ProcessName.Contains("Server Autosender") Then
                    ActiveProcessesCounter += 1
                End If
            Next
            If ActiveProcessesCounter > 1 Then
                Me.Close()
            Else
                Try
                    System.IO.File.Delete(Application.StartupPath & "\ServerIsBeingClosed.txt")
                Catch ex As Exception

                End Try

                Try
                    If System.IO.File.Exists(Application.StartupPath & "\ServerIsActive.txt") Then
                        leggi = System.IO.File.OpenText(Application.StartupPath & "\ServerIsActive.txt")
                        letto = leggi.ReadToEnd
                        leggi.Close()
                    End If
                Catch ex As Exception

                End Try

                If letto = "ServerIsActive" Then
                    ServerIsActive = 1
                Else
                    ServerIsActive = 0
                End If

                If ServerIsActive = 1 Then
                    ServerIsActive1 = 1
                    LoadAsClient()
                Else
                    ServerIsActive1 = 0
                    LoadAsServer()
                End If
                Timer2.Start()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub LoadAsServer()

        If ServerIsActive1 = 0 Then
            ServerIsActive1 = 1
            Try
                Me.Hide()
                Listener = New TcpListener(IPAddress.Any, 3818)
                Listener.Start()
                Listener.BeginAcceptTcpClient(New AsyncCallback(AddressOf AcceptClient), Listener)
                RequestConnectedClientList()
                NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
                NotifyIcon1.BalloonTipText = "[Modalità Server] - Invio automatico messaggi in coda ai client iniziato..."
                NotifyIcon1.BalloonTipTitle = "Server Autosender Avviato"
                NotifyIcon1.ShowBalloonTip(5)
                Timer1.Start()
            Catch ex As Exception

            End Try
        End If

    End Sub

    Public Sub LoadAsClient()

        If ServerIsActive1 = 1 Then
            ServerIsActive1 = 0
            Try
                Me.Hide()
                Client = New TcpClient("127.0.0.1", CInt("3818"))
                Client.GetStream.BeginRead(New Byte() {0}, 0, 0, New AsyncCallback(AddressOf Read), Nothing)
                RequestConnectedClientList()
                NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
                NotifyIcon1.BalloonTipText = "[Modalità Client] - Invio automatico messaggi in coda ai client iniziato..."
                NotifyIcon1.BalloonTipTitle = "Server Autosender Avviato"
                NotifyIcon1.ShowBalloonTip(5)
                Timer1.Start()
            Catch ex As Exception

            End Try
        End If

    End Sub

    Sub AcceptClient(ByVal ar As IAsyncResult)

        Try
            cClient = New ChatClient(Listener.EndAcceptTcpClient(ar))
            AddHandler (cClient.MessageRecieved), AddressOf MessageRecieved
            AddHandler (cClient.ClientExited), AddressOf ClientExited
            ClientList.Add(cClient)
            Listener.BeginAcceptTcpClient(New AsyncCallback(AddressOf AcceptClient), Listener)
        Catch ex As Exception

        End Try

    End Sub

    Sub MessageRecieved(ByVal Str As String)

        Try
            xUpdate(Str, True)
        Catch ex As Exception

        End Try

    End Sub

    Sub ClientExited(ByVal Client As ChatClient)

        Try
            ClientList.Remove(Client)
        Catch ex As Exception

        End Try

    End Sub

    Sub xUpdate(ByVal Str As String, ByVal Relay As Boolean)

        On Error Resume Next
        If InvokeRequired Then
            Invoke(New _xUpdate(AddressOf xUpdate), Str, Relay)
        Else
            TextBox2.AppendText(Str & vbNewLine)
            If Relay Then Send(Str)
        End If

    End Sub

    Sub Send(ByVal Str As String)

        For i As Integer = 0 To ClientList.Count - 1
            Try
                ClientList(i).Send(Str)
            Catch
                ClientList.RemoveAt(i)
            End Try
        Next

    End Sub

    Private Sub SendToClient(ByVal Str As String)

        Try
            sWriter = New StreamWriter(Client.GetStream)
            sWriter.WriteLine(Str)
            sWriter.Flush()
        Catch ex As Exception

        End Try

    End Sub

    Sub Read(ByVal ar As IAsyncResult)

        Try
            xUpdate(New StreamReader(Client.GetStream).ReadLine, True)
            Client.GetStream.BeginRead(New Byte() {0}, 0, 0, AddressOf Read, Nothing)
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ChkFilesToBeSent()

        Try
            Dim DirectoriesCollection() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\temp")
            For r = 0 To UBound(DirectoriesCollection)
                Dim DirectoryInfo As New IO.DirectoryInfo(DirectoriesCollection(r).ToString)
                If PresendingClientList.Contains(DirectoryInfo.Name.ToString) Then
                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(DirectoriesCollection(r).ToString)
                    If FilesCollection.Count > 0 Then
                        Timer1.Interval = "1000"
                        For z = 0 To UBound(FilesCollection)
                            Dim FileInfo As New FileInfo(FilesCollection(z).ToString)
                            Dim Splitter() As String = Split(FileInfo.Name.ToString, ".txt")
                            FileName = Splitter(0).ToString
                            leggi = System.IO.File.OpenText(FilesCollection(z).ToString)
                            letto = leggi.ReadToEnd
                            leggi.Close()
                            Try
                                If ServerIsActive = 0 Then
                                    xUpdate("MessaggioDaSalvareDalServer|" & DirectoryInfo.Name.ToString & "|" & FileName & "|" & letto, True)
                                Else
                                    SendToClient("MessaggioDaSalvareDalServer|" & DirectoryInfo.Name.ToString & "|" & FileName & "|" & letto)
                                End If
                            Catch ex As Exception

                            End Try
                            System.IO.File.Delete(FilesCollection(z).ToString)
                            RequestConnectedClientList()
                            Exit Sub
                        Next
                    Else
                        Timer1.Interval = "15000"
                    End If
                End If
            Next
            RequestConnectedClientList()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub RequestConnectedClientList()

        PresendingClientList = ""
        Try
            TextBox2.Clear()
            If ServerIsActive1 = 1 Then
                xUpdate("PreSendClientsListRequest", True)
            Else
                SendToClient("PreSendClientsListRequest")
            End If
            TextBox2.Clear()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

        If TextBox2.Text.Contains("PreSendingClientList") Then
            Try
                Dim mysplitter() As String
                mysplitter = Split(TextBox2.Text, "|")
                For r = 0 To UBound(mysplitter)
                    PresendingClientList += mysplitter(r).ToString & vbCrLf
                Next
                TextBox2.Clear()
            Catch ex As Exception

            End Try
        End If

        If TextBox2.Text.Contains("Email da salvare") Then
            If ServerIsActive = 0 Then
                Dim mysplitter() As String = Split(TextBox2.Text, "|")
                Try
                    If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(1).ToString & ".txt") Then
                        For r = 1 To 300
                            If System.IO.File.Exists(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(1).ToString & r & ".txt") Then
                            Else
                                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(1).ToString & r & ".txt")
                                For x = 2 To UBound(mysplitter)
                                    scrivi.Write((mysplitter(x)).ToString & vbCrLf)
                                Next
                                scrivi.Close()
                                Exit For
                            End If
                        Next
                    Else
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Messaggi Ricevuti\" & mysplitter(1).ToString & ".txt")
                        For r = 2 To UBound(mysplitter)
                            scrivi.Write((mysplitter(r)).ToString & vbCrLf)
                        Next
                        scrivi.Close()
                    End If
                    TextBox2.Clear()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        End If

    End Sub

    Private Sub NotifyIcon1_BalloonTipClicked(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.BalloonTipClicked

        Try
            If NotifyIcon1.BalloonTipText = "Invio automatico messaggi ai client completato..." Then
                Me.Close()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick

        If e.Button = Windows.Forms.MouseButtons.Left Then
            Try
                Dim FilesToBeSent As Integer = 0
                Dim DirecColl() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\temp")
                For r = 0 To UBound(DirecColl)
                    Dim FilesColl() As String = System.IO.Directory.GetFiles(DirecColl(r).ToString)
                    FilesToBeSent += FilesColl.Count
                Next
                MsgBox("Server Autosender:" & vbCrLf & "Sono ancora in attesa di essere inviati " & FilesToBeSent & " messaggi.", MsgBoxStyle.Information, "Asc Server Autosender")
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub ChiudiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChiudiToolStripMenuItem.Click

        Try
            Me.Close()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MessaggiInAttesaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MessaggiInAttesaToolStripMenuItem.Click

        Try
            Dim FilesToBeSent As Integer = 0
            Dim DirecColl() As String = System.IO.Directory.GetDirectories(Application.StartupPath & "\temp")
            For r = 0 To UBound(DirecColl)
                Dim FilesColl() As String = System.IO.Directory.GetFiles(DirecColl(r).ToString)
                FilesToBeSent += FilesColl.Count
            Next
            MsgBox("Server Autosender:" & vbCrLf & "Sono ancora in attesa di essere inviati " & FilesToBeSent & " messaggi.", MsgBoxStyle.Information, "Asc Server Autosender")
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick

        Try
            ChkServerIsActive()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ChkServerIsActive()

        Try
            If System.IO.File.Exists(Application.StartupPath & "\ServerIsBeingClosed.txt") Then
                Me.Close()
            End If

            If System.IO.File.Exists(Application.StartupPath & "\ServerIsActive.txt") Then
                leggi = System.IO.File.OpenText(Application.StartupPath & "\ServerIsActive.txt")
                letto = leggi.ReadToEnd
                leggi.Close()
            End If

            If letto = "ServerIsActive" Then
                ServerIsActive = 1
            Else
                ServerIsActive = 0
            End If

            If ServerIsActive = 1 Then
                LoadAsClient()
            Else
                LoadAsServer()
            End If

        Catch ex As Exception

        End Try

    End Sub

End Class
