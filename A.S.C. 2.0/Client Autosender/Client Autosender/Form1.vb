Imports System.IO
Imports System.Net.Sockets

Public Class Form1

    Dim leggi As System.IO.StreamReader
    Dim letto As String = ""
    Dim ServerIP As String = ""
    Dim Client As TcpClient
    Dim sWriter As StreamWriter
    Dim NickPrefix As Integer = New Random().Next(1111, 9999)
    Delegate Sub _xUpdate(ByVal Str As String)
    Dim Connected As String = ""

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Dim ActiveProcessesCounter As Integer = 0
            Dim ActiveProcessesList() As Process = Diagnostics.Process.GetProcesses
            Dim SingleProcess As Process
            For Each SingleProcess In ActiveProcessesList
                If SingleProcess.ProcessName.Contains("Client Autosender") Then
                    ActiveProcessesCounter += 1
                End If
            Next
            If ActiveProcessesCounter > 1 Then
                Me.Close()
            Else
                GetConnectionInfo()
                Me.Hide()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub GetConnectionInfo()

        Me.Hide()
        Try
            leggi = System.IO.File.OpenText(Application.StartupPath & "\ConnectionInfo.txt")
            letto = leggi.ReadToEnd
            leggi.Close()
            Dim MySplitter() As String = Split(letto, "|")
            ServerIP = MySplitter(0).ToString
            ConnectToServer()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ConnectToServer()

        AutoSender.Start()

        Try
            Client = New TcpClient(ServerIP, CInt("3818"))
            Client.GetStream.BeginRead(New Byte() {0}, 0, 0, New AsyncCallback(AddressOf Read), Nothing)
            Connected = "Connected"
        Catch ex As Exception
            Connected = "Disconnected"
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


    Sub Read(ByVal ar As IAsyncResult)

        Try
            xUpdate(New StreamReader(Client.GetStream).ReadLine)
            Client.GetStream.BeginRead(New Byte() {0}, 0, 0, AddressOf Read, Nothing)
        Catch ex As Exception
            Connected = "Disconnected"
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

        End Try

    End Sub

    Private Sub AutoSender_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoSender.Tick

        AutoSendProcess()

    End Sub

    Public Sub AutoSendProcess()

        Try
            If System.IO.File.Exists(Application.StartupPath & "\ClientBeingClosed.txt") Then
                Me.Close()
            End If
        Catch ex As Exception

        End Try

        If Connected = "Disconnected" Then
            Try
                Client = New TcpClient(ServerIP, CInt("3818"))
                Client.GetStream.BeginRead(New Byte() {0}, 0, 0, New AsyncCallback(AddressOf Read), Nothing)
                Connected = "Connected"
            Catch ex As Exception
                Connected = "Disconnected"
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

    Private Sub ChiudiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChiudiToolStripMenuItem.Click

        Me.Close()

    End Sub

    Private Sub MessaggiInAttesaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MessaggiInAttesaToolStripMenuItem.Click

        FilesRimanenti()

    End Sub

    Public Sub FilesRimanenti()

        Try
            Dim FilesCollection() As String = System.IO.Directory.GetFiles(Application.StartupPath & "\temp")
            MsgBox("Client Autosender:" & vbCrLf & "Sono ancora in attesa di essere inviati " & FilesCollection.Count.ToString & " messaggi.", MsgBoxStyle.Information, "Asc Client Autosender")
        Catch ex As Exception

        End Try

    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick

        If e.Button = Windows.Forms.MouseButtons.Left Then
            FilesRimanenti()
        End If

    End Sub
End Class
