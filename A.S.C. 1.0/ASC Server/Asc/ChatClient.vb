Imports System.Net.Sockets, System.IO

Public Class ChatClient

    Public Event MessageRecieved(ByVal Str As String)
    Public Event ClientExited(ByVal Client As ChatClient)
    Private sWriter As StreamWriter
    Private Client As TcpClient

    Sub New(ByVal xclient As TcpClient)
        Client = xclient
        client.GetStream.BeginRead(New Byte() {0}, 0, 0, AddressOf Read, Nothing)
    End Sub

    Private Sub Read()
        Try
            RaiseEvent MessageRecieved(New StreamReader(Client.GetStream).ReadLine)
            Client.GetStream.BeginRead(New Byte() {0}, 0, 0, New AsyncCallback(AddressOf Read), Nothing)
        Catch ex As Exception
            RaiseEvent ClientExited(Me)
        End Try
    End Sub

    Public Sub Send(ByVal Message As String)
        sWriter = New StreamWriter(Client.GetStream)
        sWriter.WriteLine(Message)
        sWriter.Flush()
    End Sub
End Class
