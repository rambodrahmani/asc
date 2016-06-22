Imports System.Net.Sockets
Imports System.IO
Imports System.Collections.Generic

Public Class clsListener
    Private Enum RequestTags As Byte
        Connect = 1
        Disconnect = 2
        DataTransfer = 3
        StringTransfer = 4
        ByteTransfer = 5
    End Enum

    Private Server As TcpListener
    Private Clients As Dictionary(Of Int16, tcpConnection)
    Private nClients As Byte
    Private mPacketSize As Int16

#Region "Event Declarations"
    Public Event ConnectionRequest(ByVal Requestor As TcpClient, ByRef AllowConnection As Boolean)
    Public Event DataReceived(ByVal Client As tcpConnection, ByVal msgTag As Byte, ByVal mStream As MemoryStream)
    Public Event MessageReceived(ByVal Client As tcpConnection, ByVal message As Byte)
    Public Event StringReceived(ByVal Client As tcpConnection, ByVal msgTag As Byte, ByVal message As String)
    Public Event Disconnect(ByVal Client As tcpConnection)
#End Region

    Public Sub New(ByVal nPort As Integer, Optional ByVal PacketSize As Int16 = 4096)

        Try
            'constructor:
            'Creates a dictionary to store client connections
            'and starts listening for connection requests on the
            'port specified
            mPacketSize = PacketSize
            Clients = New Dictionary(Of Int16, tcpConnection)
            Listen(nPort)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Listen(ByVal nPort As Integer)

        Try
            'instantiate the listener class
            Server = New TcpListener(System.Net.IPAddress.Any, nPort)
            'start listening for connection requests
            Server.Start()
            'asyncronously wait for connection requests
            Server.BeginAcceptTcpClient(New AsyncCallback(AddressOf DoAcceptTcpClientCallback), Server)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub DoAcceptTcpClientCallback(ByVal ar As IAsyncResult)

        Try
            ' Processes the client connection request.
            ' A maximum of 255 connections can be made.
            Dim Allowed As Boolean

            ' Get the listener that handles the client request.
            Dim listener As TcpListener = CType(ar.AsyncState, TcpListener)

            Dim client As TcpClient = listener.EndAcceptTcpClient(ar)
            'ask end user if he wants to accept this connection request
            RaiseEvent ConnectionRequest(client, Allowed)

            If Not Allowed Then
                SendMessage(RequestTags.Disconnect, 0, client)
                client.GetStream.Close()
                client.Close()
                client = Nothing
            Else
                'add client to list of connected clients
                SendMessage(RequestTags.Connect, mPacketSize, client)
                Dim newClient As tcpConnection = New tcpConnection(client, mPacketSize)
                newClient.Tag = nClients
                Clients.Add(nClients, newClient)
                nClients += CByte(1)

                'need to delegate events for client since client
                'is created dynamically and is going to be put into a 
                'collection so this is how we are going to propagate 
                'the events up to the end user
                AddHandler newClient.DataReceived, AddressOf onDataReceived
                AddHandler newClient.MessageReceived, AddressOf onMessageReceived
                AddHandler newClient.StringReceived, AddressOf onStringReceived
                AddHandler newClient.Disconnect, AddressOf onClientDisconnect
                newClient = Nothing
            End If

            'continue listening
            listener.BeginAcceptTcpClient(New AsyncCallback(AddressOf DoAcceptTcpClientCallback), listener)
        Catch ex As Exception

        End Try

    End Sub 'DoAcceptTcpClientCallback

    Private Sub SendMessage(ByVal msgTag As Byte, ByVal msg As Int16, ByVal client As TcpClient)
        ' This subroutine sends a one-byte response through a tcpClient
        ' Synclock ensure that no other threads try to use the stream at the same time.
        SyncLock client.GetStream
            Dim writer As New BinaryWriter(client.GetStream)
            writer.Write(msgTag)
            writer.Write(msg)
            ' Make sure all data is sent now.
            writer.Flush()
        End SyncLock
    End Sub

#Region "Overloaded Broadcast methods send data to all client connections"
    Public Sub Broadcast(ByVal msgTag As Byte)
        For Each myClient As KeyValuePair(Of Short, tcpConnection) In Clients
            Clients(myClient.Key).Send(msgTag)
        Next
    End Sub

    Public Sub Broadcast(ByVal msgTag As Byte, ByVal strX As String)
        For Each myClient As KeyValuePair(Of Short, tcpConnection) In Clients
            Clients(myClient.Key).Send(msgTag, strX)
        Next
    End Sub

    Public Sub Broadcast(ByVal msgTag As Byte, ByVal byteData() As Byte)
        For Each myClient As KeyValuePair(Of Short, tcpConnection) In Clients
            Clients(myClient.Key).Send(msgTag, byteData)
        Next
    End Sub
#End Region

#Region "Events handlers for the clients"
    Private Sub onClientDisconnect(ByVal Client As tcpConnection)
        RaiseEvent Disconnect(Client)
        'close client and remove from collection
        Clients.Remove(Client.Tag)
        Client = Nothing
    End Sub

    Private Sub onDataReceived(ByVal client As tcpConnection, ByVal msgTag As Byte, ByVal mstream As MemoryStream)
        RaiseEvent DataReceived(client, msgTag, mstream)
    End Sub

    Private Sub onMessageReceived(ByVal Client As tcpConnection, ByVal message As Byte)
        RaiseEvent MessageReceived(Client, message)
    End Sub

    Private Sub onStringReceived(ByVal Sender As tcpConnection, ByVal msgTag As Byte, ByVal message As String)
        RaiseEvent StringReceived(Sender, msgTag, message)
    End Sub
#End Region

End Class
