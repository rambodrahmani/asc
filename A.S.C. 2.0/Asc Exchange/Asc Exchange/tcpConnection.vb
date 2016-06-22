Imports System.Net.Sockets
Imports System.IO

Public Class tcpConnection
    Private Enum RequestTags As Byte
        Connect = 1
        Disconnect = 2
        DataTransfer = 3
        StringTransfer = 4
        ByteTransfer = 5
    End Enum

    Private PACKET_SIZE As Int16 = 4096
    Private client As TcpClient
    Private readByte(0) As Byte
    Private myTag As Byte
    Private isConnected As Boolean 'set to true when connection is granted by server

#Region "Declare events"
    Public Event DataReceived(ByVal sender As tcpConnection, ByVal msgTag As Byte, ByVal mstream As MemoryStream)
    Public Event MessageReceived(ByVal sender As tcpConnection, ByVal msgTag As Byte)
    Public Event StringReceived(ByVal Sender As tcpConnection, ByVal msgTag As Byte, ByVal message As String)
    Public Event TransferProgress(ByVal Sender As tcpConnection, ByVal msgTag As Byte, ByVal Percentage As Single)
    Public Event Connect(ByVal sender As tcpConnection)
    Public Event Disconnect(ByVal sender As tcpConnection)
#End Region

#Region "Overloaded Constructors and destructor"
    Public Sub New(ByVal client As TcpClient, ByVal PacketSize As Int16)
        'instantiates class for use in server
        'don't handle errors here, allow caller to handle them
        PACKET_SIZE = PacketSize
        Me.client = client
        'Start an asynchronous read from the NetworkStream
        Me.client.GetStream.BeginRead(readByte, 0, 1, AddressOf ReceiveOneByte, Nothing)
    End Sub

    Public Sub New(ByVal IPServer As String, ByVal Port As Integer)
        'instantiates class for use in client
        'don't handle errors here, allow caller to handle them
        'create a new TcpClient and make a synchronous connection 
        'attempt to the provided host name and port number
        Me.client = New TcpClient(IPServer, Port)
        'Start an asynchronous read from the NetworkStream
        Me.client.GetStream.BeginRead(readByte, 0, 1, AddressOf ReceiveOneByte, Nothing)
    End Sub

    Protected Overrides Sub finalize()
        'we don't care about errors in this section since we are ending
        On Error Resume Next
        If isConnected Then Send(RequestTags.Disconnect)
        client.GetStream.Close()
        client.Close()
        client = Nothing
    End Sub

#End Region

    Public Property Tag() As Byte
        Get
            Tag = myTag
        End Get
        Set(ByVal value As Byte)
            myTag = value
        End Set
    End Property

    Private Sub ReceiveOneByte(ByVal ar As IAsyncResult)
        'This is where the data is received.  All data communications 
        'begin with a one-byte identifier that tells the class how to
        'handle what is to follow.
        Dim r As BinaryReader
        Dim sreader As StreamReader
        Dim mStream As MemoryStream
        Dim lData As Int32
        Dim readBuffer(PACKET_SIZE) As Byte
        Dim StrBuffer(PACKET_SIZE) As Char
        Dim passThroughByte As Byte
        Dim nData As Integer
        Dim lenData As Integer

        SyncLock client.GetStream
            'if error occurs here then socket has closed
            Try
                client.GetStream.EndRead(ar)
            Catch
                RaiseEvent Disconnect(Me)
                Exit Sub
            End Try
        End SyncLock
        'this is the instruction on how to handle the rest of the data to come.
        Select Case readByte(0)
            Case RequestTags.Connect
                'sent from server to client informing client that a successful 
                'connection negotiation has been made and the connection now exists.
                isConnected = True 'identifies this class as on the client end
                RaiseEvent Connect(Me)
                SyncLock client.GetStream
                    r = New BinaryReader(client.GetStream)
                    PACKET_SIZE = r.ReadInt16
                    'Continue the asynchronous read from the NetworkStream
                    Me.client.GetStream.BeginRead(readByte, 0, 1, AddressOf ReceiveOneByte, Nothing)
                End SyncLock

            Case RequestTags.Disconnect
                'sent to either client or server
                'propagated forward to Listener on server end
                RaiseEvent Disconnect(Me)

            Case RequestTags.DataTransfer
                'a block of data is coming
                'Format of this transaction is
                '   DataTransfer identifier byte
                '   Pass-Through Byte contains user defined data
                '   Length of data block (max size = 2,147,483,647 bytes)
                SyncLock client.GetStream
                    r = New BinaryReader(client.GetStream)
                    'next we expect a pass-through byte
                    client.GetStream.Read(readByte, 0, 1)
                    passThroughByte = readByte(0)
                    'next expect length of data (Int32)
                    nData = r.ReadInt32
                    lenData = nData
                    'now comes the data, save it in a memory stream
                    mStream = New MemoryStream
                    While nData > 0
                        RaiseEvent TransferProgress(Me, passThroughByte, CSng(1.0 - nData / lenData))
                        lData = client.GetStream.Read(readBuffer, 0, PACKET_SIZE)
                        mStream.Write(readBuffer, 0, lData)
                        nData -= lData
                    End While
                    'Continue the asynchronous read from the NetworkStream
                    Me.client.GetStream.BeginRead(readByte, 0, 1, AddressOf ReceiveOneByte, Nothing)
                End SyncLock
                'once all data has arrived, pass it on to the end user as a stream
                RaiseEvent DataReceived(Me, passThroughByte, mStream)
                mStream.Dispose()

            Case RequestTags.StringTransfer
                'Here we receive the transfer of string data in a block
                'not to exceed PACKET_SIZE in length.
                SyncLock client.GetStream
                    sreader = New StreamReader(client.GetStream)
                    'next we expect a pass-through byte
                    client.GetStream.Read(readByte, 0, 1)
                    passThroughByte = readByte(0)
                    nData = sreader.Read(StrBuffer, 0, PACKET_SIZE)
                End SyncLock
                Dim strX As New String(StrBuffer, 0, nData)

                'pass string on to end user
                RaiseEvent StringReceived(Me, passThroughByte, strX)
                SyncLock client.GetStream
                    'Continue the asynchronous read from the NetworkStream
                    Me.client.GetStream.BeginRead(readByte, 0, 1, AddressOf ReceiveOneByte, Nothing)
                End SyncLock

            Case RequestTags.ByteTransfer
                'Here we receive a user-defined one-byte message
                SyncLock client.GetStream
                    client.GetStream.Read(readByte, 0, 1)
                End SyncLock
                'pass byte on to end user
                RaiseEvent MessageReceived(Me, readByte(0))
                SyncLock client.GetStream
                    'Continue the asynchronous read from the NetworkStream
                    Me.client.GetStream.BeginRead(readByte, 0, 1, AddressOf ReceiveOneByte, Nothing)
                End SyncLock

        End Select

    End Sub


#Region "Overloaded methods to send data between client(s) and server"

    Public Sub Send(ByVal msgTag As Byte)
        ' This subroutine sends a one-byte response
        SyncLock client.GetStream
            Dim writer As New BinaryWriter(client.GetStream)
            'notify that 1-byte is coming
            writer.Write(RequestTags.ByteTransfer)
            'Send user defined message byte
            writer.Write(msgTag)
            ' Make sure all data is sent now.
            writer.Flush()
        End SyncLock
    End Sub

    Public Sub Send(ByVal msgTag As Byte, ByVal strX As String)
        'sends a string of max length PACKET_SIZE
        SyncLock client.GetStream
            Dim writer As New StreamWriter(client.GetStream)
            'Notify other end that a string block is coming
            writer.Write(Chr(RequestTags.StringTransfer))
            'Send user defined message byte
            writer.Write(Chr(msgTag))
            'Send the string
            writer.Write(strX)
            'make sure all data gets sent now
            writer.Flush()
        End SyncLock
    End Sub

    Public Sub Send(ByVal msgTag As Byte, ByVal byteData() As Byte)
        'sends array of byte data
        SyncLock client.GetStream
            Dim writer As New BinaryWriter(client.GetStream)
            'Notify that byte data is coming
            writer.Write(RequestTags.DataTransfer)
            'send user-define message byte
            writer.Write(msgTag)
            'send length of data block
            writer.Write(byteData.Length)
            'send the data
            writer.Write(byteData)
            'make sure all data is sent
            writer.Flush()
        End SyncLock
    End Sub

    Public Sub SendFile(ByVal msgTag As Byte, ByVal FilePath As String)
        'max filesize is 2GB
        Dim byteArray() As Byte
        Dim fs As FileStream = New FileStream(FilePath, FileMode.Open, FileAccess.Read)
        Dim r As New BinaryReader(fs)

        SyncLock client.GetStream
            Dim w As New BinaryWriter(client.GetStream)
            'notify that file data is coming
            w.Write(RequestTags.DataTransfer)
            'send user-define message byte
            w.Write(msgTag)
            'send size of file
            w.Write(CInt(fs.Length))
            'Send the file data
            Do
                'read data from file
                byteArray = r.ReadBytes(PACKET_SIZE)
                'write data to Network Stream
                w.Write(byteArray)
            Loop While byteArray.Length = PACKET_SIZE
            'make sure all data is sent
            w.Flush()
            'w.Close()
        End SyncLock
        fs.Close()
    End Sub
#End Region

End Class
