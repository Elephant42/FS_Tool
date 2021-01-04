Option Explicit On
Option Strict On

#Region "Imports"
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Text
#End Region

Public Class AsyncSocket

#Region "Public Declarations"

    Public Const ASCII_EOT As Integer = 4

    Enum SocketStates
        IsConnecting = 0
        IsAccepting = 1
        IsSending = 2
        IsReceving = 3
        IsDisconnecting = 4
    End Enum

    Public ReadOnly Property IsConnected As Boolean
        Get

            If (mySocket Is Nothing) Then Return False

            Try
                Return mySocket.Connected
            Catch Exception As Exception
                Return False
            End Try

        End Get
    End Property

    Public ReadOnly Property NetworkEndPoint() As IPEndPoint
        Get

            If (mySocket Is Nothing) Then Return Nothing

            Try
                Return CType(mySocket.RemoteEndPoint, IPEndPoint)
            Catch Exception As Exception
                Return Nothing
            End Try

        End Get
    End Property

    Public ReadOnly Property NetworkAddress() As IPAddress
        Get

            If (mySocket Is Nothing) Then Return Nothing

            Try
                Return CType(mySocket.RemoteEndPoint, IPEndPoint).Address
            Catch Exception As Exception
                Return Nothing
            End Try

        End Get
    End Property

    Public ReadOnly Property NetworkPort() As Integer
        Get

            If (mySocket Is Nothing) Then Return Nothing

            Try
                Return CType(mySocket.RemoteEndPoint, IPEndPoint).Port
            Catch Exception As Exception
                Return Nothing
            End Try

        End Get
    End Property

    Public ReadOnly Property UID() As String
        Get
            Return _GUID.ToString
        End Get
    End Property

#Region "Events"

    Public Event OnListen()
    Public Event OnListenFailed(ByVal Exception As Exception)

    Public Event OnConnect()
    Public Event OnConnectFailed(ByVal Exception As Exception)

    Public Event OnSend()
    Public Event OnSendFailed(ByVal Exception As Exception)

    Public Event OnAccept(ByVal AcceptedSocket As AsyncSocket)
    Public Event OnAcceptFailed(ByVal Exception As Exception)

    Public Event OnReceive(ByVal Stream As String)
    Public Event OnReceiveFailed(ByVal Exception As Exception)

    Public Event OnDisconnect()
    Public Event OnDisconnectFailed(ByVal Exception As Exception)

#End Region

#End Region

#Region "Constructors"

    Public Sub New()
        mySocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
    End Sub

    Public Sub New(ByVal Socket As Socket, Optional ByVal ListenData As Boolean = False)
        Me.mySocket = Socket
        If ListenData Then Me.listenData()
    End Sub

    Public Sub New(ByVal Ip As String, ByVal Port As Integer)
        Connect(Ip, Port)
    End Sub

    Public Sub New(ByVal Port As Integer)
        mySocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        Listen(Port)
    End Sub

#End Region

#Region "Public Functions"

    Public Sub Connect(ByVal host As String, ByVal Port As Integer)

        If mySocket Is Nothing Then
            mySocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        End If

        Dim hostEntry As IPHostEntry = Dns.GetHostEntry(host)
        'you might get more than one ip for a hostname since 
        'DNS supports more than one record
        Dim useIP As IPAddress = Nothing
        For Each ip As IPAddress In hostEntry.AddressList
            If ip.AddressFamily = Sockets.AddressFamily.InterNetwork Then
                useIP = ip
                Exit For
            End If
        Next
        If useIP Is Nothing Then
            Throw New Exception("Invalid hostname: " & host)
        Else
            Connect(useIP, Port)
        End If

    End Sub
    Public Sub Connect(ByVal Ip As IPAddress, ByVal Port As Integer)

        If (mySocket Is Nothing) Then Exit Sub

        Try
            myStates.Add(SocketStates.IsConnecting)
            mySocket.BeginConnect(New IPEndPoint(Ip, Port), AddressOf connectCallBack, mySocket)
        Catch Exception As Exception
            myStates.Remove(SocketStates.IsConnecting)
            RaiseEvent OnConnectFailed(Exception)
        End Try

    End Sub

    Public Sub Disconnect()

        If (Not IsConnected) OrElse (mySocket Is Nothing) Then Exit Sub

        Try
            If stateIs(SocketStates.IsReceving) Then
                mySocket.Shutdown(SocketShutdown.Both)
                myStates.Remove(SocketStates.IsReceving)
            End If

            While Not myStates.Count = 0
                Thread.Sleep(1)
            End While

            myStates.Add(SocketStates.IsDisconnecting)
            mySocket.BeginDisconnect(False, AddressOf onDisconnectCallBack, Nothing)
        Catch Exception As Exception
            myStates.Remove(SocketStates.IsDisconnecting)
            RaiseEvent OnDisconnectFailed(Exception)
        End Try

    End Sub

    Public Sub Close()

        If mySocket Is Nothing Then Exit Sub

        If stateIs(SocketStates.IsConnecting) Or stateIs(SocketStates.IsReceving) Or stateIs(SocketStates.IsSending) Then
            Disconnect()
        End If

        mySocket.Close()
        mySocket = Nothing

    End Sub

    Public Sub Listen(ByVal Port As Integer)

        If (mySocket Is Nothing) Then Exit Sub

        Try
            mySocket.Bind(New IPEndPoint(IPAddress.Parse("0.0.0.0"), Port))
            mySocket.Listen(0)
            listenSocket()

            RaiseEvent OnListen()
        Catch Exception As Exception
            RaiseEvent OnListenFailed(Exception)
        End Try

    End Sub

    Public Sub ListenEvent()
        RaiseEvent OnListen()
    End Sub

    Public Sub Send(ByVal Stream As String)

        If (Not IsConnected) OrElse (mySocket Is Nothing) Then Exit Sub

        Try
            Dim t As String = Stream.Length.ToString(sTREAM_LEN_FORMAT) & Stream & ChrW(ASCII_EOT)
            Dim Buffer() As Byte = Encoding.Default.GetBytes(t)

            Me.myStates.Add(SocketStates.IsSending)
            Me.mySocket.BeginSend(Buffer, 0, Buffer.Length, SocketFlags.None, AddressOf onSendCallBack, Nothing)
        Catch Exception As Exception
            Me.myStates.Remove(SocketStates.IsSending)
            RaiseEvent OnSendFailed(Exception)
        End Try

    End Sub

#End Region

#Region "Private Functions"

    Private Sub listenSocket()

        If (mySocket Is Nothing) Then Exit Sub

        Try

            myStates.Add(SocketStates.IsAccepting)
            mySocket.BeginAccept(AddressOf onAcceptCallBack, Nothing)

        Catch Exception As Exception
            myStates.Remove(SocketStates.IsAccepting)
            RaiseEvent OnAcceptFailed(Exception)
        End Try

    End Sub

    Private Sub listenData()

        If (Not IsConnected) OrElse (mySocket Is Nothing) Then Exit Sub

        Try

            myStates.Add(SocketStates.IsReceving)
            mySocket.BeginReceive(myBuffer, 0, myBuffer.Length, SocketFlags.None, AddressOf onReceiveCallBack, Nothing)

        Catch Exception As Exception
            myStates.Remove(SocketStates.IsReceving)
            RaiseEvent OnReceiveFailed(Exception)
        End Try

    End Sub

    Private Function stateIs(ByVal SocketState As SocketStates) As Boolean
        If myStates.Contains(SocketState) Then Return True
        Return False
    End Function

#End Region

#Region "Callbacks"

    Private Sub onReceiveCallBack(ByVal IAsyncResult As IAsyncResult)

        If (Not IsConnected) OrElse (mySocket Is Nothing) OrElse (IAsyncResult Is Nothing) OrElse (Not stateIs(SocketStates.IsReceving)) Then Exit Sub

        Try

            Dim Bytes As Integer = mySocket.EndReceive(IAsyncResult)
            myStates.Remove(SocketStates.IsReceving)

            If Bytes > 0 Then
                Dim Stream As String = Encoding.Default.GetString(myBuffer)
                Array.Clear(myBuffer, 0, myBuffer.Length - 1)

                If Stream.Contains(ChrW(ASCII_EOT)) Then
                    myStrBuf &= Stream.Replace(ChrW(ASCII_EOT), "")
                    Dim rcLen As Integer = CInt(myStrBuf.Substring(0, sTREAM_LEN_FLD_SIZE))
                    Dim len As Integer = myStrBuf.Length - sTREAM_LEN_FLD_SIZE
                    Dim t As String
                    If len < rcLen Then
                        t = String.Copy(myStrBuf.Substring(sTREAM_LEN_FLD_SIZE, len))
                    Else
                        t = String.Copy(myStrBuf.Substring(sTREAM_LEN_FLD_SIZE, rcLen))
                    End If
                    myStrBuf = ""
                    RaiseEvent OnReceive(t)
                Else
                    myStrBuf &= Stream
                End If
            Else
                Dim Socket As Socket = CType(IAsyncResult.AsyncState, Socket)
                Dim AsyncSocket As New AsyncSocket(Socket)
                AsyncSocket.Disconnect()
            End If

            listenData()
        Catch Exception As Exception
            myStates.Remove(SocketStates.IsReceving)
            RaiseEvent OnReceiveFailed(Exception)
        End Try

    End Sub

    Private Sub connectCallBack(ByVal IAsyncResult As IAsyncResult)

        If (mySocket Is Nothing) OrElse (IAsyncResult Is Nothing) OrElse (Not stateIs(SocketStates.IsConnecting)) Then Exit Sub

        Try

            mySocket.EndConnect(IAsyncResult)
            myStates.Remove(SocketStates.IsConnecting)
            listenData()

            RaiseEvent OnConnect()

        Catch Exception As Exception
            Me.myStates.Remove(SocketStates.IsConnecting)
            RaiseEvent OnConnectFailed(Exception)
        End Try

    End Sub

    Private Sub onSendCallBack(ByVal IAsyncResult As IAsyncResult)

        If (Not IsConnected) OrElse (mySocket Is Nothing) OrElse (IAsyncResult Is Nothing) OrElse (Not stateIs(SocketStates.IsSending)) Then Exit Sub

        Try

            mySocket.EndSend(IAsyncResult)
            myStates.Remove(SocketStates.IsSending)

            RaiseEvent OnSend()

        Catch Exception As Exception
            myStates.Remove(SocketStates.IsSending)
            RaiseEvent OnSendFailed(Exception)
        End Try

    End Sub

    Private Sub onDisconnectCallBack(ByVal IAsyncResult As IAsyncResult)

        If (mySocket Is Nothing) OrElse (IAsyncResult Is Nothing) OrElse (Not stateIs(SocketStates.IsDisconnecting)) Then Exit Sub

        Try

            mySocket.EndDisconnect(IAsyncResult)
            myStates.Remove(SocketStates.IsDisconnecting)

            RaiseEvent OnDisconnect()

        Catch Exception As Exception
            myStates.Remove(SocketStates.IsDisconnecting)
            RaiseEvent OnDisconnectFailed(Exception)
        End Try

    End Sub

    Private Sub onAcceptCallBack(ByVal IAsyncResult As IAsyncResult)

        If (mySocket Is Nothing) OrElse (IAsyncResult Is Nothing) OrElse (Not stateIs(SocketStates.IsAccepting)) Then Exit Sub

        Try

            Dim AcceptedSocket As New AsyncSocket(mySocket.EndAccept(IAsyncResult), True)
            myStates.Remove(SocketStates.IsAccepting)
            listenSocket()

            RaiseEvent OnAccept(AcceptedSocket)

        Catch Exception As Exception
            myStates.Remove(SocketStates.IsAccepting)
            RaiseEvent OnAcceptFailed(Exception)
        End Try

    End Sub

#End Region

#Region "Private Declarations"

    Private Const sTREAM_LEN_FORMAT As String = "0000000000"
    Private Const sTREAM_LEN_FLD_SIZE As Integer = 10

    Private mySocket As Socket
    Private myBuffer(1024) As Byte
    Private myStates As New List(Of SocketStates)
    Private _GUID As Guid = Guid.NewGuid()
    Private myStrBuf As String = ""

#End Region

End Class

