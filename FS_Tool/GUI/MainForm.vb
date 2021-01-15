Option Explicit On
Option Strict On

Imports System.IO
Imports SimConLib

Public Class MainForm

#Region "Server Stuff"

    Private Const dEBUG_MODE As Boolean = False

    Private Class acceptedSocket

        Public ReadOnly Property UID As String
            Get
                Return mySocket.UID
            End Get
        End Property

        Public Sub Send(ByVal data As String)
            mySocket.Send(data)
        End Sub

        Public Sub Disconnect()
            mySocket.Disconnect()
        End Sub

        Public Sub Close()
            mySocket.Close()
        End Sub

        Public Sub New(ByVal pSocket As AsyncSocket, ByVal dataRcvCallback As dataReceivedDelegate, ByVal sockClosed As socketClosedDelegate)

            mySocket = pSocket
            AddHandler mySocket.OnReceive, AddressOf handleOnRcv
            AddHandler mySocket.OnDisconnect, AddressOf handleDisconnect
            AddHandler mySocket.OnReceiveFailed, AddressOf handleReceiveFailed
            AddHandler mySocket.OnSendFailed, AddressOf handleSendFailed
            AddHandler mySocket.OnDisconnectFailed, AddressOf handleDisconnectFailed

            dataRcvedCallback = dataRcvCallback
            sockClosedCallback = sockClosed

        End Sub

        Private Sub handleOnRcv(ByVal data As String)

            Dim tMsg As New SimMessage(mySocket, data)
            If tMsg.IsValid Then
                dataRcvedCallback(Me, tMsg)
                myLastSimMsg = tMsg
            Else
                dataRcvedCallback(Me, data)
            End If

        End Sub

        Private Sub handleDisconnect()
            sockClosedCallback(Me, Nothing)
        End Sub

        Private Sub handleSendFailed(ByVal ex As Exception)

            failEx = ex
            mySocket.Disconnect()
            sockClosedCallback(Me, ex)

        End Sub

        Private Sub handleReceiveFailed(ByVal ex As Exception)

            failEx = ex
            mySocket.Disconnect()

        End Sub

        Private Sub handleDisconnectFailed(ByVal ex As Exception)

            failEx = ex
            sockClosedCallback(Me, ex)

        End Sub

        Private mySocket As AsyncSocket = Nothing
        Private dataRcvedCallback As dataReceivedDelegate = Nothing
        Private sockClosedCallback As socketClosedDelegate = Nothing
        Private failEx As Exception = Nothing
        Private myLastSimMsg As SimMessage = Nothing

    End Class

    Private Delegate Sub dataReceivedDelegate(ByVal sock As acceptedSocket, ByVal data As Object)
    Private Delegate Sub socketClosedDelegate(ByVal sock As acceptedSocket, ByVal ex As Exception)

    Private Sub handleOnAccept(ByVal acceptedSock As AsyncSocket)

        Dim sock As New acceptedSocket(acceptedSock, AddressOf dataReceived, AddressOf socketClosed)
        mySockets.Add(sock.UID, sock)
        doStatus("Client Connected: " & Now & ", " & sock.UID)

    End Sub

    Private Sub handleOnDisconnect()
        doStatus("Server Stopped: " & Now)
    End Sub

    Private Sub handleOnDisconnectEx(ByVal ex As Exception)
        doStatus("Server Stop Exception: " & ex.Message)
    End Sub

    Private Sub dataReceived(ByVal sock As acceptedSocket, ByVal data As Object)

        Dim strTemp As String = ""
        Dim reply As String = ""
        Dim simMsg As SimMessage = Nothing

        If TypeOf data Is String Then
            strTemp = CType(data, String)
            reply = "<RCV SockUID=""" & sock.UID & """ Time=""" & Now & """ />"
        ElseIf TypeOf data Is SimMessage Then
            simMsg = CType(data, SimMessage)
            strTemp = "SimMessage: " & simMsg.UID & " from " & sock.UID
            reply = "<RCV SockUID=""" & sock.UID & """ MsgUID=""" & simMsg.UID & """ Time=""" & Now & """ />"
        Else
            Throw New Exception("Server received unknown data")
        End If

        If dEBUG_MODE Then
            doStatus("Rcv: " & strTemp)
        End If
        sock.Send(reply)

        If simMsg IsNot Nothing Then
            If simMsg.Command = SimMessage.CommandEnum.SendSimEvent Then
                executeSimEventMsg(simMsg)
            ElseIf simMsg.Command = SimMessage.CommandEnum.RegisterSimVars Then
                executeRegSimVarsMsg(simMsg)
            ElseIf simMsg.Command = SimMessage.CommandEnum.GetSimVals Then
                executeGetSimValsMsg(simMsg)
            Else
                simMsg.Execute()
            End If
        End If

    End Sub

    Private Sub executeSimEventMsg(ByVal simMsg As SimMessage)

        Dim t As SimData.SimEventEnum = SimConnectLib.GetSimEventByName(simMsg.SimEvent.SimEventCode)
        If t <> SimData.SimEventEnum.NONE Then
            If sim Is Nothing Then
                sendErrorMessage(simMsg, "SimConnect is not active")
            Else
                If sim.Connected Then
                    sim.TransmitClientEvent(t, simMsg.SimEvent.SimEventData)
                Else
                    sendErrorMessage(simMsg, "FlightSim is not running")
                End If
            End If
        End If

    End Sub

    Private Sub executeRegSimVarsMsg(ByVal simMsg As SimMessage)

        If Me.InvokeRequired Then
            Me.Invoke(Sub() executeRegSimVarsMsg(simMsg))
            Return
        End If

        myRegisteredSimVars = New List(Of String)

        For Each key As String In simMsg.SimVars.Keys
            myRegisteredSimVars.Add(key)
        Next

        If sim IsNot Nothing Then
            'sim.RegisterSimVars(myRegisteredSimVars)
            registerSimVars()
        End If

        Application.DoEvents()

    End Sub

    Private Sub executeGetSimValsMsg(ByVal simMsg As SimMessage)

        Try
            Dim t As New List(Of String)
            For Each key As String In simMsg.SimVars.Keys
                t.Add(key)
            Next

            For Each key As String In t
                Dim val As Double = 0.0
                If sim IsNot Nothing Then val = sim.GetSimValDouble(key)
                'simMsg.SimVars(key) = sim.GetSimVal(key)
                simMsg.SetSimVar(key, val)
            Next

            simMsg.Send()
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
        End Try

    End Sub

    Private Sub socketClosed(ByVal sock As acceptedSocket, ByVal ex As Exception)

        If ex Is Nothing Then
            doStatus("Socket Closed: " & Now)
        Else
            doStatus("Socket Close Exception: " & ex.Message)
        End If
        mySockets.Remove(sock.UID)
        sock = Nothing

    End Sub

    Private Sub startServer()

        myServer = New AsyncSocket()

        AddHandler myServer.OnAccept, AddressOf handleOnAccept
        AddHandler myServer.OnDisconnect, AddressOf handleOnDisconnect
        AddHandler myServer.OnDisconnectFailed, AddressOf handleOnDisconnectEx

        myServer.Listen(getPort())

        doStatus("Server Started: " & Now)

    End Sub

    Private Sub stopServer()

        myServer.Close()
        myServer = Nothing

        For Each sock As acceptedSocket In mySockets.Values
            sock.Close()
        Next

        mySockets.Clear()

        doStatus("Server Stopped: " & Now)

    End Sub

    Private Sub doEnableServer(Optional ByVal startupFlag As Boolean = False)

        If startupFlag Then
            startServer()
            Me.ctx1Item_EnableServer.Checked = True
        Else
            If Me.ctx1Item_EnableServer.Checked Then
                stopServer()
                Me.ctx1Item_EnableServer.Checked = False
            Else
                startServer()
                Me.ctx1Item_EnableServer.Checked = True
            End If
        End If

    End Sub

    Private Shared Sub sendErrorMessage(ByVal simmsg As SimMessage, ByVal msg As String)
        sendErrorMessage(simmsg.Socket, msg)
    End Sub
    Private Shared Sub sendErrorMessage(ByVal sock As AsyncSocket, ByVal msg As String)

        Dim t As New SimMessage(sock, SimMessage.CommandEnum.ErrorMessage, msg)
        t.Send()

    End Sub

#End Region

#Region "Private Declarations"

    Private Const HIDDEN_X As Integer = -16000
    Private Const WINDOW_LABEL As String = "Flight Sim Tool"

    Private Const sIM_NOT_CONNECT_MSG As String = "Sim not connected." & vbNewLine

    Private ReadOnly Property isHidden As Boolean
        Get
            Return (Me.Left < -10000)
        End Get
    End Property

    Private ReadOnly Property isQuietStart As Boolean
        Get
            For Each arg As String In My.Application.CommandLineArgs
                If arg.ToUpper.StartsWith("/Q") Then
                    Return True
                ElseIf arg.ToUpper.StartsWith("-Q") Then
                    Return True
                End If
            Next
            Return False
        End Get
    End Property

    Private bFormLoading As Boolean = False
    Private bLastConnected As Boolean = False
    Private lastLeft As Integer = 0
    Private lastSimRate As Integer = 0
    Private sim As SimConnectLib = Nothing
    Private lastCmd As String = ""
    Private lastCmdRcvd As Date = Date.MinValue
    Private cmdRcvd As Date = Date.MinValue

    Private mySockets As New Dictionary(Of String, acceptedSocket)
    Private myServer As AsyncSocket = Nothing
    Private myStatus As New CircularBuffer(3)
    Private myRegisteredSimVars As New List(Of String)
    Private simEnabled As Boolean = False
    Private lastHeartbeat As Date = Date.MinValue
    Private myJoyMaps As JoystickMappings = Nothing
    Private joyMapEvtTimes As Dictionary(Of String, Date) = Nothing
    'Private currentJoyProfile As String = ""
    Private g1000Form As Form = Nothing

#If DEBUG Then
    Private Class recordData

        Private Structure dataStruct
            Public Key As String
            Public Title As String
            Public NumPos As Double
            Public Index As Double
            Public Percent As Double
        End Structure

        Private Const dATA_FILE As String = "D:\FS2020\Temp\SimConnectRecording2.txt"

        Public Sub AddRecord(ByVal title As String, ByVal numPos As Double, ByVal index As Double, ByVal percent As Double)

            Dim key As String = title & "|" & numPos & "|" & index & "|" & percent
            'loadData()

            If Not dataList.ContainsKey(key) Then
                Dim t As dataStruct
                t.Key = key
                t.Title = title
                t.NumPos = numPos
                t.Index = index
                't.Percent = AircraftDataClass.RadiansToDegrees(percent)
                t.Percent = 0.0
                dataList.Add(key, t)

                saveData()
            End If

        End Sub

        Private Shared Sub loadData()

            'dataList.Clear()

            'If File.Exists(dATA_FILE) Then
            '    Dim sr As New StreamReader(dATA_FILE)
            '    Dim line As String
            '    Do While Not sr.EndOfStream
            '        line = sr.ReadLine()
            '        Dim t As dataStruct

            '        Dim tl As String = line.Replace("{", "")
            '        tl = tl.Replace("},", "")
            '        tl = tl.Replace("""", "")

            '        Dim fields() As String = tl.Split(","c)
            '        Dim values() As String = fields(0).Split("|"c)

            '        t.Key = fields(0)
            '        t.F1 = values(0)
            '        t.F2 = values(1)
            '        t.F3 = values(2)
            '        dataList.Add(fields(0), t)
            '    Loop
            '    sr.Close()
            '    sr.Dispose()
            '    sr = Nothing
            'End If

        End Sub

        Private Sub saveData()

            Dim tl As New Dictionary(Of String, Dictionary(Of Double, Integer))

            Dim tk = dataList.Keys.ToList
            tk.Sort()

            For Each key In tk
                Dim t = dataList(key)
                If Not tl.ContainsKey(t.Title) Then tl.Add(t.Title, New Dictionary(Of Double, Integer))
                Dim t2 = tl(t.Title)
                If Not t2.ContainsKey(t.Percent) Then t2.Add(t.Percent, CInt(t.Index))
            Next

            Dim sw As New StreamWriter(dATA_FILE)
            For Each key In tl.Keys
                Dim t = tl(key)
                Dim str As String = " {""" & key & """, New Dictionary(Of Double, Integer) From {"
                Dim s = ""
                For Each key2 In t.Keys
                    Dim st = "{" & key2 & "," & t(key2) & "}"
                    If s = "" Then
                        s = st
                    Else
                        s &= "," & st
                    End If
                Next
                str = str & s & "}},"
                sw.WriteLine(str)
            Next
            sw.Close()
            sw.Dispose()
            sw = Nothing
            '{"2", New Dictionary(Of Double, Integer) From {{0.0, 0}, {0.0, 0}, {0.0, 0}}},

        End Sub

        Private dataList As New Dictionary(Of String, dataStruct)

    End Class

    Private Sub doRecord()

        Static data As New recordData

        'Dim title = sim.AirCraftData.Title
        'Dim flapsNumPos = sim.AirCraftData.FlapsNumHandlePos
        Dim title = "Title"
        Dim flapsNumPos = 0
        Dim flapsIndex = sim.GetSimValDouble(sim.SimVarName(SimData.SimVarEnum.FLAPS_HANDLE_INDEX))
        Dim flapsPercent = sim.GetSimValDouble(sim.SimVarName(SimData.SimVarEnum.TRAILING_EDGE_FLAPS_LEFT_ANGLE))

        data.AddRecord(title, flapsNumPos, flapsIndex, flapsPercent)

    End Sub

#End If

#End Region

#Region "Private Functions"

    Private Sub openG1000PFD()

        'If Me.ctx1Item_G1000PFD.Checked Then
        '    If g1000Form IsNot Nothing Then
        '        g1000Form.Close()
        '    End If
        '    g1000Form = Nothing
        '    Me.ctx1Item_G1000PFD.Checked = False
        'Else
        '    If g1000Form IsNot Nothing Then Throw New Exception("G-1000 PFD is already open")
        '    g1000Form = New G1000PFD_Form
        '    g1000Form.Show()
        '    Me.ctx1Item_G1000PFD.Checked = True
        'End If

    End Sub

    Private Sub doEnableJoyMaps(ByVal profileName As String, Optional ByVal startupFlag As Boolean = False)

        If startupFlag Then
            joyMapsEnable(profileName)
        Else
            If myJoyMaps.CurrentlyOpenProfile = "" Then
                'no current profile so fire this one up
                joyMapsEnable(profileName)
            Else
                If profileName = myJoyMaps.CurrentlyOpenProfile Then
                    'It's the current profile being toggled off
                    joyMapsDisable()
                Else
                    'It's a new profile so toggle off the old one and fire up this one
                    joyMapsDisable()
                    joyMapsEnable(profileName)
                End If
            End If
            'If Me.ctx1Item_EnableJoyMaps.Checked Then
            '    joyMapsDisable()
            'Else
            '    joyMapsEnable(profileName)
            'End If
        End If

    End Sub

    Private Sub joyMapsEnable(ByVal profileName As String)

        Debug.WriteLine("Enabling joymaps")
        myJoyMaps.Open(profileName)
        joyMapEvtTimes = New Dictionary(Of String, Date)

        Me.ctx1Item_EnableJoyMaps.Checked = True
        For Each tsmi As ToolStripMenuItem In Me.ctx1Item_EnableJoyMaps.DropDownItems
            If tsmi.Text = profileName Then
                tsmi.Checked = True
            Else
                tsmi.Checked = False
                tsmi.Enabled = False
            End If
        Next

        My.Settings.JoyProfile = profileName
        My.Settings.Save()

    End Sub

    Private Sub joyMapsDisable()

        Debug.WriteLine("Disabling joymaps")
        myJoyMaps.Close()
        joyMapEvtTimes = Nothing

        Me.ctx1Item_EnableJoyMaps.Checked = False
        For Each tsmi As ToolStripMenuItem In Me.ctx1Item_EnableJoyMaps.DropDownItems
            tsmi.Checked = False
            tsmi.Enabled = True
        Next

    End Sub

    Private Sub sendSingleSimEvent(ByVal simEvent As String)

        If simEvent.StartsWith("KEY.") Then
            Try
                Dim t = simEvent.Replace("KEY.", "")
                Dim k As String() = t.Split("."c)
                Dim km = GetKeyModifierByName(k(0))
                Dim sc = GetScanCodeByName(k(1))
                Debug.WriteLine(vbTab & vbTab & Utility.GetTicks() & " Sending Keystroke " & simEvent)
                SimConLib.SendScanCodeStroke(sc, km, KeyStrokeTypeEnum.Full)
            Catch ex As Exception
                Throw New Exception("Invalid KeyStroke: " & simEvent)
            End Try
        Else
            Dim simEv = simEvent
            Dim simDat As String = "0"

            If simEvent.Contains(":") Then
                Dim tokens = simEvent.Split(":"c)
                If tokens.Count = 2 Then
                    simEv = tokens(0)
                    simDat = tokens(1)
                End If
            End If

            Debug.WriteLine(vbTab & vbTab & Utility.GetTicks() & " Sending SimEvent " & simEvent & " - " & simEv & ": " & simDat)
            Dim t As SimData.SimEventEnum = SimConnectLib.GetSimEventByName(simEv)
            If t <> SimData.SimEventEnum.NONE Then
                If sim IsNot Nothing Then
                    If sim.Connected Then
                        sim.TransmitClientEvent(t, simDat)
                        Debug.WriteLine(vbTab & vbTab & Utility.GetTicks() & " SimEvent sent")
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub sendCountSimEvents(ByVal simEvent As String)

        If simEvent.Contains("#") Then
            Dim tokens = simEvent.Split("#"c)
            If tokens.Count = 2 Then
                Dim evtCount As Integer = CInt(tokens(0))
                For i = 1 To evtCount
                    sendSingleSimEvent(tokens(1))
                Next
            End If
        Else
            sendSingleSimEvent(simEvent)
        End If

    End Sub

    Private Sub sendSimEvent(ByVal simEvent As String)

        If simEvent.Contains(",") Then
            Dim tokens = simEvent.Split(","c)
            For Each token In tokens
                sendCountSimEvents(token)
            Next
        Else
            sendCountSimEvents(simEvent)
        End If

    End Sub

    Private Function getSimValBool(ByVal simVar As String) As Boolean

        If sim Is Nothing Then
            Return False
        Else
            Dim b = sim.GetSimValBool(simVar)
            Threading.Thread.Sleep(100)
            b = sim.GetSimValBool(simVar)
            Return b
        End If

    End Function

    Private Sub joyEventReceived(ByVal dat As JoystickData)

        If Me.InvokeRequired Then
            Me.Invoke(Sub() joyEventReceived(dat))
            Return
        End If

        If dat IsNot Nothing Then
            Dim t As String = ""
            Dim maps = dat.Device.Mappings
            For Each d In dat.Values
                If Not d.IsAxis Then
                    'Debug.WriteLine("Joystick event received from " & dat.JoystickName)
                    't &= vbTab & d.Index & "|" & d.Name.PadRight(40) & ": " & d.PreviousValue & "->" & d.Value & vbNewLine

                    Dim md = maps.GetMapData(d.Name)
                    If md.SimEvent <> "" Then
                        If md.SimVar = "" Then
                            If md.Acceleration > 0 Then
                                If Not joyMapEvtTimes.ContainsKey(md.GUID) Then joyMapEvtTimes.Add(md.GUID, Date.MinValue)
                                If Now < joyMapEvtTimes(md.GUID).AddMilliseconds(md.Delay) Then
                                    For i = 1 To md.Acceleration
                                        sendSimEvent(md.SimEvent)
                                    Next
                                Else
                                    sendSimEvent(md.SimEvent)
                                End If
                                joyMapEvtTimes(md.GUID) = Now
                            Else
                                If md.LongPushActive Then
                                    sendSimEvent(md.LongPushSimEvent)
                                ElseIf md.ReleaseActive Then
                                    sendSimEvent(md.ReleaseSimEvent)
                                Else
                                    sendSimEvent(md.SimEvent)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            'If t <> "" Then Debug.Write(t)
        End If

        Application.DoEvents()

    End Sub

    Private Sub pollJoysticks()

        If joyMapEvtTimes Is Nothing Then Exit Sub

        For Each mds In myJoyMaps.ActiveMaps
            Dim v = getSimValBool(mds.SimVar)
            'Debug.WriteLine("JoyMap: " & mds.GUID & ", " & mds.SimVar & ", " & v & ", " & mds.ActiveState)
            If mds.ReleaseSimEvent <> "" Then
                'The gist of this block of code is such that the state of the in-sim control is read from 
                'the sim var and if the current state of the joystick control does NOT match the sim then
                'the appropriate event is sent.  If the sim and joystick match then nothing is sent.
                If mds.ReleaseActive Then
                    'If the joystick switch is OFF and the SimVar is TRUE then send the ReleaseSimEvent
                    If v Then
                        If Not joyMapEvtTimes.ContainsKey(mds.GUID) Then joyMapEvtTimes.Add(mds.GUID, Date.MinValue)
                        If Now > joyMapEvtTimes(mds.GUID).AddMilliseconds(1000) Then
                            sendSimEvent(mds.ReleaseSimEvent)
                            joyMapEvtTimes(mds.GUID) = Now
                        End If
                    End If
                Else
                    'If the joystick switch is ON and the SimVar is FALSE then send the SimEvent
                    If Not v Then
                        If Not joyMapEvtTimes.ContainsKey(mds.GUID) Then joyMapEvtTimes.Add(mds.GUID, Date.MinValue)
                        If Now > joyMapEvtTimes(mds.GUID).AddMilliseconds(1000) Then
                            sendSimEvent(mds.SimEvent)
                            joyMapEvtTimes(mds.GUID) = Now
                        End If
                    End If
                End If
            Else
                If v = (Not mds.ActiveState) Then
                    If Not joyMapEvtTimes.ContainsKey(mds.GUID) Then joyMapEvtTimes.Add(mds.GUID, Date.MinValue)
                    If Now > joyMapEvtTimes(mds.GUID).AddMilliseconds(1000) Then
                        'Debug.WriteLine("JoyMap sending SimEvent: " & mds.SimEvent)
                        sendSimEvent(mds.SimEvent)
                        joyMapEvtTimes(mds.GUID) = Now
                    End If
                End If
            End If
        Next

    End Sub

    Private Sub doEnableSimCon(Optional ByVal startupFlag As Boolean = False)

        If startupFlag Then
            simEnable()
            Me.ctx1Item_EnableSimCon.Checked = True
        Else
            If Me.ctx1Item_EnableSimCon.Checked Then
                simDisable()
                Me.ctx1Item_EnableSimCon.Checked = False
            Else
                simEnable()
                Me.ctx1Item_EnableSimCon.Checked = True
            End If
        End If

    End Sub

    Private Sub simEnable()

        If sim Is Nothing Then
            sim = New SimConnectLib()
            sim.UpdateInterval_ms = My.Settings.SimConnectPollInterval_ms
            registerSimVars()
        End If
        sim.Enable = True
        simEnabled = True
        Me.HeartbeatLabel1.Visible = True
        initTimer()

    End Sub

    Private Sub simDisable()

        If sim IsNot Nothing Then
            sim.Enable = False
            'sim = Nothing
        End If
        simEnabled = False
        Me.HeartbeatLabel1.Visible = False
        tmrUpdateTimer.Stop()

    End Sub

    Private Sub doSimConDebug()

        SimConnectDebugForm.Sim = sim
        SimConnectDebugForm.Show()

    End Sub

    Private Sub doDisplay(ByVal strParam As String)

        If Me.InvokeRequired Then
            Me.Invoke(Sub() doDisplay(strParam))
            Return
        End If

        Dim t As String = ""
        If strParam <> "" Then
            t = strParam.Trim().Replace(vbNullChar, "") & vbNewLine
            't = strParam.Trim().Replace(vbNullChar, "")
        End If

        Dim tS As String = myStatus.ToString
        If tS <> "" Then
            t &= tS
        End If

        'Me.txtDisplay.AppendText(t)
        Me.txtDisplay.Text = t

        Application.DoEvents()

    End Sub

    Private Sub doStatus(ByVal strParam As String)

        If Me.InvokeRequired Then
            Me.Invoke(Sub() doStatus(strParam))
            Return
        End If

        myStatus.Append(strParam)

        If sim Is Nothing Then
            Me.txtDisplay.Text = sIM_NOT_CONNECT_MSG & vbNewLine & myStatus.ToString
        End If

        Application.DoEvents()

    End Sub

    Private Sub doDebug()

    End Sub

    Private Sub doClip(ByVal bExecuteImmediately As Boolean)

        If bExecuteImmediately Then
            If sim Is Nothing Then Exit Sub
        End If

        Try
            Dim t As String = My.Computer.Clipboard.GetText()
            sim.DoSlew(t, bExecuteImmediately, Me.txtNewLat.Text, Me.txtNewLong.Text)
        Catch ex As Exception
            Me.txtNewLat.Text = ""
            Me.txtNewLong.Text = ""
            validateSlew()
            If Not bExecuteImmediately Then
                MessageBox.Show(Me, "Invalid coordinates format!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Try

    End Sub

    Private Function checkValidSlew() As Boolean

        Try
            Return sim.CheckValidSlew(Me.txtNewLat.Text, Me.txtNewLong.Text)
        Catch ex As Exception
            Return False
        End Try

    End Function

    Private Sub validateSlew()

        If sim Is Nothing Then
            Me.btnSlew.Enabled = False
        Else
            If sim.Connected Then
                Me.btnSlew.Enabled = checkValidSlew()
            Else
                Me.btnSlew.Enabled = False
            End If
        End If

    End Sub

    Private Sub doSlew()
        sim.DoSlew(CDbl(Me.txtNewLat.Text), CDbl(Me.txtNewLong.Text))
    End Sub

    Private Sub doSimRate(ByVal newSimRate As Integer)
        sim.DoSimRate(newSimRate)
    End Sub

    Private Sub doActivate()

        If Me.cbSimRate.Text = "" Then Exit Sub

        doSimRate(CInt(Me.cbSimRate.Text))

    End Sub

    Private Sub doSimRateChanged()

        If sim Is Nothing Then
            Me.btnActivateSimRate.Enabled = False
        Else
            If Not sim.Connected Then
                Me.btnActivateSimRate.Enabled = False
                Exit Sub
            End If

            If Me.cbSimRate.Text = "" Then
                Me.btnActivateSimRate.Enabled = False
            Else
                Me.btnActivateSimRate.Enabled = True
            End If
        End If

    End Sub

    Private Sub setSimRateMenuTicks()

        If Not sim.Connected Then
            If lastSimRate = -1 Then Exit Sub
            lastSimRate = -1
            Me.ctx1Item_Rate_x1.Checked = False
            Me.ctx1Item_Rate_x2.Checked = False
            Me.ctx1Item_Rate_x4.Checked = False
            Me.ctx1Item_Rate_x8.Checked = False
        Else
            Dim simRate As Integer = sim.CurrentSimRate
            If simRate <> lastSimRate Then
                lastSimRate = simRate
                If simRate = 8 Then
                    Me.ctx1Item_Rate_x1.Checked = False
                    Me.ctx1Item_Rate_x2.Checked = False
                    Me.ctx1Item_Rate_x4.Checked = False
                    Me.ctx1Item_Rate_x8.Checked = True
                ElseIf simRate = 4 Then
                    Me.ctx1Item_Rate_x1.Checked = False
                    Me.ctx1Item_Rate_x2.Checked = False
                    Me.ctx1Item_Rate_x4.Checked = True
                    Me.ctx1Item_Rate_x8.Checked = False
                ElseIf simRate = 2 Then
                    Me.ctx1Item_Rate_x1.Checked = False
                    Me.ctx1Item_Rate_x2.Checked = True
                    Me.ctx1Item_Rate_x4.Checked = False
                    Me.ctx1Item_Rate_x8.Checked = False
                ElseIf simRate = 1 Then
                    Me.ctx1Item_Rate_x1.Checked = True
                    Me.ctx1Item_Rate_x2.Checked = False
                    Me.ctx1Item_Rate_x4.Checked = False
                    Me.ctx1Item_Rate_x8.Checked = False
                Else
                    'This should NEVER happen
                    Me.ctx1Item_Rate_x1.Checked = False
                    Me.ctx1Item_Rate_x2.Checked = False
                    Me.ctx1Item_Rate_x4.Checked = False
                    Me.ctx1Item_Rate_x8.Checked = False
                End If
                Me.NotifyIcon1.Text = WINDOW_LABEL & " - SimRate: " & sim.CurrentSimRate
            End If
        End If

    End Sub

    Private Sub initTimer()

        tmrUpdateTimer.Interval = My.Settings.SimConnectPollInterval_ms
        tmrUpdateTimer.Enabled = True
        tmrUpdateTimer.Start()

    End Sub

    Private Sub displaySimVals()

        Dim str As String = ""

        str &= "LATITUDE: " & sim.GetSimValDouble(SimData.LATITUDE) & vbNewLine
        str &= "LONGITUDE: " & sim.GetSimValDouble(SimData.LONGITUDE) & vbNewLine
        str &= "ALTITUDE: " & sim.GetSimValDouble(SimData.ALTITUDE) & vbNewLine
        str &= "HEADING: " & sim.GetSimValDouble(SimData.HEADING) & vbNewLine
        str &= "AIRSPEED (I): " & sim.GetSimValDouble(SimData.AIRSPEED_INDICATED) & vbNewLine
        str &= "PARKING BRAKE: " & sim.GetSimValDouble(SimData.PARKING_BRAKE) & vbNewLine

        str &= "FLAPS POS: " & sim.GetSimValDouble(SimData.SimVarEnum.CUST_FLAPS_POSITION) & vbNewLine

        str &= "LIGHT NAV: " & sim.GetSimValBool("LIGHT NAV") & vbNewLine

        'str &= "TITLE: " & sim.GetSimValString(SimData.SimVarEnum.TITLE) & vbNewLine
        'str &= "ATC ID: " & sim.GetSimValString(sim.SimVarName(SimData.SimVarEnum.ATC_ID)) & vbNewLine

        'str &= "FLAPS NUM POS: " & sim.GetSimValDouble(sim.SimVarName(SimData.SimVarEnum.FLAPS_NUM_HANDLE_POSITIONS)) & vbNewLine
        'str &= "FLAPS INDEX: " & sim.GetSimValDouble(sim.SimVarName(SimData.SimVarEnum.FLAPS_HANDLE_INDEX)) & vbNewLine
        'str &= "FLAPS %: " & sim.GetSimValDouble(sim.SimVarName(SimData.SimVarEnum.TRAILING_EDGE_FLAPS_LEFT_PERCENT)) & vbNewLine
        'str &= "FLAPS deg: " & AircraftData.RadiansToDegrees(sim.GetSimValDouble(sim.SimVarName(SimData.SimVarEnum.TRAILING_EDGE_FLAPS_LEFT_ANGLE))) & vbNewLine

        'str &= "ATC MODEL: " & sim.GetSimValString(sim.SimVarName(SimData.SimVarEnum.ATC_MODEL)) & vbNewLine
        'str &= "ATC TYPE: " & sim.GetSimValString(sim.SimVarName(SimData.SimVarEnum.ATC_TYPE)) & vbNewLine

        doDisplay(str & vbNewLine)

    End Sub

    Private Sub addRegSimVar(ByVal simVar As SimData.SimVarEnum)
        addRegSimVar(sim.SimVarName(simVar))
    End Sub
    Private Sub addRegSimVar(ByVal simVar As String)

        If Not myRegisteredSimVars.Contains(simVar) Then
            myRegisteredSimVars.Add(simVar)
        End If

    End Sub

    Private Sub registerSimVars()

        For Each sv In myJoyMaps.MappedSimVars
            If Not sim.ContainsSimVar(sv) Then sim.RegisterCustomSimVar(sv)
            addRegSimVar(sv)
        Next

        addRegSimVar(SimData.SimVarEnum.CUST_FLAPS_POSITION)

        sim.RegisterSimVars(myRegisteredSimVars)

    End Sub

    Private Sub timerTick()

        Me.HeartbeatLabel1.ForeColor = Color.Red

        If simEnabled Then
            If Not sim.Enable Then sim.Enable = True
        Else
            If sim.Enable Then sim.Enable = False
        End If

        If sim.Connected Then 'We are connected, Let's try to grab the data from the Sim
            Try
                Me.HeartbeatLabel1.ForeColor = Color.Green
                If Not bLastConnected Then
                    bLastConnected = True
                    SimConnectDebugForm.SetTransmitButtonVisible(True)
                    Me.Text = Me.WindowLabel
                End If
                setSimRateMenuTicks()
                Me.txtCurrSimRate.Text = sim.CurrentSimRate.ToString
                displaySimVals()
                pollJoysticks()
            Catch ex As Exception
                doStatus("Timer exception: " & ex.Message)
            End Try
        Else
            If bLastConnected Then
                bLastConnected = False
                SimConnectDebugForm.SetTransmitButtonVisible(False)
                Dim t As String = Me.WindowLabel
                Me.Text = t
                Me.NotifyIcon1.Text = t
                Me.txtCurrSimRate.Text = ""
            End If
            If cmdRcvd > lastCmdRcvd Then
                Me.txtCurrSimRate.Text = lastCmd
            End If
            doDisplay(sIM_NOT_CONNECT_MSG)
        End If

        If Now >= lastHeartbeat.AddMilliseconds(1000) Then
            lastHeartbeat = Now
            Me.HeartbeatLabel1.Toggle()
        End If

    End Sub

    Private Sub formNotBusy()
        formBusy(False)
    End Sub

    Private Sub formBusy(Optional ByVal busyFlag As Boolean = True)

        Me.gbSlew.Enabled = (Not busyFlag)
        Me.gbSimRate.Enabled = (Not busyFlag)

    End Sub

    Private Sub doSettings()

        Dim frm As New SettingsForm
        frm.StartMinimised = My.Settings.StartMinimised
        frm.AutoStartSimConnect = My.Settings.AutoStartSimCon
        frm.AutoStartServer = My.Settings.AutoStartServer
        frm.AutoStartJoyMap = My.Settings.AutoStartJoy
        frm.ServerPort = My.Settings.ServerPort
        frm.SimConnectPollingInterval = My.Settings.SimConnectPollInterval_ms

        If frm.ShowDialog(Me) = DialogResult.OK Then
            My.Settings.StartMinimised = frm.StartMinimised
            My.Settings.AutoStartSimCon = frm.AutoStartSimConnect
            My.Settings.AutoStartServer = frm.AutoStartServer
            My.Settings.AutoStartJoy = frm.AutoStartJoyMap
            My.Settings.ServerPort = frm.ServerPort
            My.Settings.SimConnectPollInterval_ms = frm.SimConnectPollingInterval

            My.Settings.Save()
        End If

    End Sub

    Private Sub doJoysticks()

        Dim frm As New JoystickForm

        If frm.ShowDialog(Me) = DialogResult.OK Then

            'My.Settings.Save()
        End If

    End Sub

    Private Sub hideForm()

        lastLeft = Me.Left
        Me.Left = HIDDEN_X

    End Sub

    Private Sub restoreForm()
        Me.Left = lastLeft
    End Sub

    Private Function isQuiet() As Boolean

        For Each cmd As String In My.Application.CommandLineArgs
            If cmd.ToUpper.StartsWith("/Q") Or cmd.ToUpper.StartsWith("-Q") Then
                Return True
            End If
        Next

        Return False

    End Function

    Private Function getPort() As Integer

        For Each cmd As String In My.Application.CommandLineArgs
            If cmd.ToUpper.StartsWith("/P:") Then
                Return CInt(cmd.ToUpper.Replace("/P:", ""))
            End If
            If cmd.ToUpper.StartsWith("-P:") Then
                Return CInt(cmd.ToUpper.Replace("-P:", ""))
            End If
        Next

        Return My.Settings.ServerPort

    End Function

    Private ReadOnly Property WindowLabel As String
        Get
            Dim t As String = WINDOW_LABEL & "  " & Application.ProductVersion
            If sim Is Nothing Then
                t &= " - Disconnected"
            Else
                If sim.Connected Then
                    t &= " - Connected"
                Else
                    t &= " - Disconnected"
                End If
            End If
            Return t
        End Get
    End Property

#End Region

#Region "Event Handlers"

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim bQuiet As Boolean = isQuiet()

        bFormLoading = True

        Randomize()

        If My.Settings.FormX <> -1 Then Me.Left = My.Settings.FormX
        If My.Settings.FormY <> -1 Then Me.Top = My.Settings.FormY

        lastLeft = Me.Left

#If DEBUG Then
        Me.btnSimConDebug.Visible = True
        Me.btnRecord.Visible = True
#Else
        Me.btnSimConDebug.Visible = False
        Me.btnRecord.Visible = False
#End If

        formBusy()

        myJoyMaps = New JoystickMappings(AddressOf joyEventReceived)
        For Each jp In myJoyMaps.ProfileNames
            Dim tsmi = New ToolStripMenuItem(jp)
            AddHandler tsmi.Click, AddressOf ctx1Item_JoyProfile_Click
            Me.ctx1Item_EnableJoyMaps.DropDownItems.Add(tsmi)
        Next

        SimConnectDebugForm.SetTransmitButtonVisible(False)
        Me.Text = Me.WindowLabel
        validateSlew()
        doSimRateChanged()

        formNotBusy()

        If isQuietStart Then
            hideForm()
            Me.NotifyIcon1.Visible = False
            doEnableServer(True)
        Else
            If My.Settings.AutoStartServer Or bQuiet Then
                doEnableServer(True)
            End If

            If My.Settings.StartMinimised Or bQuiet Then
                hideForm()
            End If
        End If

        If bQuiet Then
            Me.NotifyIcon1.Visible = False
        End If

        bFormLoading = False

    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        joyMapsDisable()
        simDisable()

        Me.NotifyIcon1.Visible = False

        If Me.Left = HIDDEN_X Then
            My.Settings.FormX = Me.lastLeft
        Else
            My.Settings.FormX = Me.Left
        End If
        My.Settings.FormY = Me.Top

        My.Settings.Save()

    End Sub

    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

        If My.Settings.AutoStartSimCon Or isQuietStart Then
            doEnableSimCon(True)
        End If

        If My.Settings.AutoStartJoy Or isQuietStart Then
            doEnableJoyMaps(My.Settings.JoyProfile, True)
        End If

        Me.ShowInTaskbar = False

    End Sub

    Private Sub tmrUpdateTimer_Tick(sender As Object, e As EventArgs) Handles tmrUpdateTimer.Tick
        timerTick()
    End Sub

    Private Sub btnClip_Click(sender As Object, e As EventArgs) Handles btnClip.Click
        doClip(False)
    End Sub

    Private Sub btnSlew_Click(sender As Object, e As EventArgs) Handles btnSlew.Click
        doSlew()
    End Sub

    Private Sub btnActivateSimRate_Click(sender As Object, e As EventArgs) Handles btnActivateSimRate.Click
        doActivate()
    End Sub

    Private Sub txtNewLat_Validated(sender As Object, e As EventArgs) Handles txtNewLat.Validated
        validateSlew()
    End Sub

    Private Sub txtNewLong_Validated(sender As Object, e As EventArgs) Handles txtNewLong.Validated
        validateSlew()
    End Sub

    Private Sub cbSimRate_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbSimRate.SelectedIndexChanged
        doSimRateChanged()
    End Sub

    Private Sub cmTrayIcon_MouseDoubleClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles cmTrayIcon.MouseDoubleClick, NotifyIcon1.MouseDoubleClick

        If Me.isHidden Then
            restoreForm()
        Else
            hideForm()
        End If

    End Sub


    Private Sub ctx1Item_Slew_Click(sender As Object, e As EventArgs) Handles ctx1Item_Slew.Click
        doClip(True)
    End Sub

    Private Sub ctx1Item_Settings_Click(sender As Object, e As EventArgs) Handles ctx1Item_Settings.Click
        doSettings()
    End Sub

    Private Sub ctx1Item_Rate_x1_Click(sender As System.Object, e As System.EventArgs) Handles ctx1Item_Rate_x1.Click
        doSimRate(1)
    End Sub

    Private Sub ctx1Item_Rate_x2_Click(sender As Object, e As EventArgs) Handles ctx1Item_Rate_x2.Click
        doSimRate(2)
    End Sub

    Private Sub ctx1Item_Rate_x4_Click(sender As Object, e As EventArgs) Handles ctx1Item_Rate_x4.Click
        doSimRate(4)
    End Sub

    Private Sub ctx1Item_Rate_x8_Click(sender As System.Object, e As System.EventArgs) Handles ctx1Item_Rate_x8.Click
        doSimRate(8)
    End Sub

    Private Sub ctx1Item_Exit_Click(sender As System.Object, e As System.EventArgs) Handles ctx1Item_Exit.Click

        Me.WindowState = FormWindowState.Normal
        Me.Close()

    End Sub

    Private Sub ctx1Item_Minimise_Click(sender As System.Object, e As System.EventArgs) Handles ctx1Item_Minimise.Click

        If Not Me.isHidden Then
            hideForm()
        End If

    End Sub

    Private Sub ctx1Item_Restore_Click(sender As System.Object, e As System.EventArgs) Handles ctx1Item_Restore.Click

        If Me.isHidden Then
            restoreForm()
        End If

    End Sub

    Private Sub MainForm_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDoubleClick

        If Me.isHidden Then
            restoreForm()
        Else
            hideForm()
        End If

    End Sub

    Private Sub btnSimConDebug_Click(sender As Object, e As EventArgs) Handles btnSimConDebug.Click
        doSimConDebug()
    End Sub

    Private Sub ctx1Item_EnableServer_Click(sender As Object, e As EventArgs) Handles ctx1Item_EnableServer.Click
        doEnableServer()
    End Sub

    Private Sub ctx1Item_EnableSimCon_Click(sender As Object, e As EventArgs) Handles ctx1Item_EnableSimCon.Click
        doEnableSimCon()
    End Sub

    Private Sub ctx1Item_G1000PFD_Click(sender As Object, e As EventArgs) Handles ctx1Item_G1000PFD.Click
        openG1000PFD()
    End Sub

    Private Sub ctx1Item_Joysticks_Click(sender As Object, e As EventArgs) Handles ctx1Item_Joysticks.Click
        doJoysticks()
    End Sub

    Private Sub ctx1Item_JoyProfile_Click(sender As Object, e As EventArgs)

        Dim tsmi = CType(sender, ToolStripMenuItem)
        doEnableJoyMaps(tsmi.Text)

    End Sub

#If DEBUG Then
    Private Sub btnRecord_Click(sender As Object, e As EventArgs) Handles btnRecord.Click
        doRecord()
    End Sub

#End If

#End Region

End Class
