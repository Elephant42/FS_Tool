Option Explicit On
Option Strict On

Imports Microsoft.FlightSimulator.SimConnect

Public Class MainFormX

    ' User-defined win32 event => put basically any number?
    Public Const WM_USER_SIMCONNECT As Integer = &H402

    Public Const SIMCONNECT_OBJECT_ID_USER = 0

    Public Enum DUMMYENUM
        Dummy = 0
    End Enum

    Public Enum SimValsEnum
        LONGITUDE = 1
        LATITUDE = 2
        HEADING = 3
        ALTITUDE = 4
        AIRSPEED = 5
        SIM_RATE = 6
    End Enum

    Public Enum DATA_DEFINE_ID
        DEFINITION3
    End Enum

    Public Enum GROUP
        GROUP1
    End Enum

    Public Enum EventEnum
        AUTOPILOT_ON
        AUTOPILOT_OFF
        THROTTLE_SET
        AP_SPD_VAR_SET
        SIM_RATE_INCR
        SIM_RATE_DECR
    End Enum


    Private Const WINDOW_LABEL As String = "Flight Sim Tool"
    Private Const SIM_RATE_MIN As Integer = 1
    Private Const SIM_RATE_MAX As Integer = 8

    Private Structure Coordinates
        Public Latitude As Double
        Public Longitude As Double
    End Structure

    ' <summary>
    ' Contains the list of all the SimConnect properties we will read, the unit Is separated by coma by our own code.
    ' </summary>
    Private simConnectProperties As New Dictionary(Of Integer, String) From
        {{1, "PLANE LONGITUDE,degree"},
        {2, "PLANE LATITUDE,degree"},
        {3, "PLANE HEADING DEGREES MAGNETIC,degree"},
        {4, "PLANE ALTITUDE,feet"},
        {5, "AIRSPEED INDICATED,knots"},
        {6, "SIMULATION RATE,times"}}

    Private sim As SimConnect = Nothing

    Private simLabels As New Dictionary(Of Integer, String) From
        {{1, ""},
        {2, ""},
        {3, ""},
        {4, ""},
        {5, ""},
        {6, ""}}

    Private simVals As New Dictionary(Of Integer, Double) From
        {{1, 0.0},
        {2, 0.0},
        {3, 0.0},
        {4, 0.0},
        {5, 0.0},
        {6, 0.0}}

    Private simRateStepIndex As New Dictionary(Of Integer, Integer) From
        {{1, 1},
        {2, 2},
        {4, 3},
        {8, 4},
        {16, 5},
        {32, 6},
        {64, 7}}

    Private Sub updateVals()

        Dim tempDsp As String = ""
        For Each iIndex As Integer In simConnectProperties.Keys
            Dim temp As String = simLabels(iIndex) & ": " & simVals(iIndex) & vbNewLine
            If iIndex = SimValsEnum.SIM_RATE Then
                Me.txtCurrSimRate.Text = simVals(iIndex).ToString
            Else
                If tempDsp = "" Then
                    tempDsp = temp
                Else
                    tempDsp &= temp
                End If
            End If
        Next iIndex

        Me.txtDisplay.Text = tempDsp

    End Sub

    Private Sub doSetPosition(ByVal slewData As SIMCONNECT_DATA_INITPOSITION)

        Try
            sim.SetDataOnSimObject(DATA_DEFINE_ID.DEFINITION3, 0, 0, slewData)
        Catch ex As Exception
            MessageBox.Show("The SetData Request Failed: " + ex.Message)
        End Try

    End Sub

    '27.5513° S 152.9835° E
    '-27.3426258,152.8341711
    '27° 42' 16.52" S 153° 0' 14.72" E
    '27.7702° S 152.9807° E
    '33.1376° N 63.7603° W
    '27° 50.95' S 153° 8.49' E
    '-27.69973 153.05353
    '152.99722 -27.76051
    Private Function parseCoords(ByVal input As String) As Coordinates

        Dim coords As New Coordinates
        Dim sep As String()
        coords.Latitude = -1
        coords.Longitude = -1

        input = input.Replace("°", "")
        input = input.Replace("'", "")
        input = input.Replace("""", "")

        If input.Contains(",") Then
            sep = {","}
        Else
            sep = {" "}
        End If

        Dim parts As String() = input.Split(sep, StringSplitOptions.None)
        Dim numParts As Integer = parts.Count
        If numParts = 2 Then
            coords.Latitude = CDbl(parts(0))
            coords.Longitude = CDbl(parts(1))
        ElseIf numParts = 4 Then
            coords.Latitude = CDbl(parts(0))
            coords.Longitude = CDbl(parts(2))
            If parts(1).ToUpper = "S" Then coords.Latitude *= -1
            If parts(3).ToUpper = "W" Then coords.Longitude *= -1
        ElseIf numParts = 6 Then
            Dim latDeg As Integer = CInt(parts(0))
            Dim longDeg As Integer = CInt(parts(3))
            Dim latMin As Double = CDbl(parts(1))
            Dim longMin As Double = CDbl(parts(4))

            coords.Latitude = latDeg + latMin / 60
            coords.Longitude = longDeg + longMin / 60

            If parts(2).ToUpper = "S" Then coords.Latitude *= -1
            If parts(5).ToUpper = "W" Then coords.Longitude *= -1
        ElseIf numParts = 8 Then
            Dim latDeg As Integer = CInt(parts(0))
            Dim longDeg As Integer = CInt(parts(4))
            Dim latMin As Double = CDbl(parts(1))
            Dim longMin As Double = CDbl(parts(5))
            Dim latSec As Double = CDbl(parts(2))
            Dim longSec As Double = CDbl(parts(6))

            coords.Latitude = latDeg + latMin / 60 + latSec / 3600
            coords.Longitude = longDeg + longMin / 60 + longSec / 3600

            If parts(3).ToUpper = "S" Then coords.Latitude *= -1
            If parts(7).ToUpper = "W" Then coords.Longitude *= -1
        End If

        Return coords

    End Function

    Private Sub doClip()

        Try
            Dim t As String = My.Computer.Clipboard.GetText()
            If t <> "" Then
                Dim coords As Coordinates = parseCoords(t)
                If coords.Latitude = -1 Or coords.Longitude = -1 Then
                    Me.txtNewLat.Text = ""
                    Me.txtNewLong.Text = ""
                Else
                    Me.txtNewLat.Text = coords.Latitude.ToString
                    Me.txtNewLong.Text = coords.Longitude.ToString
                End If
            End If
            validateSlew()
        Catch ex As Exception
            Me.txtNewLat.Text = ""
            Me.txtNewLong.Text = ""
            validateSlew()
            MessageBox.Show(Me, "Invalid coordinates format!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Function checkValidSlew() As Boolean

        If Me.txtNewLat.Text = "" Then Return False
        If Me.txtNewLong.Text = "" Then Return False

        Try
            Dim coords As Coordinates = parseCoords(Me.txtNewLat.Text & " " & Me.txtNewLong.Text)
            If checkValidSlew(coords) Then
                Me.txtNewLat.Text = coords.Latitude.ToString
                Me.txtNewLong.Text = coords.Longitude.ToString
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function
    Private Function checkValidSlew(ByVal coords As Coordinates) As Boolean

        If coords.Latitude = -1 Or coords.Longitude = -1 Then
            Return False
        Else
            Return True
        End If

    End Function

    Private Sub validateSlew()

        If sim Is Nothing Then
            Me.btnSlew.Enabled = False
        Else
            Me.btnSlew.Enabled = checkValidSlew()
        End If

    End Sub

    Private Sub doSlew()

        Dim posData As New SIMCONNECT_DATA_INITPOSITION

        posData.Altitude = simVals(SimValsEnum.ALTITUDE)
        posData.Latitude = CDbl(Me.txtNewLat.Text)
        posData.Longitude = CDbl(Me.txtNewLong.Text)
        posData.Pitch = -0.0
        posData.Bank = -1.0
        posData.Heading = simVals(SimValsEnum.HEADING)
        posData.OnGround = 0
        posData.Airspeed = CType(simVals(SimValsEnum.AIRSPEED), UInteger)

        doSetPosition(posData)

    End Sub

    Private Sub doSimRate(ByVal simRate As Integer)

        If simRate = 0 Then Exit Sub
        If simRate > 8 Then Exit Sub

        Dim curRate As Integer = CInt(simVals(SimValsEnum.SIM_RATE))

        If curRate = simRate Then Exit Sub

        If curRate < simRate Then
            If curRate = 1 Then

            End If
            For i = curRate To simRate
                simRateIncr()
            Next
        Else
            For i = curRate To simRate Step -1
                simRateDecr()
            Next
        End If

    End Sub

    Private Sub simRateIncr()

        If simVals(SimValsEnum.SIM_RATE) < 8 Then
            TransmitEvent(EventEnum.SIM_RATE_INCR, "0")
        End If

    End Sub

    Private Sub simRateDecr()

        If simVals(SimValsEnum.SIM_RATE) > 1 Then
            TransmitEvent(EventEnum.SIM_RATE_DECR, "0")
        End If

    End Sub

    Private Sub doActivate()

        If Me.cbSimRate.Text = "" Then Exit Sub

        doSimRate(CInt(Me.cbSimRate.Text))

    End Sub

    Private Sub doSimRateChanged()

        If sim Is Nothing Then
            Me.btnActivateSimRate.Enabled = False
            Exit Sub
        End If

        If Me.cbSimRate.Text = "" Then
            Me.btnActivateSimRate.Enabled = False
        Else
            Me.btnActivateSimRate.Enabled = True
        End If

    End Sub

    Public Sub TransmitEvent(ByVal cmd As EventEnum, ByVal data As String)

        Dim Bytes As Byte() = BitConverter.GetBytes(Convert.ToInt32(data))
        Dim Param As UInt32 = BitConverter.ToUInt32(Bytes, 0)
        Try
            sim.TransmitClientEvent(0, cmd, Param, GROUP.GROUP1, SIMCONNECT_EVENT_FLAG.GROUPID_IS_PRIORITY)
        Catch ex As Exception
            MessageBox.Show("The Transmit Request Failed: " + ex.Message)
        End Try

    End Sub

    Private Sub initTimer()

        tmrUpdateTimer.Interval = My.Settings.UpdateInterval_ms
        tmrUpdateTimer.Start()

    End Sub

    Private Sub timerTick()

        If (sim Is Nothing) Then 'We Then are Not connected, Let's try to connect
            Connect()
        Else 'We are connected, Let's try to grab the data from the Sim
            Try
                For Each toConnect In simConnectProperties
                    sim.RequestDataOnSimObjectType(CType(toConnect.Key, DUMMYENUM), CType(toConnect.Key, DUMMYENUM), 0, SIMCONNECT_SIMOBJECT_TYPE.USER)
                Next
                updateVals()
            Catch ex As Exception
                Disconnect()
            End Try
        End If

    End Sub

    ' <summary>
    ' We received a disconnection from SimConnect
    ' </summary>
    ' <param name="sender"></param>
    ' <param name="data"></param>
    Private Sub Sim_OnRecvQuit(ByVal sender As SimConnect, ByVal data As SIMCONNECT_RECV)
        Me.Text = WINDOW_LABEL & " - Disconnected"
    End Sub

    ' <summary>
    ' We received a connection from SimConnect.
    ' Let's register all the properties we need.
    ' </summary>
    ' <param name="sender"></param>
    ' <param name="data"></param>
    Private Sub Sim_OnRecvOpen(ByVal sender As SimConnect, ByVal data As SIMCONNECT_RECV_OPEN)

        Me.Text = WINDOW_LABEL & " - Connected"
        Me.cbSimRate.Text = ""
        Me.btnDecr.Enabled = True
        Me.btnIncr.Enabled = True

        For Each toConnect In simConnectProperties
            Dim values() As String = toConnect.Value.Split(","c)

            simLabels(toConnect.Key) = values(0)

            ' Define a data structure
            sim.AddToDataDefinition(CType(toConnect.Key, DUMMYENUM), values(0), values(1), SIMCONNECT_DATATYPE.FLOAT64, 0.0, SimConnect.SIMCONNECT_UNUSED)

            ' IMPORTANT: Register it With the simconnect managed wrapper marshaller
            ' If you skip this step, you will only receive a uint in the .dwData field.
            sim.RegisterDataDefineStruct(Of Double)(CType(toConnect.Key, DUMMYENUM))
        Next

    End Sub

    ' <summary>
    ' Received data from SimConnect
    ' </summary>
    ' <param name="sender"></param>
    ' <param name="data"></param>
    Private Sub Sim_OnRecvSimobjectDataBytype(ByVal sender As SimConnect, ByVal data As SIMCONNECT_RECV_SIMOBJECT_DATA_BYTYPE)

        Dim iRequest As Integer = CInt(data.dwRequestID)
        Dim dValue As Double = CDbl(data.dwData(0))

        simVals(iRequest) = dValue

    End Sub

    Public Sub ReceiveSimConnectMessage()
        sim?.ReceiveMessage()
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)

        Try
            If (m.Msg = WM_USER_SIMCONNECT) Then
                ReceiveSimConnectMessage()
            End If
        Catch ex As Exception
            Disconnect()
        End Try

        MyBase.WndProc(m)

    End Sub

    Public Sub Disconnect()

        If (sim IsNot Nothing) Then
            sim.Dispose()
            sim = Nothing
            Me.Text = WINDOW_LABEL & " - Disconnected"
            Me.btnActivateSimRate.Enabled = False
            Me.btnIncr.Enabled = False
            Me.btnDecr.Enabled = False
        End If

    End Sub

    Private Sub Connect()

        'The constructor Is similar to SimConnect_Open in the native API
        Try
            'Pass the self defined ID which will be returned on WndProc
            sim = New SimConnect(Me.Text, Me.Handle, WM_USER_SIMCONNECT, Nothing, 0)
            AddHandler sim.OnRecvOpen, AddressOf Sim_OnRecvOpen
            AddHandler sim.OnRecvQuit, AddressOf Sim_OnRecvQuit
            AddHandler sim.OnRecvSimobjectDataBytype, AddressOf Sim_OnRecvSimobjectDataBytype

            For Each item As EventEnum In [Enum].GetValues(GetType([EventEnum]))
                sim.MapClientEventToSimEvent(item, item.ToString())
                sim.AddClientEventToNotificationGroup(GROUP.GROUP1, item, False)
            Next

            sim.AddToDataDefinition(DATA_DEFINE_ID.DEFINITION3, "Initial Position", "", SIMCONNECT_DATATYPE.INITPOSITION, 0, 0)
        Catch ex As Exception
            sim = Nothing
        End Try

    End Sub

    Private Sub formNotBusy()
        formBusy(False)
    End Sub

    Private Sub formBusy(Optional ByVal busyFlag As Boolean = True)

        Me.gbSlew.Enabled = (Not busyFlag)
        Me.gbSimRate.Enabled = (Not busyFlag)

        'Me.gbSimRate.Visible = False

    End Sub


#Region "Event Handlers"

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Left = My.Settings.FormX
        Me.Top = My.Settings.FormY

        formBusy()
        Me.Text = WINDOW_LABEL
        Connect()
        validateSlew()
        doSimRateChanged()
        Me.btnDecr.Enabled = False
        Me.btnIncr.Enabled = False
        formNotBusy()

    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        My.Settings.FormX = Me.Left
        My.Settings.FormY = Me.Top
        My.Settings.Save()

        Disconnect()

    End Sub

    Private Sub tmrUpdateTimer_Tick(sender As Object, e As EventArgs) Handles tmrUpdateTimer.Tick
        timerTick()
    End Sub

    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        initTimer()
    End Sub

    Private Sub btnClip_Click(sender As Object, e As EventArgs) Handles btnClip.Click
        doClip()
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

    Private Sub btnDecr_Click(sender As Object, e As EventArgs) Handles btnDecr.Click
        simRateDecr()
    End Sub

    Private Sub btnIncr_Click(sender As Object, e As EventArgs) Handles btnIncr.Click
        simRateIncr()
    End Sub

#End Region

End Class
