Option Explicit On
Option Strict On

Imports Microsoft.FlightSimulator.SimConnect

Public Class SimConnectLib

    '
    'The SimVars to be read and SimEvents to be sent are held in the SimData class.
    'Any extra SimVars and SimEvents required must be added to this class.
    '

#Region "Public Declarations"

    Public Structure Coordinates
        Public Latitude As Double
        Public Longitude As Double
    End Structure

    Public Property UpdateInterval_ms As Integer = 500

    Public ReadOnly Property Connected As Boolean
        Get
            Return (sim IsNot Nothing)
        End Get
    End Property

    Public ReadOnly Property CurrentSimRate As Integer
        Get
            Return GetSimValInt("SIMULATION RATE")
        End Get
    End Property

    Public Function GetSimValBool(ByVal simValName As String) As Boolean
        Return CBool(GetSimVal(simValName))
    End Function

    Public Function GetSimValInt(ByVal simValName As String) As Integer
        Return CInt(GetSimVal(simValName))
    End Function

    Public Function GetSimVal(ByVal simVarName As String) As Double

        If Not simDat.Contains(simVarName) Then Throw New Exception("Unknown SimVar: " & simVarName)
        Dim t As SimData.SimVar = simDat.SimVars(simVarName)
        Dim ndx As Integer = t.Index
        If Not regSimVars.ContainsKey(ndx) Then Throw New Exception("SimVar not registered: " & simVarName)
        If Not simVals.ContainsKey(ndx) Then Throw New Exception("Registered SimVar, data not found: " & simVarName)
        Return simVals(ndx)

    End Function

#End Region

#Region "Private Declarations"

    Private Const SIMCONNECT_OBJECT_ID_USER = 0
    Private Const HIDDEN_X As Integer = -16000

    Private Const SIM_RATE_MIN As Integer = 1
    Private Const SIM_RATE_MAX As Integer = 8

    ' User-defined win32 event => put basically any number?
    Private Const WM_USER_SIMCONNECT As Integer = &H402

    Private Enum DUMMYENUM
        Dummy = 0
    End Enum

    Private Enum DATA_DEFINE_ID
        DEFINITION3
    End Enum

    Private Enum GROUP
        GROUP1
    End Enum

    Private simRateStepIndex As New Dictionary(Of Integer, Integer) From
        {{1, 1},
        {2, 2},
        {4, 3},
        {8, 4},
        {16, 5},
        {32, 6},
        {64, 7}}

    Private regSimVars As New Dictionary(Of Integer, SimData.SimVar)
    Private simVals As New Dictionary(Of Integer, Double)
    Private simDat As New SimData
    Private sim As SimConnect = Nothing
    Private lastSimRate As Integer = 0
    Private myForm As Windows.Forms.Form = Nothing
    Private myTimer As Windows.Forms.Timer = Nothing

#End Region

#Region "Public Functions"

    Public Shared Function GetSimEventByName(ByVal simEvtName As String) As SimData.EventEnum

        Try
            Return CType([Enum].Parse(GetType(SimData.EventEnum), simEvtName), SimData.EventEnum)
        Catch ex As Exception
            Return SimData.EventEnum.NONE
        End Try

    End Function

    Public Sub New(ByVal pForm As Windows.Forms.Form)

        myForm = pForm

    End Sub

    Public Function CheckValidSlew(ByRef newLat As String, ByRef newLong As String) As Boolean

        If newLat = "" Then Return False
        If newLong = "" Then Return False

        Try
            Dim coords As Coordinates = parseCoords(newLat & " " & newLong)
            If checkValidSlew(coords) Then
                newLat = coords.Latitude.ToString
                newLong = coords.Longitude.ToString
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Sub RegisterSimVars(ByVal varNames As List(Of String))

        Disconnect()

        regSimVars.Clear()
        simVals.Clear()

        For Each str As String In varNames
            If Not simDat.SimVars.ContainsKey(str) Then
                Throw New Exception("SimVar not known: " & str)
            End If
            Dim t As SimData.SimVar = simDat.SimVars(str)
            regSimVars.Add(t.Index, t)
            simVals.Add(t.Index, 0.0)
        Next

        Connect()

    End Sub

    Public Sub DoSlew(ByVal coords As Coordinates)
        DoSlew(coords.Latitude, coords.Longitude)
    End Sub
    Public Sub DoSlew(ByVal newLat As Double, ByVal newLong As Double)

        Dim posData As New SIMCONNECT_DATA_INITPOSITION

        posData.Altitude = GetSimVal(SimData.ALTITUDE)
        posData.Latitude = newLat
        posData.Longitude = newLong
        posData.Pitch = -0.0
        posData.Bank = -1.0
        posData.Heading = GetSimVal(SimData.HEADING)
        posData.OnGround = 0
        posData.Airspeed = CType(GetSimVal(SimData.AIRSPEED_INDICATED), UInteger)

        doSetPosition(posData)

    End Sub
    Public Sub DoSlew(ByVal coordsText As String, ByVal bExecuteImmediately As Boolean, ByRef newLat As String, ByRef newLong As String)

        If bExecuteImmediately Then
            If sim Is Nothing Then Exit Sub
        End If

        Try
            Dim t As String = coordsText
            If t <> "" Then
                Dim coords As Coordinates = parseCoords(t)

                If bExecuteImmediately Then
                    If checkValidSlew(coords) Then
                        DoSlew(coords)
                    Else
                        Exit Sub
                    End If
                Else
                    If checkValidSlew(coords) Then
                        newLat = coords.Latitude.ToString
                        newLong = coords.Longitude.ToString
                    Else
                        newLat = ""
                        newLong = ""
                    End If
                End If
            End If
        Catch ex As Exception
            newLat = ""
            newLong = ""
            If Not bExecuteImmediately Then
                Throw New Exception("Invalid coordinates format!")
            End If
        End Try

    End Sub

    Public Sub DoSimRate(ByVal newSimRate As Integer)

        If sim Is Nothing Then Exit Sub

        If newSimRate = 0 Then Exit Sub
        If newSimRate > 8 Then Exit Sub

        Dim curRateIndex As Integer = simRateStepIndex(GetSimValInt("SIMULATION RATE"))
        Dim newRateIndex As Integer = simRateStepIndex(newSimRate)

        If curRateIndex = newRateIndex Then Exit Sub
        Debug.WriteLine("curRateIndex :" & curRateIndex)
        Debug.WriteLine("newRateIndex :" & newRateIndex)

        If curRateIndex < newRateIndex Then
            For i = curRateIndex To newRateIndex - 1
                TransmitClientEvent(SimData.EventEnum.SIM_RATE_INCR, "0")
                Debug.WriteLine("SIM_RATE_INCR " & Now)
            Next
        Else
            For i = curRateIndex To newRateIndex + 1 Step -1
                TransmitClientEvent(SimData.EventEnum.SIM_RATE_DECR, "0")
                Debug.WriteLine("SIM_RATE_DECR " & Now)
            Next
        End If

    End Sub

    Public Sub TransmitClientEvent(ByVal eventID As [Enum])
        TransmitClientEvent(eventID, "0")
    End Sub
    Public Sub TransmitClientEvent(ByVal eventID As [Enum], ByVal data As String)

        Dim Bytes As Byte() = BitConverter.GetBytes(Convert.ToInt32(data))
        Dim Param As UInt32 = BitConverter.ToUInt32(Bytes, 0)
        Try
            sim.TransmitClientEvent(0, eventID, Param, GROUP.GROUP1, SIMCONNECT_EVENT_FLAG.GROUPID_IS_PRIORITY)
        Catch ex As Exception
            Throw New Exception("The Transmit Request Failed: " + ex.Message)
        End Try

    End Sub

    Public Sub InitTimer(ByVal pTimer As Windows.Forms.Timer)

        myTimer = pTimer
        pTimer.Interval = UpdateInterval_ms
        pTimer.Enabled = True
        pTimer.Start()

    End Sub

    Public Sub TimerTick()

        If (sim Is Nothing) Then 'We Then are Not connected, Let's try to connect
            Connect()
        Else 'We are connected, Let's try to grab the data from the Sim
            Try
                For Each simVal In regSimVars.Values
                    sim.RequestDataOnSimObjectType(CType(simVal.Index, DUMMYENUM), CType(simVal.Index, DUMMYENUM), 0, SIMCONNECT_SIMOBJECT_TYPE.USER)
                Next
            Catch ex As Exception
                Disconnect()
            End Try
        End If

    End Sub

    Public Sub WndProcHandler(ByRef m As Windows.Forms.Message)

        Try
            If (m.Msg = WM_USER_SIMCONNECT) Then
                receiveSimConnectMessage()
            End If
        Catch ex As Exception
            Disconnect()
        End Try

    End Sub

    Public Sub Disconnect()

        If (sim IsNot Nothing) Then
            If myTimer IsNot Nothing Then
                myTimer.Stop()
                myTimer.Enabled = False
                myTimer = Nothing
            End If

            sim.Dispose()
            sim = Nothing
        End If

    End Sub

    Public Sub MapEventEnums(ByVal pEnums As List(Of [Enum]))

        Try
            For Each item As [Enum] In pEnums
                sim.MapClientEventToSimEvent(item, item.ToString())
            Next
        Catch ex As Exception
            sim = Nothing
        End Try

    End Sub

    Public Sub Connect()

        Try
            sim = New SimConnect(myForm.Text, myForm.Handle, WM_USER_SIMCONNECT, Nothing, 0)
            AddHandler sim.OnRecvOpen, AddressOf sim_OnRecvOpen
            AddHandler sim.OnRecvQuit, AddressOf sim_OnRecvQuit
            AddHandler sim.OnRecvSimobjectDataBytype, AddressOf sim_OnRecvSimobjectDataBytype
        Catch ex As Exception
            sim = Nothing
        End Try

    End Sub

#End Region

#Region "Private Functions"

    Private Sub doSetPosition(ByVal slewData As SIMCONNECT_DATA_INITPOSITION)

        Try
            sim.SetDataOnSimObject(DATA_DEFINE_ID.DEFINITION3, 0, 0, slewData)
        Catch ex As Exception
            Throw New Exception("The SetData Request Failed: " + ex.Message)
        End Try

    End Sub

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

    Private Function checkValidSlew(ByVal coords As Coordinates) As Boolean

        If coords.Latitude = -1 Or coords.Longitude = -1 Then
            Return False
        Else
            Return True
        End If

    End Function

    Private Sub sim_OnRecvQuit(ByVal sender As SimConnect, ByVal data As SIMCONNECT_RECV)
    End Sub

    Private Sub sim_OnRecvOpen(ByVal sender As SimConnect, ByVal data As SIMCONNECT_RECV_OPEN)

        For Each t As SimData.SimVar In simDat.SimVars.Values
            ' Define a data structure
            sim.AddToDataDefinition(CType(t.Index, DUMMYENUM), t.Name, t.Units, SIMCONNECT_DATATYPE.FLOAT64, 0.0, SimConnect.SIMCONNECT_UNUSED)

            ' IMPORTANT: Register it With the simconnect managed wrapper marshaller
            ' If you skip this step, you will only receive a uint in the .dwData field.
            sim.RegisterDataDefineStruct(Of Double)(CType(t.Index, DUMMYENUM))
        Next

        For Each item As SimData.EventEnum In [Enum].GetValues(GetType(SimData.EventEnum))
            If item <> SimData.EventEnum.NONE Then
                sim.MapClientEventToSimEvent(item, item.ToString())
                sim.AddClientEventToNotificationGroup(GROUP.GROUP1, item, False)
            End If
        Next

        sim.AddToDataDefinition(DATA_DEFINE_ID.DEFINITION3, "Initial Position", "", SIMCONNECT_DATATYPE.INITPOSITION, 0, 0)

    End Sub

    Private Sub sim_OnRecvSimobjectDataBytype(ByVal sender As SimConnect, ByVal data As SIMCONNECT_RECV_SIMOBJECT_DATA_BYTYPE)

        Dim iRequest As Integer = CInt(data.dwRequestID)
        Dim dValue As Double = CDbl(data.dwData(0))

        If Not simVals.ContainsKey(iRequest) Then Throw New Exception("Registered SimVar index not found: " & iRequest)
        simVals(iRequest) = dValue

    End Sub

    Private Sub receiveSimConnectMessage()
        sim?.ReceiveMessage()
    End Sub

#End Region

End Class

