Option Explicit On
Option Strict On

Imports System.Diagnostics.Eventing.Reader
Imports Microsoft.FlightSimulator.SimConnect
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Threading
Imports System.ComponentModel

Public Class AircraftData

    Public Const FLAPS_MOVING As Integer = -1

    Public Property Title As String
        Get
            Return _Title
        End Get
        Set(value As String)
            _Title = value
            newTitle(value)
        End Set
    End Property

    Public ReadOnly Property CurrentAircraft As AircraftEnum
        Get
            Return _Enum
        End Get
    End Property

    Public Enum AircraftEnum
        A320
        A321
        G58
        SAVAGE_CUB
        XCUB
        BC350I
        B747_400
        B747_8I
        B787
        CRJ700
        G36
        C152
        C172
        C208B
        CJ4
        C_LONGITUDE
        DA40
        DA62
        DR400
        DV20
        E330
        FDCT
        IA5
        C10C
        PAE
        PITTS
        SAVAGE_CARBON
        SAVAGE_SHOCK_ULTRA
        SR22
        TBM930
        VL3
        NONE
        NOTSET
    End Enum

    Public Sub New()
    End Sub

    Public Sub SimDisconnected()
        _Enum = AircraftEnum.NOTSET
    End Sub

    Public Shared Function RadiansToDegrees(ByVal radians As Double) As Double
        Return 180 * radians / Math.PI
    End Function

    Public Function GetFlapsIndex(ByVal flapsAngle As Double) As Integer

        Dim fa As Double = RadiansToDegrees(flapsAngle)

        If flapIndexes Is Nothing Then
            Return 0
        Else
            If flapIndexes.ContainsKey(fa) Then
                Return flapIndexes(fa)
            Else
                Return FLAPS_MOVING
            End If
        End If

    End Function

    Public Shared Function GetAircraft(ByVal sTitle As String) As AircraftEnum

        If sTitle.ToLower.Contains("a320") Then
            Return AircraftEnum.DA62
        ElseIf sTitle.ToLower.Contains("a321") Then
            Return AircraftEnum.A320
        ElseIf sTitle.ToLower.Contains("g58") Then
            Return AircraftEnum.A321
        ElseIf sTitle.ToLower.Contains("savage cub") Then
            Return AircraftEnum.SAVAGE_CUB
        ElseIf sTitle.ToLower.Contains("350i") Then
            Return AircraftEnum.BC350I
        ElseIf sTitle.ToLower.Contains("747-400") Then
            Return AircraftEnum.B747_400
        ElseIf sTitle.ToLower.Contains("747-8i") Then
            Return AircraftEnum.B747_8I
        ElseIf sTitle.ToLower.Contains("787-10") Then
            Return AircraftEnum.B787
        ElseIf sTitle.ToLower.Contains("crj 700") Then
            Return AircraftEnum.CRJ700
        ElseIf sTitle.ToLower.Contains("208b") Then
            Return AircraftEnum.C208B
        ElseIf sTitle.ToLower.Contains("g36") Then
            Return AircraftEnum.G36
        ElseIf sTitle.ToLower.Contains("cessna 152") Then
            Return AircraftEnum.C152
        ElseIf sTitle.ToLower.Contains("cj4") Then
            Return AircraftEnum.CJ4
        ElseIf sTitle.ToLower.Contains("cessna longitude") Then
            Return AircraftEnum.C_LONGITUDE
        ElseIf sTitle.ToLower.Contains("cessna skyhawk") Then
            Return AircraftEnum.C172
        ElseIf sTitle.ToLower.Contains("da40-ng") Then
            Return AircraftEnum.DA40
        ElseIf sTitle.ToLower.Contains("da62") Then
            Return AircraftEnum.DA62
        ElseIf sTitle.ToLower.Contains("dr400") Then
            Return AircraftEnum.DR400
        ElseIf sTitle.ToLower.Contains("dv20") Then
            Return AircraftEnum.DV20
        ElseIf sTitle.ToLower.Contains("extra 330") Then
            Return AircraftEnum.E330
        ElseIf sTitle.ToLower.Contains("flightdesignct") Then
            Return AircraftEnum.FDCT
        ElseIf sTitle.ToLower.Contains("icon a5") Then
            Return AircraftEnum.IA5
        ElseIf sTitle.ToLower.Contains("cap 10 c") Then
            Return AircraftEnum.C10C
        ElseIf sTitle.ToLower.Contains("pipistrel alpha electro") Then
            Return AircraftEnum.PAE
        ElseIf sTitle.ToLower.Contains("pitts") Then
            Return AircraftEnum.PITTS
        ElseIf sTitle.ToLower.Contains("savage carbon") Then
            Return AircraftEnum.SAVAGE_CARBON
        ElseIf sTitle.ToLower.Contains("savage shock ultra") Then
            Return AircraftEnum.SAVAGE_SHOCK_ULTRA
        ElseIf sTitle.ToLower.Contains("sr22") Then
            Return AircraftEnum.SR22
        ElseIf sTitle.ToLower.Contains("tbm 930") Then
            Return AircraftEnum.TBM930
        ElseIf sTitle.ToLower.Contains("vl3") Then
            Return AircraftEnum.VL3
        ElseIf sTitle.ToLower.Contains("xcub") Then
            Return AircraftEnum.XCUB
        Else
            Return AircraftEnum.NONE
        End If

    End Function

    Private Sub newTitle(ByVal sTitle As String)

        If sTitle <> "" Then
            flapIndexes = Nothing
            _Enum = GetAircraft(sTitle)

            If _Enum <> AircraftEnum.NONE Then
                If allFlapIndexes.ContainsKey(_Enum.ToString) Then
                    flapIndexes = allFlapIndexes(_Enum.ToString)
                End If
            End If
        End If

    End Sub

    Private _Title As String = ""
    Private _Enum As AircraftEnum = AircraftEnum.NOTSET
    Private flapIndexes As Dictionary(Of Double, Integer)

    Private allFlapIndexes As New Dictionary(Of String, Dictionary(Of Double, Integer)) From {
            {"A320", New Dictionary(Of Double, Integer) From {{0, 0}, {10, 1}, {15, 2}, {20, 3}, {35, 4}}},
            {"A321", New Dictionary(Of Double, Integer) From {{0, 0}, {10, 1}, {15, 2}, {20, 3}, {35, 4}}},
            {"G58", New Dictionary(Of Double, Integer) From {{0, 0}, {10, 1}, {30, 2}}},
            {"SAVAGE_CUB", New Dictionary(Of Double, Integer) From {{0, 0}, {15, 1}, {42, 2}}},
            {"XCUB", New Dictionary(Of Double, Integer) From {{0, 0}, {15, 1}, {23, 2}, {46, 3}}},
            {"BC350I", New Dictionary(Of Double, Integer) From {{0, 0}, {20, 1}, {40, 2}}},
            {"B747_400", New Dictionary(Of Double, Integer) From {{0, 0}, {1, 1}, {5, 2}, {10, 3}, {20, 4}, {25, 5}, {30, 6}}},
            {"B747_8i", New Dictionary(Of Double, Integer) From {{0, 0}, {1, 1}, {5, 2}, {10, 3}, {20, 4}, {25, 5}, {30, 6}}},
            {"B787", New Dictionary(Of Double, Integer) From {{0, 0}, {1, 1}, {5, 2}, {10, 3}, {15, 4}, {17, 5}, {18, 6}, {20, 7}, {25, 8}, {30, 9}}},
            {"CRJ700", New Dictionary(Of Double, Integer) From {{0, 0}, {10, 1}, {15, 2}, {20, 3}, {35, 4}}},
            {"G36", New Dictionary(Of Double, Integer) From {{0, 0}, {12, 1}, {30, 2}}},
            {"C152", New Dictionary(Of Double, Integer) From {{0, 0}, {10, 1}, {20, 2}, {30, 3}}},
            {"C208B", New Dictionary(Of Double, Integer) From {{0, 0}, {15, 1}, {30, 2}}},
            {"CJ4", New Dictionary(Of Double, Integer) From {{0, 0}, {15, 1}, {35, 2}}},
            {"C_LONGITUDE", New Dictionary(Of Double, Integer) From {{0, 0}, {7, 1}, {15, 2}, {35, 3}}},
            {"C172", New Dictionary(Of Double, Integer) From {{0, 0}, {10, 1}, {20, 2}, {30, 3}}},
            {"DA40", New Dictionary(Of Double, Integer) From {{0, 0}, {20, 1}, {42, 2}}},
            {"DA62", New Dictionary(Of Double, Integer) From {{0, 0}, {20, 1}, {42, 2}}},
            {"DR400", New Dictionary(Of Double, Integer) From {{0, 0}, {15, 1}, {60, 2}}},
            {"DV20", New Dictionary(Of Double, Integer) From {{0, 0}, {15, 1}, {40.5, 2}}},
            {"E330", New Dictionary(Of Double, Integer) From {{0, 0}}},
            {"FDCT", New Dictionary(Of Double, Integer) From {{-12, 0}, {0, 1}, {15, 2}, {30, 3}, {35, 4}}},
            {"IA5", New Dictionary(Of Double, Integer) From {{0, 0}, {15, 1}, {30, 2}}},
            {"C10C", New Dictionary(Of Double, Integer) From {{0, 0}, {15, 1}, {40, 2}}},
            {"PAE", New Dictionary(Of Double, Integer) From {{-5, 0}, {0, 1}, {9, 2}, {20, 3}}},
            {"PITTS", New Dictionary(Of Double, Integer) From {{0, 0}}},
            {"SAVAGE_CARBON", New Dictionary(Of Double, Integer) From {{0, 0}, {15, 1}, {35, 2}}},
            {"SAVAGE_SHOCK_ULTRA", New Dictionary(Of Double, Integer) From {{0, 0}, {15, 1}, {42, 2}}},
            {"SR22", New Dictionary(Of Double, Integer) From {{0, 0}, {16, 1}, {32, 2}}},
            {"TBM930", New Dictionary(Of Double, Integer) From {{0, 0}, {10, 1}, {34, 2}}},
            {"VL3", New Dictionary(Of Double, Integer) From {{0, 0}, {15, 1}, {37, 2}, {55, 3}}}
        }

End Class

Public Class SimConnectLib

    Private debugFlag As Boolean = False

    Private Class simWindow
        Inherits Form

        Public Sub New(ByVal pSim As SimConnectLib)

            sim = pSim
            Me.ShowIcon = False
            Me.ShowInTaskbar = False

        End Sub

        Protected Overrides Sub WndProc(ByRef m As Message)

            If sim IsNot Nothing Then
                If sim.Connected Then
                    Try
                        sim.WndProcHandler(m)
                    Catch ex As Exception
                        Debug.WriteLine("WndProc exception: " & ex.Message)
                        sim.Enable = False
                    End Try
                End If
            End If

            MyBase.WndProc(m)

        End Sub

        Private sim As SimConnectLib = Nothing

    End Class

#Region "Public Declarations"

    Public Structure Coordinates
        Public Latitude As Double
        Public Longitude As Double
    End Structure

    Public Property Enable As Boolean
        Get
            Return _Enabled
        End Get
        Set(value As Boolean)
            _Enabled = value
            If _Enabled Then
                connect()
            Else
                disconnect()
            End If
        End Set
    End Property

    Public Property UpdateInterval_ms As Integer = 500

    Public ReadOnly Property Connected As Boolean
        Get
            Return (sim IsNot Nothing)
        End Get
    End Property

    Public ReadOnly Property CurrentSimRate As Integer
        Get
            Return GetSimValInt(SimData.SIM_RATE)
        End Get
    End Property

    Public ReadOnly Property RegisteredSimVals As Dictionary(Of String, Double)
        Get
            Dim t As New Dictionary(Of String, Double)
            For Each i In regSimVars.Keys
                t.Add(regSimVars(i).Name, simValsDouble(i))
            Next
            Return t
        End Get
    End Property

#End Region

#Region "Private Declarations"

    Private Const SIMCONNECT_OBJECT_ID_USER = 0
    Private Const HIDDEN_X As Integer = -16000

    Private Const SIM_RATE_MIN As Integer = 1
    Private Const SIM_RATE_MAX As Integer = 8

    ' User-defined win32 event => put basically any number?
    Private Const WM_USER_SIMCONNECTED As Integer = &H402
    Private Const WM_USER_SIMDATA As Integer = &H403

    Private Enum SYS_EVENTS
        SIMSTART = 2000
        SIMSTOP
        FOURSECS
        AIRCRAFTLOADED
    End Enum

    Private Enum DUMMYENUM
        Dummy = 0
    End Enum

    Private Enum DataDefID
        Definition1 = 1000
        Definition2
        Definition3
        Definition4
        Definition5
    End Enum

    Private Enum SimEventGroupID
        Group1
        Group2
        Group3
    End Enum

    Private Structure string1024Struct
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)> Public strVal As String
    End Structure
    Private Structure string256Struct
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)> Public strVal As String
    End Structure
    Private Structure string64Struct
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=64)> Public strVal As String
    End Structure
    Private Structure string8Struct
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=8)> Public strVal As String
    End Structure

    Private regSimVars As New Dictionary(Of Integer, SimData.SimVar)
    Private simValsDouble As New Dictionary(Of Integer, Double)
    Private simValsString As New Dictionary(Of Integer, String)
    Private simDat As New SimData
    Private sim As SimConnect = Nothing
    Private lastSimRate As Integer = 0
    Private myForm As Windows.Forms.Form = Nothing
    Private myTimer As Windows.Forms.Timer = Nothing
    Private simVarTypes As New Dictionary(Of Integer, SIMCONNECT_DATATYPE)
    Private _Enabled As Boolean = False
    Private baseSimVarsRegistered As Boolean = False
    Private aircraftDat As New AircraftData

    Private simRateStepIndex As New Dictionary(Of Integer, Integer) From
        {{1, 1},
        {2, 2},
        {4, 3},
        {8, 4},
        {16, 5},
        {32, 6},
        {64, 7}}

#End Region

#Region "Public Functions"

    Public Sub New()

        myForm = New simWindow(Me)
        myForm.Show()
        myForm.Left = -16000

        myTimer = New Windows.Forms.Timer
        AddHandler myTimer.Tick, AddressOf timerTick

        intSimVarTypes()

    End Sub

    Public Function SimVarName(ByVal simVarEnm As SimData.SimVarEnum) As String
        Return simDat.SimVarNameFromEnum(simVarEnm)
    End Function

    Public Function GetSimValBool(ByVal simValName As String) As Boolean

        'Dim d = GetSimValDouble(simValName)
        'Dim b = CBool(d)
        'Return b

        Return CBool(GetSimValDouble(simValName))
    End Function

    Public Function GetSimValInt(ByVal simValName As String) As Integer
        Return CInt(GetSimValDouble(simValName))
    End Function

    Public Function GetSimValDouble(ByVal simVarName As String) As Double

        Try
            If Not simDat.ContainsSimVar(simVarName) Then Throw New Exception("Unknown SimVar: " & simVarName)
            Dim t As SimData.SimVar = simDat.SimVarsName(simVarName)
            If t.IsString Then Throw New Exception("GetSimValDouble - SimVar is String type: " & simVarName)
            Dim ndx As Integer = t.Index
            If Not regSimVars.ContainsKey(ndx) Then
                Throw New Exception("SimVar not registered: " & simVarName)
            End If
            If Not simValsDouble.ContainsKey(ndx) Then
                Throw New Exception("Registered SimVar, data not found: " & simVarName)
            End If
            requestBaseSimVars()
            doRequestSimData(ndx)
            Return simValsDouble(ndx)
        Catch ex As Exception
            Debug.WriteLine("GetSimValDouble Exception: " & ex.Message)
            Return 0.0
        End Try

    End Function
    Public Function GetSimValDouble(ByVal simVarIndex As Integer) As Double

        Try
            If Not regSimVars.ContainsKey(simVarIndex) Then Throw New Exception("SimVar index not registered: " & simVarIndex)
            If Not simValsDouble.ContainsKey(simVarIndex) Then Throw New Exception("Registered SimVar index, data not found: " & simVarIndex)
            If regSimVars(simVarIndex).IsString Then Throw New Exception("GetSimValDouble - SimVar is String type: " & regSimVars(simVarIndex).Name)
            requestBaseSimVars()
            doRequestSimData(simVarIndex)
            Return simValsDouble(simVarIndex)
        Catch ex As Exception
            Debug.WriteLine("GetSimValDouble Exception: " & ex.Message)
            Return 0.0
        End Try

    End Function

    Public Function GetSimValDebug(ByVal simVarName As String) As Double

        debugFlag = True
        Dim v = GetSimValDouble(simVarName)
        debugFlag = False

        Return v

    End Function

    Public Function GetSimValString(ByVal simVarName As String) As String

        Try
            If Not simDat.ContainsSimVar(simVarName) Then Throw New Exception("Unknown SimVar: " & simVarName)
            Dim t As SimData.SimVar = simDat.SimVarsName(simVarName)
            If Not t.IsString Then Throw New Exception("GetSimValString - SimVar is NOT String type: " & simVarName)
            Dim ndx As Integer = t.Index
            If Not regSimVars.ContainsKey(ndx) Then Throw New Exception("SimVar not registered: " & simVarName)
            If Not simValsString.ContainsKey(ndx) Then Throw New Exception("Registered SimVar, data not found: " & simVarName)
            requestBaseSimVars()
            doRequestSimData(ndx)
            Return simValsString(ndx)
        Catch ex As Exception
            Debug.WriteLine("GetSimValString Exception: " & ex.Message)
            Return ""
        End Try

    End Function
    Public Function GetSimValString(ByVal simVarIndex As Integer) As String

        Try
            If Not regSimVars.ContainsKey(simVarIndex) Then Throw New Exception("SimVar index not registered: " & simVarIndex)
            If Not simValsString.ContainsKey(simVarIndex) Then Throw New Exception("Registered SimVar index, data not found: " & simVarIndex)
            If Not regSimVars(simVarIndex).IsString Then Throw New Exception("GetSimValString - SimVar is NOT String type: " & regSimVars(simVarIndex).Name)
            requestBaseSimVars()
            doRequestSimData(simVarIndex)
            Return simValsString(simVarIndex)
        Catch ex As Exception
            Debug.WriteLine("GetSimValString Exception: " & ex.Message)
            Return ""
        End Try

    End Function

    Public Shared Function GetSimEventByName(ByVal simEvtName As String) As SimData.SimEventEnum

        Try
            Return CType([Enum].Parse(GetType(SimData.SimEventEnum), simEvtName), SimData.SimEventEnum)
        Catch ex As Exception
            Return SimData.SimEventEnum.NONE
        End Try

    End Function

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

    Public Sub DoSlew(ByVal coords As Coordinates)
        DoSlew(coords.Latitude, coords.Longitude)
    End Sub
    Public Sub DoSlew(ByVal newLat As Double, ByVal newLong As Double)

        Dim posData As New SIMCONNECT_DATA_INITPOSITION

        posData.Altitude = GetSimValDouble(SimData.ALTITUDE)
        posData.Latitude = newLat
        posData.Longitude = newLong
        posData.Pitch = -0.0
        posData.Bank = -1.0
        posData.Heading = GetSimValDouble(SimData.HEADING)
        posData.OnGround = 0
        posData.Airspeed = CType(GetSimValDouble(SimData.AIRSPEED_INDICATED), UInteger)

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

        Dim curRateIndex As Integer = simRateStepIndex(GetSimValInt(SimData.SIM_RATE))
        Dim newRateIndex As Integer = simRateStepIndex(newSimRate)

        If curRateIndex = newRateIndex Then Exit Sub
        Debug.WriteLine("curRateIndex :" & curRateIndex)
        Debug.WriteLine("newRateIndex :" & newRateIndex)

        If curRateIndex < newRateIndex Then
            For i = curRateIndex To newRateIndex - 1
                TransmitClientEvent(SimData.SimEventEnum.SIM_RATE_INCR, "0")
                Debug.WriteLine("SIM_RATE_INCR " & Now)
            Next
        Else
            For i = curRateIndex To newRateIndex + 1 Step -1
                TransmitClientEvent(SimData.SimEventEnum.SIM_RATE_DECR, "0")
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
        If sim IsNot Nothing Then
            Try
                sim.TransmitClientEvent(0, eventID, Param, SimEventGroupID.Group1, SIMCONNECT_EVENT_FLAG.GROUPID_IS_PRIORITY)
            Catch ex As Exception
                Throw New Exception("The Transmit Request Failed: " + ex.Message)
            End Try
        End If

    End Sub

    Public Sub WndProcHandler(ByRef m As Windows.Forms.Message)

        Try
            If (m.Msg = WM_USER_SIMDATA) Then
                sim?.ReceiveMessage()
            End If
        Catch ex As Exception
            disconnect()
        End Try

    End Sub

    Public Sub RegisterCustomSimVar(ByVal name As String, ByVal units As String)
        doRegisterCustomSimVar(name, units)
    End Sub

    Public Sub RegisterSimVars(ByVal varNames As List(Of String))

        disconnect()

        'Always register the base SimVars
        registerBaseSimVars()

        For Each str As String In varNames
            If Not simDat.SimVarsName.ContainsKey(str) Then
                Throw New Exception("SimVar not known: " & str)
            End If
            Dim t As SimData.SimVar = simDat.SimVarsName(str)
            If Not regSimVars.ContainsKey(t.Index) Then
                regSimVars.Add(t.Index, t)
                If t.IsString Then
                    simValsString.Add(t.Index, "")
                Else
                    simValsDouble.Add(t.Index, 0.0)
                End If
            End If
        Next

        connect()

    End Sub

    Public Sub RegisterCommonSimVars()

        disconnect()

        'Always register the base SimVars
        registerBaseSimVars()

        registerSimVars(simDat.CommonSimVars)

        connect()

    End Sub

    Public Sub AddRegisteredSimVar(ByVal sve As SimData.SimVarEnum)
        doAddRegisteredSimVar(sve)
    End Sub

#End Region

#Region "Private Functions"

    Private Sub debugMsg(ByVal msg As String)

        If debugFlag Then
            Debug.WriteLine(Utility.GetTicks() & " " & msg)
        End If

    End Sub

    Private Sub doRegisterCustomSimVar(ByVal simVarName As String, ByVal simVarUnits As String)

        Dim t As SimData.SimVar
        If simDat.ContainsSimVar(simVarName) Then
            t = simDat.SimVarsName(simVarName)
        Else
            t = simDat.SimVars(simDat.AddCustomSimVar(simVarName, simVarUnits))
        End If

        doAddRegisteredSimVar(t.Index)

        If Connected Then
            registerSingleSimVar(t)
        End If

    End Sub

    Private Sub doAddRegisteredSimVar(ByVal sve As SimData.SimVarEnum)
        doAddRegisteredSimVarInt(sve)
    End Sub
    Private Sub doAddRegisteredSimVar(ByVal svi As Integer)
        doAddRegisteredSimVarInt(svi)
    End Sub

    Private Sub doAddRegisteredSimVarInt(ByVal svi As Integer)

        Dim t As SimData.SimVar = simDat.SimVars(svi)

        If Not regSimVars.ContainsKey(t.Index) Then
            regSimVars.Add(t.Index, t)
            If t.IsString Then
                simValsString.Add(t.Index, "")
            Else
                simValsDouble.Add(t.Index, 0.0)
            End If
        End If

    End Sub

    Private Sub registerSimVars(ByVal varList As List(Of SimData.SimVarEnum))

        For Each sve As SimData.SimVarEnum In varList
            doAddRegisteredSimVar(sve)
        Next

    End Sub

    Private Sub registerBaseSimVars()

        If Not baseSimVarsRegistered Then
            regSimVars.Clear()
            simValsDouble.Clear()
            simValsString.Clear()

            registerSimVars(simDat.BaseSimVars)

            'For Each i As SimData.SimVarEnum In simDat.BaseSimVars
            '    Dim t As SimData.SimVar = simDat.SimVars(i)
            '    regSimVars.Add(i, t)
            '    If t.IsString Then
            '        simValsString.Add(i, "")
            '    Else
            '        simValsDouble.Add(i, 0.0)
            '    End If
            'Next

            baseSimVarsRegistered = True
        End If

    End Sub

    Private Sub registerNotificationEvents()

        sim.AddClientEventToNotificationGroup(SimEventGroupID.Group1, SimData.SimEventEnum.PARKING_BRAKES, False)


        sim.SetNotificationGroupPriority(SimEventGroupID.Group1, SimConnect.SIMCONNECT_GROUP_PRIORITY_HIGHEST)

    End Sub

    Private Sub subscribeToSystemEvents()

        'Subscribe to system events
        sim.SubscribeToSystemEvent(SYS_EVENTS.FOURSECS, "4sec")
        sim.SubscribeToSystemEvent(SYS_EVENTS.SIMSTART, "SimStart")
        sim.SubscribeToSystemEvent(SYS_EVENTS.SIMSTOP, "SimStop")
        sim.SubscribeToSystemEvent(SYS_EVENTS.AIRCRAFTLOADED, "AircraftLoaded")

        'Initially turn the events off
        sim.SetSystemEventState(SYS_EVENTS.AIRCRAFTLOADED, SIMCONNECT_STATE.OFF)
        'sim.SetSystemEventState(SYS_EVENTS.SIMSTART, SIMCONNECT_STATE.OFF)
        'sim.SetSystemEventState(SYS_EVENTS.SIMSTOP, SIMCONNECT_STATE.OFF)
        sim.SetSystemEventState(SYS_EVENTS.FOURSECS, SIMCONNECT_STATE.OFF)

        'Initially turn the events on
        'sim.SetSystemEventState(SYS_EVENTS.FOURSECS, SIMCONNECT_STATE.ON)
        sim.SetSystemEventState(SYS_EVENTS.SIMSTART, SIMCONNECT_STATE.ON)
        sim.SetSystemEventState(SYS_EVENTS.SIMSTOP, SIMCONNECT_STATE.ON)
        'sim.SetSystemEventState(SYS_EVENTS.AIRCRAFTLOADED, SIMCONNECT_STATE.ON)

    End Sub

    Private Sub intSimVarTypes()

        For Each t In simDat.SimVars.Values

            Select Case t.Units.ToLower
                Case "string"
                    simVarTypes.Add(t.Index, SIMCONNECT_DATATYPE.STRINGV)
                Case "variable length string"
                    simVarTypes.Add(t.Index, SIMCONNECT_DATATYPE.STRINGV)
                Case "string128"
                    simVarTypes.Add(t.Index, SIMCONNECT_DATATYPE.STRING128)
                Case "string256"
                    simVarTypes.Add(t.Index, SIMCONNECT_DATATYPE.STRING256)
                Case "string260"
                    simVarTypes.Add(t.Index, SIMCONNECT_DATATYPE.STRING260)
                Case "string32"
                    simVarTypes.Add(t.Index, SIMCONNECT_DATATYPE.STRING32)
                Case "string64"
                    simVarTypes.Add(t.Index, SIMCONNECT_DATATYPE.STRING64)
                Case "string8"
                    simVarTypes.Add(t.Index, SIMCONNECT_DATATYPE.STRING8)
                Case Else
                    simVarTypes.Add(t.Index, SIMCONNECT_DATATYPE.FLOAT64)
            End Select
            '
        Next
    End Sub

    Private Sub doRequestSimData(ByVal simVarIndex As Integer)
        doRequestSimData(CType(simVarIndex, SimData.SimVarEnum))
    End Sub
    Private Sub doRequestSimData(ByVal simVar As SimData.SimVarEnum)
        doRequestSimData(simDat.SimVars(simVar))
    End Sub
    Private Sub doRequestSimData(ByVal simVar As SimData.SimVar)

        If Me.Connected Then
            If isCustomSimVar(simVar) Then
                debugMsg("doRequestSimData() Custom: " & simVar.Name)
                pollOneCustomSimVar(simVar)
            Else
                debugMsg("RequestDataOnSimObjectType()-" & simVar.Name)
                sim.RequestDataOnSimObjectType(CType(simVar.Index, DUMMYENUM), CType(simVar.Index, DUMMYENUM), 0, SIMCONNECT_SIMOBJECT_TYPE.USER)
            End If
        End If

    End Sub

    Private Sub pollFlapsPos()

        doRequestSimData(SimData.SimVarEnum.TRAILING_EDGE_FLAPS_LEFT_ANGLE)
        Dim fa = simValsDouble(SimData.SimVarEnum.TRAILING_EDGE_FLAPS_LEFT_ANGLE)
        Dim fi = aircraftDat.GetFlapsIndex(fa)

        If fi <> AircraftData.FLAPS_MOVING Then
            simValsDouble(SimData.SimVarEnum.CUST_FLAPS_POSITION) = fi
        End If

    End Sub

    Private Sub pollOneCustomSimVar(ByVal simVar As SimData.SimVar)

        Select Case simVar.Index
            Case SimData.SimVarEnum.CUST_VAR
                simValsDouble(simVar.Index) = 0.0
            Case SimData.SimVarEnum.CUST_FLAPS_POSITION
                pollFlapsPos()
        End Select

    End Sub

    Private Sub requestBaseSimVars()

        For Each simVar In simDat.BaseSimVars
            'Debug.WriteLine(Now.Ticks & " TimerTick: " & simVar.Index)
            doRequestSimData(simVar)
        Next

        If aircraftDat.CurrentAircraft = AircraftData.AircraftEnum.NOTSET Then
            aircraftDat.Title = simValsString(SimData.SimVarEnum.TITLE)
        End If

    End Sub

    Private Function isCustomSimVar(ByVal simVar As SimData.SimVar) As Boolean

        Select Case simVar.Index
            Case SimData.SimVarEnum.CUST_VAR
                Return True
            Case SimData.SimVarEnum.CUST_FLAPS_POSITION
                Return True
            Case Else
                Return False
        End Select

    End Function

    Private Sub timerTick(sender As Object, e As EventArgs)

        myTimer.Stop()

        If Me.Connected Then 'We are connected, Let's try to grab the data from the Sim
            ''myTimer.Interval = Me.UpdateInterval_ms
            'Try
            '    For Each simVar In regSimVars.Values
            '        If isCustomSimVar(simVar) Then
            '            pollOneCustomSimVar(simVar)
            '        Else
            '            'Debug.WriteLine(Now.Ticks & " TimerTick: " & simVar.Index)
            '            doRequestSimData(simVar.Index)
            '        End If
            '    Next
            '    myTimer.Start()
            'Catch ex As Exception
            '    disconnect()
            'End Try
        Else 'We Then are Not connected, Let's try to connect
            connect()
        End If

    End Sub

    Private Sub startTimer()

        If myTimer IsNot Nothing Then
            If myTimer.Enabled Then myTimer.Stop()
            If _Enabled Then
                If Me.Connected Then
                    myTimer.Interval = UpdateInterval_ms
                Else
                    myTimer.Interval = 2000
                End If
                Debug.WriteLine(Now.Ticks & " sim.Starting timer with " & myTimer.Interval)
                myTimer.Start()
            End If
        End If

    End Sub

    Private Sub stopTimer()

        If myTimer IsNot Nothing Then
            If myTimer.Enabled Then myTimer.Stop()
            Debug.WriteLine(Now.Ticks & " sim.Timer stopped")
        End If

    End Sub

    Private Sub disconnect()

        If Me.Connected Then
            stopTimer()
            Try
                If (sim IsNot Nothing) Then
                    'Unsubscribe from all the system events
                    sim.UnsubscribeFromSystemEvent(SYS_EVENTS.FOURSECS)
                    sim.UnsubscribeFromSystemEvent(SYS_EVENTS.SIMSTART)
                    sim.UnsubscribeFromSystemEvent(SYS_EVENTS.SIMSTOP)

                    sim.Dispose()
                    sim = Nothing
                End If
                Debug.WriteLine(Now.Ticks & " sim.Disconnected")
            Catch ex As Exception
                sim = Nothing
                Debug.WriteLine("sim.Disconnect exception: " & ex.Message)
            End Try
            startTimer()
        End If

    End Sub

    Private Sub connect()

        stopTimer()

        If _Enabled Then
            If Not Me.Connected Then
                Try
                    sim = New SimConnect("SimConnectLib", myForm.Handle, WM_USER_SIMDATA, Nothing, 0)

                    AddHandler sim.OnRecvOpen, AddressOf sim_OnRecvOpen
                    AddHandler sim.OnRecvQuit, AddressOf sim_OnRecvQuit
                    AddHandler sim.OnRecvSimobjectDataBytype, AddressOf sim_OnRecvSimobjectDataBytype

                    AddHandler sim.OnRecvException, AddressOf sim_OnRecvException

                    'listen to events
                    AddHandler sim.OnRecvEvent, AddressOf sim_OnRecvEvent

                    subscribeToSystemEvents()

                    registerBaseSimVars()

                    Debug.WriteLine(Now.Ticks & " sim.Connected")
                Catch ex As Exception
                    sim = Nothing
                    Debug.WriteLine("sim.Connect exception: " & ex.Message)
                    startTimer()
                End Try
            End If
        Else
            If Me.Connected Then
                disconnect()
            End If
        End If

    End Sub

    Private Sub doSetPosition(ByVal slewData As SIMCONNECT_DATA_INITPOSITION)

        Try
            sim.SetDataOnSimObject(DataDefID.Definition1, 0, 0, slewData)
        Catch ex As Exception
            Throw New Exception("The SetData Request Failed: " + ex.Message)
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

    Private Function checkValidSlew(ByVal coords As Coordinates) As Boolean

        If coords.Latitude = -1 Or coords.Longitude = -1 Then
            Return False
        Else
            Return True
        End If

    End Function

#Region "SimConnect Callbacks"

    Private Sub sim_OnRecvEvent(ByVal sender As SimConnect, ByVal recEvent As SIMCONNECT_RECV_EVENT)

        Select Case recEvent.uEventID
            Case CUInt(SYS_EVENTS.SIMSTART)
                aircraftDat.Title = GetSimValString(SimData.SimVarEnum.TITLE)
                'Debug.WriteLine(Now.Ticks & " Sim running")
                Debug.WriteLine(Now.Ticks & " Sim running: " & GetSimValString(SimData.SimVarEnum.TITLE))
            Case CUInt(SYS_EVENTS.SIMSTOP)
                aircraftDat.SimDisconnected()
                Debug.WriteLine(Now.Ticks & " Sim stopped")
            Case CUInt(SYS_EVENTS.FOURSECS)
                Debug.WriteLine(Now.Ticks & " Sim 4s tick")
            Case CUInt(SYS_EVENTS.AIRCRAFTLOADED)
                doRequestSimData(SimData.SimVarEnum.TITLE)
                Debug.WriteLine(Now.Ticks & " Sim Aicraft Loaded: " & GetSimValString(SimData.SimVarEnum.TITLE))
            Case CUInt(SimData.SimEventEnum.PARKING_BRAKES)
                'doRequestSimData(SimData.SimVarEnum.BRAKE_PARKING_INDICATOR)
                'Debug.WriteLine(Now.Ticks & " Parking Brakes toggle: " & GetSimValString(SimData.SimVarEnum.BRAKE_PARKING_INDICATOR))
                Debug.WriteLine(Now.Ticks & " Parking Brakes toggle")
        End Select

    End Sub

    Private Sub sim_OnRecvException(ByVal sender As SimConnect, ByVal data As SIMCONNECT_RECV_EXCEPTION)
        'Debug.WriteLine("sim_OnRecvException: " & data.dwException)
    End Sub

    ' <summary>
    ' We received a disconnection from SimConnect
    ' </summary>
    ' <param name="sender"></param>
    ' <param name="data"></param>
    Private Sub sim_OnRecvQuit(ByVal sender As SimConnect, ByVal data As SIMCONNECT_RECV)

        Debug.WriteLine(Now.Ticks & " sim_OnRecvQuit")
        disconnect()

    End Sub

    Private Sub registerSingleSimVar(ByVal simv As SimData.SimVar)

        If simv.IsString Then
            ' Define a data structure
            sim.AddToDataDefinition(CType(simv.Index, DUMMYENUM), simv.Name, "number", simVarTypes(simv.Index), 0, SimConnect.SIMCONNECT_UNUSED)

            ' IMPORTANT: Register it With the simconnect managed wrapper marshaller
            ' If you skip this step, you will only receive a uint in the .dwData field.
            Select Case simVarTypes(simv.Index)
                Case SIMCONNECT_DATATYPE.STRING8
                    sim.RegisterDataDefineStruct(Of string1024Struct)(CType(simv.Index, DUMMYENUM))
                Case SIMCONNECT_DATATYPE.STRING64
                    sim.RegisterDataDefineStruct(Of string64Struct)(CType(simv.Index, DUMMYENUM))
                Case SIMCONNECT_DATATYPE.STRING256
                    sim.RegisterDataDefineStruct(Of string256Struct)(CType(simv.Index, DUMMYENUM))
                Case Else
                    sim.RegisterDataDefineStruct(Of string1024Struct)(CType(simv.Index, DUMMYENUM))
            End Select
        Else
            sim.AddToDataDefinition(CType(simv.Index, DUMMYENUM), simv.Name, simv.Units, simVarTypes(simv.Index), 0.0, SimConnect.SIMCONNECT_UNUSED)
            sim.RegisterDataDefineStruct(Of Double)(CType(simv.Index, DUMMYENUM))
        End If

    End Sub

    ' <summary>
    ' We received a connection from SimConnect.
    ' Let's register all the properties we need.
    ' </summary>
    ' <param name="sender"></param>
    ' <param name="data"></param>
    Private Sub sim_OnRecvOpen(ByVal sender As SimConnect, ByVal data As SIMCONNECT_RECV_OPEN)

        Try
            For Each t As SimData.SimVar In simDat.SimVars.Values
                registerSingleSimVar(t)
            Next

            For Each item As SimData.SimEventEnum In [Enum].GetValues(GetType(SimData.SimEventEnum))
                If item <> SimData.SimEventEnum.NONE Then
                    sim.MapClientEventToSimEvent(item, item.ToString().Replace("MobiFlight_", "MobiFlight."))
                End If
            Next

            registerNotificationEvents()

            sim.AddToDataDefinition(DataDefID.Definition1, "Initial Position", "", SIMCONNECT_DATATYPE.INITPOSITION, 0, 0)
        Catch ex As Exception
            Debug.WriteLine("sim_OnRecvOpen.Exception: " & ex.Message)
        End Try

    End Sub

    ' <summary>
    ' Received data from SimConnect
    ' </summary>
    ' <param name="sender"></param>
    ' <param name="data"></param>
    Private Sub sim_OnRecvSimobjectDataBytype(ByVal sender As SimConnect, ByVal data As SIMCONNECT_RECV_SIMOBJECT_DATA_BYTYPE)

        Dim iRequest As Integer = -1

        Try
            iRequest = CInt(data.dwRequestID)
            'Debug.WriteLine(Now.Ticks & " Start: " & iRequest)
            Dim sv = simDat.SimVars(CType(iRequest, SimData.SimVarEnum))
            debugMsg("OnRecvSimobjectDataBytype()-" & sv.Name)
            Dim dt = simVarTypes(iRequest)
            If sv.IsString Then
                Dim sValue As String
                Select Case dt
                    Case SIMCONNECT_DATATYPE.STRING8
                        Dim t As string8Struct = CType(data.dwData(0), string8Struct)
                        sValue = t.strVal
                    Case SIMCONNECT_DATATYPE.STRING64
                        Dim t As string64Struct = CType(data.dwData(0), string64Struct)
                        sValue = t.strVal
                    Case SIMCONNECT_DATATYPE.STRING256
                        Dim t As string256Struct = CType(data.dwData(0), string256Struct)
                        sValue = t.strVal
                    Case Else
                        Dim t As string1024Struct = CType(data.dwData(0), string1024Struct)
                        sValue = t.strVal
                End Select

                If Not simValsString.ContainsKey(iRequest) Then simValsString.Add(iRequest, "")
                simValsString(iRequest) = sValue
                'Debug.WriteLine(Now.Ticks & " sim_OnRecvSimobjectDataBytype.Received String: " & iRequest)
            Else
                Dim dValue As Double = CDbl(data.dwData(0))
                'If Not simVals.ContainsKey(iRequest) Then Throw New Exception("Registered SimVar index not found: " & iRequest)
                If Not simValsDouble.ContainsKey(iRequest) Then simValsDouble.Add(iRequest, 0.0)
                simValsDouble(iRequest) = dValue
                'Debug.WriteLine(Now.Ticks & " sim_OnRecvSimobjectDataBytype.Received Double: " & iRequest)
            End If
        Catch ex As Exception
            Debug.WriteLine("sim_OnRecvSimobjectDataBytype.Exception: " & ex.Message)
        End Try

        'Debug.WriteLine(Now.Ticks & " Finish: " & iRequest)

    End Sub

#End Region

#End Region

End Class

