Option Explicit On
Option Strict On

Imports System.Threading
Imports HidSharp
Imports HidSharp.Reports
Imports HidSharp.Reports.Encodings
Imports System.Xml
Imports System.ComponentModel

Public Class JoystickEventMap

    Public Structure MappingDataStruct
        Public GUID As String
        Public SimEvent As String
        Public LongPushSimEvent As String
        Public ReleaseSimEvent As String
        Public SimVar As String
        Public ActiveState As Boolean
        Public LongPushActive As Boolean
        Public ReleaseActive As Boolean
        Public Acceleration As Integer
        Public Delay As Integer
        Public EnableEventCount As Integer
        Public MapIsEnabled As Boolean
        Public LastEventTime As Date
    End Structure

    Public Shared ReadOnly Property EmptyMappingData() As MappingDataStruct
        Get
            Dim md As New MappingDataStruct

            md.GUID = ""
            md.SimEvent = ""
            md.SimVar = ""
            md.ActiveState = False
            md.MapIsEnabled = False
            md.Acceleration = 0
            md.Delay = 0
            md.EnableEventCount = 0
            md.LastEventTime = Date.MinValue

            Return md
        End Get
    End Property

    Public ReadOnly Property Device As Joystick
        Get
            Return myJoyMaps.Device
        End Get
    End Property

    Public ReadOnly Property Key As String
        Get
            Return _Key
        End Get
    End Property

    Public ReadOnly Property MappingData As MappingDataStruct
        Get
            Return getMapData()
        End Get
    End Property

    Public ReadOnly Property JoyEvent As String
        Get
            Return _JoyEvent
        End Get
    End Property

    Public ReadOnly Property Acceleration As Integer
        Get
            Return _Accel
        End Get
    End Property

    Public ReadOnly Property Delay As Integer
        Get
            Return _Delay
        End Get
    End Property

    Public ReadOnly Property SimEvent As String
        Get
            Return _SimEvent
        End Get
    End Property

    Public ReadOnly Property LongPushSimEvent As String
        Get
            Return _LongPushSimEvent
        End Get
    End Property

    Public ReadOnly Property ReleaseSimEvent As String
        Get
            Return _ReleaseSimEvent
        End Get
    End Property

    Public ReadOnly Property SimVar As String
        Get
            Return _SimVar
        End Get
    End Property

    Public ReadOnly Property EnableEventCount As Integer
        Get
            Return _EnableEvents.Count
        End Get
    End Property

    Public ReadOnly Property ActiveState As Boolean
        Get
            Return _ActiveState
        End Get
    End Property

    Public ReadOnly Property LongPushActive As Boolean
        Get
            Return _LongPushActive
        End Get
    End Property

    Public Sub New(ByVal elem As XmlElement, ByVal joyMaps As JoystickMapping)

        Dim jeEvents As String = elem.GetAttribute("JoyEnableEvents")
        myJoyMaps = joyMaps
        _JoyEvent = elem.GetAttribute("JoyEvent")
        _SimEvent = elem.GetAttribute("SimEvent")
        _LongPushSimEvent = elem.GetAttribute("LongPushSimEvent")
        _ReleaseSimEvent = elem.GetAttribute("ReleaseSimEvent")
        _SimVar = elem.GetAttribute("SimVar")
        _Accel = getIntAttr(elem.GetAttribute("Accel"))
        _Delay = getIntAttr(elem.GetAttribute("Delay"))
        _Key = _JoyEvent & jeEvents

        Dim t = elem.GetAttribute("ActiveState")
        If t = "" Then t = "False"
        _ActiveState = CBool(t)

        If jeEvents <> "" Then
            Dim ee As String() = jeEvents.Split(","c)
            For Each t In ee
                _EnableEvents.Add(t)
            Next
        End If

        '<JoyMap  = "Button31" ="Button8" ="AP_VS_VAR_INC" />
        '<JoyMap JoyEvent = "Button34" SimEvent="TOGGLE_NAV_LIGHTS" ="LIGHT_NAV" ="False" />

    End Sub

    Private Function getMapData() As MappingDataStruct

        Dim md = EmptyMappingData
        md.GUID = _GUID

        If mapIsEnabled() Then
            md.SimEvent = SimEvent
            md.LongPushSimEvent = LongPushSimEvent
            md.ReleaseSimEvent = ReleaseSimEvent
            md.SimVar = SimVar
            md.ActiveState = ActiveState
            md.Acceleration = _Accel
            md.Delay = _Delay
            md.MapIsEnabled = True
            md.EnableEventCount = _EnableEvents.Count
            md.LongPushActive = _LongPushActive
            md.ReleaseActive = _ReleaseActive
        End If

        Return md

    End Function

    Private Function checkLongPush(ByVal joyEventIsTrue As Boolean) As Boolean

        If joyEventIsTrue Then
            'Start the long press timer and do nothing else
            eventTrueTime = Now
            Return False
        Else
            If Now > eventTrueTime.AddMilliseconds(myJoyMaps.LongPressTimeout) Then
                'If the long press timer has expired make it a long press event
                _LongPushActive = True
            Else
                'Else it's a normal press event
                _LongPushActive = False
            End If
            Return True
        End If

    End Function

    Private Function mapIsEnabled() As Boolean

        Dim joyEventIsTrue = myJoyMaps.Device.Data.GetValueBool(JoyEvent)

        'First we check that all required enable events are True
        If _EnableEvents.Count > 0 Then
            For Each evt In _EnableEvents
                If Not myJoyMaps.Device.Data.GetValueBool(evt) Then
                    'One of the enable events is not pressed so bad luck, it's not enabled
                    Return False
                End If
            Next
        End If
        'if we get to the end of the enable event list then they are all pressed and bingo,
        'it's enabled so we can continue processing the event.

        If joyEventIsTrue Then
            'Maps are always enabled on button press events unless they are long push or Specified as a press and release event.
            _ReleaseActive = False
            If _LongPushSimEvent <> "" Then
                'Starts the long push timer.
                Return checkLongPush(True)
            Else
                Return True
            End If
        Else
            If _LongPushSimEvent <> "" Then
                'Enables either a long push or normal event depending on whether the long push timer has expired.
                _ReleaseActive = False
                Return checkLongPush(False)
            ElseIf ReleaseSimEvent <> "" Then
                'Enables a release event as specified.
                _ReleaseActive = True
                Return True
            Else
                Return False
            End If
        End If

    End Function

    Private Function getIntAttr(ByVal attrVal As String) As Integer

        If attrVal = "" Then
            Return 0
        Else
            If IsNumeric(attrVal) Then
                Return CInt(attrVal)
            Else
                Return 0
            End If
        End If

    End Function

    Private _Key As String = ""
    Private _JoyEvent As String = ""
    Private _SimEvent As String = ""
    Private _LongPushSimEvent As String = ""
    Private _ReleaseSimEvent As String = ""
    Private _SimVar As String = ""
    Private _Accel As Integer = 0
    Private _Delay As Integer = 0
    Private _ActiveState As Boolean = False
    Private _LongPushActive As Boolean = False
    Private _ReleaseActive As Boolean = False
    Private _EnableEvents As New List(Of String)
    Private myJoyMaps As JoystickMapping = Nothing
    Private _GUID As String = Guid.NewGuid.ToString
    Private eventTrueTime As Date = Date.MinValue

End Class

Public Class JoystickEvent

    Public ReadOnly Property ActiveMaps As List(Of JoystickEventMap.MappingDataStruct)
        Get
            Dim l As New List(Of JoystickEventMap.MappingDataStruct)
            For Each m In myJoyMaps.Values
                If m.SimVar <> "" Then
                    Dim md = m.MappingData
                    If md.MapIsEnabled Then
                        l.Add(md)
                    End If
                End If
            Next
            Return l
        End Get
    End Property

    Public ReadOnly Property MappedSimVars As List(Of String)
        Get
            Dim l As New List(Of String)
            For Each m In myJoyMaps.Values
                If m.SimVar <> "" Then
                    If Not l.Contains(m.SimVar) Then l.Add(m.SimVar)
                End If
            Next
            Return l
        End Get
    End Property

    Public ReadOnly Property Name As String
        Get
            Return _Name
        End Get
    End Property

    Public ReadOnly Property Device As Joystick
        Get
            Return myJoyMapping.Device
        End Get
    End Property

    Public ReadOnly Property MappingData() As JoystickEventMap.MappingDataStruct
        Get
            Dim md = JoystickEventMap.EmptyMappingData

            For Each jem As JoystickEventMap In myJoyMaps.Values
                Dim temp = jem.MappingData
                If temp.MapIsEnabled Then
                    If md.MapIsEnabled Then
                        'If the joystick event has multiple mappings then the one with the highest number of enable events wins
                        If temp.EnableEventCount > md.EnableEventCount Then
                            md = temp
                        End If
                    Else
                        md = temp
                    End If
                End If
            Next

            Return md
        End Get
    End Property

    Public Sub AddMap(ByVal elem As XmlElement)

        Dim jm As New JoystickEventMap(elem, myJoyMapping)
        myJoyMaps.Add(jm.Key, jm)

    End Sub

    Public Sub New(ByVal elem As XmlElement, ByVal joyMaps As JoystickMapping)

        myJoyMapping = joyMaps
        _Name = elem.GetAttribute("JoyEvent")
        AddMap(elem)

    End Sub

    Private _Name As String = ""
    Private myJoyMapping As JoystickMapping = Nothing
    Private myJoyMaps As New Dictionary(Of String, JoystickEventMap)

End Class

Public Class JoystickMapping

    Public ReadOnly Property ActiveMaps As List(Of JoystickEventMap.MappingDataStruct)
        Get
            Dim l As New List(Of JoystickEventMap.MappingDataStruct)
            For Each je In myEvents.Values
                For Each mds In je.ActiveMaps
                    l.Add(mds)
                Next
            Next
            Return l
        End Get
    End Property

    Public ReadOnly Property MappedSimVars As List(Of String)
        Get
            Dim l As New List(Of String)
            For Each je In myEvents.Values
                For Each m In je.MappedSimVars
                    If Not l.Contains(m) Then l.Add(m)
                Next
            Next
            Return l
        End Get
    End Property

    Public Property Device As Joystick
        Get
            Return myJoy
        End Get
        Set(value As Joystick)
            myJoy = value
        End Set
    End Property

    Public Property LongPressTimeout As Integer = 200

    Public ReadOnly Property Name As String
        Get
            Return _Name
        End Get
    End Property

    Public ReadOnly Property JoystickName As String
        Get
            Return _JoyName
        End Get
    End Property

    Public Sub EventReceived(ByVal eventName As String, ByVal evenValue As Double)

    End Sub

    Public Function HasEvent(ByVal eventName As String) As Boolean
        Return myEvents.ContainsKey(eventName)
    End Function

    Public Function GetEvent(ByVal eventName As String) As JoystickEvent
        Return myEvents(eventName)
    End Function

    Public Function GetMapData(ByVal eventName As String) As JoystickEventMap.MappingDataStruct

        If HasEvent(eventName) Then
            Return GetEvent(eventName).MappingData
        Else
            Return JoystickEventMap.EmptyMappingData
        End If

    End Function

    Public Sub Open()

        If myJoy Is Nothing Then Throw New Exception("Joystick device not set for " & Name)
        If myJoy.Busy Then Throw New Exception("Joystick device already open: " & myJoy.DisplayName)
        myJoy.Mappings = Me
        myJoy.OpenDevice()

    End Sub

    Public Sub Close()

        If myJoy IsNot Nothing Then
            Debug.WriteLine("Closing " & myJoy.DisplayName & "...")
            myJoy.CloseDevice()
            Debug.WriteLine("Closed")
        End If

    End Sub

    Public Sub New(ByVal elem As XmlElement)
        doNew(elem)
    End Sub

    Private Sub doNew(ByVal elem As XmlElement)

        _Name = elem.GetAttribute("MappingName")
        _JoyName = elem.GetAttribute("Name")
        Dim lpt = elem.GetAttribute("LongPressTimeout")
        If lpt <> "" And IsNumeric(lpt) Then
            Me.LongPressTimeout = CInt(lpt)
        End If

        For Each el As XmlElement In elem.GetElementsByTagName("JoyMap")
            Dim evtName = el.GetAttribute("JoyEvent")
            If myEvents.ContainsKey(evtName) Then
                myEvents(evtName).AddMap(el)
            Else
                Dim je As New JoystickEvent(el, Me)
                myEvents.Add(je.Name, je)
            End If
        Next

    End Sub

    Private _Name As String = ""
    Private _JoyName As String = ""
    Private myJoy As Joystick = Nothing
    Private myEvents As New Dictionary(Of String, JoystickEvent)

End Class

Public Class JoystickMappings

    Public ReadOnly Property CurrentlyOpenProfile As String
        Get
            Return _CurrentlyOpenProfile
        End Get
    End Property

    Public ReadOnly Property ActiveMaps As List(Of JoystickEventMap.MappingDataStruct)
        Get
            Dim l As New List(Of JoystickEventMap.MappingDataStruct)
            If _CurrentProfile IsNot Nothing Then
                For Each name In _CurrentProfile
                    Dim m = GetMapping(name)
                    'If m Is Nothing Then Throw New Exception("Joystick Mapping " & name & " does not exist")
                    If m IsNot Nothing Then
                        For Each mds In m.ActiveMaps
                            l.Add(mds)
                        Next
                    End If
                Next
            End If
            Return l
        End Get
    End Property

    Public ReadOnly Property MappedSimVars As List(Of String)
        Get
            Dim l As New List(Of String)

            For Each jm In myMappings.Values
                For Each m In jm.MappedSimVars
                    If Not l.Contains(m) Then l.Add(m)
                Next
            Next

            Return l
        End Get
    End Property

    Public ReadOnly Property MappingNames As List(Of String)
        Get
            Dim l As New List(Of String)
            For Each t In myMappings.Keys
                l.Add(t)
            Next
            Return l
        End Get
    End Property

    Public Function GetProfileJoystickNames(ByVal profileName As String) As List(Of String)

        Dim joyMaps As List(Of String) = GetProfile(profileName)
        Dim l As New List(Of String)
        For Each s In joyMaps
            Dim j = GetMapping(s)
            If Not l.Contains(j.JoystickName) Then
                l.Add(j.JoystickName)
            End If
        Next
        Return l

    End Function

    Public ReadOnly Property ProfileNames As List(Of String)
        Get
            Dim l As New List(Of String)
            For Each t In myProfiles.Keys
                l.Add(t)
            Next
            Return l
        End Get
    End Property

    Public Function GetMapping(ByVal mapName As String) As JoystickMapping

        If myMappings.ContainsKey(mapName) Then
            Return myMappings(mapName)
        Else
            Return Nothing
        End If

    End Function

    Public Function GetProfile(ByVal profileName As String) As List(Of String)

        If myProfiles.ContainsKey(profileName) Then
            Return myProfiles(profileName)
        Else
            Return Nothing
        End If

    End Function

    Public Sub Open(ByVal profileName As String)

        Dim l = GetProfile(profileName)
        If l Is Nothing Then Throw New Exception("Profile " & profileName & " does not exist")

        For Each name In l
            Dim m = GetMapping(name)
            If m Is Nothing Then
                'Throw New Exception("Joystick Mapping " & name & " does not exist")
                Debug.WriteLine("Joystick Mapping Open() - Device " & name & " is not available")
            Else
                m.Open()
            End If
        Next

        _CurrentProfile = l
        _CurrentlyOpenProfile = profileName

    End Sub

    Public Sub Close()

        'If currentlyOpenProfile Is Nothing Then Throw New Exception("No joystick profile is open")
        If _CurrentProfile Is Nothing Then Exit Sub

        For Each name In _CurrentProfile
            Dim m = GetMapping(name)
            If m Is Nothing Then
                Debug.WriteLine("Joystick Mapping Close() - Device " & name & " is not available")
            Else
                m.Close()
            End If
        Next

        _CurrentlyOpenProfile = ""

    End Sub

    Public Sub New(ByVal eventCallback As Joystick.DeviceReportCallbackDelegate)

        Dim joysticks = Joystick.GetJoysticksDictionary()

        Dim appDir As String = IO.Path.GetDirectoryName(Application.ExecutablePath)
        Dim dom As New XmlDocument
        dom.Load(IO.Path.Combine(appDir, "JoystickMappings.xml"))

        For Each el As XmlElement In dom.GetElementsByTagName("Joystick")
            Dim jm = New JoystickMapping(el)
            If joysticks.ContainsKey(jm.JoystickName) Then
                If Not myJoysticks.ContainsKey(jm.JoystickName) Then myJoysticks.Add(jm.JoystickName, joysticks(jm.JoystickName))
                jm.Device = joysticks(jm.JoystickName)
                jm.Device.DeviceReportCallback = eventCallback
                jm.Device.Mappings = jm
                myMappings.Add(jm.Name, jm)
            End If
        Next

        For Each el As XmlElement In dom.GetElementsByTagName("Profile")
            Dim l As New List(Of String)
            For Each el2 As XmlElement In el.GetElementsByTagName("JoyMapping")
                l.Add(el2.GetAttribute("Name"))
            Next
            myProfiles.Add(el.GetAttribute("Name"), l)
        Next

        Try
        Catch ex As Exception
            Debug.WriteLine("JoystickMappings exception: " & ex.Message)
        End Try

    End Sub

    Private myMappings As New Dictionary(Of String, JoystickMapping)
    Private myProfiles As New Dictionary(Of String, List(Of String))
    Private myJoysticks As New Dictionary(Of String, Joystick)
    Private _CurrentProfile As List(Of String) = Nothing
    Private _CurrentlyOpenProfile As String = ""

End Class

Public Class JoystickValue

    Public Property Name As String = ""
    Public Property Index As Integer = -1
    Public Property Value As Double = 0
    Public Property PreviousValue As Double = 0

    Public ReadOnly Property IsAxis As Boolean
        Get
            If Name.ToLower = "genericdesktopx" Then Return True
            If Name.ToLower = "genericdesktopy" Then Return True
            If Name.ToLower = "genericdesktopz" Then Return True
            If Name.ToLower = "genericdesktoprx" Then Return True
            If Name.ToLower = "genericdesktopry" Then Return True
            If Name.ToLower = "genericdesktoprz" Then Return True
            If Name.ToLower = "genericdesktopslider" Then Return True
            If Name.ToLower = "genericdesktopdial" Then Return True
            'If Name.ToLower = "" Then Return True
            'If Name.ToLower.StartsWith("") Then Return True

            Return False
        End Get
    End Property

    Public Sub New()

    End Sub
    Public Sub New(ByVal iIndex As Integer, ByVal strName As String, ByVal dValue As Double, ByVal dPreviousValue As Double)

        Name = strName
        Index = iIndex
        Value = dValue
        PreviousValue = dPreviousValue
    End Sub

End Class

Public Class JoystickData

    Public ReadOnly Property JoystickName As String
        Get
            Return myJoy.DisplayName
        End Get
    End Property

    Public ReadOnly Property Device As Joystick
        Get
            Return myJoy
        End Get
    End Property

    Public ReadOnly Property Values As List(Of JoystickValue)
        Get
            Dim l As New List(Of JoystickValue)
            For Each i In datIndexes.Keys
                l.Add(datIndexes(i))
            Next
            Return l
        End Get
    End Property

    Public Function GetValue(ByVal strName As String) As Double

        If datNames.ContainsKey(strName) Then
            Return datNames(strName).Value
        Else
            Return 0
        End If

    End Function
    Public Function GetValue(ByVal iIndex As Integer) As Double

        If datIndexes.ContainsKey(iIndex) Then
            Return datIndexes(iIndex).Value
        Else
            Return 0
        End If

    End Function

    Public Function GetValueBool(ByVal strName As String) As Boolean

        Dim t = GetValue(strName)
        Dim b = CBool(t)
        Return b
        'Return CBool(GetValue(strName))

    End Function
    Public Function GetValueBool(ByVal iIndex As Integer) As Boolean
        Return CBool(GetValue(iIndex))
    End Function

    Public Sub SetValue(ByVal iIndex As Integer, strName As String, ByVal dValue As Double, ByVal dPreviousValue As Double)

        If datIndexes.ContainsKey(iIndex) Then
            datIndexes(iIndex).Value = dValue
        Else
            Dim v = New JoystickValue(iIndex, strName, dValue, dPreviousValue)
            datIndexes.Add(iIndex, v)
            datNames.Add(strName, v)
        End If

    End Sub
    Public Sub SetValue(ByVal val As JoystickValue)

        If datIndexes.ContainsKey(val.Index) Then
            datIndexes(val.Index).Value = val.Value
        Else
            datIndexes.Add(val.Index, val)
            datNames.Add(val.Name, val)
        End If

    End Sub

    Public Sub Clear()

        datIndexes.Clear()
        datNames.Clear()

    End Sub

    Public Sub New(ByVal joy As Joystick)
        myJoy = joy
    End Sub

    Private datIndexes As New Dictionary(Of Integer, JoystickValue)
    Private datNames As New Dictionary(Of String, JoystickValue)
    Private myJoy As Joystick = Nothing

End Class

Public Class Joystick

    Public Delegate Sub DeviceReportCallbackDelegate(ByVal reportData As JoystickData)

    Public Property CancelFlag As Boolean = False

    Public Property Mappings As JoystickMapping = Nothing

    Public ReadOnly Property Data As JoystickData
        Get
            Return myJoyData
        End Get
    End Property

    Public ReadOnly Property Worker As BackgroundWorker
        Get
            Return myBW
        End Get
    End Property

    Public ReadOnly Property Busy As Boolean
        Get
            Return _Busy
        End Get
    End Property

    Public WriteOnly Property DeviceReportCallback As DeviceReportCallbackDelegate
        Set(value As DeviceReportCallbackDelegate)
            _reportCallback = value
        End Set
    End Property

    Public ReadOnly Property IsJoystickOrGamepad As Boolean
        Get
            Return (myHID IsNot Nothing And myReportDescriptor IsNot Nothing And myDeviceItem IsNot Nothing)
        End Get
    End Property

    Public ReadOnly Property DisplayName() As String
        Get
            Try
                Dim t As String = myHID.GetFriendlyName
                If t = "" Then
                    t = myHID.ToString
                End If
                Return t
            Catch ex As Exception
                Try
                    Return myHID.ToString
                Catch ex2 As Exception
                    Return ex2.Message
                End Try
            End Try
        End Get
    End Property

    Public Shared Function GetJoysticks() As List(Of Joystick)

        Dim tempList As New List(Of Joystick)
        Dim hidDevList = getHID_Devlist()
        For Each dev In hidDevList
            Dim tempJoy = New Joystick(dev)
            If tempJoy.IsJoystickOrGamepad Then
                tempList.Add(tempJoy)
            End If
        Next
        Return tempList

    End Function

    Public Shared Function GetJoysticksDictionary() As Dictionary(Of String, Joystick)

        Dim d As HidDevice

        Dim tempList As New Dictionary(Of String, Joystick)
        Dim hidDevList = getHID_Devlist()
        For Each dev In hidDevList
            Debug.WriteLine(dev.ToString)
            Dim tempJoy = New Joystick(dev)
            If tempJoy.IsJoystickOrGamepad Then
                If tempList.ContainsKey(tempJoy.DisplayName) Then
                    Debug.WriteLine("Duplicate HID Device: " & tempJoy.DisplayName)
                Else
                    tempList.Add(tempJoy.DisplayName, tempJoy)
                End If
            End If
        Next
        Return tempList

    End Function

    Private Shared Function getHID_Devlist() As HidDevice()

        Dim devList = DeviceList.Local
        AddHandler devList.Changed, AddressOf devListChanged
        Return devList.GetHidDevices().ToArray()

    End Function

    Private Shared Function writeDeviceItemInputParserResult(ByVal parser As Reports.Input.DeviceItemInputParser, ByVal joy As Joystick) As String

        Dim t As String = ""

        While (parser.HasChanged)
            Dim changedIndex As Integer = parser.GetNextChangedIndex()
            Dim previousDataValue = parser.GetPreviousValue(changedIndex)
            Dim dataValue = parser.GetValue(changedIndex)

            Dim tt As String = Now & " " & String.Format("  {0}: {1} -> {2}", CType(dataValue.Usages.FirstOrDefault(), Usage),
                                                  previousDataValue.GetPhysicalValue(), dataValue.GetPhysicalValue())
            t &= tt & vbNewLine
        End While

        Return t

    End Function

    Private Shared Function writeDeviceItemInputParserResultData(ByVal parser As Reports.Input.DeviceItemInputParser, ByVal joy As Joystick) As JoystickData

        Dim jd As New JoystickData(joy)

        While (parser.HasChanged)
            Dim changedIndex As Integer = parser.GetNextChangedIndex()
            Dim previousDataValue = parser.GetPreviousValue(changedIndex)
            Dim dataValue = parser.GetValue(changedIndex)

            jd.SetValue(changedIndex, CType(dataValue.Usages.FirstOrDefault(), Usage).ToString, dataValue.GetPhysicalValue(), previousDataValue.GetPhysicalValue)
        End While

        For Each d In jd.Values
            joy.Data.SetValue(d)
        Next

        Return jd

    End Function

    Private Shared Function writeAllDeviceItemInputParserResults(ByVal parser As Reports.Input.DeviceItemInputParser, ByVal joy As Joystick) As String

        Dim t As String = Now & " New HID report" & vbNewLine

        For idx = 0 To parser.ValueCount - 1
            Dim previousDataValue = parser.GetPreviousValue(idx)
            Dim dataValue = parser.GetValue(idx)

            Dim tt As String = Now & " " & String.Format("  {0}: {1} -> {2}", CType(dataValue.Usages.FirstOrDefault(), Usage),
                                                  previousDataValue.GetPhysicalValue(), dataValue.GetPhysicalValue())
            t &= tt & vbNewLine
        Next

        Return t & vbNewLine

    End Function

    Private Shared Sub devListChanged(ByVal sender As Object, ByVal e As Object)
        Debug.WriteLine("Device list changed.")
    End Sub

    Public Sub OpenDevice()

        _Busy = True
        CancelFlag = False
        myJoyData = New JoystickData(Me)
        myBW = New BackgroundWorker()
        AddHandler myBW.DoWork, AddressOf doOpenDevice
        'myBW.WorkerSupportsCancellation = True
        myBW.RunWorkerAsync()

    End Sub

    Public Sub CloseDevice()

        If myBW IsNot Nothing Then
            CancelFlag = True
            'myBW.CancelAsync()
            'While Busy
            '    Thread.Sleep(100)
            '    Application.DoEvents()
            'End While
        End If

    End Sub

    Public Sub New(ByVal hDev As HidDevice)
        newDevice(hDev)
    End Sub

    Private Sub doReportCallback(ByVal dat As JoystickData)

        Static lastCallback As Integer = 0

        If _reportCallback IsNot Nothing Then
            _reportCallback.Invoke(dat)
        End If

    End Sub

    Private Sub doOpenDevice(ByVal sender As Object, ByVal e As DoWorkEventArgs)

        Try
            Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
            Dim inputReportBuffer() As Byte = Nothing
            Dim inputReceiver As Input.HidDeviceInputReceiver = Nothing
            Dim inputParser As Input.DeviceItemInputParser = Nothing
            Dim hidStrm As HidStream = Nothing

            If myHID.TryOpen(hidStrm) Then
                Debug.WriteLine("Opened device.")
                hidStrm.ReadTimeout = Timeout.Infinite
                Array.Resize(inputReportBuffer, myHID.GetMaxInputReportLength())
                inputReceiver = myReportDescriptor.CreateHidDeviceInputReceiver()
                inputParser = myDeviceItem.CreateDeviceItemInputParser()
                inputReceiver.Start(hidStrm)

                'Debug.WriteLine(Now.Ticks & " Waiting...")
                Do While True
                    'If worker.CancellationPending Then e.Cancel = True : Exit Do
                    If CancelFlag Then Exit Do

                    'Debug.Write(Now.Ticks & " Waiting...")
                    If (inputReceiver.WaitHandle.WaitOne(10)) Then
                        If Not inputReceiver.IsRunning Then Exit Sub 'Disconnected?
                        'If worker.CancellationPending Then e.Cancel = True : Exit Do
                        If CancelFlag Then Exit Do

                        Dim rpt As Report = Nothing
                        'Debug.WriteLine(Now.Ticks & " TryRead...")
                        While inputReceiver.TryRead(inputReportBuffer, 0, rpt)
                            'Parse the report if possible.
                            'This will return false if (for example) the report applies to a different DeviceItem.
                            'Debug.Write(Now.Ticks & " Parsing...")
                            'If worker.CancellationPending Then e.Cancel = True : Exit Do
                            If CancelFlag Then Exit Do
                            If inputParser.TryParseReport(inputReportBuffer, 0, rpt) Then
                                'Debug.WriteLine("Ok")
                                'If worker.CancellationPending Then e.Cancel = True : Exit Do
                                If CancelFlag Then Exit Do
                                Dim jd As JoystickData = writeDeviceItemInputParserResultData(inputParser, Me)

                                'If worker.CancellationPending Then e.Cancel = True : Exit Do
                                If CancelFlag Then Exit Do
                                doReportCallback(jd)
                            Else
                                Debug.WriteLine(Now & " Bad parse")
                                'If worker.CancellationPending Then e.Cancel = True : Exit Do
                                If CancelFlag Then Exit Do
                                doReportCallback(Nothing)
                            End If
                        End While
                    Else
                        'If worker.CancellationPending Then e.Cancel = True : Exit Do
                        If CancelFlag Then Exit Do
                        doReportCallback(Nothing)
                        'Debug.WriteLine(Now.Ticks & " timeout")
                    End If
                Loop

                hidStrm.Close()
                hidStrm.Dispose()
                hidStrm = Nothing
                inputParser = Nothing
                inputReceiver = Nothing
                myJoyData = Nothing

                Debug.WriteLine(Now & " Device closed.")
            Else
                Debug.WriteLine("Failed to open device.")
            End If
        Catch ex As Exception
            Debug.WriteLine("Exception: " & ex.Message)
        Finally
            _Busy = False
        End Try

    End Sub

    Private Function isValidUsage(ByVal tUsage As UInteger, ByVal hdev As HidDevice) As Boolean

        If CType(tUsage, Usage) = Usage.GenericDesktopJoystick Then Return True
        If CType(tUsage, Usage) = Usage.GenericDesktopGamepad Then Return True
        'If hdev.GetProductName().ToLower.Contains("flight") Then Return True
        If tUsage = &H10000 Then Return True 'Logitech Flight Multi Panel

        Return False

    End Function

    Private Sub newDevice(ByVal hDev As HidDevice)

        Try
            Dim rawReportDescriptor = hDev.GetRawReportDescriptor()
            Debug.WriteLine("Report Descriptor: " & rawReportDescriptor.ToString())
            '            Debug.WriteLine("  {0} ({1} bytes)", String.Join(" ", rawReportDescriptor.Select(d >= d.ToString("X2"))), rawReportDescriptor.Length)

            Dim indent As Integer = 0
            For Each element In EncodedItem.DecodeItems(rawReportDescriptor, 0, rawReportDescriptor.Length)
                If (element.ItemType = ItemType.Main And element.TagForMain = MainItemTag.EndCollection) Then indent -= 2
                Debug.WriteLine("  {0}{1}", New String(" "c, indent), element)
                If (element.ItemType = ItemType.Main And element.TagForMain = MainItemTag.Collection) Then indent += 2
            Next

            Dim reportDescriptor = hDev.GetReportDescriptor()

            'Lengths should match.
            Debug.Assert(hDev.GetMaxInputReportLength() = reportDescriptor.MaxInputReportLength)
            Debug.Assert(hDev.GetMaxOutputReportLength() = reportDescriptor.MaxOutputReportLength)
            Debug.Assert(hDev.GetMaxFeatureReportLength() = reportDescriptor.MaxFeatureReportLength)

            Debug.WriteLine("Num DeviceItems: " & reportDescriptor.DeviceItems.Count)
            'If reportDescriptor.DeviceItems.Count > 1 Then
            '    Exit Sub
            'End If

            For Each deviceItem In reportDescriptor.DeviceItems
                Dim gotJoy As Boolean = False
                Debug.WriteLine("Device: " & hDev.GetProductName())
                Debug.WriteLine("Num Usages: " & deviceItem.Usages.Count)

                For Each tUsage In deviceItem.Usages.GetAllValues()
                    Debug.WriteLine(String.Format("Usage: {0:X4} {1}", tUsage, CType(tUsage, Usage)))
                    'If CType(tUsage, Usage) = Usage.GenericDesktopJoystick Or CType(tUsage, Usage) = Usage.GenericDesktopGamepad Or hDev.GetFriendlyName().ToLower.Contains("flight") Then
                    If isValidUsage(tUsage, hDev) Then
                        myHID = hDev
                        myReportDescriptor = reportDescriptor
                        myDeviceItem = deviceItem
                        Exit For
                    End If
                    'myHID = hDev
                    'myReportDescriptor = reportDescriptor
                    'myDeviceItem = deviceItem
                    'GenericDesktopJoystick = 0x00010004,
                    'GenericDesktopGamepad = 0x00010005,
                Next

                Debug.WriteLine("Num Reports: " & deviceItem.Reports.Count)

                For Each report In deviceItem.Reports
                    Debug.WriteLine(String.Format("{0}: ReportID={1}, Length={2}, Items={3}",
                                                    report.ReportType, report.ReportID, report.Length, report.DataItems.Count))
                    For Each dataItem In report.DataItems
                        'Debug.WriteLine(String.Format("  {0} Elements x {1} Bits, Units: {2}, Expected Usage Type: {3}, Flags: {4}, Usages: {5}",
                        '                dataItem.ElementCount, dataItem.ElementBits, dataItem.Unit.System, dataItem.ExpectedUsageType, dataItem.Flags,
                        '                String.Join(", ", dataItem.Usages.GetAllValues().Select(tUsage >= Usage.ToString("X4") + " " + (CType(tusage, Usage)).ToString()))))
                    Next
                Next
            Next
        Catch e As Exception
            Debug.WriteLine("Exception")
            Debug.WriteLine(e)
        End Try

    End Sub

    Private myHID As HidDevice = Nothing
    Private myReportDescriptor As ReportDescriptor = Nothing
    Private myDeviceItem As DeviceItem = Nothing
    Private _reportCallback As DeviceReportCallbackDelegate = Nothing
    Private myBW As BackgroundWorker
    Private _Busy As Boolean = False
    Private myJoyData As JoystickData = Nothing

End Class
