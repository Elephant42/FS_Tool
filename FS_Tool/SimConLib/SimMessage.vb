Option Explicit On
Option Strict On

Imports System.Xml

Public Class SimMessage

#Region "Public Declarations"

    Public Enum CommandEnum
        Null
        SysCmd
        RegisterSimVars
        GetSimVals
        SendSimEvent
        SendKeyEvent
        GetStatus
        ErrorMessage
    End Enum

    Public Structure SimEventStruct
        Public SimEventCode As String
        Public SimEventData As String
    End Structure

    Public Property Command As CommandEnum = CommandEnum.Null

    Public ReadOnly Property Socket As AsyncSocket
        Get
            Return mySocket
        End Get
    End Property

    Public Property KeyEvent As KeyEventData
        Get
            Return _KeyEvent
        End Get
        Set(value As KeyEventData)
            _KeyEvent = value
        End Set
    End Property

    Public Property SimEvent As SimEventStruct
        Get
            Return _SimEvent
        End Get
        Set(value As SimEventStruct)
            _SimEvent = value
        End Set
    End Property

    Public ReadOnly Property SimVars As Dictionary(Of String, Double)
        Get
            Return mysimVars
        End Get
    End Property

    Public ReadOnly Property UID As String
        Get
            Return _UID
        End Get
    End Property

    Public ReadOnly Property IsValid As Boolean
        Get
            Return _IsValid
        End Get
    End Property

    Public Property XML As String
        Get
            Return generateXML()
        End Get
        Set(value As String)
            parseXML(value)
        End Set
    End Property

#End Region

#Region "Public Functions"

    Public Sub New(pSocket As AsyncSocket)
        Me.New(pSocket, "")
    End Sub
    Public Sub New(pSocket As AsyncSocket, ByVal xml As String)

        initData()
        mySocket = pSocket
        If xml <> "" Then
            parseXML(xml)
        End If

    End Sub
    Public Sub New(pSocket As AsyncSocket, ByVal command As String, ByVal data As Object)
        Me.New(pSocket, getCommandByName(command), data)
    End Sub
    Public Sub New(pSocket As AsyncSocket, ByVal command As CommandEnum, ByVal data As Object)

        initData()
        mySocket = pSocket
        Me.Command = command
        Select Case Me.Command
            Case CommandEnum.ErrorMessage
                If TypeOf data Is String Then
                    _Message = CType(data, String)
                Else
                    Throw New Exception("Data does not match command")
                End If
            Case CommandEnum.SendKeyEvent
                If TypeOf data Is KeyEventData Then
                    _KeyEvent = CType(data, KeyEventData)
                Else
                    Throw New Exception("Data does not match command")
                End If
            Case CommandEnum.SendSimEvent
                If TypeOf data Is SimEventStruct Then
                    _SimEvent = CType(data, SimEventStruct)
                Else
                    Throw New Exception("Data does not match command")
                End If
            Case CommandEnum.RegisterSimVars
                If TypeOf data Is List(Of String) Then
                    mysimVars.Clear()
                    For Each temp As String In CType(data, List(Of String))
                        mysimVars.Add(temp, 0.0)
                    Next
                Else
                    Throw New Exception("Data does not match command")
                End If
            Case CommandEnum.GetSimVals
                If TypeOf data Is Dictionary(Of String, Double) Then
                    mysimVars.Clear()
                    Dim t As Dictionary(Of String, Double) = CType(data, Dictionary(Of String, Double))
                    For Each key As String In t.Keys
                        mysimVars.Add(key, t(key))
                    Next
                Else
                    Throw New Exception("Data does not match command")
                End If
            Case Else

        End Select

    End Sub

    Public Sub Send()

        If mySocket Is Nothing Then
            Throw New Exception("Not connected!")
        Else
            Dim t As String = Me.XML
            mySocket.Send(Me.XML)
        End If

    End Sub

    Public Sub Execute()

        Select Case Command
            Case CommandEnum.SysCmd

            Case CommandEnum.RegisterSimVars
                executeRegisterSimVars()
            Case CommandEnum.GetSimVals
                executeSimVarsMsg()
            Case CommandEnum.SendKeyEvent
                executeKeyMsg()
            Case Else
        End Select
    End Sub

    Public Function GetSimVar(ByVal simVarCode As String) As Double

        If mysimVars.ContainsKey(simVarCode) Then
            Return mysimVars(simVarCode)
        Else
            Return 0.0
        End If

    End Function

    Public Sub SetSimVar(ByVal simVarCode As String, ByVal simVarVal As Double)

        If mysimVars.ContainsKey(simVarCode) Then
            mysimVars(simVarCode) = simVarVal
        Else
            mysimVars.Add(simVarCode, simVarVal)
        End If

    End Sub


#End Region

#Region "Private Functions"

    Private Sub initData()

        _KeyEvent.ScanCode = ScanCodeEnum.KEY_NONE
        _KeyEvent.KeyModifier = KeyModifierEnum.None

        _SimEvent.SimEventCode = ""
        _SimEvent.SimEventData = "0"

    End Sub

    Private Sub executeKeyMsg()

        If Me.KeyEvent.ScanCode <> ScanCodeEnum.KEY_NONE Then
            SendScanCodeStroke(Me.KeyEvent)
        End If

    End Sub

    Private Sub executeRegisterSimVars()

    End Sub

    Private Sub executeSimVarsMsg()

    End Sub

    Private Sub parseXML(ByVal xml As String)

        Try
            Dim tempDOM As New XmlDocument
            tempDOM.LoadXml(xml)
            Dim rootNode As XmlElement = CType(tempDOM.ChildNodes(0), XmlElement)

            Me.Command = getCommandByName(rootNode.GetAttribute("Command"))
            Dim t As String = rootNode.GetAttribute("UID")
            If t <> "" Then _UID = t

            Select Case Me.Command
                Case CommandEnum.ErrorMessage
                    _Message = rootNode.GetAttribute("Message")
                    _IsValid = True
                Case CommandEnum.SendKeyEvent
                    Dim temp As XmlElement = CType(rootNode.ChildNodes(0), XmlElement)
                    _KeyEvent.ScanCode = GetScanCodeByName(temp.GetAttribute("ScanCode"))
                    _KeyEvent.KeyModifier = GetKeyModifierByName(temp.GetAttribute("KeyModifier"))
                    _KeyEvent.StrokeType = GetKeyStrokeTypeByName(temp.GetAttribute("StrokeType"))
                    _IsValid = True
                Case CommandEnum.SendSimEvent
                    Dim temp As XmlElement = CType(rootNode.ChildNodes(0), XmlElement)
                    _SimEvent.SimEventCode = temp.GetAttribute("SimEventCode")
                    _SimEvent.SimEventData = temp.GetAttribute("SimEventData")
                    _IsValid = True
                Case CommandEnum.GetSimVals
                    For Each tempElem As XmlElement In rootNode.ChildNodes
                        Dim key As String = tempElem.GetAttribute("SimVarCode")
                        If Not mysimVars.ContainsKey(key) Then
                            mysimVars.Add(key, 0.0)
                        End If
                        mysimVars(key) = CDbl(tempElem.GetAttribute("SimVarValue"))
                    Next
                    _IsValid = True
                Case CommandEnum.RegisterSimVars
                    mysimVars.Clear()
                    For Each tempElem As XmlElement In rootNode.ChildNodes
                        mysimVars.Add(tempElem.InnerText, 0.0)
                    Next
                    _IsValid = True
                Case Else
                    _IsValid = False
            End Select
        Catch ex As Exception
            Debug.WriteLine(xml)
            _IsValid = False
        End Try

    End Sub

    Private Function generateXML() As String

        Dim tempDOM As New XmlDocument

        Dim rootNode As XmlElement = tempDOM.CreateElement("SimMessage")
        rootNode.SetAttribute("Command", Me.Command.ToString)
        rootNode.SetAttribute("UID", Me.UID)

        Select Case Me.Command
            Case CommandEnum.ErrorMessage
                rootNode.SetAttribute("Message", _Message)
            Case CommandEnum.SendKeyEvent
                Dim temp As XmlElement = tempDOM.CreateElement("KeyEvent")
                temp.SetAttribute("ScanCode", _KeyEvent.ScanCode.ToString)
                temp.SetAttribute("KeyModifier", _KeyEvent.KeyModifier.ToString)
                temp.SetAttribute("StrokeType", _KeyEvent.StrokeType.ToString)
                rootNode.AppendChild(temp)
            Case CommandEnum.SendSimEvent
                Dim temp As XmlElement = tempDOM.CreateElement("SimEvent")
                temp.SetAttribute("SimEventCode", _SimEvent.SimEventCode)
                temp.SetAttribute("SimEventData", _SimEvent.SimEventData)
                rootNode.AppendChild(temp)
            Case CommandEnum.GetSimVals
                For Each key As String In mysimVars.Keys
                    Dim temp As XmlElement = tempDOM.CreateElement("SimVar")
                    temp.SetAttribute("SimVarCode", key)
                    temp.SetAttribute("SimVarValue", mysimVars(key).ToString)
                    rootNode.AppendChild(temp)
                Next
            Case CommandEnum.RegisterSimVars
                For Each key As String In mysimVars.Keys
                    Dim temp As XmlElement = tempDOM.CreateElement("SimVar")
                    temp.InnerText = key
                    rootNode.AppendChild(temp)
                Next
            Case Else

        End Select

        tempDOM.AppendChild(rootNode)

        Return tempDOM.OuterXml

    End Function

    Private Shared Function getCommandByName(ByVal commandName As String) As CommandEnum

        Try
            Return CType([Enum].Parse(GetType(CommandEnum), commandName), CommandEnum)
        Catch ex As Exception
            Return CommandEnum.Null
        End Try

    End Function

#End Region

#Region "Private Declarations"

    Private mysimVars As New Dictionary(Of String, Double)
    Private mySocket As AsyncSocket = Nothing
    Private _KeyEvent As KeyEventData
    Private _SimEvent As SimEventStruct
    Private _UID As String = Guid.NewGuid.ToString
    Private _IsValid As Boolean = False
    Private _Message As String = ""

#End Region

End Class
