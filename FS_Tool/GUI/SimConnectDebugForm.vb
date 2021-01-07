Option Explicit On
Option Strict On

Imports System.Runtime.InteropServices
Imports SimConLib
Imports System.Text

Public Class SimConnectDebugForm

    Public Sub SetTransmitButtonVisible(ByVal visibleFlag As Boolean)
        Me.btnTransmit.Visible = visibleFlag
    End Sub

    Private mySim As SimConnectLib = Nothing
    Private bEnumsMapped As Boolean = False

    Public WriteOnly Property Sim As SimConnectLib
        Set(value As SimConnectLib)
            mySim = value
        End Set
    End Property

    Private Sub moveInstWin(ByVal hWnd As IntPtr)

        If hWnd <> IntPtr.Zero Then
            Dim X = Me.Left + Me.pnlWinBox.Left
            Dim Y = Me.Top + Me.pnlWinBox.Top
            Dim nWidth = Me.pnlWinBox.Width
            Dim nHeight = Me.pnlWinBox.Height

            HidSharp.Win32API.MoveWindow(hWnd, X, Y, nWidth, nHeight, True)
            HidSharp.Win32API.SetForegroundWindow(hWnd)
        End If

    End Sub

    Private Sub checkInstWin()

        Dim localByName As Process() = Process.GetProcessesByName("FlightSimulator")
        If localByName.Count = 1 Then
            Dim fsp = localByName(0)
            Dim _windowHandles = New List(Of IntPtr)
            Dim listHandle As GCHandle = CType(Nothing, GCHandle)
            listHandle = GCHandle.Alloc(_windowHandles)
            For Each pt As ProcessThread In fsp.Threads
                HidSharp.Win32API.EnumThreadWindows(CType(pt.Id, UInteger), New HidSharp.Win32API.EnumThreadDelegate(AddressOf EnumThreadCallback), GCHandle.ToIntPtr(listHandle))
            Next

            For Each hWnd As IntPtr In _windowHandles
                Dim wt As New StringBuilder(128)
                Dim cn As New StringBuilder(128)

                Dim wtl = HidSharp.Win32API.GetWindowText(hWnd, wt, 128)
                HidSharp.Win32API.GetClassName(hWnd, cn, 128)

                If cn.ToString = "AceApp" And wtl = 0 Then
                    moveInstWin(hWnd)
                    Exit For
                End If

                'Debug.WriteLine(hWnd.ToString("X8") & ": " & wt.ToString & "," & cn.ToString & "")
            Next
        End If

    End Sub

    Private Sub doPollWin()

        'MessageBox.Show("Left: " & Me.Left & ", Top: " & Me.Top)


    End Sub

    Private Shared Function EnumThreadCallback(ByVal hWnd As IntPtr, ByVal lParam As IntPtr) As Boolean

        Dim gch As GCHandle = GCHandle.FromIntPtr(lParam)
        Dim list As List(Of IntPtr) = CType(gch.Target, List(Of IntPtr))
        If list Is Nothing Then
            Throw New InvalidCastException("GCHandle Target could not be cast as List(Of IntPtr)")
        End If

        list.Add(hWnd)

        Return True

    End Function

    Private Sub winPollContinuous(ByVal activeFlag As Boolean)

    End Sub

    Private Sub doTransmit()

        Try
            'If Not bEnumsMapped Then
            '    bEnumsMapped = True
            '    Dim eL As New List(Of [Enum])
            '    For Each item As ENUM_EVENTS In [Enum].GetValues(GetType(ENUM_EVENTS))
            '        eL.Add(item)
            '    Next
            '    mySim.MapEventEnums(eL)
            'End If
            'Dim t As ENUM_EVENTS = getEnumByName(Me.cbEventID.Text)
            Dim t As SimData.SimEventEnum = SimConnectLib.GetSimEventByName(Me.cbEventID.Text)
            If t <> SimData.SimEventEnum.NONE Then
                mySim.TransmitClientEvent(t, "0")
            End If
        Catch ex As Exception
            MessageBox.Show("Exception: " & ex.Message)
        End Try

    End Sub

    Private Sub doHeadingFlip(Optional ByVal flipDegrees As Double = 180.0)

        Dim curHdg As Double = mySim.GetSimValDouble("HEADING INDICATOR") + flipDegrees
        If curHdg > 360.0 Then
            curHdg -= 360
        End If

        mySim.TransmitClientEvent(SimData.SimEventEnum.HEADING_BUG_SET, CInt(curHdg).ToString)

    End Sub

    Private Sub doTimerTick()

        If mySim IsNot Nothing Then
            Dim str As String = Now & vbNewLine

            str = str & vbNewLine & "PLANE HEADING DEGREES MAGNETIC: " & mySim.GetSimValDouble("PLANE HEADING DEGREES MAGNETIC")
            str = str & vbNewLine & "PLANE HEADING DEGREES GYRO: " & mySim.GetSimValDouble("PLANE HEADING DEGREES GYRO")
            str = str & vbNewLine & "PLANE HEADING DEGREES: " & mySim.GetSimValDouble("PLANE HEADING DEGREES")
            str = str & vbNewLine & "HEADING INDICATOR: " & mySim.GetSimValDouble("HEADING INDICATOR")
            'str = str & vbNewLine & "HSI BEARING: " & mySim.GetSimValDouble("HSI BEARING")
            'str = str & vbNewLine & "AUTOPILOT HEADING LOCK DIR: " & mySim.GetSimValDouble("AUTOPILOT HEADING LOCK DIR")

            str = str & vbNewLine & "GENERAL ENG MASTER ALTERNATOR1: " & mySim.GetSimValDouble("GENERAL ENG MASTER ALTERNATOR:1")
            str = str & vbNewLine & "GENERAL ENG MASTER ALTERNATOR2: " & mySim.GetSimValDouble("GENERAL ENG MASTER ALTERNATOR:2")
            str = str & vbNewLine & "GENERAL ENG MASTER ALTERNATOR3: " & mySim.GetSimValDouble("GENERAL ENG MASTER ALTERNATOR:3")
            str = str & vbNewLine & "GENERAL ENG MASTER ALTERNATOR4: " & mySim.GetSimValDouble("GENERAL ENG MASTER ALTERNATOR:4")

            Me.txtDisplay.Text = str

            If Me.chkContinuous.Checked Then
                doSimVarPoll()
            End If
        End If

    End Sub

    Private Sub doSimVarPoll()

        If Me.txtSimVar.Text = "" Then
            Me.txtSimVarVal.Text = ""
        Else
            Me.txtSimVarVal.Text = mySim.GetSimValDebug(Me.txtSimVar.Text).ToString
        End If

    End Sub

    Private Sub registerSimVars()

        If mySim Is Nothing Then
            Me.txtDisplay.Text = "SimConnect not active"
        Else
            'Dim tl As New List(Of String)
            'tl.Add("PLANE HEADING DEGREES GYRO")
            'tl.Add("PLANE HEADING DEGREES")
            'tl.Add("HEADING INDICATOR")
            'tl.Add("HSI BEARING")
            'tl.Add("AUTOPILOT HEADING LOCK DIR")
            'mySim.RegisterSimVars(tl)

            mySim.RegisterCommonSimVars()

            mySim.RegisterCustomSimVar("GENERAL ENG MASTER ALTERNATOR:1", "Bool")
            mySim.RegisterCustomSimVar("GENERAL ENG MASTER ALTERNATOR:2", "Bool")
            mySim.RegisterCustomSimVar("GENERAL ENG MASTER ALTERNATOR:3", "Bool")
            mySim.RegisterCustomSimVar("GENERAL ENG MASTER ALTERNATOR:4", "Bool")

        End If

    End Sub

    Private Sub formInit()

        Me.Left = My.Settings.DebugX
        Me.Top = My.Settings.DebugY

        'MessageBox.Show("PARKING_BRAKES: " & CInt(ENUM_EVENTS.PARKING_BRAKES))

        Me.cbEventID.Items.Clear()
        For Each item As SimData.SimEventEnum In [Enum].GetValues(GetType(SimData.SimEventEnum))
            Me.cbEventID.Items.Add(item)
            'If item <> SimConnectLib.EventEnum.NONE Then
            '    'Sim.MapClientEventToSimEvent(item, item.ToString())
            '    'Sim.AddClientEventToNotificationGroup(GROUP.GROUP1, item, False)
            'End If
        Next
        Me.cbEventID.Sorted = True

        registerSimVars()

        tmrPoll.Enabled = True
        tmrPoll.Start()

    End Sub

    Private Sub formClose()

        tmrPoll.Stop()
        tmrPoll.Enabled = False

        My.Settings.DebugX = Me.Left
        My.Settings.DebugY = Me.Top
        My.Settings.Save()

    End Sub

#Region "Event Handlers"

    Private Sub SimConnectDebugForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        formInit()
    End Sub

    Private Sub SimConnectDebugForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        formClose()
    End Sub

    Private Sub btnTransmit_Click(sender As Object, e As EventArgs) Handles btnTransmit.Click
        doTransmit()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Threading.Thread.Sleep(2000)
        'SendKey(ScanCodeEnum.KEY_F4, KeyModifierEnum.SH_CTL)
        SendScanCodeStroke(ScanCodeEnum.KEY_F4, KeyModifierEnum.SH_CTL, KeyStrokeTypeEnum.Full)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        'Threading.Thread.Sleep(3)
        'SendKey(Keys.B)

        'SetWindowFocus("Microsoft Flight Simulator - 1.9.3.0")
        Threading.Thread.Sleep(2000)
        'SendScanCode(ScanCodeEnum.KEY_B, KeyModifierEnum.CTL)
        'SendScanCode(ScanCodeEnum.KEY_F8)
        'SendScanCode(ScanCodeEnum.KEY_B)
        'SendKey(VK_CodeEnum.VK_F8)
        'SetWindowFocus("Untitled - Notepad")
        'SetWindowFocus("*Untitled - Notepad")
        '
        'SendKey(VK_CodeEnum.VK_L)

        SendScanCodeStroke(ScanCodeEnum.KEY_L, KeyStrokeTypeEnum.Full)
        'InputHelper.Keyboard.PressKey(Keys.L, True) 'True = Send key as hardware scan code.



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles btnTestHello.Click
        TestSendInput()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Threading.Thread.Sleep(2000)
        SendScanCodeStroke(ScanCodeEnum.KEY_TAB, KeyStrokeTypeEnum.Full)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Threading.Thread.Sleep(2000)
        SendScanCodeStroke(ScanCodeEnum.KEY_SPACE, KeyStrokeTypeEnum.Full)

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Threading.Thread.Sleep(2000)
        SendScanCodeStroke(ScanCodeEnum.KEY_ENTER, KeyStrokeTypeEnum.Full)

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Threading.Thread.Sleep(2000)
        SendScanCodeStroke(ScanCodeEnum.KEY_KP0, KeyStrokeTypeEnum.Full)

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Threading.Thread.Sleep(2000)
        SendScanCodeStroke(ScanCodeEnum.KEY_KPSLASH, KeyStrokeTypeEnum.Full)

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        Threading.Thread.Sleep(2000)
        SendScanCodeStroke(ScanCodeEnum.KEY_UP, KeyStrokeTypeEnum.Full)

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        Threading.Thread.Sleep(2000)
        SendScanCodeStroke(ScanCodeEnum.KEY_DOWN, KeyStrokeTypeEnum.Full)

    End Sub

    Private Sub cbEventID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbEventID.SelectedIndexChanged

    End Sub

    Private Sub tmrPoll_Tick(sender As Object, e As EventArgs) Handles tmrPoll.Tick
        doTimerTick()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        doHeadingFlip()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        doHeadingFlip(90.0)
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        doHeadingFlip(-90.0)
    End Sub

    Private Sub btnPoll_Click(sender As Object, e As EventArgs) Handles btnPoll.Click
        doSimVarPoll()
    End Sub

    Private Sub btnPollWin_Click(sender As Object, e As EventArgs) Handles btnPollWin.Click
        doPollWin()
    End Sub

    Private Sub chkContinuousWin_CheckedChanged(sender As Object, e As EventArgs) Handles chkContinuousWin.CheckedChanged
        winPollContinuous(CType(sender, CheckBox).Checked)
    End Sub

    Private Sub SimConnectDebugForm_Move(sender As Object, e As EventArgs) Handles MyBase.Move
        checkInstWin()
    End Sub

#End Region

End Class
