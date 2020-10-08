Option Explicit On
Option Strict On

Imports SimConLib

Public Class MainForm

#Region "Private Declarations"

    Private Const HIDDEN_X As Integer = -16000
    Private Const WINDOW_LABEL As String = "Flight Sim Tool"

    Private Const sIM_NOT_CONNECT_MSG As String = "Sim not connected." & vbNewLine

    Private ReadOnly Property isHidden As Boolean
        Get
            Return (Me.Left < -10000)
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

#End Region

#Region "Private Functions"

    Private Sub simEnable()

        If sim Is Nothing Then
            sim = New SimConnectLib(Me)
            registerSimVars()
        End If
        initTimer()

    End Sub

    Private Sub simDisable()

        If sim IsNot Nothing Then
            sim.Disconnect()
            sim = Nothing
        End If

    End Sub

    Private Sub doDisplay(ByVal strParam As String)

        If Me.InvokeRequired Then
            Me.Invoke(Sub() doDisplay(strParam))
            Return
        End If

        Dim t As String = ""
        If strParam <> "" Then
            t = strParam.Trim().Replace(vbNullChar, "") & vbNewLine
        End If

        Me.txtDisplay.Text = t

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

        If sim IsNot Nothing Then
            sim.InitTimer(tmrUpdateTimer)
        End If

    End Sub

    Private Sub displaySimVals()

        Dim str As String = ""

        str &= "LATITUDE: " & sim.GetSimVal(SimData.LATITUDE) & vbNewLine
        str &= "LONGITUDE: " & sim.GetSimVal(SimData.LONGITUDE) & vbNewLine
        str &= "ALTITUDE: " & sim.GetSimVal(SimData.ALTITUDE) & vbNewLine
        str &= "HEADING: " & sim.GetSimVal(SimData.HEADING) & vbNewLine
        str &= "AIRSPEED (I): " & sim.GetSimVal(SimData.AIRSPEED_INDICATED) & vbNewLine
        str &= "AUTOPILOT MASTER: " & sim.GetSimVal(SimData.AUTOPILOT_MASTER) & vbNewLine
        str &= "PARKING BRAKE: " & sim.GetSimVal(SimData.PARKING_BRAKE) & vbNewLine

        doDisplay(str & vbNewLine)

    End Sub

    Private Sub doSettings()

        Dim frm As New SettingsForm
        frm.StartMinimised = My.Settings.StartMinimised
        If frm.ShowDialog(Me) = DialogResult.OK Then
            My.Settings.StartMinimised = frm.StartMinimised
            My.Settings.Save()
        End If

    End Sub

    Private Sub registerSimVars()

        Dim vars As New List(Of String)
        vars.Add(SimData.AIRSPEED_INDICATED)
        vars.Add(SimData.ALTITUDE)
        vars.Add(SimData.AUTOPILOT_MASTER)
        vars.Add(SimData.HEADING)
        vars.Add(SimData.LATITUDE)
        vars.Add(SimData.LONGITUDE)
        vars.Add(SimData.PARKING_BRAKE)
        vars.Add(SimData.SIM_RATE)
        sim.RegisterSimVars(vars)

    End Sub

    Private Sub timerTick()

        If sim IsNot Nothing Then
            If sim.Connected Then 'We are connected, Let's try to grab the data from the Sim
                Try
                    If Not bLastConnected Then
                        bLastConnected = True
                        Me.Text = WINDOW_LABEL & " - Connected"
                    End If
                    sim.TimerTick()
                    setSimRateMenuTicks()
                    Me.txtCurrSimRate.Text = sim.CurrentSimRate.ToString
                    'Me.txtDisplay.Text = sim.CurrentVals
                    'Me.txtDisplay.AppendText(vbNewLine & vbNewLine & "")
                    displaySimVals()
                Catch ex As Exception
                    sim.Disconnect()
                End Try
            Else 'We Then are Not connected, Let's try to connect
                If bLastConnected Then
                    bLastConnected = False
                    Dim t As String = WINDOW_LABEL & " - Disconnected"
                    Me.Text = t
                    Me.NotifyIcon1.Text = t
                    'Me.txtDisplay.Text = ""
                    Me.txtCurrSimRate.Text = ""
                End If
                If cmdRcvd > lastCmdRcvd Then
                    Me.txtCurrSimRate.Text = lastCmd
                End If
                doDisplay(sIM_NOT_CONNECT_MSG)
                sim.Connect()
            End If
        End If

    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)

        If sim IsNot Nothing Then
            If sim.Connected Then
                Try
                    sim.WndProcHandler(m)
                Catch ex As Exception
                    sim.Disconnect()
                End Try
            End If
        End If

        MyBase.WndProc(m)

    End Sub

    Private Sub formNotBusy()
        formBusy(False)
    End Sub

    Private Sub formBusy(Optional ByVal busyFlag As Boolean = True)

        Me.gbSlew.Enabled = (Not busyFlag)
        Me.gbSimRate.Enabled = (Not busyFlag)

    End Sub

    Private Sub hideForm()

        lastLeft = Me.Left
        Me.Left = HIDDEN_X

    End Sub

    Private Sub restoreForm()
        Me.Left = lastLeft
    End Sub

#End Region

#Region "Event Handlers"

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        bFormLoading = True

        Randomize()

        If My.Settings.FormX <> -1 Then Me.Left = My.Settings.FormX
        If My.Settings.FormY <> -1 Then Me.Top = My.Settings.FormY

        lastLeft = Me.Left

        formBusy()

        simEnable()

        Me.Text = WINDOW_LABEL
        validateSlew()
        doSimRateChanged()

        formNotBusy()

        If My.Settings.StartMinimised Then
            hideForm()
        End If

        bFormLoading = False

    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If sim IsNot Nothing Then
            sim.TransmitClientEvent(SimData.EventEnum.PARKING_BRAKES)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If sim IsNot Nothing Then
            sim.TransmitClientEvent(SimData.EventEnum.AP_MASTER)
        End If
    End Sub

    Private Sub ctx1Item_Settings_Click(sender As Object, e As EventArgs) Handles ctx1Item_Settings.Click
        doSettings()
    End Sub

#End Region

End Class

