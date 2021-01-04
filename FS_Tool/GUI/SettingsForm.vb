Option Explicit On
Option Strict On
Imports System.IO

Public Class SettingsForm

    Public Property StartMinimised As Boolean
        Get
            Return Me.chkStartMinim.Checked
        End Get
        Set(value As Boolean)
            Me.chkStartMinim.Checked = value
        End Set
    End Property

    Public Property ServerPort As Integer
        Get
            Return CInt(Me.txtPort.Text)
        End Get
        Set(value As Integer)
            Me.txtPort.Text = value.ToString
        End Set
    End Property

    Public Property AutoStartJoyMap As Boolean
        Get
            Return Me.chkAutoJoy.Checked
        End Get
        Set(value As Boolean)
            Me.chkAutoJoy.Checked = value
        End Set
    End Property

    Public Property AutoStartServer As Boolean
        Get
            Return Me.chkStartServer.Checked
        End Get
        Set(value As Boolean)
            Me.chkStartServer.Checked = value
        End Set
    End Property

    Public Property AutoStartSimConnect As Boolean
        Get
            Return Me.chkAutoSim.Checked
        End Get
        Set(value As Boolean)
            Me.chkAutoSim.Checked = value
        End Set
    End Property

    Public Property SimConnectPollingInterval As Integer
        Get
            Return CInt(Me.txtSimPoll_ms.Text)
        End Get
        Set(value As Integer)
            Me.txtSimPoll_ms.Text = value.ToString
        End Set
    End Property

    Private Sub doSave()

        If Not IsNumeric(Me.txtPort.Text) Then
            MessageBox.Show("Port Number is not numeric", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If Not IsNumeric(Me.txtSimPoll_ms.Text) Then
            MessageBox.Show("SimConnect Poll Interval is not numeric", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Me.DialogResult = DialogResult.OK

    End Sub

    Private Sub doSaveGrunt()

        Try
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MessageBox.Show("Exception: " & ex.Message)
        End Try

    End Sub

    Private Sub formNotBusy()
        formBusy(False)
    End Sub

    Private Sub formBusy()
        formBusy(True)
    End Sub
    Private Sub formBusy(ByVal pBusyFlag As Boolean)

        Me.btnSave.Enabled = Not pBusyFlag

    End Sub

    Private Sub doFormLoad()

    End Sub

#Region "Event Handlers"

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        doFormLoad()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        doSave()
    End Sub

    Private Sub chkAutoJoy_CheckedChanged(sender As Object, e As EventArgs) Handles chkAutoJoy.CheckedChanged
        If chkAutoJoy.Checked Then chkAutoSim.Checked = True
    End Sub

    Private Sub chkAutoSim_CheckedChanged(sender As Object, e As EventArgs) Handles chkAutoSim.CheckedChanged
        If Not chkAutoSim.Checked Then chkAutoJoy.Checked = False
    End Sub

#End Region

End Class
