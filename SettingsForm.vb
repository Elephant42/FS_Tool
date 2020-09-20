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

    Private Sub doSave()

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

#End Region

End Class
