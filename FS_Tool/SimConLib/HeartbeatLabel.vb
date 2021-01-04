Option Strict On
Option Explicit On

Imports System.Windows.Forms

Public Class HeartbeatLabel
    Inherits Label

    Public Property HeartbeatString As String = "*"

    Public Property HeartbeatMilliSeconds As Integer = 1

    Public Sub Tick()

        If Now > lastBeat.AddMilliseconds(Me.HeartbeatMilliSeconds) Then
            lastBeat = Now
            Me.Toggle()
        End If

    End Sub

    Public Sub Toggle()

        If Me.Text = "" Then
            Me.Text = Me.HeartbeatString
        Else
            Me.Text = ""
        End If

    End Sub

    Private lastBeat As Date = Date.MinValue

End Class

