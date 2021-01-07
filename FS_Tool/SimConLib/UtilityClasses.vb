Option Explicit On
Option Strict On

Imports System.ComponentModel
Imports System.Threading
Imports System.Windows.Forms

Public Class CircularBuffer

    Public Sub New(ByVal bufSz As Integer)
        Resize(bufSz)
    End Sub

    Public Sub Append(ByVal strVal As String)

        myBuffer(currentIndex) = strVal
        currentIndex += 1
        If currentIndex > mySize - 1 Then currentIndex = 0

    End Sub

    Public Sub Resize(ByVal bufSz As Integer)

        ReDim Preserve myBuffer(bufSz - 1)

        Dim ti As Integer = 0
        If mySize > 0 Then
            If bufSz <= mySize Then
                Exit Sub
            Else
                ti = mySize
            End If
        End If

        For i = ti To bufSz - 1
            myBuffer(i) = ""
        Next

        mySize = bufSz

    End Sub

    Public Shadows Function ToString() As String

        Dim t As String = ""

        Dim ndx As Integer = currentIndex
        For i = 0 To mySize - 1
            Dim temp As String = myBuffer(ndx)
            If temp <> "" Then
                If t = "" Then
                    t = temp & vbNewLine
                Else
                    t &= temp & vbNewLine
                End If
            End If

            ndx += 1
            If ndx > mySize - 1 Then
                ndx = 0
            End If
        Next

        Return t

    End Function

    Private currentIndex As Integer = 0
    Private myBuffer(0) As String
    Private mySize As Integer = 0

End Class

Public Class WaitTimer

    Public Shared Sub Wait(ByVal ms As Integer)

        Dim are As New AutoResetEvent(False)
        Dim tmr As New Threading.Timer(New TimerCallback(AddressOf timerThread), are, ms, 0)
        are.WaitOne()
        tmr.Dispose()

    End Sub

    Private Shared Sub timerThread(ByVal state As Object)

        Dim are = CType(state, AutoResetEvent)
        are.Set()

    End Sub

End Class
