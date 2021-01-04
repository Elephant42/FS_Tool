Option Explicit On
Option Strict On

Imports System.Threading
Imports HidSharp
Imports HidSharp.Reports
Imports HidSharp.Reports.Encodings
Imports System.Xml

Public Class JoystickForm

    Private Sub doDisplay(ByVal dat As JoystickData)

        If Me.InvokeRequired Then
            Me.Invoke(Sub() doDisplay(dat))
            Return
        End If

        Dim enableDisplay As Boolean = True
        Dim ignoreAxes As Boolean = Me.cbIgnoreAxes.Checked

        If myJoy.Worker.CancellationPending Then Exit Sub

        If dat IsNot Nothing Then
            If Me.cbThrottle.Checked Then
                If Now < lastDisplay.AddMilliseconds(100) Then
                    enableDisplay = False
                End If
            End If

            If enableDisplay Then
                Dim t As String = ""
                For Each d In dat.Values
                    Dim b As Boolean = True
                    If ignoreAxes Then
                        If d.IsAxis Then b = False
                    End If
                    If b Then t &= d.Index & vbTab & d.Name.PadRight(40) & ": " & d.PreviousValue & "->" & d.Value & vbNewLine
                Next
                Me.txtDisplay.AppendText(t)
                lastDisplay = Now
            End If
        End If

        'If strParam <> "" Then
        '    Dim t As String = ""
        '    If strParam <> "" Then
        '        t = strParam.Trim().Replace(vbNullChar, "") & vbNewLine
        '    End If
        '    Me.txtDisplay.AppendText(t)
        'End If

        Application.DoEvents()

    End Sub

    Private Sub loadJoysticks()

        myJoysticks = Joystick.GetJoysticks()
        Me.cbJoysticks.Items.Clear()

        Me.cbJoysticks.Items.Add("")
        For Each j In myJoysticks
            j.DeviceReportCallback = AddressOf doDisplay
            Me.cbJoysticks.Items.Add(j)
        Next
        Me.cbJoysticks.DisplayMember = "DisplayName"
        Me.cbJoysticks.Sorted = True

    End Sub

    Private Sub closeJoy()

        If myJoy IsNot Nothing Then
            myJoy.CloseDevice()
            myJoy = Nothing
        End If

        Me.cbJoysticks.Text = ""
        Me.btnClose.Visible = False
        Me.btnClipboard.Visible = False
        Me.cbJoysticks.Enabled = True

    End Sub

    Private Sub copyToClip()

        Dim t As String = "<Joystick MappingName=""" & Me.cbJoysticks.Text & " - Default"" Name=""" & Me.cbJoysticks.Text & """>" & vbNewLine
        For Each d In myJoy.Data.Values
            If Not d.IsAxis Then
                Dim tt As String = "<JoyMap JoyEvent=""" & d.Name & """ " & "JoyEnableEvents="""" SimEvent="""" SimVar="""" ActiveState=""True"" /><!---->"
                t &= vbTab & tt & vbNewLine
            End If
        Next
        t &= "</Joystick>" & vbNewLine

        Try
            Clipboard.SetText(t)
        Catch ex As Exception
            Debug.WriteLine("Clipboard exception: " & ex.Message)
        End Try

    End Sub

    Private myJoy As Joystick = Nothing
    Private lastDisplay As Date = Date.MinValue
    Private myJoysticks As List(Of Joystick) = Nothing

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbJoysticks.SelectedIndexChanged

        If Me.cbJoysticks.Text <> "" Then
            Me.cbJoysticks.Enabled = False
            Me.btnClose.Visible = True
            Me.btnClipboard.Visible = True
            myJoy = CType(Me.cbJoysticks.SelectedItem, Joystick)
            If myJoy.Busy Then
                MessageBox.Show("Joystick device is busy", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                myJoy.OpenDevice()
            End If
        End If

    End Sub

    Private Sub JoystickForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Left = My.Settings.JoyX
        Me.Top = My.Settings.JoyY
        Me.cbThrottle.Checked = My.Settings.JoyThrottle
        Me.cbIgnoreAxes.Checked = My.Settings.JoyIgnoreAxes

        loadJoysticks()
        Me.btnClose.Visible = False
        Me.btnClipboard.Visible = False
        Me.cbJoysticks.Enabled = True

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Me.Timer1.Stop()

        'If myJoy IsNot Nothing Then
        '    Debug.WriteLine(Now & " Start poll")
        '    myJoy.PollDevice()
        '    Debug.WriteLine(Now & " Finish poll")
        'End If

        'Application.DoEvents()

        'Me.Timer1.Start()

    End Sub

    Private Sub JoystickForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If myJoy IsNot Nothing Then
            If myJoy.Busy Then e.Cancel = True : Exit Sub
        End If

        My.Settings.JoyX = Me.Left
        My.Settings.JoyY = Me.Top
        My.Settings.JoyThrottle = Me.cbThrottle.Checked
        My.Settings.JoyIgnoreAxes = Me.cbIgnoreAxes.Checked
        My.Settings.Save()

    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click

        If myJoy IsNot Nothing Then
            closeJoy()
        End If

    End Sub

    Private Sub btnClipboard_Click(sender As Object, e As EventArgs) Handles btnClipboard.Click
        copyToClip()
    End Sub

End Class
