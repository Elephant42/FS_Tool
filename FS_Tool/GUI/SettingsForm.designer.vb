<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SettingsForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.chkStartMinim = New System.Windows.Forms.CheckBox()
        Me.lblPort = New System.Windows.Forms.Label()
        Me.txtPort = New System.Windows.Forms.TextBox()
        Me.chkStartServer = New System.Windows.Forms.CheckBox()
        Me.chkAutoSim = New System.Windows.Forms.CheckBox()
        Me.lblSimPoll_ms = New System.Windows.Forms.Label()
        Me.txtSimPoll_ms = New System.Windows.Forms.TextBox()
        Me.chkAutoJoy = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(205, 196)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(66, 27)
        Me.btnSave.TabIndex = 12
        Me.btnSave.Text = "&Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'chkStartMinim
        '
        Me.chkStartMinim.AutoSize = True
        Me.chkStartMinim.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkStartMinim.Location = New System.Drawing.Point(12, 12)
        Me.chkStartMinim.MinimumSize = New System.Drawing.Size(179, 0)
        Me.chkStartMinim.Name = "chkStartMinim"
        Me.chkStartMinim.Size = New System.Drawing.Size(179, 17)
        Me.chkStartMinim.TabIndex = 33
        Me.chkStartMinim.Text = "Start Minimised:"
        Me.chkStartMinim.UseVisualStyleBackColor = True
        '
        'lblPort
        '
        Me.lblPort.AutoSize = True
        Me.lblPort.Location = New System.Drawing.Point(10, 94)
        Me.lblPort.Name = "lblPort"
        Me.lblPort.Size = New System.Drawing.Size(63, 13)
        Me.lblPort.TabIndex = 35
        Me.lblPort.Text = "Server Port:"
        '
        'txtPort
        '
        Me.txtPort.Location = New System.Drawing.Point(177, 91)
        Me.txtPort.Name = "txtPort"
        Me.txtPort.Size = New System.Drawing.Size(95, 20)
        Me.txtPort.TabIndex = 34
        '
        'chkStartServer
        '
        Me.chkStartServer.AutoSize = True
        Me.chkStartServer.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkStartServer.Location = New System.Drawing.Point(12, 117)
        Me.chkStartServer.MinimumSize = New System.Drawing.Size(179, 0)
        Me.chkStartServer.Name = "chkStartServer"
        Me.chkStartServer.Size = New System.Drawing.Size(179, 17)
        Me.chkStartServer.TabIndex = 36
        Me.chkStartServer.Text = "AutoStart Server:"
        Me.chkStartServer.UseVisualStyleBackColor = True
        '
        'chkAutoSim
        '
        Me.chkAutoSim.AutoSize = True
        Me.chkAutoSim.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkAutoSim.Location = New System.Drawing.Point(12, 35)
        Me.chkAutoSim.MinimumSize = New System.Drawing.Size(179, 0)
        Me.chkAutoSim.Name = "chkAutoSim"
        Me.chkAutoSim.Size = New System.Drawing.Size(179, 17)
        Me.chkAutoSim.TabIndex = 37
        Me.chkAutoSim.Text = "AutoStart SimConnect:"
        Me.chkAutoSim.UseVisualStyleBackColor = True
        '
        'lblSimPoll_ms
        '
        Me.lblSimPoll_ms.AutoSize = True
        Me.lblSimPoll_ms.Location = New System.Drawing.Point(10, 157)
        Me.lblSimPoll_ms.Name = "lblSimPoll_ms"
        Me.lblSimPoll_ms.Size = New System.Drawing.Size(161, 13)
        Me.lblSimPoll_ms.TabIndex = 39
        Me.lblSimPoll_ms.Text = "Sim Connect Polling Interval(ms):"
        '
        'txtSimPoll_ms
        '
        Me.txtSimPoll_ms.Location = New System.Drawing.Point(177, 154)
        Me.txtSimPoll_ms.Name = "txtSimPoll_ms"
        Me.txtSimPoll_ms.Size = New System.Drawing.Size(95, 20)
        Me.txtSimPoll_ms.TabIndex = 38
        '
        'chkAutoJoy
        '
        Me.chkAutoJoy.AutoSize = True
        Me.chkAutoJoy.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkAutoJoy.Location = New System.Drawing.Point(12, 58)
        Me.chkAutoJoy.MinimumSize = New System.Drawing.Size(179, 0)
        Me.chkAutoJoy.Name = "chkAutoJoy"
        Me.chkAutoJoy.Size = New System.Drawing.Size(179, 17)
        Me.chkAutoJoy.TabIndex = 40
        Me.chkAutoJoy.Text = "AutoStart Joystick Mapping:"
        Me.chkAutoJoy.UseVisualStyleBackColor = True
        '
        'SettingsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(283, 235)
        Me.Controls.Add(Me.chkAutoJoy)
        Me.Controls.Add(Me.lblSimPoll_ms)
        Me.Controls.Add(Me.txtSimPoll_ms)
        Me.Controls.Add(Me.chkAutoSim)
        Me.Controls.Add(Me.chkStartServer)
        Me.Controls.Add(Me.lblPort)
        Me.Controls.Add(Me.txtPort)
        Me.Controls.Add(Me.chkStartMinim)
        Me.Controls.Add(Me.btnSave)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SettingsForm"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSave As Button
    Friend WithEvents chkStartMinim As CheckBox
    Friend WithEvents lblPort As Label
    Friend WithEvents txtPort As TextBox
    Friend WithEvents chkStartServer As CheckBox
    Friend WithEvents chkAutoSim As CheckBox
    Friend WithEvents lblSimPoll_ms As Label
    Friend WithEvents txtSimPoll_ms As TextBox
    Friend WithEvents chkAutoJoy As CheckBox
End Class
