<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainFormX
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.tmrUpdateTimer = New System.Windows.Forms.Timer(Me.components)
        Me.txtDisplay = New System.Windows.Forms.TextBox()
        Me.btnClip = New System.Windows.Forms.Button()
        Me.btnSlew = New System.Windows.Forms.Button()
        Me.lblNewLat = New System.Windows.Forms.Label()
        Me.txtNewLat = New System.Windows.Forms.TextBox()
        Me.lblNewLong = New System.Windows.Forms.Label()
        Me.txtNewLong = New System.Windows.Forms.TextBox()
        Me.gbSlew = New System.Windows.Forms.GroupBox()
        Me.gbSimRate = New System.Windows.Forms.GroupBox()
        Me.btnActivateSimRate = New System.Windows.Forms.Button()
        Me.lblNewSimRate = New System.Windows.Forms.Label()
        Me.lblCurrSimRate = New System.Windows.Forms.Label()
        Me.txtCurrSimRate = New System.Windows.Forms.TextBox()
        Me.cbSimRate = New System.Windows.Forms.ComboBox()
        Me.btnIncr = New System.Windows.Forms.Button()
        Me.btnDecr = New System.Windows.Forms.Button()
        Me.gbSlew.SuspendLayout()
        Me.gbSimRate.SuspendLayout()
        Me.SuspendLayout()
        '
        'tmrUpdateTimer
        '
        '
        'txtDisplay
        '
        Me.txtDisplay.Enabled = False
        Me.txtDisplay.Location = New System.Drawing.Point(12, 12)
        Me.txtDisplay.Multiline = True
        Me.txtDisplay.Name = "txtDisplay"
        Me.txtDisplay.ReadOnly = True
        Me.txtDisplay.Size = New System.Drawing.Size(422, 153)
        Me.txtDisplay.TabIndex = 1
        '
        'btnClip
        '
        Me.btnClip.Location = New System.Drawing.Point(331, 19)
        Me.btnClip.Name = "btnClip"
        Me.btnClip.Size = New System.Drawing.Size(73, 46)
        Me.btnClip.TabIndex = 2
        Me.btnClip.Text = "&Copy From Clipboard"
        Me.btnClip.UseVisualStyleBackColor = True
        '
        'btnSlew
        '
        Me.btnSlew.Location = New System.Drawing.Point(96, 71)
        Me.btnSlew.Name = "btnSlew"
        Me.btnSlew.Size = New System.Drawing.Size(104, 25)
        Me.btnSlew.TabIndex = 3
        Me.btnSlew.Text = "&Slew"
        Me.btnSlew.UseVisualStyleBackColor = True
        '
        'lblNewLat
        '
        Me.lblNewLat.AutoSize = True
        Me.lblNewLat.Location = New System.Drawing.Point(8, 22)
        Me.lblNewLat.Name = "lblNewLat"
        Me.lblNewLat.Size = New System.Drawing.Size(73, 13)
        Me.lblNewLat.TabIndex = 11
        Me.lblNewLat.Text = "New Latitude:"
        '
        'txtNewLat
        '
        Me.txtNewLat.Location = New System.Drawing.Point(96, 19)
        Me.txtNewLat.Name = "txtNewLat"
        Me.txtNewLat.Size = New System.Drawing.Size(229, 20)
        Me.txtNewLat.TabIndex = 10
        '
        'lblNewLong
        '
        Me.lblNewLong.AutoSize = True
        Me.lblNewLong.Location = New System.Drawing.Point(8, 48)
        Me.lblNewLong.Name = "lblNewLong"
        Me.lblNewLong.Size = New System.Drawing.Size(82, 13)
        Me.lblNewLong.TabIndex = 13
        Me.lblNewLong.Text = "New Longitude:"
        '
        'txtNewLong
        '
        Me.txtNewLong.Location = New System.Drawing.Point(96, 45)
        Me.txtNewLong.Name = "txtNewLong"
        Me.txtNewLong.Size = New System.Drawing.Size(229, 20)
        Me.txtNewLong.TabIndex = 12
        '
        'gbSlew
        '
        Me.gbSlew.Controls.Add(Me.btnSlew)
        Me.gbSlew.Controls.Add(Me.btnClip)
        Me.gbSlew.Controls.Add(Me.lblNewLong)
        Me.gbSlew.Controls.Add(Me.txtNewLat)
        Me.gbSlew.Controls.Add(Me.txtNewLong)
        Me.gbSlew.Controls.Add(Me.lblNewLat)
        Me.gbSlew.Location = New System.Drawing.Point(12, 171)
        Me.gbSlew.Name = "gbSlew"
        Me.gbSlew.Size = New System.Drawing.Size(422, 110)
        Me.gbSlew.TabIndex = 14
        Me.gbSlew.TabStop = False
        Me.gbSlew.Text = "Slew Coordinates"
        '
        'gbSimRate
        '
        Me.gbSimRate.Controls.Add(Me.btnIncr)
        Me.gbSimRate.Controls.Add(Me.btnDecr)
        Me.gbSimRate.Controls.Add(Me.cbSimRate)
        Me.gbSimRate.Controls.Add(Me.btnActivateSimRate)
        Me.gbSimRate.Controls.Add(Me.lblNewSimRate)
        Me.gbSimRate.Controls.Add(Me.lblCurrSimRate)
        Me.gbSimRate.Controls.Add(Me.txtCurrSimRate)
        Me.gbSimRate.Location = New System.Drawing.Point(12, 287)
        Me.gbSimRate.Name = "gbSimRate"
        Me.gbSimRate.Size = New System.Drawing.Size(420, 56)
        Me.gbSimRate.TabIndex = 15
        Me.gbSimRate.TabStop = False
        Me.gbSimRate.Text = "Sim Rate"
        '
        'btnActivateSimRate
        '
        Me.btnActivateSimRate.Location = New System.Drawing.Point(344, 16)
        Me.btnActivateSimRate.Name = "btnActivateSimRate"
        Me.btnActivateSimRate.Size = New System.Drawing.Size(70, 25)
        Me.btnActivateSimRate.TabIndex = 18
        Me.btnActivateSimRate.Text = "&Activate"
        Me.btnActivateSimRate.UseVisualStyleBackColor = True
        '
        'lblNewSimRate
        '
        Me.lblNewSimRate.AutoSize = True
        Me.lblNewSimRate.Location = New System.Drawing.Point(141, 22)
        Me.lblNewSimRate.Name = "lblNewSimRate"
        Me.lblNewSimRate.Size = New System.Drawing.Size(32, 13)
        Me.lblNewSimRate.TabIndex = 17
        Me.lblNewSimRate.Text = "New:"
        '
        'lblCurrSimRate
        '
        Me.lblCurrSimRate.AutoSize = True
        Me.lblCurrSimRate.Location = New System.Drawing.Point(8, 22)
        Me.lblCurrSimRate.Name = "lblCurrSimRate"
        Me.lblCurrSimRate.Size = New System.Drawing.Size(44, 13)
        Me.lblCurrSimRate.TabIndex = 15
        Me.lblCurrSimRate.Text = "Current:"
        '
        'txtCurrSimRate
        '
        Me.txtCurrSimRate.Enabled = False
        Me.txtCurrSimRate.Location = New System.Drawing.Point(58, 19)
        Me.txtCurrSimRate.Name = "txtCurrSimRate"
        Me.txtCurrSimRate.ReadOnly = True
        Me.txtCurrSimRate.Size = New System.Drawing.Size(65, 20)
        Me.txtCurrSimRate.TabIndex = 14
        '
        'cbSimRate
        '
        Me.cbSimRate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbSimRate.FormattingEnabled = True
        Me.cbSimRate.Items.AddRange(New Object() {"", "1", "2", "4", "8"})
        Me.cbSimRate.Location = New System.Drawing.Point(179, 19)
        Me.cbSimRate.Name = "cbSimRate"
        Me.cbSimRate.Size = New System.Drawing.Size(67, 21)
        Me.cbSimRate.TabIndex = 14
        '
        'btnIncr
        '
        Me.btnIncr.Location = New System.Drawing.Point(303, 16)
        Me.btnIncr.Name = "btnIncr"
        Me.btnIncr.Size = New System.Drawing.Size(28, 25)
        Me.btnIncr.TabIndex = 19
        Me.btnIncr.Text = "+"
        Me.btnIncr.UseVisualStyleBackColor = True
        '
        'btnDecr
        '
        Me.btnDecr.Location = New System.Drawing.Point(269, 16)
        Me.btnDecr.Name = "btnDecr"
        Me.btnDecr.Size = New System.Drawing.Size(28, 25)
        Me.btnDecr.TabIndex = 20
        Me.btnDecr.Text = "-"
        Me.btnDecr.UseVisualStyleBackColor = True
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(444, 353)
        Me.Controls.Add(Me.gbSimRate)
        Me.Controls.Add(Me.gbSlew)
        Me.Controls.Add(Me.txtDisplay)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "MainForm"
        Me.Text = "Flight Sim Tool"
        Me.gbSlew.ResumeLayout(False)
        Me.gbSlew.PerformLayout()
        Me.gbSimRate.ResumeLayout(False)
        Me.gbSimRate.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tmrUpdateTimer As Timer
    Friend WithEvents txtDisplay As TextBox
    Friend WithEvents btnClip As Button
    Friend WithEvents btnSlew As Button
    Friend WithEvents lblNewLat As Label
    Friend WithEvents txtNewLat As TextBox
    Friend WithEvents lblNewLong As Label
    Friend WithEvents txtNewLong As TextBox
    Friend WithEvents gbSlew As GroupBox
    Friend WithEvents gbSimRate As GroupBox
    Friend WithEvents btnActivateSimRate As Button
    Friend WithEvents lblNewSimRate As Label
    Friend WithEvents lblCurrSimRate As Label
    Friend WithEvents txtCurrSimRate As TextBox
    Friend WithEvents cbSimRate As ComboBox
    Friend WithEvents btnIncr As Button
    Friend WithEvents btnDecr As Button
End Class
