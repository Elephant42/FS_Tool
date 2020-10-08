<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.ttDesc = New System.Windows.Forms.ToolTip(Me.components)
        Me.tmrUpdateTimer = New System.Windows.Forms.Timer(Me.components)
        Me.ttDesc2 = New System.Windows.Forms.ToolTip(Me.components)
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.cmTrayIcon = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ctx1Item_Slew = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.ctx1Item_Rate_x8 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ctx1Item_Rate_x4 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ctx1Item_Rate_x2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ctx1Item_Rate_x1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.ctx1Item_Restore = New System.Windows.Forms.ToolStripMenuItem()
        Me.ctx1Item_Minimise = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator8 = New System.Windows.Forms.ToolStripSeparator()
        Me.ctx1Item_Exit = New System.Windows.Forms.ToolStripMenuItem()
        Me.gbSimRate = New System.Windows.Forms.GroupBox()
        Me.cbSimRate = New System.Windows.Forms.ComboBox()
        Me.btnActivateSimRate = New System.Windows.Forms.Button()
        Me.lblNewSimRate = New System.Windows.Forms.Label()
        Me.lblCurrSimRate = New System.Windows.Forms.Label()
        Me.txtCurrSimRate = New System.Windows.Forms.TextBox()
        Me.gbSlew = New System.Windows.Forms.GroupBox()
        Me.btnSlew = New System.Windows.Forms.Button()
        Me.btnClip = New System.Windows.Forms.Button()
        Me.lblNewLong = New System.Windows.Forms.Label()
        Me.txtNewLat = New System.Windows.Forms.TextBox()
        Me.txtNewLong = New System.Windows.Forms.TextBox()
        Me.lblNewLat = New System.Windows.Forms.Label()
        Me.txtDisplay = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ctx1Item_Settings = New System.Windows.Forms.ToolStripMenuItem()
        Me.cmTrayIcon.SuspendLayout()
        Me.gbSimRate.SuspendLayout()
        Me.gbSlew.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tmrUpdateTimer
        '
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenuStrip = Me.cmTrayIcon
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Flight Sim Tool"
        Me.NotifyIcon1.Visible = True
        '
        'cmTrayIcon
        '
        Me.cmTrayIcon.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ctx1Item_Slew, Me.ToolStripSeparator6, Me.ctx1Item_Rate_x8, Me.ctx1Item_Rate_x4, Me.ctx1Item_Rate_x2, Me.ctx1Item_Rate_x1, Me.ToolStripSeparator5, Me.ctx1Item_Settings, Me.ToolStripSeparator1, Me.ctx1Item_Restore, Me.ctx1Item_Minimise, Me.ToolStripSeparator8, Me.ctx1Item_Exit})
        Me.cmTrayIcon.Name = "cmTrayIcon"
        Me.cmTrayIcon.Size = New System.Drawing.Size(181, 248)
        Me.cmTrayIcon.Text = "Actions"
        '
        'ctx1Item_Slew
        '
        Me.ctx1Item_Slew.Name = "ctx1Item_Slew"
        Me.ctx1Item_Slew.Size = New System.Drawing.Size(180, 22)
        Me.ctx1Item_Slew.Text = "Slew To Clipboard"
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        Me.ToolStripSeparator6.Size = New System.Drawing.Size(177, 6)
        '
        'ctx1Item_Rate_x8
        '
        Me.ctx1Item_Rate_x8.Name = "ctx1Item_Rate_x8"
        Me.ctx1Item_Rate_x8.Size = New System.Drawing.Size(180, 22)
        Me.ctx1Item_Rate_x8.Text = "Sim Rate x8"
        '
        'ctx1Item_Rate_x4
        '
        Me.ctx1Item_Rate_x4.Name = "ctx1Item_Rate_x4"
        Me.ctx1Item_Rate_x4.Size = New System.Drawing.Size(180, 22)
        Me.ctx1Item_Rate_x4.Text = "Sim Rate x4"
        '
        'ctx1Item_Rate_x2
        '
        Me.ctx1Item_Rate_x2.Name = "ctx1Item_Rate_x2"
        Me.ctx1Item_Rate_x2.Size = New System.Drawing.Size(180, 22)
        Me.ctx1Item_Rate_x2.Text = "Sim Rate x2"
        '
        'ctx1Item_Rate_x1
        '
        Me.ctx1Item_Rate_x1.Name = "ctx1Item_Rate_x1"
        Me.ctx1Item_Rate_x1.Size = New System.Drawing.Size(180, 22)
        Me.ctx1Item_Rate_x1.Text = "Sim Rate x1"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(177, 6)
        '
        'ctx1Item_Restore
        '
        Me.ctx1Item_Restore.Name = "ctx1Item_Restore"
        Me.ctx1Item_Restore.Size = New System.Drawing.Size(180, 22)
        Me.ctx1Item_Restore.Text = "Restore"
        '
        'ctx1Item_Minimise
        '
        Me.ctx1Item_Minimise.Name = "ctx1Item_Minimise"
        Me.ctx1Item_Minimise.Size = New System.Drawing.Size(180, 22)
        Me.ctx1Item_Minimise.Text = "Minimise"
        '
        'ToolStripSeparator8
        '
        Me.ToolStripSeparator8.Name = "ToolStripSeparator8"
        Me.ToolStripSeparator8.Size = New System.Drawing.Size(177, 6)
        '
        'ctx1Item_Exit
        '
        Me.ctx1Item_Exit.Name = "ctx1Item_Exit"
        Me.ctx1Item_Exit.Size = New System.Drawing.Size(180, 22)
        Me.ctx1Item_Exit.Text = "Exit"
        '
        'gbSimRate
        '
        Me.gbSimRate.Controls.Add(Me.cbSimRate)
        Me.gbSimRate.Controls.Add(Me.btnActivateSimRate)
        Me.gbSimRate.Controls.Add(Me.lblNewSimRate)
        Me.gbSimRate.Controls.Add(Me.lblCurrSimRate)
        Me.gbSimRate.Controls.Add(Me.txtCurrSimRate)
        Me.gbSimRate.Location = New System.Drawing.Point(12, 349)
        Me.gbSimRate.Name = "gbSimRate"
        Me.gbSimRate.Size = New System.Drawing.Size(420, 56)
        Me.gbSimRate.TabIndex = 20
        Me.gbSimRate.TabStop = False
        Me.gbSimRate.Text = "Sim Rate"
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
        Me.gbSlew.TabIndex = 19
        Me.gbSlew.TabStop = False
        Me.gbSlew.Text = "Slew Coordinates"
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
        'btnClip
        '
        Me.btnClip.Location = New System.Drawing.Point(331, 19)
        Me.btnClip.Name = "btnClip"
        Me.btnClip.Size = New System.Drawing.Size(73, 46)
        Me.btnClip.TabIndex = 2
        Me.btnClip.Text = "&Copy From Clipboard"
        Me.btnClip.UseVisualStyleBackColor = True
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
        'txtNewLat
        '
        Me.txtNewLat.Location = New System.Drawing.Point(96, 19)
        Me.txtNewLat.Name = "txtNewLat"
        Me.txtNewLat.Size = New System.Drawing.Size(229, 20)
        Me.txtNewLat.TabIndex = 10
        '
        'txtNewLong
        '
        Me.txtNewLong.Location = New System.Drawing.Point(96, 45)
        Me.txtNewLong.Name = "txtNewLong"
        Me.txtNewLong.Size = New System.Drawing.Size(229, 20)
        Me.txtNewLong.TabIndex = 12
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
        'txtDisplay
        '
        Me.txtDisplay.Enabled = False
        Me.txtDisplay.Location = New System.Drawing.Point(12, 12)
        Me.txtDisplay.Multiline = True
        Me.txtDisplay.Name = "txtDisplay"
        Me.txtDisplay.ReadOnly = True
        Me.txtDisplay.Size = New System.Drawing.Size(422, 153)
        Me.txtDisplay.TabIndex = 18
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 287)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(420, 56)
        Me.GroupBox1.TabIndex = 21
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Sample Commands"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(144, 19)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(124, 25)
        Me.Button2.TabIndex = 19
        Me.Button2.Text = "Toggle Auto Pilot"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(11, 19)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(124, 25)
        Me.Button1.TabIndex = 18
        Me.Button1.Text = "Toggle Parking Brake"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(177, 6)
        '
        'ctx1Item_Settings
        '
        Me.ctx1Item_Settings.Name = "ctx1Item_Settings"
        Me.ctx1Item_Settings.Size = New System.Drawing.Size(180, 22)
        Me.ctx1Item_Settings.Text = "Settings"
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(442, 412)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.gbSimRate)
        Me.Controls.Add(Me.gbSlew)
        Me.Controls.Add(Me.txtDisplay)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "MainForm"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "Flight Sim Tool"
        Me.cmTrayIcon.ResumeLayout(False)
        Me.gbSimRate.ResumeLayout(False)
        Me.gbSimRate.PerformLayout()
        Me.gbSlew.ResumeLayout(False)
        Me.gbSlew.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ttDesc As System.Windows.Forms.ToolTip
    Friend WithEvents tmrUpdateTimer As System.Windows.Forms.Timer
    Friend WithEvents ttDesc2 As System.Windows.Forms.ToolTip
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents cmTrayIcon As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ctx1Item_Exit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator6 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator8 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ctx1Item_Minimise As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ctx1Item_Restore As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ctx1Item_Rate_x8 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ctx1Item_Rate_x2 As ToolStripMenuItem
    Friend WithEvents ctx1Item_Rate_x1 As ToolStripMenuItem
    Friend WithEvents ctx1Item_Rate_x4 As ToolStripMenuItem
    Friend WithEvents gbSimRate As GroupBox
    Friend WithEvents cbSimRate As ComboBox
    Friend WithEvents btnActivateSimRate As Button
    Friend WithEvents lblNewSimRate As Label
    Friend WithEvents lblCurrSimRate As Label
    Friend WithEvents txtCurrSimRate As TextBox
    Friend WithEvents gbSlew As GroupBox
    Friend WithEvents btnSlew As Button
    Friend WithEvents btnClip As Button
    Friend WithEvents lblNewLong As Label
    Friend WithEvents txtNewLat As TextBox
    Friend WithEvents txtNewLong As TextBox
    Friend WithEvents lblNewLat As Label
    Friend WithEvents txtDisplay As TextBox
    Friend WithEvents ctx1Item_Slew As ToolStripMenuItem
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Button2 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents ctx1Item_Settings As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
End Class
