<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SimConnectDebugForm
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
        Me.btnTransmit = New System.Windows.Forms.Button()
        Me.lblEventID = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.btnTestHello = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.Button10 = New System.Windows.Forms.Button()
        Me.cbEventID = New System.Windows.Forms.ComboBox()
        Me.txtDisplay = New System.Windows.Forms.TextBox()
        Me.tmrPoll = New System.Windows.Forms.Timer(Me.components)
        Me.Button11 = New System.Windows.Forms.Button()
        Me.Button12 = New System.Windows.Forms.Button()
        Me.Button13 = New System.Windows.Forms.Button()
        Me.lblSimVar = New System.Windows.Forms.Label()
        Me.txtSimVar = New System.Windows.Forms.TextBox()
        Me.txtSimVarVal = New System.Windows.Forms.TextBox()
        Me.chkContinuous = New System.Windows.Forms.CheckBox()
        Me.btnPoll = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnTransmit
        '
        Me.btnTransmit.Location = New System.Drawing.Point(338, 9)
        Me.btnTransmit.Name = "btnTransmit"
        Me.btnTransmit.Size = New System.Drawing.Size(104, 25)
        Me.btnTransmit.TabIndex = 14
        Me.btnTransmit.Text = "&Transmit"
        Me.btnTransmit.UseVisualStyleBackColor = True
        '
        'lblEventID
        '
        Me.lblEventID.AutoSize = True
        Me.lblEventID.Location = New System.Drawing.Point(12, 15)
        Me.lblEventID.Name = "lblEventID"
        Me.lblEventID.Size = New System.Drawing.Size(52, 13)
        Me.lblEventID.TabIndex = 16
        Me.lblEventID.Text = "Event ID:"
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(135, 224)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(104, 25)
        Me.Button1.TabIndex = 17
        Me.Button1.Text = "L"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(11, 224)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(104, 25)
        Me.Button2.TabIndex = 18
        Me.Button2.Text = "SH-CTL-F4"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'btnTestHello
        '
        Me.btnTestHello.Location = New System.Drawing.Point(463, 82)
        Me.btnTestHello.Name = "btnTestHello"
        Me.btnTestHello.Size = New System.Drawing.Size(104, 25)
        Me.btnTestHello.TabIndex = 19
        Me.btnTestHello.Text = "Hello"
        Me.btnTestHello.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button4.Location = New System.Drawing.Point(11, 255)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(104, 25)
        Me.Button4.TabIndex = 20
        Me.Button4.Text = "Tab"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button5.Location = New System.Drawing.Point(135, 255)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(104, 25)
        Me.Button5.TabIndex = 21
        Me.Button5.Text = "Space"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button6.Location = New System.Drawing.Point(11, 286)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(104, 25)
        Me.Button6.TabIndex = 22
        Me.Button6.Text = "Enter"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button7.Location = New System.Drawing.Point(135, 286)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(104, 25)
        Me.Button7.TabIndex = 23
        Me.Button7.Text = "KP0"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button8.Location = New System.Drawing.Point(258, 286)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(104, 25)
        Me.Button8.TabIndex = 24
        Me.Button8.Text = "KP/"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Button9
        '
        Me.Button9.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button9.Location = New System.Drawing.Point(258, 224)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(104, 25)
        Me.Button9.TabIndex = 25
        Me.Button9.Text = "Up Arrow"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'Button10
        '
        Me.Button10.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button10.Location = New System.Drawing.Point(258, 255)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(104, 25)
        Me.Button10.TabIndex = 26
        Me.Button10.Text = "Down Arrow"
        Me.Button10.UseVisualStyleBackColor = True
        '
        'cbEventID
        '
        Me.cbEventID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbEventID.FormattingEnabled = True
        Me.cbEventID.Location = New System.Drawing.Point(70, 12)
        Me.cbEventID.Name = "cbEventID"
        Me.cbEventID.Size = New System.Drawing.Size(253, 21)
        Me.cbEventID.TabIndex = 27
        '
        'txtDisplay
        '
        Me.txtDisplay.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDisplay.Enabled = False
        Me.txtDisplay.Location = New System.Drawing.Point(15, 72)
        Me.txtDisplay.Multiline = True
        Me.txtDisplay.Name = "txtDisplay"
        Me.txtDisplay.ReadOnly = True
        Me.txtDisplay.Size = New System.Drawing.Size(442, 144)
        Me.txtDisplay.TabIndex = 30
        '
        'tmrPoll
        '
        Me.tmrPoll.Interval = 500
        '
        'Button11
        '
        Me.Button11.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button11.Location = New System.Drawing.Point(463, 113)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(104, 25)
        Me.Button11.TabIndex = 31
        Me.Button11.Text = "Flip Heading 180"
        Me.Button11.UseVisualStyleBackColor = True
        '
        'Button12
        '
        Me.Button12.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button12.Location = New System.Drawing.Point(463, 144)
        Me.Button12.Name = "Button12"
        Me.Button12.Size = New System.Drawing.Size(104, 25)
        Me.Button12.TabIndex = 32
        Me.Button12.Text = "Flip Heading 90"
        Me.Button12.UseVisualStyleBackColor = True
        '
        'Button13
        '
        Me.Button13.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button13.Location = New System.Drawing.Point(463, 175)
        Me.Button13.Name = "Button13"
        Me.Button13.Size = New System.Drawing.Size(104, 25)
        Me.Button13.TabIndex = 33
        Me.Button13.Text = "Flip Heading -90"
        Me.Button13.UseVisualStyleBackColor = True
        '
        'lblSimVar
        '
        Me.lblSimVar.AutoSize = True
        Me.lblSimVar.Location = New System.Drawing.Point(12, 42)
        Me.lblSimVar.Name = "lblSimVar"
        Me.lblSimVar.Size = New System.Drawing.Size(43, 13)
        Me.lblSimVar.TabIndex = 34
        Me.lblSimVar.Text = "SimVar:"
        '
        'txtSimVar
        '
        Me.txtSimVar.Location = New System.Drawing.Point(70, 39)
        Me.txtSimVar.Name = "txtSimVar"
        Me.txtSimVar.Size = New System.Drawing.Size(253, 20)
        Me.txtSimVar.TabIndex = 35
        '
        'txtSimVarVal
        '
        Me.txtSimVarVal.Location = New System.Drawing.Point(329, 39)
        Me.txtSimVarVal.Name = "txtSimVarVal"
        Me.txtSimVarVal.ReadOnly = True
        Me.txtSimVarVal.Size = New System.Drawing.Size(113, 20)
        Me.txtSimVarVal.TabIndex = 36
        '
        'chkContinuous
        '
        Me.chkContinuous.AutoSize = True
        Me.chkContinuous.Location = New System.Drawing.Point(491, 41)
        Me.chkContinuous.Name = "chkContinuous"
        Me.chkContinuous.Size = New System.Drawing.Size(79, 17)
        Me.chkContinuous.TabIndex = 37
        Me.chkContinuous.Text = "Continuous"
        Me.chkContinuous.UseVisualStyleBackColor = True
        '
        'btnPoll
        '
        Me.btnPoll.Location = New System.Drawing.Point(448, 36)
        Me.btnPoll.Name = "btnPoll"
        Me.btnPoll.Size = New System.Drawing.Size(38, 25)
        Me.btnPoll.TabIndex = 38
        Me.btnPoll.Text = "Poll"
        Me.btnPoll.UseVisualStyleBackColor = True
        '
        'SimConnectDebugForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(580, 321)
        Me.Controls.Add(Me.btnPoll)
        Me.Controls.Add(Me.chkContinuous)
        Me.Controls.Add(Me.txtSimVarVal)
        Me.Controls.Add(Me.txtSimVar)
        Me.Controls.Add(Me.lblSimVar)
        Me.Controls.Add(Me.Button13)
        Me.Controls.Add(Me.Button12)
        Me.Controls.Add(Me.Button11)
        Me.Controls.Add(Me.txtDisplay)
        Me.Controls.Add(Me.cbEventID)
        Me.Controls.Add(Me.Button10)
        Me.Controls.Add(Me.Button9)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.btnTestHello)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnTransmit)
        Me.Controls.Add(Me.lblEventID)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "SimConnectDebugForm"
        Me.Text = "SimConnect Debug Form"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnTransmit As Button
    Friend WithEvents lblEventID As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents btnTestHello As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents Button8 As Button
    Friend WithEvents Button9 As Button
    Friend WithEvents Button10 As Button
    Friend WithEvents cbEventID As ComboBox
    Friend WithEvents txtDisplay As TextBox
    Friend WithEvents tmrPoll As Timer
    Friend WithEvents Button11 As Button
    Friend WithEvents Button12 As Button
    Friend WithEvents Button13 As Button
    Friend WithEvents lblSimVar As Label
    Friend WithEvents txtSimVar As TextBox
    Friend WithEvents txtSimVarVal As TextBox
    Friend WithEvents chkContinuous As CheckBox
    Friend WithEvents btnPoll As Button
End Class
