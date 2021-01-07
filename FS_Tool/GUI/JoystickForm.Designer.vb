<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class JoystickForm
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
        Me.cbJoysticks = New System.Windows.Forms.ComboBox()
        Me.txtDisplay = New System.Windows.Forms.TextBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.btnClose = New System.Windows.Forms.Button()
        Me.cbThrottle = New System.Windows.Forms.CheckBox()
        Me.cbIgnoreAxes = New System.Windows.Forms.CheckBox()
        Me.btnClipboard = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cbJoysticks
        '
        Me.cbJoysticks.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbJoysticks.FormattingEnabled = True
        Me.cbJoysticks.Location = New System.Drawing.Point(12, 12)
        Me.cbJoysticks.Name = "cbJoysticks"
        Me.cbJoysticks.Size = New System.Drawing.Size(578, 21)
        Me.cbJoysticks.TabIndex = 0
        '
        'txtDisplay
        '
        Me.txtDisplay.Font = New System.Drawing.Font("Cascadia Code", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDisplay.Location = New System.Drawing.Point(12, 64)
        Me.txtDisplay.Multiline = True
        Me.txtDisplay.Name = "txtDisplay"
        Me.txtDisplay.ReadOnly = True
        Me.txtDisplay.Size = New System.Drawing.Size(675, 586)
        Me.txtDisplay.TabIndex = 1
        '
        'Timer1
        '
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(612, 11)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 21)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'cbThrottle
        '
        Me.cbThrottle.AutoSize = True
        Me.cbThrottle.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cbThrottle.Location = New System.Drawing.Point(12, 39)
        Me.cbThrottle.Name = "cbThrottle"
        Me.cbThrottle.Size = New System.Drawing.Size(190, 17)
        Me.cbThrottle.TabIndex = 3
        Me.cbThrottle.Text = "Throttle display updates to 10/sec:"
        Me.cbThrottle.UseVisualStyleBackColor = True
        '
        'cbIgnoreAxes
        '
        Me.cbIgnoreAxes.AutoSize = True
        Me.cbIgnoreAxes.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cbIgnoreAxes.Location = New System.Drawing.Point(281, 39)
        Me.cbIgnoreAxes.Name = "cbIgnoreAxes"
        Me.cbIgnoreAxes.Size = New System.Drawing.Size(84, 17)
        Me.cbIgnoreAxes.TabIndex = 4
        Me.cbIgnoreAxes.Text = "Ignore axes:"
        Me.cbIgnoreAxes.UseVisualStyleBackColor = True
        '
        'btnClipboard
        '
        Me.btnClipboard.Location = New System.Drawing.Point(414, 39)
        Me.btnClipboard.Name = "btnClipboard"
        Me.btnClipboard.Size = New System.Drawing.Size(176, 21)
        Me.btnClipboard.TabIndex = 5
        Me.btnClipboard.Text = "Copy Blank Profile to Clipboard"
        Me.btnClipboard.UseVisualStyleBackColor = True
        '
        'JoystickForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(694, 662)
        Me.Controls.Add(Me.btnClipboard)
        Me.Controls.Add(Me.cbIgnoreAxes)
        Me.Controls.Add(Me.cbThrottle)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.txtDisplay)
        Me.Controls.Add(Me.cbJoysticks)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "JoystickForm"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "JoystickForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cbJoysticks As ComboBox
    Friend WithEvents txtDisplay As TextBox
    Friend WithEvents Timer1 As Timer
    Friend WithEvents btnClose As Button
    Friend WithEvents cbThrottle As CheckBox
    Friend WithEvents cbIgnoreAxes As CheckBox
    Friend WithEvents btnClipboard As Button
End Class
