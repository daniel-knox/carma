<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SignalDashboard
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
        Me.txtbx_Current = New System.Windows.Forms.TextBox()
        Me.lbl_Current = New System.Windows.Forms.Label()
        Me.lbl_Phase = New System.Windows.Forms.Label()
        Me.txtbx_Phase_Degrees = New System.Windows.Forms.TextBox()
        Me.lbl_PowerFactor = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.lbl_Rail_v5 = New System.Windows.Forms.Label()
        Me.txtbx_Voltage_v5 = New System.Windows.Forms.TextBox()
        Me.lbl_Rail_v3 = New System.Windows.Forms.Label()
        Me.txtbx_Voltage_v3 = New System.Windows.Forms.TextBox()
        Me.lbl_Rail_v1 = New System.Windows.Forms.Label()
        Me.txtbox_Voltage_v1 = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtbx_Current
        '
        Me.txtbx_Current.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtbx_Current.Location = New System.Drawing.Point(197, 44)
        Me.txtbx_Current.Name = "txtbx_Current"
        Me.txtbx_Current.Size = New System.Drawing.Size(120, 23)
        Me.txtbx_Current.TabIndex = 0
        '
        'lbl_Current
        '
        Me.lbl_Current.AutoSize = True
        Me.lbl_Current.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Current.Location = New System.Drawing.Point(6, 47)
        Me.lbl_Current.Name = "lbl_Current"
        Me.lbl_Current.Size = New System.Drawing.Size(81, 16)
        Me.lbl_Current.TabIndex = 1
        Me.lbl_Current.Text = "Current (A)"
        '
        'lbl_Phase
        '
        Me.lbl_Phase.AutoSize = True
        Me.lbl_Phase.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Phase.Location = New System.Drawing.Point(6, 82)
        Me.lbl_Phase.Name = "lbl_Phase"
        Me.lbl_Phase.Size = New System.Drawing.Size(156, 16)
        Me.lbl_Phase.TabIndex = 3
        Me.lbl_Phase.Text = "Phase Degrees (Lag)"
        '
        'txtbx_Phase_Degrees
        '
        Me.txtbx_Phase_Degrees.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtbx_Phase_Degrees.Location = New System.Drawing.Point(197, 79)
        Me.txtbx_Phase_Degrees.Name = "txtbx_Phase_Degrees"
        Me.txtbx_Phase_Degrees.Size = New System.Drawing.Size(120, 23)
        Me.txtbx_Phase_Degrees.TabIndex = 2
        '
        'lbl_PowerFactor
        '
        Me.lbl_PowerFactor.AutoSize = True
        Me.lbl_PowerFactor.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_PowerFactor.Location = New System.Drawing.Point(6, 120)
        Me.lbl_PowerFactor.Name = "lbl_PowerFactor"
        Me.lbl_PowerFactor.Size = New System.Drawing.Size(132, 16)
        Me.lbl_PowerFactor.TabIndex = 5
        Me.lbl_PowerFactor.Text = "Power Factor (PF)"
        '
        'TextBox2
        '
        Me.TextBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(197, 117)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(120, 23)
        Me.TextBox2.TabIndex = 4
        '
        'lbl_Rail_v5
        '
        Me.lbl_Rail_v5.AutoSize = True
        Me.lbl_Rail_v5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Rail_v5.Location = New System.Drawing.Point(353, 47)
        Me.lbl_Rail_v5.Name = "lbl_Rail_v5"
        Me.lbl_Rail_v5.Size = New System.Drawing.Size(30, 16)
        Me.lbl_Rail_v5.TabIndex = 7
        Me.lbl_Rail_v5.Text = "V.5"
        '
        'txtbx_Voltage_v5
        '
        Me.txtbx_Voltage_v5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtbx_Voltage_v5.Location = New System.Drawing.Point(403, 44)
        Me.txtbx_Voltage_v5.Name = "txtbx_Voltage_v5"
        Me.txtbx_Voltage_v5.Size = New System.Drawing.Size(120, 23)
        Me.txtbx_Voltage_v5.TabIndex = 6
        '
        'lbl_Rail_v3
        '
        Me.lbl_Rail_v3.AutoSize = True
        Me.lbl_Rail_v3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Rail_v3.Location = New System.Drawing.Point(353, 85)
        Me.lbl_Rail_v3.Name = "lbl_Rail_v3"
        Me.lbl_Rail_v3.Size = New System.Drawing.Size(30, 16)
        Me.lbl_Rail_v3.TabIndex = 9
        Me.lbl_Rail_v3.Text = "V.3"
        '
        'txtbx_Voltage_v3
        '
        Me.txtbx_Voltage_v3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtbx_Voltage_v3.Location = New System.Drawing.Point(403, 82)
        Me.txtbx_Voltage_v3.Name = "txtbx_Voltage_v3"
        Me.txtbx_Voltage_v3.Size = New System.Drawing.Size(120, 23)
        Me.txtbx_Voltage_v3.TabIndex = 8
        '
        'lbl_Rail_v1
        '
        Me.lbl_Rail_v1.AutoSize = True
        Me.lbl_Rail_v1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Rail_v1.Location = New System.Drawing.Point(353, 123)
        Me.lbl_Rail_v1.Name = "lbl_Rail_v1"
        Me.lbl_Rail_v1.Size = New System.Drawing.Size(30, 16)
        Me.lbl_Rail_v1.TabIndex = 11
        Me.lbl_Rail_v1.Text = "V.1"
        '
        'txtbox_Voltage_v1
        '
        Me.txtbox_Voltage_v1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtbox_Voltage_v1.Location = New System.Drawing.Point(403, 120)
        Me.txtbox_Voltage_v1.Name = "txtbox_Voltage_v1"
        Me.txtbox_Voltage_v1.Size = New System.Drawing.Size(120, 23)
        Me.txtbox_Voltage_v1.TabIndex = 10
        '
        'SignalDashboard
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(833, 260)
        Me.Controls.Add(Me.lbl_Rail_v1)
        Me.Controls.Add(Me.txtbox_Voltage_v1)
        Me.Controls.Add(Me.lbl_Rail_v3)
        Me.Controls.Add(Me.txtbx_Voltage_v3)
        Me.Controls.Add(Me.lbl_Rail_v5)
        Me.Controls.Add(Me.txtbx_Voltage_v5)
        Me.Controls.Add(Me.lbl_PowerFactor)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.lbl_Phase)
        Me.Controls.Add(Me.txtbx_Phase_Degrees)
        Me.Controls.Add(Me.lbl_Current)
        Me.Controls.Add(Me.txtbx_Current)
        Me.Name = "SignalDashboard"
        Me.Text = "SignalDashboard"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtbx_Current As System.Windows.Forms.TextBox
    Friend WithEvents lbl_Current As System.Windows.Forms.Label
    Friend WithEvents lbl_Phase As System.Windows.Forms.Label
    Friend WithEvents txtbx_Phase_Degrees As System.Windows.Forms.TextBox
    Friend WithEvents lbl_PowerFactor As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents lbl_Rail_v5 As System.Windows.Forms.Label
    Friend WithEvents txtbx_Voltage_v5 As System.Windows.Forms.TextBox
    Friend WithEvents lbl_Rail_v3 As System.Windows.Forms.Label
    Friend WithEvents txtbx_Voltage_v3 As System.Windows.Forms.TextBox
    Friend WithEvents lbl_Rail_v1 As System.Windows.Forms.Label
    Friend WithEvents txtbox_Voltage_v1 As System.Windows.Forms.TextBox
End Class
