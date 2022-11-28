<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmModNOA
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents txtToDate As System.Windows.Forms.TextBox
	Public WithEvents txtFromDate As System.Windows.Forms.TextBox
	Public WithEvents txtFT As System.Windows.Forms.TextBox
	Public WithEvents txtNOA As System.Windows.Forms.TextBox
	Public WithEvents cmbDuplex As System.Windows.Forms.ComboBox
	Public WithEvents cmbPurge As System.Windows.Forms.ComboBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdModify As System.Windows.Forms.Button
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents label6 As System.Windows.Forms.Label
	Public WithEvents label5 As System.Windows.Forms.Label
	Public WithEvents label3 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtToDate = New System.Windows.Forms.TextBox
        Me.txtFromDate = New System.Windows.Forms.TextBox
        Me.txtFT = New System.Windows.Forms.TextBox
        Me.txtNOA = New System.Windows.Forms.TextBox
        Me.cmbDuplex = New System.Windows.Forms.ComboBox
        Me.cmbPurge = New System.Windows.Forms.ComboBox
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdModify = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.label6 = New System.Windows.Forms.Label
        Me.label5 = New System.Windows.Forms.Label
        Me.label3 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'txtToDate
        '
        Me.txtToDate.AcceptsReturn = True
        Me.txtToDate.BackColor = System.Drawing.SystemColors.Window
        Me.txtToDate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtToDate.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtToDate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtToDate.Location = New System.Drawing.Point(16, 136)
        Me.txtToDate.MaxLength = 0
        Me.txtToDate.Name = "txtToDate"
        Me.txtToDate.ReadOnly = True
        Me.txtToDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtToDate.Size = New System.Drawing.Size(105, 28)
        Me.txtToDate.TabIndex = 3
        '
        'txtFromDate
        '
        Me.txtFromDate.AcceptsReturn = True
        Me.txtFromDate.BackColor = System.Drawing.SystemColors.Window
        Me.txtFromDate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFromDate.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFromDate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFromDate.Location = New System.Drawing.Point(16, 96)
        Me.txtFromDate.MaxLength = 0
        Me.txtFromDate.Name = "txtFromDate"
        Me.txtFromDate.ReadOnly = True
        Me.txtFromDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFromDate.Size = New System.Drawing.Size(105, 28)
        Me.txtFromDate.TabIndex = 2
        '
        'txtFT
        '
        Me.txtFT.AcceptsReturn = True
        Me.txtFT.BackColor = System.Drawing.SystemColors.Window
        Me.txtFT.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFT.Enabled = False
        Me.txtFT.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFT.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFT.Location = New System.Drawing.Point(16, 56)
        Me.txtFT.MaxLength = 0
        Me.txtFT.Name = "txtFT"
        Me.txtFT.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFT.Size = New System.Drawing.Size(353, 28)
        Me.txtFT.TabIndex = 1
        Me.txtFT.Text = "NOA"
        '
        'txtNOA
        '
        Me.txtNOA.AcceptsReturn = True
        Me.txtNOA.BackColor = System.Drawing.SystemColors.Window
        Me.txtNOA.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNOA.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNOA.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNOA.Location = New System.Drawing.Point(16, 16)
        Me.txtNOA.MaxLength = 0
        Me.txtNOA.Name = "txtNOA"
        Me.txtNOA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNOA.Size = New System.Drawing.Size(49, 28)
        Me.txtNOA.TabIndex = 0
        '
        'cmbDuplex
        '
        Me.cmbDuplex.BackColor = System.Drawing.SystemColors.Window
        Me.cmbDuplex.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbDuplex.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbDuplex.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbDuplex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbDuplex.Items.AddRange(New Object() {"Yes", "No"})
        Me.cmbDuplex.Location = New System.Drawing.Point(16, 216)
        Me.cmbDuplex.Name = "cmbDuplex"
        Me.cmbDuplex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbDuplex.Size = New System.Drawing.Size(81, 26)
        Me.cmbDuplex.TabIndex = 5
        '
        'cmbPurge
        '
        Me.cmbPurge.BackColor = System.Drawing.SystemColors.Window
        Me.cmbPurge.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbPurge.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPurge.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbPurge.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbPurge.Items.AddRange(New Object() {"Yes", "No"})
        Me.cmbPurge.Location = New System.Drawing.Point(16, 176)
        Me.cmbPurge.Name = "cmbPurge"
        Me.cmbPurge.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbPurge.Size = New System.Drawing.Size(81, 26)
        Me.cmbPurge.TabIndex = 4
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(360, 208)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(81, 33)
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdModify
        '
        Me.cmdModify.BackColor = System.Drawing.SystemColors.Control
        Me.cmdModify.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdModify.Enabled = False
        Me.cmdModify.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdModify.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdModify.Location = New System.Drawing.Point(248, 208)
        Me.cmdModify.Name = "cmdModify"
        Me.cmdModify.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdModify.Size = New System.Drawing.Size(81, 33)
        Me.cmdModify.TabIndex = 6
        Me.cmdModify.Text = "Modify"
        Me.cmdModify.UseVisualStyleBackColor = False
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(128, 144)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(185, 17)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "Effective To Date (Not Editable)"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(128, 104)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(185, 17)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Effective From Date (Not Editable)"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(72, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(145, 17)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "NOA Code"
        '
        'label6
        '
        Me.label6.BackColor = System.Drawing.SystemColors.Control
        Me.label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.label6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label6.Location = New System.Drawing.Point(104, 224)
        Me.label6.Name = "label6"
        Me.label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.label6.Size = New System.Drawing.Size(81, 25)
        Me.label6.TabIndex = 10
        Me.label6.Text = "Duplex (Yes/No)"
        '
        'label5
        '
        Me.label5.BackColor = System.Drawing.SystemColors.Control
        Me.label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.label5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label5.Location = New System.Drawing.Point(104, 184)
        Me.label5.Name = "label5"
        Me.label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.label5.Size = New System.Drawing.Size(81, 25)
        Me.label5.TabIndex = 9
        Me.label5.Text = "Purge (Yes/No)"
        '
        'label3
        '
        Me.label3.BackColor = System.Drawing.SystemColors.Control
        Me.label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label3.Location = New System.Drawing.Point(376, 64)
        Me.label3.Name = "label3"
        Me.label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.label3.Size = New System.Drawing.Size(81, 25)
        Me.label3.TabIndex = 8
        Me.label3.Text = "Form Type"
        '
        'frmModNOA
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(460, 257)
        Me.Controls.Add(Me.txtToDate)
        Me.Controls.Add(Me.txtFromDate)
        Me.Controls.Add(Me.txtFT)
        Me.Controls.Add(Me.txtNOA)
        Me.Controls.Add(Me.cmbDuplex)
        Me.Controls.Add(Me.cmbPurge)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdModify)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.label6)
        Me.Controls.Add(Me.label5)
        Me.Controls.Add(Me.label3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Name = "frmModNOA"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Modify Nature of Action Record"
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class