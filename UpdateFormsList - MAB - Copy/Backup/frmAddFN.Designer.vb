<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAddFN
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
	Public WithEvents cmbFT As System.Windows.Forms.ComboBox
	Public WithEvents cmbVS As System.Windows.Forms.ComboBox
	Public WithEvents cmbDuplex As System.Windows.Forms.ComboBox
	Public WithEvents cmbPurge As System.Windows.Forms.ComboBox
	Public WithEvents txtOFN As System.Windows.Forms.TextBox
	Public WithEvents txtFN As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdAdd As System.Windows.Forms.Button
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents label6 As System.Windows.Forms.Label
	Public WithEvents label5 As System.Windows.Forms.Label
	Public WithEvents label4 As System.Windows.Forms.Label
	Public WithEvents label3 As System.Windows.Forms.Label
	Public WithEvents label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAddFN))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmbFT = New System.Windows.Forms.ComboBox
		Me.cmbVS = New System.Windows.Forms.ComboBox
		Me.cmbDuplex = New System.Windows.Forms.ComboBox
		Me.cmbPurge = New System.Windows.Forms.ComboBox
		Me.txtOFN = New System.Windows.Forms.TextBox
		Me.txtFN = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdAdd = New System.Windows.Forms.Button
		Me.Label7 = New System.Windows.Forms.Label
		Me.label6 = New System.Windows.Forms.Label
		Me.label5 = New System.Windows.Forms.Label
		Me.label4 = New System.Windows.Forms.Label
		Me.label3 = New System.Windows.Forms.Label
		Me.label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Add Form Number Record"
		Me.ClientSize = New System.Drawing.Size(666, 225)
		Me.Location = New System.Drawing.Point(4, 30)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmAddFN"
		Me.cmbFT.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbFT.Size = New System.Drawing.Size(353, 28)
		Me.cmbFT.Location = New System.Drawing.Point(16, 80)
		Me.cmbFT.Items.AddRange(New Object(){"BENEFITS", "EMPLOYEE", "INVESTIGATIONS/SECURITY CLEARENCE", "PAYROLL", "PERFORMANCE APPRAISAL", "PERSONNEL ACTION/SUPPORT DOC", "POSITION", "TRAINING"})
		Me.cmbFT.TabIndex = 2
		Me.cmbFT.BackColor = System.Drawing.SystemColors.Window
		Me.cmbFT.CausesValidation = True
		Me.cmbFT.Enabled = True
		Me.cmbFT.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbFT.IntegralHeight = True
		Me.cmbFT.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbFT.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbFT.Sorted = False
		Me.cmbFT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbFT.TabStop = True
		Me.cmbFT.Visible = True
		Me.cmbFT.Name = "cmbFT"
		Me.cmbVS.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbVS.Size = New System.Drawing.Size(137, 28)
		Me.cmbVS.Location = New System.Drawing.Point(16, 112)
		Me.cmbVS.Items.AddRange(New Object(){"Deleted", "Miscellaneous", "Payroll", "Performance", "Permanent", "Temporary", "Training"})
		Me.cmbVS.TabIndex = 3
		Me.cmbVS.BackColor = System.Drawing.SystemColors.Window
		Me.cmbVS.CausesValidation = True
		Me.cmbVS.Enabled = True
		Me.cmbVS.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbVS.IntegralHeight = True
		Me.cmbVS.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbVS.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbVS.Sorted = False
		Me.cmbVS.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbVS.TabStop = True
		Me.cmbVS.Visible = True
		Me.cmbVS.Name = "cmbVS"
		Me.cmbDuplex.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbDuplex.Size = New System.Drawing.Size(81, 28)
		Me.cmbDuplex.Location = New System.Drawing.Point(16, 176)
		Me.cmbDuplex.Items.AddRange(New Object(){"Yes", "No"})
		Me.cmbDuplex.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbDuplex.TabIndex = 5
		Me.cmbDuplex.BackColor = System.Drawing.SystemColors.Window
		Me.cmbDuplex.CausesValidation = True
		Me.cmbDuplex.Enabled = True
		Me.cmbDuplex.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbDuplex.IntegralHeight = True
		Me.cmbDuplex.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbDuplex.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbDuplex.Sorted = False
		Me.cmbDuplex.TabStop = True
		Me.cmbDuplex.Visible = True
		Me.cmbDuplex.Name = "cmbDuplex"
		Me.cmbPurge.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbPurge.Size = New System.Drawing.Size(81, 28)
		Me.cmbPurge.Location = New System.Drawing.Point(16, 144)
		Me.cmbPurge.Items.AddRange(New Object(){"Yes", "No"})
		Me.cmbPurge.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbPurge.TabIndex = 4
		Me.cmbPurge.BackColor = System.Drawing.SystemColors.Window
		Me.cmbPurge.CausesValidation = True
		Me.cmbPurge.Enabled = True
		Me.cmbPurge.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbPurge.IntegralHeight = True
		Me.cmbPurge.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbPurge.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbPurge.Sorted = False
		Me.cmbPurge.TabStop = True
		Me.cmbPurge.Visible = True
		Me.cmbPurge.Name = "cmbPurge"
		Me.txtOFN.AutoSize = False
		Me.txtOFN.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtOFN.Size = New System.Drawing.Size(353, 25)
		Me.txtOFN.Location = New System.Drawing.Point(16, 48)
		Me.txtOFN.TabIndex = 1
		Me.txtOFN.AcceptsReturn = True
		Me.txtOFN.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtOFN.BackColor = System.Drawing.SystemColors.Window
		Me.txtOFN.CausesValidation = True
		Me.txtOFN.Enabled = True
		Me.txtOFN.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtOFN.HideSelection = True
		Me.txtOFN.ReadOnly = False
		Me.txtOFN.Maxlength = 0
		Me.txtOFN.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtOFN.MultiLine = False
		Me.txtOFN.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtOFN.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtOFN.TabStop = True
		Me.txtOFN.Visible = True
		Me.txtOFN.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtOFN.Name = "txtOFN"
		Me.txtFN.AutoSize = False
		Me.txtFN.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtFN.Size = New System.Drawing.Size(353, 25)
		Me.txtFN.Location = New System.Drawing.Point(16, 16)
		Me.txtFN.TabIndex = 0
		Me.txtFN.AcceptsReturn = True
		Me.txtFN.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtFN.BackColor = System.Drawing.SystemColors.Window
		Me.txtFN.CausesValidation = True
		Me.txtFN.Enabled = True
		Me.txtFN.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtFN.HideSelection = True
		Me.txtFN.ReadOnly = False
		Me.txtFN.Maxlength = 0
		Me.txtFN.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtFN.MultiLine = False
		Me.txtFN.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtFN.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtFN.TabStop = True
		Me.txtFN.Visible = True
		Me.txtFN.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtFN.Name = "txtFN"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.CausesValidation = False
		Me.cmdCancel.Size = New System.Drawing.Size(81, 33)
		Me.cmdCancel.Location = New System.Drawing.Point(432, 168)
		Me.cmdCancel.TabIndex = 7
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdAdd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAdd.Text = "Add"
		Me.cmdAdd.Size = New System.Drawing.Size(81, 33)
		Me.cmdAdd.Location = New System.Drawing.Point(296, 168)
		Me.cmdAdd.TabIndex = 6
		Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAdd.CausesValidation = True
		Me.cmdAdd.Enabled = True
		Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAdd.TabStop = True
		Me.cmdAdd.Name = "cmdAdd"
		Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label7.Text = "All Fields are required"
		Me.Label7.Size = New System.Drawing.Size(281, 17)
		Me.Label7.Location = New System.Drawing.Point(264, 136)
		Me.Label7.TabIndex = 14
		Me.Label7.BackColor = System.Drawing.SystemColors.Control
		Me.Label7.Enabled = True
		Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label7.UseMnemonic = True
		Me.Label7.Visible = True
		Me.Label7.AutoSize = False
		Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label7.Name = "Label7"
		Me.label6.Text = "Duplex (Yes/No)"
		Me.label6.Size = New System.Drawing.Size(81, 25)
		Me.label6.Location = New System.Drawing.Point(104, 176)
		Me.label6.TabIndex = 13
		Me.label6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.label6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.label6.BackColor = System.Drawing.SystemColors.Control
		Me.label6.Enabled = True
		Me.label6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.label6.Cursor = System.Windows.Forms.Cursors.Default
		Me.label6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.label6.UseMnemonic = True
		Me.label6.Visible = True
		Me.label6.AutoSize = False
		Me.label6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.label6.Name = "label6"
		Me.label5.Text = "Purge (Yes/No)"
		Me.label5.Size = New System.Drawing.Size(81, 25)
		Me.label5.Location = New System.Drawing.Point(104, 144)
		Me.label5.TabIndex = 12
		Me.label5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.label5.BackColor = System.Drawing.SystemColors.Control
		Me.label5.Enabled = True
		Me.label5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.label5.UseMnemonic = True
		Me.label5.Visible = True
		Me.label5.AutoSize = False
		Me.label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.label5.Name = "label5"
		Me.label4.Text = "Virtual Side"
		Me.label4.Size = New System.Drawing.Size(81, 25)
		Me.label4.Location = New System.Drawing.Point(160, 112)
		Me.label4.TabIndex = 11
		Me.label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.label4.BackColor = System.Drawing.SystemColors.Control
		Me.label4.Enabled = True
		Me.label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.label4.UseMnemonic = True
		Me.label4.Visible = True
		Me.label4.AutoSize = False
		Me.label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.label4.Name = "label4"
		Me.label3.Text = "Form Type"
		Me.label3.Size = New System.Drawing.Size(81, 25)
		Me.label3.Location = New System.Drawing.Point(376, 80)
		Me.label3.TabIndex = 10
		Me.label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.label3.BackColor = System.Drawing.SystemColors.Control
		Me.label3.Enabled = True
		Me.label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.label3.UseMnemonic = True
		Me.label3.Visible = True
		Me.label3.AutoSize = False
		Me.label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.label3.Name = "label3"
		Me.label2.Text = "Original Form Name/Number (as seen in eOPF) "
		Me.label2.Size = New System.Drawing.Size(273, 25)
		Me.label2.Location = New System.Drawing.Point(376, 48)
		Me.label2.TabIndex = 9
		Me.label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.label2.BackColor = System.Drawing.SystemColors.Control
		Me.label2.Enabled = True
		Me.label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.label2.UseMnemonic = True
		Me.label2.Visible = True
		Me.label2.AutoSize = False
		Me.label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.label2.Name = "label2"
		Me.Label1.Text = "Form Name/Number (no spaces or special characters) "
		Me.Label1.Size = New System.Drawing.Size(265, 25)
		Me.Label1.Location = New System.Drawing.Point(376, 16)
		Me.Label1.TabIndex = 8
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(cmbFT)
		Me.Controls.Add(cmbVS)
		Me.Controls.Add(cmbDuplex)
		Me.Controls.Add(cmbPurge)
		Me.Controls.Add(txtOFN)
		Me.Controls.Add(txtFN)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdAdd)
		Me.Controls.Add(Label7)
		Me.Controls.Add(label6)
		Me.Controls.Add(label5)
		Me.Controls.Add(label4)
		Me.Controls.Add(label3)
		Me.Controls.Add(label2)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class