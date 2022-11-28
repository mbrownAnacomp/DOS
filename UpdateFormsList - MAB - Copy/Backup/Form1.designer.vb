<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmOPMXML
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
	Public WithEvents chkBackUp As System.Windows.Forms.CheckBox
	Public WithEvents txtInfo As System.Windows.Forms.TextBox
	Public WithEvents cboData As System.Windows.Forms.ComboBox
	Public WithEvents cmdStartDoc As System.Windows.Forms.Button
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents cmdWorkDir As System.Windows.Forms.Button
	Public WithEvents txtWorkDir As System.Windows.Forms.TextBox
	Public WithEvents cmdClose As System.Windows.Forms.Button
	Public WithEvents cmdDir As System.Windows.Forms.Button
	Public CommonDialog1Open As System.Windows.Forms.OpenFileDialog
	Public WithEvents prgBar1 As System.Windows.Forms.ProgressBar
	Public WithEvents _txtDir_0 As System.Windows.Forms.TextBox
	Public WithEvents Zip1 As AxDartZip.AxZip
	Public WithEvents lblSearch As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents lblWorkDir As System.Windows.Forms.Label
	Public WithEvents lblSelectDir As System.Windows.Forms.Label
	Public WithEvents txtDir As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOPMXML))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkBackUp = New System.Windows.Forms.CheckBox
        Me.txtInfo = New System.Windows.Forms.TextBox
        Me.cboData = New System.Windows.Forms.ComboBox
        Me.cmdStartDoc = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdWorkDir = New System.Windows.Forms.Button
        Me.txtWorkDir = New System.Windows.Forms.TextBox
        Me.cmdClose = New System.Windows.Forms.Button
        Me.cmdDir = New System.Windows.Forms.Button
        Me.CommonDialog1Open = New System.Windows.Forms.OpenFileDialog
        Me.prgBar1 = New System.Windows.Forms.ProgressBar
        Me._txtDir_0 = New System.Windows.Forms.TextBox
        Me.Zip1 = New AxDartZip.AxZip
        Me.lblSearch = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblWorkDir = New System.Windows.Forms.Label
        Me.lblSelectDir = New System.Windows.Forms.Label
        Me.txtDir = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        CType(Me.Zip1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDir, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'chkBackUp
        '
        Me.chkBackUp.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.chkBackUp.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkBackUp.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBackUp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkBackUp.Location = New System.Drawing.Point(504, 120)
        Me.chkBackUp.Name = "chkBackUp"
        Me.chkBackUp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkBackUp.Size = New System.Drawing.Size(121, 17)
        Me.chkBackUp.TabIndex = 14
        Me.chkBackUp.Text = "Backup Databases"
        Me.chkBackUp.UseVisualStyleBackColor = False
        '
        'txtInfo
        '
        Me.txtInfo.AcceptsReturn = True
        Me.txtInfo.BackColor = System.Drawing.SystemColors.Window
        Me.txtInfo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInfo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInfo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInfo.Location = New System.Drawing.Point(40, 216)
        Me.txtInfo.MaxLength = 0
        Me.txtInfo.Name = "txtInfo"
        Me.txtInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInfo.Size = New System.Drawing.Size(633, 65)
        Me.txtInfo.TabIndex = 13
        '
        'cboData
        '
        Me.cboData.BackColor = System.Drawing.SystemColors.Window
        Me.cboData.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboData.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboData.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboData.Items.AddRange(New Object() {"1", "2", "5", "10", "20", "30", "60", "90", "120", "150", "180"})
        Me.cboData.Location = New System.Drawing.Point(504, 48)
        Me.cboData.Name = "cboData"
        Me.cboData.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboData.Size = New System.Drawing.Size(89, 22)
        Me.cboData.TabIndex = 11
        '
        'cmdStartDoc
        '
        Me.cmdStartDoc.BackColor = System.Drawing.SystemColors.Control
        Me.cmdStartDoc.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdStartDoc.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdStartDoc.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStartDoc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStartDoc.Location = New System.Drawing.Point(312, 392)
        Me.cmdStartDoc.Name = "cmdStartDoc"
        Me.cmdStartDoc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdStartDoc.Size = New System.Drawing.Size(113, 41)
        Me.cmdStartDoc.TabIndex = 10
        Me.cmdStartDoc.Text = "START"
        Me.cmdStartDoc.UseVisualStyleBackColor = False
        '
        'Timer1
        '
        Me.Timer1.Interval = 1
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSave.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSave.Location = New System.Drawing.Point(32, 392)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSave.Size = New System.Drawing.Size(145, 41)
        Me.cmdSave.TabIndex = 8
        Me.cmdSave.Text = "Save Configuration"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'cmdWorkDir
        '
        Me.cmdWorkDir.BackColor = System.Drawing.SystemColors.Control
        Me.cmdWorkDir.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdWorkDir.Font = New System.Drawing.Font("Arial", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdWorkDir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdWorkDir.Location = New System.Drawing.Point(368, 136)
        Me.cmdWorkDir.Name = "cmdWorkDir"
        Me.cmdWorkDir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdWorkDir.Size = New System.Drawing.Size(73, 33)
        Me.cmdWorkDir.TabIndex = 6
        Me.cmdWorkDir.Text = "......"
        Me.cmdWorkDir.UseVisualStyleBackColor = False
        '
        'txtWorkDir
        '
        Me.txtWorkDir.AcceptsReturn = True
        Me.txtWorkDir.BackColor = System.Drawing.SystemColors.Window
        Me.txtWorkDir.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtWorkDir.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWorkDir.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtWorkDir.Location = New System.Drawing.Point(40, 136)
        Me.txtWorkDir.MaxLength = 0
        Me.txtWorkDir.Name = "txtWorkDir"
        Me.txtWorkDir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtWorkDir.Size = New System.Drawing.Size(321, 33)
        Me.txtWorkDir.TabIndex = 5
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(584, 384)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(105, 41)
        Me.cmdClose.TabIndex = 4
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'cmdDir
        '
        Me.cmdDir.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDir.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDir.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDir.Location = New System.Drawing.Point(368, 56)
        Me.cmdDir.Name = "cmdDir"
        Me.cmdDir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDir.Size = New System.Drawing.Size(65, 29)
        Me.cmdDir.TabIndex = 2
        Me.cmdDir.Text = "......"
        Me.cmdDir.UseVisualStyleBackColor = False
        '
        'prgBar1
        '
        Me.prgBar1.Location = New System.Drawing.Point(24, 336)
        Me.prgBar1.Name = "prgBar1"
        Me.prgBar1.Size = New System.Drawing.Size(657, 33)
        Me.prgBar1.TabIndex = 1
        '
        '_txtDir_0
        '
        Me._txtDir_0.AcceptsReturn = True
        Me._txtDir_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtDir_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtDir_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._txtDir_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDir.SetIndex(Me._txtDir_0, CType(0, Short))
        Me._txtDir_0.Location = New System.Drawing.Point(40, 56)
        Me._txtDir_0.MaxLength = 0
        Me._txtDir_0.Name = "_txtDir_0"
        Me._txtDir_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtDir_0.Size = New System.Drawing.Size(321, 25)
        Me._txtDir_0.TabIndex = 0
        '
        'Zip1
        '
        Me.Zip1.Enabled = True
        Me.Zip1.Location = New System.Drawing.Point(336, 208)
        Me.Zip1.Name = "Zip1"
        Me.Zip1.OcxState = CType(resources.GetObject("Zip1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Zip1.Size = New System.Drawing.Size(28, 28)
        Me.Zip1.TabIndex = 15
        '
        'lblSearch
        '
        Me.lblSearch.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.lblSearch.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSearch.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSearch.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblSearch.Location = New System.Drawing.Point(448, 16)
        Me.lblSearch.Name = "lblSearch"
        Me.lblSearch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSearch.Size = New System.Drawing.Size(225, 25)
        Me.lblSearch.TabIndex = 12
        Me.lblSearch.Text = "Select  Search Poll Time (in seconds)"
        Me.lblSearch.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Red
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(24, 304)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(657, 33)
        Me.Label1.TabIndex = 9
        '
        'lblWorkDir
        '
        Me.lblWorkDir.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.lblWorkDir.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblWorkDir.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWorkDir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWorkDir.Location = New System.Drawing.Point(40, 104)
        Me.lblWorkDir.Name = "lblWorkDir"
        Me.lblWorkDir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWorkDir.Size = New System.Drawing.Size(321, 33)
        Me.lblWorkDir.TabIndex = 7
        Me.lblWorkDir.Text = "Select Output Directory"
        '
        'lblSelectDir
        '
        Me.lblSelectDir.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.lblSelectDir.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSelectDir.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelectDir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSelectDir.Location = New System.Drawing.Point(40, 24)
        Me.lblSelectDir.Name = "lblSelectDir"
        Me.lblSelectDir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSelectDir.Size = New System.Drawing.Size(321, 33)
        Me.lblSelectDir.TabIndex = 3
        Me.lblSelectDir.Text = "Select Input Directory"
        '
        'frmOPMXML
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.CancelButton = Me.cmdStartDoc
        Me.ClientSize = New System.Drawing.Size(703, 447)
        Me.Controls.Add(Me.chkBackUp)
        Me.Controls.Add(Me.txtInfo)
        Me.Controls.Add(Me.cboData)
        Me.Controls.Add(Me.cmdStartDoc)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdWorkDir)
        Me.Controls.Add(Me.txtWorkDir)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdDir)
        Me.Controls.Add(Me.prgBar1)
        Me.Controls.Add(Me._txtDir_0)
        Me.Controls.Add(Me.Zip1)
        Me.Controls.Add(Me.lblSearch)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblWorkDir)
        Me.Controls.Add(Me.lblSelectDir)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Name = "frmOPMXML"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "DoS XML Creator"
        CType(Me.Zip1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDir, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class