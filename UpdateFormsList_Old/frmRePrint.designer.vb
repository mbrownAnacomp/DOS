<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmRePrint
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
	Public WithEvents cmdExit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents cmdFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents cmdRePrint As System.Windows.Forms.Button
	Public WithEvents txtBatchName As System.Windows.Forms.TextBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmRePrint))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.cmdFile = New System.Windows.Forms.ToolStripMenuItem
		Me.cmdExit = New System.Windows.Forms.ToolStripMenuItem
		Me.cmdRePrint = New System.Windows.Forms.Button
		Me.txtBatchName = New System.Windows.Forms.TextBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.MainMenu1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Re-Print Batch Cover Sheet"
		Me.ClientSize = New System.Drawing.Size(312, 165)
		Me.Location = New System.Drawing.Point(11, 57)
		Me.Icon = CType(resources.GetObject("frmRePrint.Icon"), System.Drawing.Icon)
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
		Me.Name = "frmRePrint"
		Me.cmdFile.Name = "cmdFile"
		Me.cmdFile.Text = "&File"
		Me.cmdFile.Checked = False
		Me.cmdFile.Enabled = True
		Me.cmdFile.Visible = True
		Me.cmdExit.Name = "cmdExit"
		Me.cmdExit.Text = "&Exit"
		Me.cmdExit.Checked = False
		Me.cmdExit.Enabled = True
		Me.cmdExit.Visible = True
		Me.cmdRePrint.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdRePrint.Text = "Re-Print"
		Me.cmdRePrint.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdRePrint.Size = New System.Drawing.Size(81, 33)
		Me.cmdRePrint.Location = New System.Drawing.Point(112, 112)
		Me.cmdRePrint.TabIndex = 2
		Me.cmdRePrint.BackColor = System.Drawing.SystemColors.Control
		Me.cmdRePrint.CausesValidation = True
		Me.cmdRePrint.Enabled = True
		Me.cmdRePrint.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdRePrint.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdRePrint.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdRePrint.TabStop = True
		Me.cmdRePrint.Name = "cmdRePrint"
		Me.txtBatchName.AutoSize = False
		Me.txtBatchName.BackColor = System.Drawing.Color.Black
		Me.txtBatchName.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtBatchName.ForeColor = System.Drawing.Color.Red
		Me.txtBatchName.Size = New System.Drawing.Size(209, 25)
		Me.txtBatchName.Location = New System.Drawing.Point(48, 80)
		Me.txtBatchName.TabIndex = 0
		Me.txtBatchName.AcceptsReturn = True
		Me.txtBatchName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtBatchName.CausesValidation = True
		Me.txtBatchName.Enabled = True
		Me.txtBatchName.HideSelection = True
		Me.txtBatchName.ReadOnly = False
		Me.txtBatchName.Maxlength = 0
		Me.txtBatchName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtBatchName.MultiLine = False
		Me.txtBatchName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtBatchName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtBatchName.TabStop = True
		Me.txtBatchName.Visible = True
		Me.txtBatchName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtBatchName.Name = "txtBatchName"
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label1.Text = "Enter Batch Name to be Printed"
		Me.Label1.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(305, 25)
		Me.Label1.Location = New System.Drawing.Point(0, 48)
		Me.Label1.TabIndex = 1
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
		Me.Controls.Add(cmdRePrint)
		Me.Controls.Add(txtBatchName)
		Me.Controls.Add(Label1)
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.cmdFile})
		cmdFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.cmdExit})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class