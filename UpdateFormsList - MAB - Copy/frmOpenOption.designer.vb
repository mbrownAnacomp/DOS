<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmOpenOption
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
    Public WithEvents mnuAbout As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents cmdBackfile As System.Windows.Forms.Button
	Public WithEvents cmdLP As System.Windows.Forms.Button
	Public WithEvents lblSelect As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAbout = New System.Windows.Forms.ToolStripMenuItem
        Me.cmdBackfile = New System.Windows.Forms.Button
        Me.cmdLP = New System.Windows.Forms.Button
        Me.lblSelect = New System.Windows.Forms.Label
        Me._mnuFile_1 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDeleteBatch = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuResetBatchStatus = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRePrint = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem
        Me.MainMenu1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuFile_1, Me.mnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(312, 24)
        Me.MainMenu1.TabIndex = 3
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAbout})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(40, 20)
        Me.mnuHelp.Text = "&Help"
        '
        'mnuAbout
        '
        Me.mnuAbout.Name = "mnuAbout"
        Me.mnuAbout.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
        Me.mnuAbout.Size = New System.Drawing.Size(153, 22)
        Me.mnuAbout.Text = "About"
        '
        'cmdBackfile
        '
        Me.cmdBackfile.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBackfile.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBackfile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBackfile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBackfile.Location = New System.Drawing.Point(184, 80)
        Me.cmdBackfile.Name = "cmdBackfile"
        Me.cmdBackfile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBackfile.Size = New System.Drawing.Size(81, 33)
        Me.cmdBackfile.TabIndex = 2
        Me.cmdBackfile.Text = "Backfile"
        Me.cmdBackfile.UseVisualStyleBackColor = False
        '
        'cmdLP
        '
        Me.cmdLP.BackColor = System.Drawing.SystemColors.Control
        Me.cmdLP.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdLP.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLP.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdLP.Location = New System.Drawing.Point(40, 80)
        Me.cmdLP.Name = "cmdLP"
        Me.cmdLP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdLP.Size = New System.Drawing.Size(81, 33)
        Me.cmdLP.TabIndex = 1
        Me.cmdLP.Text = "Loose Paper"
        Me.cmdLP.UseVisualStyleBackColor = False
        '
        'lblSelect
        '
        Me.lblSelect.BackColor = System.Drawing.SystemColors.Control
        Me.lblSelect.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSelect.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelect.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSelect.Location = New System.Drawing.Point(8, 48)
        Me.lblSelect.Name = "lblSelect"
        Me.lblSelect.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSelect.Size = New System.Drawing.Size(289, 25)
        Me.lblSelect.TabIndex = 0
        Me.lblSelect.Text = "Choose Loose Paper or Backfile for batch type"
        Me.lblSelect.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_mnuFile_1
        '
        Me._mnuFile_1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuDeleteBatch, Me.mnuResetBatchStatus, Me.mnuRePrint, Me.mnuExit})
        Me._mnuFile_1.Name = "_mnuFile_1"
        Me._mnuFile_1.Size = New System.Drawing.Size(35, 20)
        Me._mnuFile_1.Text = "&File"
        '
        'mnuDeleteBatch
        '
        Me.mnuDeleteBatch.Name = "mnuDeleteBatch"
        Me.mnuDeleteBatch.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.D), System.Windows.Forms.Keys)
        Me.mnuDeleteBatch.Size = New System.Drawing.Size(216, 22)
        Me.mnuDeleteBatch.Text = "Delete Batch"
        '
        'mnuResetBatchStatus
        '
        Me.mnuResetBatchStatus.Name = "mnuResetBatchStatus"
        Me.mnuResetBatchStatus.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
        Me.mnuResetBatchStatus.Size = New System.Drawing.Size(216, 22)
        Me.mnuResetBatchStatus.Text = "Reset Batch Status"
        '
        'mnuRePrint
        '
        Me.mnuRePrint.Name = "mnuRePrint"
        Me.mnuRePrint.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P), System.Windows.Forms.Keys)
        Me.mnuRePrint.Size = New System.Drawing.Size(216, 22)
        Me.mnuRePrint.Text = "Re-Print"
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.E), System.Windows.Forms.Keys)
        Me.mnuExit.Size = New System.Drawing.Size(216, 22)
        Me.mnuExit.Text = "Exit"
        '
        'frmOpenOption
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(312, 139)
        Me.Controls.Add(Me.cmdBackfile)
        Me.Controls.Add(Me.cmdLP)
        Me.Controls.Add(Me.lblSelect)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(11, 57)
        Me.Name = "frmOpenOption"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Select Batch Type"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents _mnuFile_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDeleteBatch As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuResetBatchStatus As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuRePrint As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
#End Region 
End Class