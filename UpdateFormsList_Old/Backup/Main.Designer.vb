<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMain
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
    Public WithEvents mnuFormDict As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents cmdHyphen As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuBackup As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents cmdHyphen2 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents cmdExit As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents cmdFile As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuAbout As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public WithEvents cmdAddNOA As System.Windows.Forms.Button
    Public WithEvents cmdModifyNOA As System.Windows.Forms.Button
    Public WithEvents cmdAddFN As System.Windows.Forms.Button
    Public WithEvents cmdModifyFN As System.Windows.Forms.Button
    Public WithEvents Image1 As System.Windows.Forms.PictureBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.cmdFile = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuPS50 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAllOther = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuFormDict = New System.Windows.Forms.ToolStripMenuItem
        Me.cmdHyphen = New System.Windows.Forms.ToolStripSeparator
        Me.mnuBackup = New System.Windows.Forms.ToolStripMenuItem
        Me.cmdHyphen2 = New System.Windows.Forms.ToolStripSeparator
        Me.cmdExit = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAbout = New System.Windows.Forms.ToolStripMenuItem
        Me.cmdAddNOA = New System.Windows.Forms.Button
        Me.cmdModifyNOA = New System.Windows.Forms.Button
        Me.cmdAddFN = New System.Windows.Forms.Button
        Me.cmdModifyFN = New System.Windows.Forms.Button
        Me.Image1 = New System.Windows.Forms.PictureBox
        Me.MainMenu1.SuspendLayout()
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.cmdFile, Me.mnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(391, 24)
        Me.MainMenu1.TabIndex = 5
        '
        'cmdFile
        '
        Me.cmdFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1, Me.mnuPS50, Me.mnuAllOther, Me.mnuFormDict, Me.cmdHyphen, Me.mnuBackup, Me.cmdHyphen2, Me.cmdExit})
        Me.cmdFile.Name = "cmdFile"
        Me.cmdFile.Size = New System.Drawing.Size(35, 20)
        Me.cmdFile.Text = "&File"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(273, 22)
        Me.ToolStripMenuItem1.Text = "Find SF50 NOA Code Dictionary File"
        '
        'mnuPS50
        '
        Me.mnuPS50.Name = "mnuPS50"
        Me.mnuPS50.Size = New System.Drawing.Size(273, 22)
        Me.mnuPS50.Text = "Find PS50 NOA Code Dictionary File"
        '
        'mnuAllOther
        '
        Me.mnuAllOther.Name = "mnuAllOther"
        Me.mnuAllOther.Size = New System.Drawing.Size(273, 22)
        Me.mnuAllOther.Text = "Find All Other NOA Code Dictionary File"
        '
        'mnuFormDict
        '
        Me.mnuFormDict.Name = "mnuFormDict"
        Me.mnuFormDict.Size = New System.Drawing.Size(273, 22)
        Me.mnuFormDict.Text = "Find Form Dictionary File"
        '
        'cmdHyphen
        '
        Me.cmdHyphen.Name = "cmdHyphen"
        Me.cmdHyphen.Size = New System.Drawing.Size(270, 6)
        '
        'mnuBackup
        '
        Me.mnuBackup.Name = "mnuBackup"
        Me.mnuBackup.Size = New System.Drawing.Size(273, 22)
        Me.mnuBackup.Text = "Backup Databases"
        '
        'cmdHyphen2
        '
        Me.cmdHyphen2.Name = "cmdHyphen2"
        Me.cmdHyphen2.Size = New System.Drawing.Size(270, 6)
        '
        'cmdExit
        '
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(273, 22)
        Me.cmdExit.Text = "E&xit"
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
        Me.mnuAbout.Size = New System.Drawing.Size(203, 22)
        Me.mnuAbout.Text = "&About Update Forms List"
        '
        'cmdAddNOA
        '
        Me.cmdAddNOA.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAddNOA.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAddNOA.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddNOA.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAddNOA.Location = New System.Drawing.Point(272, 240)
        Me.cmdAddNOA.Name = "cmdAddNOA"
        Me.cmdAddNOA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAddNOA.Size = New System.Drawing.Size(113, 33)
        Me.cmdAddNOA.TabIndex = 3
        Me.cmdAddNOA.Text = "Add NOA Record"
        Me.cmdAddNOA.UseVisualStyleBackColor = False
        '
        'cmdModifyNOA
        '
        Me.cmdModifyNOA.BackColor = System.Drawing.SystemColors.Control
        Me.cmdModifyNOA.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdModifyNOA.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdModifyNOA.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdModifyNOA.Location = New System.Drawing.Point(8, 240)
        Me.cmdModifyNOA.Name = "cmdModifyNOA"
        Me.cmdModifyNOA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdModifyNOA.Size = New System.Drawing.Size(113, 33)
        Me.cmdModifyNOA.TabIndex = 2
        Me.cmdModifyNOA.Text = "Modify NOA Record"
        Me.cmdModifyNOA.UseVisualStyleBackColor = False
        '
        'cmdAddFN
        '
        Me.cmdAddFN.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAddFN.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAddFN.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAddFN.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAddFN.Location = New System.Drawing.Point(272, 56)
        Me.cmdAddFN.Name = "cmdAddFN"
        Me.cmdAddFN.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAddFN.Size = New System.Drawing.Size(113, 33)
        Me.cmdAddFN.TabIndex = 1
        Me.cmdAddFN.Text = "Add Form Number Record"
        Me.cmdAddFN.UseVisualStyleBackColor = False
        '
        'cmdModifyFN
        '
        Me.cmdModifyFN.BackColor = System.Drawing.SystemColors.Control
        Me.cmdModifyFN.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdModifyFN.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdModifyFN.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdModifyFN.Location = New System.Drawing.Point(8, 56)
        Me.cmdModifyFN.Name = "cmdModifyFN"
        Me.cmdModifyFN.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdModifyFN.Size = New System.Drawing.Size(113, 33)
        Me.cmdModifyFN.TabIndex = 0
        Me.cmdModifyFN.Text = "Modify Form Number Record"
        Me.cmdModifyFN.UseVisualStyleBackColor = False
        '
        'Image1
        '
        Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image1.Image = CType(resources.GetObject("Image1.Image"), System.Drawing.Image)
        Me.Image1.Location = New System.Drawing.Point(120, 88)
        Me.Image1.Name = "Image1"
        Me.Image1.Size = New System.Drawing.Size(150, 155)
        Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Image1.TabIndex = 4
        Me.Image1.TabStop = False
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(391, 298)
        Me.Controls.Add(Me.cmdAddNOA)
        Me.Controls.Add(Me.cmdModifyNOA)
        Me.Controls.Add(Me.cmdAddFN)
        Me.Controls.Add(Me.cmdModifyFN)
        Me.Controls.Add(Me.Image1)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(11, 57)
        Me.Name = "frmMain"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Update Forms List"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuPS50 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAllOther As System.Windows.Forms.ToolStripMenuItem
#End Region 
End Class