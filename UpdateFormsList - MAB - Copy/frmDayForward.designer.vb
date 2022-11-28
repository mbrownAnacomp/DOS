<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmDayForward
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
    Public WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents Picture1 As System.Windows.Forms.PictureBox
	Public WithEvents cmdPrint As System.Windows.Forms.Button
	Public WithEvents lblPrint As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDayForward))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem
        Me.Picture1 = New System.Windows.Forms.PictureBox
        Me.cmdPrint = New System.Windows.Forms.Button
        Me.lblPrint = New System.Windows.Forms.Label
        Me.MainMenu1.SuspendLayout()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(291, 24)
        Me.MainMenu1.TabIndex = 3
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(35, 20)
        Me.mnuFile.Text = "&File"
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.E), System.Windows.Forms.Keys)
        Me.mnuExit.Size = New System.Drawing.Size(141, 22)
        Me.mnuExit.Text = "Exit"
        '
        'Picture1
        '
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Picture1.Image = CType(resources.GetObject("Picture1.Image"), System.Drawing.Image)
        Me.Picture1.InitialImage = Nothing
        Me.Picture1.Location = New System.Drawing.Point(80, 224)
        Me.Picture1.MaximumSize = New System.Drawing.Size(100, 100)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(100, 100)
        Me.Picture1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.Picture1.TabIndex = 2
        Me.Picture1.TabStop = False
        '
        'cmdPrint
        '
        Me.cmdPrint.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPrint.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPrint.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPrint.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPrint.Location = New System.Drawing.Point(106, 63)
        Me.cmdPrint.Name = "cmdPrint"
        Me.cmdPrint.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPrint.Size = New System.Drawing.Size(81, 33)
        Me.cmdPrint.TabIndex = 0
        Me.cmdPrint.Text = "Print"
        Me.cmdPrint.UseVisualStyleBackColor = False
        '
        'lblPrint
        '
        Me.lblPrint.BackColor = System.Drawing.SystemColors.Control
        Me.lblPrint.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPrint.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrint.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPrint.Location = New System.Drawing.Point(-2, 34)
        Me.lblPrint.Name = "lblPrint"
        Me.lblPrint.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPrint.Size = New System.Drawing.Size(293, 17)
        Me.lblPrint.TabIndex = 1
        Me.lblPrint.Text = "Print next in sequence Loose Paper Bundle Control Sheet"
        Me.lblPrint.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmDayForward
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(291, 108)
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.cmdPrint)
        Me.Controls.Add(Me.lblPrint)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(11, 57)
        Me.Name = "frmDayForward"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Create Loose Paper BCSheets"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class