<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmDoSResetBatchStatus
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
    Public WithEvents txtStatus As System.Windows.Forms.TextBox
    Public WithEvents cmdFindBatch As System.Windows.Forms.Button
    Public WithEvents lstBatchStatus As System.Windows.Forms.ListBox
    Public WithEvents txtBatchName As System.Windows.Forms.TextBox
    Public WithEvents cmdReset As System.Windows.Forms.Button
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDoSResetBatchStatus))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem
        Me.txtStatus = New System.Windows.Forms.TextBox
        Me.cmdFindBatch = New System.Windows.Forms.Button
        Me.lstBatchStatus = New System.Windows.Forms.ListBox
        Me.txtBatchName = New System.Windows.Forms.TextBox
        Me.cmdReset = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.MainMenu1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(235, 24)
        Me.MainMenu1.TabIndex = 8
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
        Me.mnuExit.Size = New System.Drawing.Size(103, 22)
        Me.mnuExit.Text = "E&xit"
        '
        'txtStatus
        '
        Me.txtStatus.AcceptsReturn = True
        Me.txtStatus.BackColor = System.Drawing.Color.Black
        Me.txtStatus.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStatus.ForeColor = System.Drawing.Color.Yellow
        Me.txtStatus.Location = New System.Drawing.Point(192, 88)
        Me.txtStatus.MaxLength = 0
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStatus.Size = New System.Drawing.Size(25, 25)
        Me.txtStatus.TabIndex = 5
        '
        'cmdFindBatch
        '
        Me.cmdFindBatch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdFindBatch.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdFindBatch.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFindBatch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdFindBatch.Location = New System.Drawing.Point(24, 48)
        Me.cmdFindBatch.Name = "cmdFindBatch"
        Me.cmdFindBatch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdFindBatch.Size = New System.Drawing.Size(137, 33)
        Me.cmdFindBatch.TabIndex = 4
        Me.cmdFindBatch.Text = "Find Batch"
        Me.cmdFindBatch.UseVisualStyleBackColor = False
        '
        'lstBatchStatus
        '
        Me.lstBatchStatus.BackColor = System.Drawing.SystemColors.Window
        Me.lstBatchStatus.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstBatchStatus.Enabled = False
        Me.lstBatchStatus.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstBatchStatus.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstBatchStatus.ItemHeight = 14
        Me.lstBatchStatus.Items.AddRange(New Object() {"01 - Ready For Scan", "02 - Scanning", "03 - Suspended From Scan", "04 - Scan Complete", "05 - Released", "88 - Issue Folder"})
        Me.lstBatchStatus.Location = New System.Drawing.Point(24, 200)
        Me.lstBatchStatus.Name = "lstBatchStatus"
        Me.lstBatchStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstBatchStatus.Size = New System.Drawing.Size(169, 88)
        Me.lstBatchStatus.TabIndex = 3
        '
        'txtBatchName
        '
        Me.txtBatchName.AcceptsReturn = True
        Me.txtBatchName.BackColor = System.Drawing.Color.Black
        Me.txtBatchName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBatchName.ForeColor = System.Drawing.Color.Red
        Me.txtBatchName.Location = New System.Drawing.Point(24, 88)
        Me.txtBatchName.MaxLength = 0
        Me.txtBatchName.Name = "txtBatchName"
        Me.txtBatchName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBatchName.Size = New System.Drawing.Size(161, 25)
        Me.txtBatchName.TabIndex = 1
        '
        'cmdReset
        '
        Me.cmdReset.BackColor = System.Drawing.SystemColors.Control
        Me.cmdReset.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdReset.Enabled = False
        Me.cmdReset.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReset.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdReset.Location = New System.Drawing.Point(27, 161)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdReset.Size = New System.Drawing.Size(101, 33)
        Me.cmdReset.TabIndex = 0
        Me.cmdReset.Text = "Reset"
        Me.cmdReset.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(184, 61)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(51, 17)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Status"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(16, 300)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(207, 36)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Once Batch has been found, Select desired batch status and Click Reset"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(24, 112)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(177, 33)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Enter batch Name to be Reset    Click ""Find Batch"" when ready"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmDoSResetBatchStatus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(235, 345)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.cmdFindBatch)
        Me.Controls.Add(Me.lstBatchStatus)
        Me.Controls.Add(Me.txtBatchName)
        Me.Controls.Add(Me.cmdReset)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(11, 57)
        Me.Name = "frmDoSResetBatchStatus"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Reset Batch Status"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class