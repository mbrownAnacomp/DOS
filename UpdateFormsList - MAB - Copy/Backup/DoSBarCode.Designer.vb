<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmDoSBarCode
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
    Public WithEvents _mnuFile_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public WithEvents Picture1 As System.Windows.Forms.PictureBox
    Public WithEvents Text1 As System.Windows.Forms.TextBox
    Public WithEvents lstSSNPrint As System.Windows.Forms.ListBox
    Public WithEvents cmdPrintDoS_Barcode As System.Windows.Forms.Button
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents mnuFile As Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDoSBarCode))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip
        Me._mnuFile_1 = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem
        Me.Picture1 = New System.Windows.Forms.PictureBox
        Me.Text1 = New System.Windows.Forms.TextBox
        Me.lstSSNPrint = New System.Windows.Forms.ListBox
        Me.cmdPrintDoS_Barcode = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.mnuFile = New Microsoft.VisualBasic.Compatibility.VB6.ToolStripMenuItemArray(Me.components)
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.MainMenu1.SuspendLayout()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuFile, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuFile_1})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(308, 24)
        Me.MainMenu1.TabIndex = 9
        '
        '_mnuFile_1
        '
        Me._mnuFile_1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuExit})
        Me.mnuFile.SetIndex(Me._mnuFile_1, CType(1, Short))
        Me._mnuFile_1.Name = "_mnuFile_1"
        Me._mnuFile_1.Size = New System.Drawing.Size(35, 20)
        Me._mnuFile_1.Text = "&File"
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
        Me.Picture1.Location = New System.Drawing.Point(136, 208)
        Me.Picture1.Name = "Picture1"
        Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture1.Size = New System.Drawing.Size(153, 153)
        Me.Picture1.TabIndex = 5
        Me.Picture1.TabStop = False
        '
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.BackColor = System.Drawing.Color.Black
        Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Text1.ForeColor = System.Drawing.Color.Red
        Me.Text1.Location = New System.Drawing.Point(136, 72)
        Me.Text1.MaxLength = 0
        Me.Text1.Name = "Text1"
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(153, 20)
        Me.Text1.TabIndex = 0
        '
        'lstSSNPrint
        '
        Me.lstSSNPrint.BackColor = System.Drawing.Color.Black
        Me.lstSSNPrint.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstSSNPrint.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstSSNPrint.ForeColor = System.Drawing.Color.Red
        Me.lstSSNPrint.ItemHeight = 14
        Me.lstSSNPrint.Location = New System.Drawing.Point(32, 72)
        Me.lstSSNPrint.Name = "lstSSNPrint"
        Me.lstSSNPrint.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstSSNPrint.Size = New System.Drawing.Size(81, 508)
        Me.lstSSNPrint.TabIndex = 2
        '
        'cmdPrintDoS_Barcode
        '
        Me.cmdPrintDoS_Barcode.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPrintDoS_Barcode.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPrintDoS_Barcode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPrintDoS_Barcode.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPrintDoS_Barcode.Location = New System.Drawing.Point(152, 128)
        Me.cmdPrintDoS_Barcode.Name = "cmdPrintDoS_Barcode"
        Me.cmdPrintDoS_Barcode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPrintDoS_Barcode.Size = New System.Drawing.Size(121, 49)
        Me.cmdPrintDoS_Barcode.TabIndex = 1
        Me.cmdPrintDoS_Barcode.Text = "Print Batch Control Number(s)"
        Me.cmdPrintDoS_Barcode.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.White
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(136, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(153, 17)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Enter SSN to be printed"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(24, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(97, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "SSNs to be printed"
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'frmDoSBarCode
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(308, 599)
        Me.Controls.Add(Me.Picture1)
        Me.Controls.Add(Me.Text1)
        Me.Controls.Add(Me.lstSSNPrint)
        Me.Controls.Add(Me.cmdPrintDoS_Barcode)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.Control
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(4, 50)
        Me.Name = "frmDoSBarCode"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Create Backfile BCSheets"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuFile, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
#End Region
End Class