<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmBrowseOtherNOA
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
    Public WithEvents cmdCancel As System.Windows.Forms.Button
    Public WithEvents cmdSave As System.Windows.Forms.Button
    Public CommonDialog1 As System.Windows.Forms.OpenFileDialog
    Public WithEvents cmdBrowseNOA As System.Windows.Forms.Button
    Public WithEvents txtBrowseNOA As System.Windows.Forms.TextBox
    Public WithEvents lblBrowse As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdSave = New System.Windows.Forms.Button
        Me.CommonDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.cmdBrowseNOA = New System.Windows.Forms.Button
        Me.txtBrowseNOA = New System.Windows.Forms.TextBox
        Me.lblBrowse = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(304, 72)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(81, 33)
        Me.cmdCancel.TabIndex = 4
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSave.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSave.Location = New System.Drawing.Point(184, 72)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSave.Size = New System.Drawing.Size(81, 33)
        Me.cmdSave.TabIndex = 3
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'cmdBrowseNOA
        '
        Me.cmdBrowseNOA.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBrowseNOA.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBrowseNOA.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseNOA.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowseNOA.Location = New System.Drawing.Point(512, 32)
        Me.cmdBrowseNOA.Name = "cmdBrowseNOA"
        Me.cmdBrowseNOA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBrowseNOA.Size = New System.Drawing.Size(81, 25)
        Me.cmdBrowseNOA.TabIndex = 2
        Me.cmdBrowseNOA.Text = "Browse"
        Me.cmdBrowseNOA.UseVisualStyleBackColor = False
        '
        'txtBrowseNOA
        '
        Me.txtBrowseNOA.AcceptsReturn = True
        Me.txtBrowseNOA.BackColor = System.Drawing.Color.Black
        Me.txtBrowseNOA.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBrowseNOA.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBrowseNOA.ForeColor = System.Drawing.Color.Red
        Me.txtBrowseNOA.Location = New System.Drawing.Point(8, 32)
        Me.txtBrowseNOA.MaxLength = 0
        Me.txtBrowseNOA.Name = "txtBrowseNOA"
        Me.txtBrowseNOA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBrowseNOA.Size = New System.Drawing.Size(497, 22)
        Me.txtBrowseNOA.TabIndex = 0
        '
        'lblBrowse
        '
        Me.lblBrowse.BackColor = System.Drawing.SystemColors.Control
        Me.lblBrowse.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblBrowse.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBrowse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBrowse.Location = New System.Drawing.Point(8, 8)
        Me.lblBrowse.Name = "lblBrowse"
        Me.lblBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblBrowse.Size = New System.Drawing.Size(497, 25)
        Me.lblBrowse.TabIndex = 1
        Me.lblBrowse.Text = "Select All Other NOA Code Dictionary File"
        Me.lblBrowse.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmBrowseOtherNOA
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(603, 127)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdBrowseNOA)
        Me.Controls.Add(Me.txtBrowseNOA)
        Me.Controls.Add(Me.lblBrowse)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(3, 29)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBrowseOtherNOA"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.Text = "Select All Other NOA Code Dictionary File"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class