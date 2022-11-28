<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStart
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.mnuODBC = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDR = New System.Windows.Forms.ToolStripMenuItem
        Me.cmdHyphen = New System.Windows.Forms.ToolStripSeparator
        Me.mnuBackup = New System.Windows.Forms.ToolStripMenuItem
        Me.cmdHyphen2 = New System.Windows.Forms.ToolStripSeparator
        Me.cmdExit = New System.Windows.Forms.ToolStripMenuItem
        Me.cmdUFL = New System.Windows.Forms.Button
        Me.cmdBC = New System.Windows.Forms.Button
        Me.cmdLPXML = New System.Windows.Forms.Button
        Me.cmdXML = New System.Windows.Forms.Button
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ODBCConnectToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRecon = New System.Windows.Forms.ToolStripMenuItem
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuImportSSN = New System.Windows.Forms.ToolStripMenuItem
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'mnuODBC
        '
        Me.mnuODBC.Name = "mnuODBC"
        Me.mnuODBC.Size = New System.Drawing.Size(220, 22)
        Me.mnuODBC.Text = "Define ODBC Connection"
        '
        'mnuDR
        '
        Me.mnuDR.Name = "mnuDR"
        Me.mnuDR.Size = New System.Drawing.Size(220, 22)
        Me.mnuDR.Text = "Find Doc Review Project File"
        '
        'cmdHyphen
        '
        Me.cmdHyphen.Name = "cmdHyphen"
        Me.cmdHyphen.Size = New System.Drawing.Size(217, 6)
        '
        'mnuBackup
        '
        Me.mnuBackup.Name = "mnuBackup"
        Me.mnuBackup.Size = New System.Drawing.Size(220, 22)
        Me.mnuBackup.Text = "Backup Databases"
        '
        'cmdHyphen2
        '
        Me.cmdHyphen2.Name = "cmdHyphen2"
        Me.cmdHyphen2.Size = New System.Drawing.Size(217, 6)
        '
        'cmdExit
        '
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(220, 22)
        Me.cmdExit.Text = "E&xit"
        '
        'cmdUFL
        '
        Me.cmdUFL.Location = New System.Drawing.Point(34, 65)
        Me.cmdUFL.Name = "cmdUFL"
        Me.cmdUFL.Size = New System.Drawing.Size(184, 23)
        Me.cmdUFL.TabIndex = 1
        Me.cmdUFL.Text = "Update Forms List"
        Me.cmdUFL.UseVisualStyleBackColor = True
        '
        'cmdBC
        '
        Me.cmdBC.Location = New System.Drawing.Point(34, 115)
        Me.cmdBC.Name = "cmdBC"
        Me.cmdBC.Size = New System.Drawing.Size(184, 23)
        Me.cmdBC.TabIndex = 2
        Me.cmdBC.Text = "Bar Code Creator"
        Me.cmdBC.UseVisualStyleBackColor = True
        '
        'cmdLPXML
        '
        Me.cmdLPXML.Location = New System.Drawing.Point(34, 171)
        Me.cmdLPXML.Name = "cmdLPXML"
        Me.cmdLPXML.Size = New System.Drawing.Size(184, 23)
        Me.cmdLPXML.TabIndex = 3
        Me.cmdLPXML.Text = "LP XML Generator"
        Me.cmdLPXML.UseVisualStyleBackColor = True
        '
        'cmdXML
        '
        Me.cmdXML.Location = New System.Drawing.Point(34, 221)
        Me.cmdXML.Name = "cmdXML"
        Me.cmdXML.Size = New System.Drawing.Size(184, 23)
        Me.cmdXML.TabIndex = 4
        Me.cmdXML.Text = "XML Generator"
        Me.cmdXML.UseVisualStyleBackColor = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(258, 24)
        Me.MenuStrip1.TabIndex = 5
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ODBCConnectToolStripMenuItem, Me.mnuRecon, Me.mnuImportSSN, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(35, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'ODBCConnectToolStripMenuItem
        '
        Me.ODBCConnectToolStripMenuItem.Name = "ODBCConnectToolStripMenuItem"
        Me.ODBCConnectToolStripMenuItem.Size = New System.Drawing.Size(202, 22)
        Me.ODBCConnectToolStripMenuItem.Text = "&ODBC Connect"
        '
        'mnuRecon
        '
        Me.mnuRecon.Name = "mnuRecon"
        Me.mnuRecon.Size = New System.Drawing.Size(202, 22)
        Me.mnuRecon.Text = "Create reconcilliation file"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(202, 22)
        Me.ExitToolStripMenuItem.Text = "E&xit"
        '
        'mnuImportSSN
        '
        Me.mnuImportSSN.Name = "mnuImportSSN"
        Me.mnuImportSSN.Size = New System.Drawing.Size(202, 22)
        Me.mnuImportSSN.Text = "Import SSN from file"
        '
        'frmStart
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(258, 277)
        Me.Controls.Add(Me.cmdXML)
        Me.Controls.Add(Me.cmdLPXML)
        Me.Controls.Add(Me.cmdBC)
        Me.Controls.Add(Me.cmdUFL)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmStart"
        Me.Text = "DoS Scan Tracking System"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents mnuODBC As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuDR As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents cmdHyphen As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuBackup As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents cmdHyphen2 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents cmdExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmdUFL As System.Windows.Forms.Button
    Friend WithEvents cmdBC As System.Windows.Forms.Button
    Friend WithEvents cmdLPXML As System.Windows.Forms.Button
    Friend WithEvents cmdXML As System.Windows.Forms.Button
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ODBCConnectToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuRecon As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuImportSSN As System.Windows.Forms.ToolStripMenuItem
End Class
