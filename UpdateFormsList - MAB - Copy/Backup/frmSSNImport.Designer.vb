<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSSNImport
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
        Me.cmdImport = New System.Windows.Forms.Button
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.OpenSSNDialog = New System.Windows.Forms.OpenFileDialog
        Me.txtBrowseSSN = New System.Windows.Forms.TextBox
        Me.cmdBrowseSSN = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'cmdImport
        '
        Me.cmdImport.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cmdImport.Location = New System.Drawing.Point(171, 100)
        Me.cmdImport.Name = "cmdImport"
        Me.cmdImport.Size = New System.Drawing.Size(67, 23)
        Me.cmdImport.TabIndex = 0
        Me.cmdImport.Text = "Import"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(264, 100)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'OpenSSNDialog
        '
        Me.OpenSSNDialog.Filter = "CSV File|*.csv|Text File|*.txt"
        Me.OpenSSNDialog.Title = "Select Import File"
        '
        'txtBrowseSSN
        '
        Me.txtBrowseSSN.Location = New System.Drawing.Point(12, 44)
        Me.txtBrowseSSN.Name = "txtBrowseSSN"
        Me.txtBrowseSSN.Size = New System.Drawing.Size(433, 20)
        Me.txtBrowseSSN.TabIndex = 1
        '
        'cmdBrowseSSN
        '
        Me.cmdBrowseSSN.Location = New System.Drawing.Point(451, 44)
        Me.cmdBrowseSSN.Name = "cmdBrowseSSN"
        Me.cmdBrowseSSN.Size = New System.Drawing.Size(75, 23)
        Me.cmdBrowseSSN.TabIndex = 2
        Me.cmdBrowseSSN.Text = "Browse"
        Me.cmdBrowseSSN.UseVisualStyleBackColor = True
        '
        'frmSSNImport
        '
        Me.AcceptButton = Me.cmdImport
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(562, 150)
        Me.Controls.Add(Me.cmdBrowseSSN)
        Me.Controls.Add(Me.cmdImport)
        Me.Controls.Add(Me.txtBrowseSSN)
        Me.Controls.Add(Me.Cancel_Button)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSSNImport"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Select SSN file for import"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdImport As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents OpenSSNDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtBrowseSSN As System.Windows.Forms.TextBox
    Friend WithEvents cmdBrowseSSN As System.Windows.Forms.Button

End Class
