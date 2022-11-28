<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmODBC
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtDSN = New System.Windows.Forms.TextBox()
        Me.txtDB = New System.Windows.Forms.TextBox()
        Me.txtUID = New System.Windows.Forms.TextBox()
        Me.txtPWD = New System.Windows.Forms.TextBox()
        Me.chkAuthentication = New System.Windows.Forms.CheckBox()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ODBC DSN"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 60)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Database"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(16, 147)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "SQL User ID"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(16, 188)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(77, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "SQL Password"
        '
        'txtDSN
        '
        Me.txtDSN.AcceptsReturn = True
        Me.txtDSN.Location = New System.Drawing.Point(99, 26)
        Me.txtDSN.Name = "txtDSN"
        Me.txtDSN.Size = New System.Drawing.Size(148, 20)
        Me.txtDSN.TabIndex = 0
        '
        'txtDB
        '
        Me.txtDB.AcceptsReturn = True
        Me.txtDB.Location = New System.Drawing.Point(99, 60)
        Me.txtDB.Name = "txtDB"
        Me.txtDB.Size = New System.Drawing.Size(148, 20)
        Me.txtDB.TabIndex = 1
        '
        'txtUID
        '
        Me.txtUID.Location = New System.Drawing.Point(99, 140)
        Me.txtUID.Name = "txtUID"
        Me.txtUID.Size = New System.Drawing.Size(148, 20)
        Me.txtUID.TabIndex = 3
        Me.txtUID.UseSystemPasswordChar = True
        '
        'txtPWD
        '
        Me.txtPWD.Location = New System.Drawing.Point(99, 181)
        Me.txtPWD.Name = "txtPWD"
        Me.txtPWD.Size = New System.Drawing.Size(148, 20)
        Me.txtPWD.TabIndex = 4
        Me.txtPWD.UseSystemPasswordChar = True
        Me.txtPWD.UseWaitCursor = True
        '
        'chkAuthentication
        '
        Me.chkAuthentication.AutoSize = True
        Me.chkAuthentication.Checked = True
        Me.chkAuthentication.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAuthentication.Location = New System.Drawing.Point(48, 99)
        Me.chkAuthentication.Name = "chkAuthentication"
        Me.chkAuthentication.Size = New System.Drawing.Size(163, 17)
        Me.chkAuthentication.TabIndex = 2
        Me.chkAuthentication.Text = "Use Windows Authentication"
        Me.chkAuthentication.UseVisualStyleBackColor = True
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(39, 240)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(75, 23)
        Me.cmdSave.TabIndex = 5
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(160, 240)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 6
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'frmODBC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CausesValidation = False
        Me.ClientSize = New System.Drawing.Size(278, 293)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.chkAuthentication)
        Me.Controls.Add(Me.txtPWD)
        Me.Controls.Add(Me.txtUID)
        Me.Controls.Add(Me.txtDB)
        Me.Controls.Add(Me.txtDSN)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmODBC"
        Me.Tag = ""
        Me.Text = "Input ODBC Conncection"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtDSN As System.Windows.Forms.TextBox
    Friend WithEvents txtDB As System.Windows.Forms.TextBox
    Friend WithEvents txtUID As System.Windows.Forms.TextBox
    Friend WithEvents txtPWD As System.Windows.Forms.TextBox
    Friend WithEvents chkAuthentication As System.Windows.Forms.CheckBox
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
End Class
