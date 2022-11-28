Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmBrowseDR
	Inherits System.Windows.Forms.Form
	
	Private Sub cmdBrowseDR_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowseDR.Click
        Dim nTextLen As Short
        Dim MyResult As System.Windows.Forms.DialogResult

        ' Initialize the dialog to the file and path specified in the text box.
		' If nothing is specified and no previous default directory has been
		' set, default to C:\
		nTextLen = Len(txtBrowseDR.Text)
		If nTextLen <> 0 Then
            CommonDialog1.InitialDirectory = txtBrowseDR.Text
        Else
            If CommonDialog1.InitialDirectory = "" Then
                CommonDialog1.InitialDirectory = My.Application.Info.DirectoryPath & "\"
            End If
        End If
        ' Set flags
        CommonDialog1.ShowReadOnly = False
        CommonDialog1.ValidateNames = False
        CommonDialog1.CheckPathExists = True
        ' Set filters
        CommonDialog1.Filter = "Dictionary File (*.txt)|*.txt"
        ' Specify default filter
        CommonDialog1.FilterIndex = 1
        CommonDialog1.Title = "Select Form Dictionary File"
        ' Display the Open dialog box
        MyResult = CommonDialog1.ShowDialog()
        If MyResult = DialogResult.Cancel Then GoTo Error_Dir
        ' Display name of selected file
        txtBrowseDR.Text = CommonDialog1.FileName
        Exit Sub

Error_Dir:
        'User pressed the Cancel button
        ' Always reset the dialog for next time
        CommonDialog1.FileName = ""
        CommonDialog1.InitialDirectory = ""
    End Sub
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		Dim lRet As Short
        If UCase(VB.Right(Me.txtBrowseDR.Text, 4)) = ".TXT" Then
            If UCase(Me.txtBrowseDR.Text) <> UCase(sHoldDRP) Then
                lRet = MsgBox("New Form Dictionary file location will be active once program has been restarted", MsgBoxStyle.OkCancel, "Warning!")
                If lRet = MsgBoxResult.Cancel Then
                    Exit Sub
                Else
                    sDRPDir = Me.txtBrowseDR.Text
                    Me.Close()
                End If
            Else 'No change
                Me.Close()
            End If
        Else
            lRet = MsgBox("You have not browsed to a valid Dictionary file" & vbCrLf & "Do you want to try Again?", MsgBoxStyle.YesNo, "Invald....")
            If lRet = MsgBoxResult.Yes Then
                Exit Sub
            Else
                Me.Close()
            End If
        End If
	End Sub
    Private Sub frmBrowseDR_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.txtBrowseDR.Text = sDRPDir
        sHoldDRP = Me.txtBrowseDR.Text
    End Sub
End Class