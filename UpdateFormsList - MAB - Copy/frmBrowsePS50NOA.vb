Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmBrowsePS50NOA
    Inherits System.Windows.Forms.Form
    Private Sub cmdBrowseNOA_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowseNOA.Click
        Dim nTextLen As Short
        Dim MyResult As System.Windows.Forms.DialogResult
        ' Initialize the dialog to the file and path specified in the text box.
        ' If nothing is specified and no previous default directory has been
        ' set, default to C:\
        nTextLen = Len(txtBrowseNOA.Text)
        If nTextLen <> 0 Then
            CommonDialog1.InitialDirectory = txtBrowseNOA.Text
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
        CommonDialog1.Filter = "PS50 NOA Code Dictionary File (*.txt)|*.txt"
        ' Specify default filter
        CommonDialog1.FilterIndex = 1
        CommonDialog1.Title = "Select PS50 NOA Code Dictionary File"
        ' Display the Open dialog box
        MyResult = CommonDialog1.ShowDialog()
        If MyResult = DialogResult.Cancel Then GoTo Error_Dir
        ' Display name of selected file
        txtBrowseNOA.Text = CommonDialog1.FileName
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
        If UCase(VB.Right(Me.txtBrowseNOA.Text, 4)) = ".TXT" Then
            If UCase(Me.txtBrowseNOA.Text) <> UCase(sHoldPS50NOA) Then
                lRet = MsgBox("New PS50 NOA Code Dictionary file location will be active once program has been restarted", MsgBoxStyle.OkCancel, "Warning!")
                If lRet = MsgBoxResult.Cancel Then
                    Exit Sub
                Else
                    sPS50NOADir = Me.txtBrowseNOA.Text
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
    Private Sub frmBrowseNOA_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.txtBrowseNOA.Text = sPS50NOADir
        sHoldPS50NOA = Me.txtBrowseNOA.Text
    End Sub
End Class