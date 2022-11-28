Option Strict Off
Option Explicit On
Friend Class frmDoSDeleteBatch
    Inherits System.Windows.Forms.Form

    Private Sub cmdDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDelete.Click
        Dim lRet As Short, sQuery As String
        On Error GoTo Err_cmdDelete_Click
        'See if the BatchName exists
        sQuery = "Select * From Batch Where BatchName = '" & UCase(Trim(txtBatchName.Text)) & "'"
        rs1 = db.OpenRecordset(sQuery, DAO.RecordsetTypeEnum.dbOpenDynaset, 512) '512 = dbSeeChanges
        If rs1.RecordCount > 0 Then
            lRet = MsgBox("This Batch Name and all its stats will be deleted" & vbCrLf & "Are you sure?", MsgBoxStyle.YesNo, "Delete Confirm")
            rs1.Close()
            If lRet = MsgBoxResult.Yes Then
                'rs1.Delete()
                sQuery = "DELETE FROM Batch WHERE BatchName = '" & UCase(Trim(txtBatchName.Text)) & "'"
                db.Execute(sQuery, 64)
                MsgBox("Record for Batch Name " & Trim(txtBatchName.Text) & " has been deleted", MsgBoxStyle.OkOnly, "Delete Complete")
                txtBatchName.Text = ""
                txtBatchName.Focus()
            Else
                txtBatchName.Focus()
                txtBatchName.SelectionStart = 0
                txtBatchName.SelectionLength = Len(Trim(txtBatchName.Text))
            End If
        Else
            MsgBox("Batch Name as entered not found", MsgBoxStyle.OkOnly, "Record Not Found")
            rs1.Close()
            txtBatchName.Focus()
            txtBatchName.SelectionStart = 0
            txtBatchName.SelectionLength = Len(Trim(txtBatchName.Text))
        End If
Exit_cmdDelete_Click:

        On Error GoTo 0
        Exit Sub

Err_cmdDelete_Click:
        MsgBox("An Error Has occured" & vbCrLf & Err.Description & "Error Number:" & Err.Number, MsgBoxStyle.OKOnly, "Delete Batch Error")
        Resume Exit_cmdDelete_Click
    End Sub

    Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click
        Me.Close()
    End Sub
End Class