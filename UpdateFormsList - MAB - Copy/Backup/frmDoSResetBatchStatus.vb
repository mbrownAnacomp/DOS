Option Strict Off
Option Explicit On
Friend Class frmDoSResetBatchStatus
    Inherits System.Windows.Forms.Form
    Private Sub cmdReset_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReset.Click
        Dim bSelected As Boolean, sQuery As String
        Dim iStatus, x As Short

        On Error GoTo Err_cmdReset_Click
        'See which box is checked
        With lstBatchStatus
            iStatus = .Items.Count
            For x = 0 To iStatus - 1
                bSelected = .GetSelected(x)
                If bSelected = True Then
                    txtStatus.Text = Mid(.SelectedItem, 1, 2)
                    'Master list Release db/table
                    sQuery = "UPDATE Batch SET BatchStatus ='" & txtStatus.Text & "'"
                    sQuery = sQuery & " Where BatchName = '" & UCase(Trim(txtBatchName.Text)) & "'"
                    db.Execute(sQuery, 64) '64 = dbSQLPassThrough
                    MsgBox("BatchStatus of " & UCase(Trim(txtBatchName.Text)) & " reset to " & txtStatus.Text)
                    Exit For
                End If
            Next
        End With
Exit_cmdReset_Click:

        On Error GoTo 0
        Exit Sub

Err_cmdReset_Click:
        MsgBox("An Error Has occured in Reset" & vbCrLf & Err.Description & "Error Number:" & Err.Number, MsgBoxStyle.OKOnly, "Reset Batch Status Error")
        Resume Exit_cmdReset_Click
    End Sub
    Private Sub cmdFindBatch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFindBatch.Click

        On Error GoTo Err_cmdFindBatch_Click
        Dim iStatus As Short, sQuery As String
        'See if the BatchName exists
        sQuery = "Select * From Batch Where BatchName = '" & UCase(Trim(txtBatchName.Text)) & "'"
        rs1 = db.OpenRecordset(sQuery, DAO.RecordsetTypeEnum.dbOpenDynaset, 512) '512 = dbSeeChanges

        If rs1.RecordCount > 0 Then
            iStatus = rs1.Fields("BatchStatus").Value
            txtStatus.Text = VB6.Format(iStatus, "00")
            cmdReset.Enabled = True
            lstBatchStatus.Enabled = True
            lstBatchStatus.SetSelected(iStatus - 1, True)
        Else
            MsgBox("Batch Name as entered not found", MsgBoxStyle.OkOnly, "Record Not Found")
            txtBatchName.Focus()
            txtBatchName.SelectionStart = 0
            txtBatchName.SelectionLength = Len(Trim(txtBatchName.Text))
            cmdReset.Enabled = False
            txtStatus.Text = ""
            lstBatchStatus.Enabled = False
        End If
        rs1.Close()

Exit_cmdFindBatch_Click:

        On Error GoTo 0
        Exit Sub

Err_cmdFindBatch_Click:
        MsgBox("An Error Has occured in Find" & vbCrLf & Err.Description & "Error Number:" & Err.Number, MsgBoxStyle.OKOnly, "Find Batch Error")
        Resume Exit_cmdFindBatch_Click
    End Sub

    Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click
        Me.Close()
    End Sub
End Class