Imports System.Windows.Forms
Imports VB = Microsoft.VisualBasic

Public Class frmRecon

    Private Sub cmdCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreate.Click
        Dim sQuery, sBeginDate, sEndDate As String, rsRecon As DAO.Recordset, iFile As Short, iFieldCount As Short
        Dim x As Short, sPrintRecord As String, sReconFileName As String
        sBeginDate = Me.dtpStart.Text
        sEndDate = Me.dtpEnd.Text
        If CDate(sEndDate) < CDate(sBeginDate) Then
            MsgBox("Begin date must be prior to End date", MsgBoxStyle.Exclamation, "Error")
            Exit Sub
        End If
        sQuery = "Select * From Reconcilliation Where GUIDDate >= '" & sBeginDate & "' And GUIDDate <= '" & sEndDate & "'"
        rsRecon = db.OpenRecordset(sQuery, DAO.RecordsetTypeEnum.dbOpenDynaset, 64) '64 = dbPassThrough
        If rsRecon.RecordCount < 1 Then
            MsgBox("No reconcilliation records found - contact your administrator", MsgBoxStyle.Critical, "Error!")
            Exit Sub
        End If
        sReconFileName = "ReconRpt_" & VB6.Format(Str(Today.ToOADate), "yymmdd") & "_" & VB6.Format(Str(TimeOfDay.ToOADate), "hhmmss") & ".csv"
        iFile = FreeFile()
        If VB.Right(My.Application.Info.DirectoryPath, 1) = "\" Then
            FileOpen(iFile, My.Application.Info.DirectoryPath & sReconFileName, OpenMode.Output)
        Else
            FileOpen(iFile, My.Application.Info.DirectoryPath & "\" & sReconFileName, OpenMode.Output)
        End If
        rsRecon.MoveFirst()
        iFieldCount = rsRecon.Fields.Count
        Do Until rsRecon.EOF
            sPrintRecord = ""
            For x = 1 To iFieldCount - 1
                sPrintRecord = sPrintRecord & CStr(rsRecon.Fields(x - 1).Value) & ","
            Next
            sPrintRecord = sPrintRecord & CStr(rsRecon.Fields(iFieldCount - 1).Value)
            PrintLine(iFile, sPrintRecord)
            rsRecon.MoveNext()
        Loop
        FileClose()
        rsRecon.Close()
        MsgBox("Reconcilliation report has been created as " & vbCrLf & _
               My.Application.Info.DirectoryPath & "\" & sReconFileName, MsgBoxStyle.Information, "Report Complete")
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub frmRecon_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        frmStart.Enabled = True
    End Sub
End Class
