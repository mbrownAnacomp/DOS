Imports System.Windows.Forms

Public Class frmSSNImport
    Dim bGoodFile As Boolean
    Private Sub cmdImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImport.Click
        If bGoodFile Then
            Dim sr As New System.IO.StreamReader(OpenSSNDialog.FileName), lDupRecords As Long
            Dim sSSN As String, lNumRecs As Long, lBadRecs As Long, sQuery As String, rsSSN As DAO.Recordset
            Do Until sr.EndOfStream
                sSSN = sr.ReadLine()
                sSSN = Trim(Replace(sSSN, "-", ""))
                If IsNumeric(sSSN) And Len(sSSN) = 9 Then
                    'Check for duplicates
                    sQuery = "Select * From DoS_SSN Where SSN='" & sSSN & "'"
                    rsSSN = db.OpenRecordset(sQuery, DAO.RecordsetTypeEnum.dbOpenDynaset)
                    If rsSSN.RecordCount > 0 Then
                        'Add to duplicate record count
                        lDupRecords = lDupRecords + 1
                    Else
                        'Add Record if new
                        sQuery = "INSERT INTO DoS_SSN(SSN) VALUES('" & sSSN & "')"
                        db.Execute(sQuery)
                        lNumRecs = lNumRecs + 1
                    End If
                    rsSSN.Close()
                Else
                    lBadRecs = lBadRecs + 1
                End If
            Loop
            sr.Close()
            If lBadRecs < 1 Then lBadRecs = 0
            If lNumRecs < 1 Then lNumRecs = 0
            If lDupRecords < 1 Then lDupRecords = 0
            MsgBox(Str(lNumRecs) & " SSN records imported" & vbCrLf & Str(lBadRecs) & " Bad Records Found" & vbCrLf & Str(lDupRecords) & " Duplicate Records Found", MsgBoxStyle.OkOnly)
        Else
            MsgBox("You have not selected a valid import file/no action taken")
        End If
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub cmdBrowseSSN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseSSN.Click
        If OpenSSNDialog.ShowDialog() = DialogResult.OK Then
            bGoodFile = True
            txtBrowseSSN.Text = OpenSSNDialog.FileName
        End If
    End Sub

    Private Sub frmSSNImport_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        frmStart.Enabled = True
    End Sub
End Class
