Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class frmDoSBarCode
    Inherits System.Windows.Forms.Form

    Dim strBatchName As String

    Private Sub cmdPrintDoS_Barcode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrintDoS_Barcode.Click
        Dim Printer As New Printer
        Dim sQuery As String
        On Error GoTo Err_cmdPrintDoS_Barcode_Click

        Dim x, iSSNCount, Y As Short
        iSSNCount = lstSSNPrint.Items.Count
        Dim Z As System.Drawing.Font = System.Windows.Forms.Control.DefaultFont.Clone()
        Dim Q As System.Drawing.Font = System.Windows.Forms.Control.DefaultFont.Clone()
        If iSSNCount > 0 Then
            Z = VB6.FontChangeName(Z, "Ariel")
            Z = VB6.FontChangeItalic(Z, False)
            Q = VB6.FontChangeItalic(Q, False)
            Q = VB6.FontChangeName(Q, "Free 3 of 9 Extended")
            For x = 0 To iSSNCount - 1
                'Check the status of this BatchName - never re-use
                sQuery = "Select * From Batch Where SSN='" & VB6.GetItemString(lstSSNPrint, 0) & "' Order By BatchName"
                rs1 = db.OpenRecordset(sQuery, DAO.RecordsetTypeEnum.dbOpenDynaset, 512) '512 = dbSeeChanges
                If rs1.RecordCount < 1 Then
                    strBatchName = "*DoS" & VB6.GetItemString(lstSSNPrint, 0) & "000001M*"
                Else
                    rs1.MoveLast()
                    Y = Val(Mid(rs1.Fields("BatchName").Value, 13, 6)) + 1
                    strBatchName = "*DOS" & VB6.GetItemString(lstSSNPrint, 0) & VB6.Format(Y, "000000") & "M*"
                End If
                Printer.Font = Z
                Printer.FontSize = 32
                Printer.CurrentX = 2048
                Printer.CurrentY = 1048
                Printer.Print(Mid(strBatchName, 2, 19))
                Printer.FontSize = 32
                Printer.CurrentX = 2048
                Printer.CurrentY = 5048
                Printer.Print("Department of State")
                Printer.FontSize = 32
                Printer.CurrentX = 2048
                Printer.CurrentY = 11048
                Printer.Print("Batch Control Sheet")
                Printer.FontSize = 32
                Printer.CurrentX = 2048
                Printer.CurrentY = 13048
                Printer.Print("Created On " & VB6.Format(Today, "mm/dd/yyyy"))
                Printer.FontSize = 12
                Printer.CurrentX = 5048
                Printer.CurrentY = 6000
                Printer.Print("Prep __________")
                Printer.FontSize = 12
                Printer.CurrentX = 5048
                Printer.CurrentY = 7000
                Printer.Print("Scan __________")
                Printer.FontSize = 12
                Printer.CurrentX = 5048
                Printer.CurrentY = 8000
                Printer.Print("DR __________")
                Printer.FontSize = 12
                Printer.CurrentX = 5048
                Printer.CurrentY = 9000
                Printer.Print("Validation1 __________")
                Printer.FontSize = 12
                Printer.CurrentX = 5048
                Printer.CurrentY = 10000
                Printer.Print("Validation2 __________")
                'Print the Barcode
                Printer.Font = Q
                Printer.FontSize = 48
                Printer.CurrentX = 2048
                Printer.CurrentY = 3024
                Printer.Print(strBatchName)
                'Stick the DoS Seal out there
                'Printer.PaintPicture(Me.Picture1.Image, 2048, 7000)
                'Eject the page...
                Printer.EndDoc()
                lstSSNPrint.Items.RemoveAt((0))
                sQuery = "INSERT INTO Batch(BatchName, SSN, CreateDateTime, BatchStatus) "
                sQuery = sQuery & "VALUES('" & Mid(strBatchName, 2, 19) & "', '" & Mid(strBatchName, 5, 9) & "', '" & Now & "', 1)"
                db.Execute(sQuery)
                'With rs1
                '    .AddNew()
                '    .Fields("BatchName").Value = Mid(strBatchName, 2, 19)
                '    .Fields("SSN").Value = Mid(strBatchName, 5, 9)
                '    .Fields("CreateDateTime").Value = Now
                '    .Fields("BatchStatus").Value = 1
                '    .Update()
                'End With
            Next
            rs1.Close()
        Else
            MsgBox("There are no items to print")
        End If

Exit_cmdPrintDoS_Barcode_Click:

        On Error GoTo 0
        Exit Sub

Err_cmdPrintDoS_Barcode_Click:
        MsgBox("An Error Has occured" & vbCrLf & Err.Description & "Error Number:" & Err.Number, MsgBoxStyle.OkOnly, "Print Barcode Error")
        Resume Exit_cmdPrintDoS_Barcode_Click

    End Sub

    Private Sub frmDoSBarCode_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        frmStart.Enabled = True
    End Sub

    Private Sub frmDoSBarCode_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        On Error GoTo Err_frmDiaBarCode_Load
        'Dim sQuery As String

        KeyPreview = True
        Me.Text1.Focus()

Exit_frmDiaBarCode_Load:

        On Error GoTo 0
        Exit Sub

Err_frmDiaBarCode_Load:
        MsgBox("An Error Has occured" & vbCrLf & Err.Description & "Error Number:" & Err.Number, MsgBoxStyle.OkOnly, "Main Form Load Error")
        Resume Exit_frmDiaBarCode_Load

    End Sub

    Private Sub frmDoSBarCode_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        FileClose()
    End Sub

    Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click
        Me.Close()
    End Sub

    Private Sub Text1_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles Text1.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Dim nRet As Short
        Dim bGoAhead As Boolean
        Dim sSSNFind, sQuery As String
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If Len(Trim(Text1.Text)) <> 9 Or Not IsNumeric(Trim(Text1.Text)) Then
                MsgBox("You must enter a 9 character SSN - All Characters must be numeric")
            Else
                bGoAhead = False
                sSSNFind = Trim(Text1.Text)
                sQuery = "Select * From Batch Where SSN='" & sSSNFind & "'"
                rs1 = db.OpenRecordset(sQuery, DAO.RecordsetTypeEnum.dbOpenDynaset, 512) '512 = dbSeeChanges
                If rs1.RecordCount > 0 Then
                    nRet = MsgBox("That SSN has already been entered in the system" & vbCrLf & "Are you sure you want it entered again?", MsgBoxStyle.YesNo, "Duplicate SSN")
                    If nRet = MsgBoxResult.Yes Then
                        bGoAhead = True
                    End If
                Else
                    bGoAhead = True
                End If
                rs1.Close()
                If bGoAhead Then
                    sQuery = "Select * From DoS_SSN Where SSN='" & sSSNFind & "'"
                    rs = db.OpenRecordset(sQuery, DAO.RecordsetTypeEnum.dbOpenDynaset, 512) '512 = dbSeeChanges
                    If rs.RecordCount < 1 Then
                        nRet = MsgBox("That SSN was not in the DoS SSN List - add it anyway?", MsgBoxStyle.YesNo, "SSN Error")
                        If nRet = MsgBoxResult.Yes Then
                            lstSSNPrint.Items.Add(UCase(Trim(Text1.Text)))
                            'Add new SSN to DoS SSN List
                            sQuery = "INSERT INTO DoS_SSN (SSN) "
                            sQuery = sQuery & "VALUES('" & sSSNFind & "')"
                            db.Execute(sQuery)
                            Text1.Text = ""
                        End If
                    Else
                        lstSSNPrint.Items.Add(UCase(Trim(Text1.Text)))
                        Text1.Text = ""
                    End If
                    rs.Close()
                End If
            End If
        End If
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class