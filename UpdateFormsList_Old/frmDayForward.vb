Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Imports VB = Microsoft.VisualBasic
Friend Class frmDayForward
	Inherits System.Windows.Forms.Form
	
	Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
        Dim Printer As New Printer
		On Error GoTo Err_cmdPrint_Click
		
        Dim sUniqueID, sCurrentUser, sBatchName, sUser, sQuery As String
		Dim Z As System.Drawing.Font = System.Windows.Forms.Control.DefaultFont.Clone()
		Dim Q As System.Drawing.Font = System.Windows.Forms.Control.DefaultFont.Clone()
		
		sCurrentUser = Space(60)
		GetUserName(sCurrentUser, Len(sCurrentUser))
        sUser = Trim(sCurrentUser)
        sUser = VB.Left(sUser, Len(sUser) - 1) 'Gets rid of Cr/Lf
        'sQuery = "INSERT INTO Batch(BatchName, SSN, CreateDateTime, BatchStatus) "
        'sQuery = sQuery & "VALUES('" & Mid(sBatchName, 2, 19) & "', '" & Mid(sBatchName, 5, 9) & "', '" & Now & "', 1)"
        sQuery = "INSERT INTO DoS_UniqueID(DateTimeCreated, Operator) "
        sQuery = sQuery & "Values('" & VB6.Format(Today, "mm/dd/yyyy") & " " & VB6.Format(TimeOfDay, "hh:mm:ss") & "', '" & Trim(sUser) & "')"
        db.Execute(sQuery, 64)
        sQuery = "Select * From DoS_UniqueID Order By UniqueID"
        rs1 = db.OpenRecordset(sQuery, DAO.RecordsetTypeEnum.dbOpenDynaset, 512) '512 = dbSeeChanges
        With rs1
            .MoveLast()
            sUniqueID = Trim(Str(.Fields("UniqueID").Value))
        End With
        rs1.Close()
		'Prepare to Print
		sUniqueID = VB6.Format(sUniqueID, "000000000")
        sBatchName = "*DOF" & sUniqueID & "000001M*"
		
        Z = VB6.FontChangeName(Z, "Ariel")
		Z = VB6.FontChangeItalic(Z, False)
		Q = VB6.FontChangeItalic(Q, False)
		Q = VB6.FontChangeName(Q, "Free 3 of 9 Extended")
		Printer.Font = Z
		Printer.FontSize = 32
		Printer.CurrentX = 2048
		Printer.CurrentY = 1048
		Printer.Print(Mid(sBatchName, 2, 19))
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
		Printer.CurrentY = 12048
		Printer.Print("Loose Paper")
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
		Printer.Print(sBatchName)
        'Stick the DoS Seal out there
        'Printer.PaintPicture(Me.Picture1.Image, 1500, 7000)
		'Eject the page...
		Printer.EndDoc()

        sQuery = "INSERT INTO Batch(BatchName, SSN, CreateDateTime, BatchStatus) "
        sQuery = sQuery & "VALUES('" & Mid(sBatchName, 2, 19) & "', '" & Mid(sBatchName, 5, 9) & "', '" & Now & "', 1)"
        db.Execute(sQuery, 64)

Exit_cmdPrint_Click: 
		
		On Error GoTo 0
		Exit Sub
		
Err_cmdPrint_Click: 
		If Err.Number = 482 Then 'Operator canceled in printer panel
			MsgBox("An Error Has occured" & vbCrLf & "Description: " & Err.Description & vbCrLf & "Error Number:" & Err.Number & vbCrLf & "Print operation canceled by operator - no BCS generated", MsgBoxStyle.OKOnly, "Day Forward Print Error")
			Resume Exit_cmdPrint_Click
		Else
			MsgBox("An Error Has occured" & vbCrLf & "Description: " & Err.Description & vbCrLf & "Error Number:" & Err.Number, MsgBoxStyle.OKOnly, "Day Forward Print Error")
			Resume Exit_cmdPrint_Click
		End If
		
	End Sub
		
	Private Sub frmDayForward_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error GoTo Err_frmDiaBarCode_Load
		KeyPreview = True

Exit_frmDiaBarCode_Load: 
		
		On Error GoTo 0
		Exit Sub
		
Err_frmDiaBarCode_Load: 
		MsgBox("An Error Has occured" & vbCrLf & Err.Description & "Error Number:" & Err.Number, MsgBoxStyle.OKOnly, "Day Forward Form Load Error")
		Resume Exit_frmDiaBarCode_Load
	End Sub
	
	Private Sub frmDayForward_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        frmOpenOption.Show()
	End Sub
	Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click
		Me.Close()
	End Sub
End Class