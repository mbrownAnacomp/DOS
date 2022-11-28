Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class frmRePrint
	Inherits System.Windows.Forms.Form
	
	Public Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
		Me.Close()
	End Sub
	
	Private Sub cmdRePrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRePrint.Click
		Dim Printer As New Printer
		On Error GoTo Err_cmdRePrint_Click
		
        Dim strRPBatchName, sQuery As String
		Dim Z As System.Drawing.Font = System.Windows.Forms.Control.DefaultFont.Clone()
		Dim Q As System.Drawing.Font = System.Windows.Forms.Control.DefaultFont.Clone()
		'UPGRADE_WARNING: Only TrueType and OpenType fonts are supported in Windows Forms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="971F4DF4-254E-44F4-861D-3AA0031FE361"'
		Z = VB6.FontChangeName(Z, "Ariel")
		Z = VB6.FontChangeItalic(Z, False)
		Q = VB6.FontChangeItalic(Q, False)
		Q = VB6.FontChangeName(Q, "Free 3 of 9 Extended")
		
		strRPBatchName = UCase(txtBatchName.Text)
		If Trim(strRPBatchName) = "" Then
			MsgBox("Please enter an existing Batch Name for Re-Print", MsgBoxStyle.OKOnly, "No data entered")
		Else
			If Len(Trim(strRPBatchName)) <> 19 Then
				MsgBox("The Batch Name must be 19 characters" & vbCrLf & "Your entry was only " & Str(Len(Trim(strRPBatchName))) & " characters", MsgBoxStyle.OKOnly, "Wrong number of characters entered")
			Else
                sQuery = "Select * From Batch Where BatchName = '" & Trim(strRPBatchName) & "'"
                rs1 = db.OpenRecordset(sQuery, DAO.RecordsetTypeEnum.dbOpenDynaset, 512) '512 = dbSeeChanges
                'rs1.Seek("=", Trim(strRPBatchName))
                If rs1.RecordCount < 1 Then
                    MsgBox("Batch Name " & Trim(strRPBatchName) & " not found", MsgBoxStyle.OkOnly, "Batch Name Not Found")
                Else
                    'Print the sucker
                    strRPBatchName = "*" & strRPBatchName & "*" 'add * for 3 of 9 format barcode
                    'UPGRADE_WARNING: Only TrueType and OpenType fonts are supported in Windows Forms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="971F4DF4-254E-44F4-861D-3AA0031FE361"'
                    Z = VB6.FontChangeName(Z, "Ariel")
                    Printer.Font = Z
                    Printer.FontSize = 32
                    Printer.CurrentX = 2048
                    Printer.CurrentY = 1048
                    Printer.Print(Mid(strRPBatchName, 2, 19))
                    Printer.FontSize = 32
                    Printer.CurrentX = 2048
                    Printer.CurrentY = 5048
                    Printer.Print("Defense Intelligence Agency")
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
                    Printer.Print("Completion __________")
                    Printer.FontSize = 12
                    Printer.CurrentX = 5048
                    Printer.CurrentY = 10000
                    Printer.Print("Validation __________")
                    'Print the Barcode
                    Printer.Font = Q
                    Printer.FontSize = 48
                    Printer.CurrentX = 2048
                    Printer.CurrentY = 3024
                    Printer.Print(strRPBatchName)
                    'Stick the DoS Seal out there
                    Printer.PaintPicture(frmDoSBarCode.Picture1.Image, 2048, 7000)
                    'Eject the page...
                    Printer.EndDoc()
                    txtBatchName.Text = ""
                End If
                rs1.Close()
			End If
		End If
Exit_cmdRePrint_Click: 
		
		On Error GoTo 0
		Exit Sub
		
Err_cmdRePrint_Click: 
        MsgBox("An Error Has occured in Re-Print" & vbCrLf & Err.Description & "Error Number:" & Err.Number, MsgBoxStyle.OkOnly, "Re-Print Error!")
		Resume Exit_cmdRePrint_Click
	End Sub
End Class