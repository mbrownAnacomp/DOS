Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmAddNOA
	Inherits System.Windows.Forms.Form
	Private Sub txtFromDate_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtFromDate.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		If Not IsDate(txtFromDate.Text) Then
			MsgBox("Effective From Date field must be a valid date (mm/dd/yyyy)")
			txtFromDate.Text = ""
			Cancel = True
			GoTo EventExitSub
		ElseIf Trim(txtFromDate.Text) = "" Then 
			MsgBox("Effective From Date field cannot be blank")
			Cancel = True
			GoTo EventExitSub
		End If
EventExitSub: 
		eventArgs.Cancel = Cancel
	End Sub
    '	Private Sub txtFT_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtFT.Validating
    '		Dim Cancel As Boolean = eventArgs.Cancel
    '		If Trim(txtFT.Text) = "" Then
    '			MsgBox("Form Type cannot be left blank")
    '			Cancel = True
    '			GoTo EventExitSub
    '		End If
    '		txtFT.Text = UCase(txtFT.Text)

    '		'Check the db's to see if it already exists
    '        Dim sFormType, sNoaQuery, sNOACode As String
    '		Dim qRs As dao.Recordset
    '		sFormType = "'" & Trim(txtFT.Text) & "'"
    '		sNOACode = "'" & Trim(txtNOA.Text) & "'"
    '        sNoaQuery = "SELECT Type, NOAC FROM VMList_NewMaster WHERE Type=" & sFormType & " And NOAC=" & sNOACode
    '        qRs = db.OpenRecordset(sNoaQuery, DAO.RecordsetTypeEnum.dbOpenDynaset)
    '		If qRs.RecordCount > 0 Then
    '			MsgBox("A NOA Code with that Form Type already exists in the system", MsgBoxStyle.Critical, "Duplicates not allowed")
    '			qRs.Close()
    '			Cancel = True
    '		End If
    'EventExitSub: 
    '		eventArgs.Cancel = Cancel
    '	End Sub
	Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
        Dim sCurrentUser, sPurge, sFT, sVS, sDuplex, sUser As String, rsForms As DAO.Recordset
        Dim sFromDate, sNOA, sToDate, sQuery, sNOAOutFile As String, iNOAOut As Short
		
		'Check for blanks
		sNOA = Trim(txtNOA.Text)
		If sNOA = "" Then
			MsgBox("NOA Code field cannot be blank")
			txtNOA.Focus()
			Exit Sub
		End If
		sFT = Trim(txtFT.Text)
		'Check for blanks
		If sFT = "" Then
			MsgBox("Form Type field cannot be blank")
			txtFT.Focus()
			Exit Sub
		End If
		sFromDate = Trim(txtFromDate.Text)
		If sFromDate = "" Then
			MsgBox("Effective From Date field cannot be blank")
			txtFromDate.Focus()
			Exit Sub
		End If
		sToDate = Trim(txtToDate.Text)
		If sToDate = "" Then
			MsgBox("Effective To Date field cannot be blank")
			txtToDate.Focus()
			Exit Sub
		End If
		'Check to see that dates are logical
		If CDate(sFromDate) > CDate(sToDate) Then 'can't be
			MsgBox("Effective From Date must be earlier than Efffective To Date", MsgBoxStyle.Critical, "Date Error!")
			txtFromDate.Focus()
			Exit Sub
		End If
		sVS = Trim(cmbVS.Text)
		'Check for blanks
		If sVS = "" Then
			MsgBox("Virtual Side field cannot be blank")
			cmbVS.Focus()
			Exit Sub
		End If
        sNOA = IIf(IsNumeric(sNOA), VB6.Format(sNOA, "000"), sNOA)
        sPurge = Trim(cmbPurge.Text)
        sPurge = IIf(sPurge = "Yes", "Yes", "")
        sDuplex = Trim(cmbDuplex.Text)
        sDuplex = IIf(sDuplex = "Yes", "X", "")
        sFromDate = VB6.Format(sFromDate, "m/d/yyyy")
        sToDate = VB6.Format(sToDate, "m/d/yyyy")
		sCurrentUser = Space(60)
		GetUserName(sCurrentUser, Len(sCurrentUser))
		'Get rid of extra spaces and cr/lf from API
		sUser = sCurrentUser
		sUser = Trim(sUser)
		sUser = VB.Left(sUser, Len(sUser) - 1) & " " & "@" & VB6.Format(Today, "mm/dd/yyyy") & " " & VB6.Format(TimeOfDay, "hh:mm:ss AMPM")
		
		'======================Add new NOA info to db
        'Master list Release db/table
        If IsNumeric(sNOA) Then 'update all NOA Forms
            sQuery = "INSERT INTO NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'PS50', 'PS 50', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'SF50', 'SF 50', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'SF50A', 'SF 50-A', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'SF50B', 'SF 50-B', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'SF52', 'SF 52', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'AD347', 'AD 347', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'AD350', 'AD 350', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'AD3501', 'AD 350-1', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'AD350A', 'AD 350A', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'AD350B', 'AD 350B', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'AO250', 'AO 250', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'DL50', 'DL50', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'FSA50', 'FSA 50', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'VA546444', 'VA 5-4644-4', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'VA54650', 'VA 5-4650', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
        End If

        'Validate/Completion VMList db/table
        If IsNumeric(sNOA) Then 'Update all NOA forms
            sQuery = "INSERT INTO VMList_NewMAster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'PS50', 'PS 50', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO VMList_NewMAster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'SF50', 'SF 50', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO VMList_NewMAster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'SF50A', 'SF 50-A', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO VMList_NewMAster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'SF50B', 'SF 50-B', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO VMList_NewMAster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'SF52', 'SF 52', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO VMList_NewMAster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'AD347', 'AD 347', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO VMList_NewMAster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'AD350', 'AD 350', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO VMList_NewMAster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'AD3501', 'AD 350-1', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO VMList_NewMAster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'AD350A', 'AD 350A', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO VMList_NewMAster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'AD350B', 'AD 350B', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO VMList_NewMAster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'AO250', 'AO 250', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO VMList_NewMAster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'DL50', 'DL 50', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO VMList_NewMAster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'FSA50', 'FSA 50', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO VMList_NewMAster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'VA546444', 'VA 5-4644-4', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
            sQuery = "INSERT INTO VMList_NewMAster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, NOAC, "
            sQuery = sQuery & "[NOA Eff From Date], [NOA Eff To Date], Duplex, Purge) "
            sQuery = sQuery & "VALUES('" & sUser & "', 'VA54650', 'VA 5-4650', '" & sFT & "', '" & sVS & "', '"
            sQuery = sQuery & sNOA & "', '" & sFromDate & "', '" & sToDate & "', '" & sDuplex & "', '" & sPurge & "')"
            db.Execute(sQuery)
        End If
        'For KTM - must update Dictionary files for Validation - 3 NOA types
        sQuery = "Select Distinct NOAC From NewMaster Where NOAC Is Not NULL Order By NOAC" 'SF50 Query - all NOAC
        rsForms = db.OpenRecordset(sQuery, DAO.RecordsetTypeEnum.dbOpenDynaset)
        rsForms.MoveFirst()
        iNOAOut = FreeFile()
        sNOAOutFile = sNOADir & "Tmp"
        FileOpen(iNOAOut, sNOAOutFile, OpenMode.Output)
        'Print NOA records from DB Table
        'Add a blank in there is no NOA Code
        PrintLine(iNOAOut, "   ")
        Do Until rsForms.EOF
            PrintLine(iNOAOut, rsForms.Fields("NOAC").Value)
            rsForms.MoveNext()
        Loop

        rsForms.Close()
        FileClose()
        If fs.FileExists(sNOADir) Then fs.MoveFile(sNOADir, VB.Left(sNOADir, Len(sNOADir) - 4) & "_" & VB6.Format(Str(Today.ToOADate), "yymmdd") & "_" & VB6.Format(Str(TimeOfDay.ToOADate), "hhmmss") & ".txt")
        'Rename new Dictionary file to original Dictionary file name
        fs.MoveFile(sNOAOutFile, sNOADir)

        sQuery = "Select Distinct NOAC From NewMaster Where NOAC Is Not NULL AND FormNameNumber = 'PS50' Order By NOAC" 'PS50 Query 
        rsForms = db.OpenRecordset(sQuery, DAO.RecordsetTypeEnum.dbOpenDynaset)
        rsForms.MoveFirst()
        iNOAOut = FreeFile()
        sNOAOutFile = sPS50NOADir & "Tmp"
        FileOpen(iNOAOut, sNOAOutFile, OpenMode.Output)
        'Print NOA records from DB Table
        Do Until rsForms.EOF
            PrintLine(iNOAOut, rsForms.Fields("NOAC").Value)
            rsForms.MoveNext()
        Loop
        rsForms.Close()
        FileClose()
        If fs.FileExists(sPS50NOADir) Then fs.MoveFile(sPS50NOADir, VB.Left(sPS50NOADir, Len(sNOADir) - 4) & "_" & VB6.Format(Str(Today.ToOADate), "yymmdd") & "_" & VB6.Format(Str(TimeOfDay.ToOADate), "hhmmss") & ".txt")
        'Rename new Dictionary file to original Dictionary file name
        fs.MoveFile(sNOAOutFile, sPS50NOADir)

        sQuery = "Select Distinct NOAC From NewMaster Where NOAC Is Not NULL AND FormNameNumber = 'SF52' Order By NOAC" 'All Other Query 
        rsForms = db.OpenRecordset(sQuery, DAO.RecordsetTypeEnum.dbOpenDynaset)
        rsForms.MoveFirst()
        iNOAOut = FreeFile()
        sNOAOutFile = sOtherNOADir & "Tmp"
        FileOpen(iNOAOut, sNOAOutFile, OpenMode.Output)
        'Print NOA records from DB Table
        Do Until rsForms.EOF
            PrintLine(iNOAOut, rsForms.Fields("NOAC").Value)
            rsForms.MoveNext()
        Loop
        rsForms.Close()
        FileClose()
        If fs.FileExists(sOtherNOADir) Then fs.MoveFile(sOtherNOADir, VB.Left(sOtherNOADir, Len(sNOADir) - 4) & "_" & VB6.Format(Str(Today.ToOADate), "yymmdd") & "_" & VB6.Format(Str(TimeOfDay.ToOADate), "hhmmss") & ".txt")
        'Rename new Dictionary file to original Dictionary file name
        fs.MoveFile(sNOAOutFile, sOtherNOADir)

        MsgBox("New record added = success!")
        Call ClearFields()
    End Sub
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
    Private Sub frmAddNOA_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Me.txtNOA.Focus()
    End Sub
	Private Sub frmAddNOA_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		KeyPreview = True
        rsML = db.OpenRecordset("NewMaster", DAO.RecordsetTypeEnum.dbOpenDynaset)
		rsML.MoveLast()
		rsML.MoveFirst()
        rsVM = db.OpenRecordset("VMList_NewMaster", DAO.RecordsetTypeEnum.dbOpenDynaset)
		rsVM.MoveLast()
		rsVM.MoveFirst()
		cmbVS.SelectedIndex = 4
		cmbPurge.SelectedIndex = 1
		cmbDuplex.SelectedIndex = 1
	End Sub
	Private Sub frmAddNOA_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		rsVM.Close()
		rsML.Close()
	End Sub
	Private Sub txtNOA_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtNOA.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		'Check the contents of the NOA Code box before proceeding
		If Trim(txtNOA.Text) = "" Then
			MsgBox("You must enter an Valid NOA code before proceeding", MsgBoxStyle.Critical, "Missing required field")
			Cancel = True
			GoTo EventExitSub
		End If
        If Not IsNumeric(txtNOA.Text) <> 3 Then
            MsgBox("The NOA Code entered is not valid, it must be a numeric between 1 and 999")
            Cancel = True
            GoTo EventExitSub
        End If
		If IsNumeric(txtNOA.Text) And (Val(txtNOA.Text) <= 0 Or Val(txtNOA.Text) > 999) Then
			MsgBox("A numeric NOA Code must have a value between 1 and 999")
			Cancel = True
			GoTo EventExitSub
		End If
		
		If Not IsNumeric(txtNOA.Text) Then
			txtNOA.Text = UCase(txtNOA.Text)
		Else
			txtNOA.Text = VB6.Format(txtNOA.Text, "000")
		End If
		
EventExitSub: 
		eventArgs.Cancel = Cancel
	End Sub
	Private Sub ClearFields()
		txtNOA.Text = ""
		txtFromDate.Text = ""
		txtToDate.Text = ""
		txtFT.Text = ""
		cmbVS.SelectedIndex = 4
		cmbPurge.SelectedIndex = 1
		cmbDuplex.SelectedIndex = 1
	End Sub
	Private Sub txtToDate_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtToDate.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		If Not IsDate(txtToDate.Text) Then
			MsgBox("Effective To Date field must be a valid date (mm/dd/yyyy)")
			txtToDate.Text = ""
			Cancel = True
			GoTo EventExitSub
		ElseIf Trim(txtToDate.Text) = "" Then 
			MsgBox("Effective To Date cannot be blank (mm/dd/yyyy)")
			Cancel = True
			GoTo EventExitSub
		End If
EventExitSub: 
		eventArgs.Cancel = Cancel
	End Sub
End Class