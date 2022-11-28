Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmModNOA
	Inherits System.Windows.Forms.Form
	Dim sOrigFrom, sOrigNOAC, sOrigFT, sOrigTo As String
	
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
    '        Dim sFormType, sNOAQuery, sNOACode As String
    '		Dim qRs As dao.Recordset
    '		sFormType = "'" & Trim(txtFT.Text) & "'"
    '		sNOACode = "'" & Trim(txtNOA.Text) & "'"
    '        sNOAQuery = "SELECT Type, NOAC FROM VMList_NewMaster WHERE Type=" & sFormType & " And NOAC=" & sNOACode
    '        qRs = db.OpenRecordset(sNOAQuery, DAO.RecordsetTypeEnum.dbOpenDynaset)
    '		If qRs.RecordCount > 0 Then
    '			MsgBox("A NOA Code with that Form Type already exists in the system", MsgBoxStyle.Critical, "Duplicates not allowed")
    '			qRs.Close()
    '			Cancel = True
    '		End If
    'EventExitSub: 
    '		eventArgs.Cancel = Cancel
    '	End Sub
	Private Sub cmdModify_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdModify.Click
        Dim sCurrentUser, sPurge, sFT, sDuplex, sUser As String
        Dim sFromDate, sNOA, sToDate As String
        Dim sNOAQuery As String
		
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

        sPurge = Trim(cmbPurge.Text)
		sDuplex = Trim(cmbDuplex.Text)
		
		sCurrentUser = Space(60)
		GetUserName(sCurrentUser, Len(sCurrentUser))
		'Get rid of extra spaces and cr/lf from API
		sUser = sCurrentUser
		sUser = Trim(sUser)
		sUser = VB.Left(sUser, Len(sUser) - 1) & " " & "@" & VB6.Format(Today, "mm/dd/yyyy") & " " & VB6.Format(TimeOfDay, "hh:mm:ss AMPM")

        'Master list Release db/table
        sNOAQuery = "UPDATE NewMaster SET Source ='" & sUser & "', Type='" & sFT & "', Purge='" & sPurge & "', Duplex='" & sDuplex & "' "
        sNOAQuery = sNOAQuery & "WHERE Type='" & sOrigFT & "' And NOAC='" & txtNOA.Text & "'"
        sNOAQuery = sNOAQuery & " And [NOA Eff From Date]='" & Trim(txtFromDate.Text)
        sNOAQuery = sNOAQuery & "' And [NOA Eff To Date]='" & Trim(txtToDate.Text) & "'"
        db.Execute(sNOAQuery, 64) '64 = dbSQLPassThrough

        'Validate/Completion VMList db/table
        sNOAQuery = "UPDATE VMList_NewMaster SET Source ='" & sUser & "', Type='" & sFT & "', Purge='" & sPurge & "', Duplex='" & sDuplex & "' "
        sNOAQuery = sNOAQuery & "WHERE Type='" & sOrigFT & "' And NOAC='" & txtNOA.Text & "'"
        sNOAQuery = sNOAQuery & " And [NOA Eff From Date]='" & Trim(txtFromDate.Text)
        sNOAQuery = sNOAQuery & "' And [NOA Eff To Date]='" & Trim(txtToDate.Text) & "'"
        db.Execute(sNOAQuery, 64) '64 = dbSQLPassThrough

        MsgBox("Records modified = success!")
		Call ClearFields()
	End Sub
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
    Private Sub frmModNOA_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Me.txtNOA.Focus()
    End Sub
	Private Sub frmModNOA_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		KeyPreview = True
        rsML = db.OpenRecordset("NewMaster", DAO.RecordsetTypeEnum.dbOpenDynaset)
		rsML.MoveLast()
		rsML.MoveFirst()
        rsVM = db.OpenRecordset("VMList_NewMaster", DAO.RecordsetTypeEnum.dbOpenDynaset)
		rsVM.MoveLast()
		rsVM.MoveFirst()
        cmbPurge.SelectedIndex = 1
		cmbDuplex.SelectedIndex = 1
	End Sub
	Private Sub frmModNOA_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		rsVM.Close()
		rsML.Close()
	End Sub
	Private Sub txtNOA_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtNOA.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim sNOACheckQuery, sNOACheck As String
		
		'Check the contents of the NOA Code box before proceeding
		If Trim(txtNOA.Text) = "" Then
			MsgBox("You must enter an Valid NOA code before proceeding", MsgBoxStyle.Critical, "Missing required field")
			Cancel = True
			GoTo EventExitSub
		End If
        'If Not IsNumeric(txtNOA.Text) And Len(Trim(txtNOA.Text)) <> 3 Then
        '          MsgBox("The NOA Code entered is not valid, it must be a numeric between 1 and 999")
        '	Cancel = True
        '	GoTo EventExitSub
        'End If
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
		
        'Check against VMTable
		sNOACheck = "'" & Trim(txtNOA.Text) & "'"
        sNOACheckQuery = "SELECT DISTINCT NOAC, Type, [NOA Eff From Date], [NOA Eff To Date], Purge, Duplex "
        sNOACheckQuery = sNOACheckQuery & "FROM VMList_NewMaster WHERE NOAC =" & sNOACheck
        rsNOACheck = db.OpenRecordset(sNOACheckQuery, DAO.RecordsetTypeEnum.dbOpenDynaset)
		If rsNOACheck.RecordCount > 0 Then
			rsNOACheck.MoveLast()
			rsNOACheck.MoveFirst()
			If rsNOACheck.RecordCount > 1 Then 'Ask Op to select correct NOAC record
				frmSelectNOA.ShowDialog()
			Else 'Only 1 possible NOAC - Fill the Fields
                txtFT.Text = MNS(rsNOACheck.Fields("Type").Value)
                txtFromDate.Text = MNS(rsNOACheck.Fields("NOA Eff From Date").Value)
                txtToDate.Text = MNS(rsNOACheck.Fields("NOA Eff To Date").Value)
                cmbDuplex.Text = IIf(MNS(rsNOACheck.Fields("Duplex").Value) = "X", "Yes", "No")
                cmbPurge.Text = IIf(MNS(rsNOACheck.Fields("Purge").Value) = "Yes", "Yes", "No")
			End If
			'Collect Original value to be modified for mod query
			sOrigFT = txtFT.Text
			cmdModify.Enabled = True
			Cancel = False
		Else
			MsgBox("NOA Code not found", MsgBoxStyle.Critical, "Error")
			Cancel = True
		End If
		rsNOACheck.Close()
		System.Windows.Forms.Application.DoEvents()
EventExitSub: 
		eventArgs.Cancel = Cancel
	End Sub
	Private Sub ClearFields()
		txtNOA.Text = ""
		txtFromDate.Text = ""
		txtToDate.Text = ""
		txtFT.Text = ""
		'cmbVS.ListIndex = 4
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
    ''' Return String if object is not null, else return empty.string
    Public Function MNS(ByVal s As Object) As String
        If IsDBNull(s) Then
            Return String.Empty
        Else
            Return Trim(CType(s, String))
        End If
    End Function
End Class