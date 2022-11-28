Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmModFN
	Inherits System.Windows.Forms.Form
	Private Sub cmbFT_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles cmbFT.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		cmbFT.Text = UCase(cmbFT.Text)
		eventArgs.Cancel = Cancel
	End Sub
	Private Sub cmdModify_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdModify.Click
		Dim sDuplex, sVS, sOFN, sFN, sFT, sPurge, sCurrentUser As String
        Dim sUser As String, sQuery As String

		sFN = Trim(txtFN.Text)
		sOFN = Trim(txtOFN.Text)
		'Check for blanks
		If sOFN = "" Then
			MsgBox("Original Form Name Number field cannot be blank")
			txtOFN.Focus()
			Exit Sub
		End If
		'Check to see that it matches FN with blanks and special characters removed, give option to proceed if not
        sFT = Trim(cmbFT.Text)
		'Check for blanks
		If sFT = "" Then
			MsgBox("Form Type field cannot be blank")
			cmbFT.Focus()
			Exit Sub
		End If
		sVS = Trim(cmbVS.Text)
		'Check for blanks
		If sVS = "" Then
			MsgBox("Virtual Side field cannot be blank")
			cmbVS.Focus()
			Exit Sub
		End If
        sPurge = Trim(cmbPurge.Text)
        sPurge = IIf(sPurge = "Yes", "Yes", "")
        sDuplex = Trim(cmbDuplex.Text)
        sDuplex = IIf(sDuplex = "Yes", "X", "")
        sCurrentUser = Space(60)
        GetUserName(sCurrentUser, Len(sCurrentUser))
        'Get rid of extra spaces and cr/lf from API
        sUser = sCurrentUser
        sUser = Trim(sUser)
        sUser = VB.Left(sUser, Len(sUser) - 1) & " " & "@" & VB6.Format(Today, "mm/dd/yyyy") & " " & VB6.Format(TimeOfDay, "hh:mm:ss AMPM")

        '======================Add new form info to db's ================================
        'Master list Release db/table
        sQuery = "UPDATE NewMaster SET Source ='" & sUser & "', OriginalFormNameNumber='" & sOFN & "', Type='" & sFT & "', VirtualSide='"
        sQuery = sQuery & sVS & "', Duplex='" & sDuplex & "', Purge='" & sPurge & "'"
        sQuery = sQuery & " WHERE FormNameNumber ='" & sFN & "'"
        db.Execute(sQuery, 64) '64 = dbSQLPassThrough

        'Validate/Completion VMList db/table
        sQuery = "UPDATE VMList_NewMaster SET Source ='" & sUser & "', OriginalFormNameNumber='" & sOFN & "', Type='" & sFT & "', VirtualSide='"
        sQuery = sQuery & sVS & "', Duplex='" & sDuplex & "', Purge='" & sPurge & "'"
        sQuery = sQuery & " WHERE FormNameNumber ='" & sFN & "'"
        db.Execute(sQuery, 64) '64 = dbSQLPassThrough
        MsgBox("Record modified = success!")
        Call ClearFields()
    End Sub
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
    Private Sub frmModFN_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Me.txtFN.Focus()
    End Sub
	Private Sub frmModFN_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		KeyPreview = True
        rsML = db.OpenRecordset("NewMaster", DAO.RecordsetTypeEnum.dbOpenDynaset)
		rsML.MoveLast()
		rsML.MoveFirst()
        rsVM = db.OpenRecordset("VMList_NewMaster", DAO.RecordsetTypeEnum.dbOpenDynaset)
        rsVM.MoveLast()
        rsVM.MoveFirst()
		cmbFT.SelectedIndex = 1
		cmbVS.SelectedIndex = 4
		cmbPurge.SelectedIndex = 1
		cmbDuplex.SelectedIndex = 1
	End Sub
	Private Sub frmModFN_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		rsVM.Close()
		rsML.Close()
	End Sub
	Private Sub frmModFN_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		'MsgBox KeyCode
	End Sub
	Private Sub txtFN_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtFN.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		'Check the contents of the Form Number box before proceeding
		If Trim(txtFN.Text) = "" Then
			MsgBox("You must enter a Form Number before proceeding", MsgBoxStyle.Critical, "Missing required field")
			Cancel = True
			GoTo EventExitSub
		End If
		
		'No NOA Reports Allowed
		If UCase(Trim(txtFN.Text)) = "SF50" Or UCase(Trim(txtFN.Text)) = "SF50B" Or UCase(Trim(txtFN.Text)) = "SF52" Or UCase(Trim(txtFN.Text)) = "PS50" Or UCase(Trim(txtFN.Text)) = "SF50A" Then
			MsgBox(UCase(Trim(txtFN.Text)) & " is a NOA Form. Use Modify NOA Record button for this", MsgBoxStyle.Critical, "NOA Code Form")
			txtFN.Text = ""
			Cancel = True
			GoTo EventExitSub
		End If
		
		Dim sFNFind, sTest As String
		Dim bFNVMNotFound, bFNMLNotFound As Boolean
		Dim bSpecial As Boolean
		Dim i As Short
		Dim strNew As String
		
		'Test for special characters
		bSpecial = False
		sTest = Trim(txtFN.Text)
        strNew = ""
        For i = 1 To Len(sTest)
            Select Case UCase(Mid(sTest, i, 1))
                Case "A" To "Z", CStr(0) To CStr(9)
                    strNew = strNew & Mid(sTest, i, 1)
                Case Else
                    bSpecial = True
            End Select
        Next
		If bSpecial Then
			MsgBox("Special characters and spaces not permitted in Form Number" & vbCrLf & "They have been removed.", MsgBoxStyle.Critical, "Invalid characters detected!")
			txtFN.Text = strNew
		End If
		
		txtFN.Text = UCase(txtFN.Text)
		
		'Check the db's to see if it exists
		sFNFind = "FormNameNumber = '" & Trim(UCase(txtFN.Text)) & "'"
		
		'Check the Master List
		rsML.MoveFirst()
		rsML.FindFirst(sFNFind)
		bFNMLNotFound = rsML.NoMatch
		
		'Check the VMList
		rsVM.MoveFirst()
		rsVM.FindFirst(sFNFind)
		bFNVMNotFound = rsVM.NoMatch
		
		If bFNMLNotFound = True And bFNVMNotFound = True Then
			MsgBox("That Form Number you want to modify is not in system." & vbCrLf & "If you want to add a Form Number Record use the Add Form Number Record button")
			txtFN.Text = ""
			Cancel = True
		Else
			'Fill out the fields from the db
			If bFNMLNotFound = False Then 'Use masterlist as a preference
                txtOFN.Text = MNS(rsML.Fields("OriginalFormNameNumber").Value)
                cmbFT.Text = MNS(rsML.Fields("Type").Value)
                cmbVS.Text = MNS(rsML.Fields("VirtualSide").Value)
                If MNS(rsML.Fields("Duplex").Value) = "X" Then
                    cmbDuplex.Text = "Yes"
                Else
                    cmbDuplex.Text = "No"
                End If
                'If Not IsDBNull(rsML.Fields("Purge").Value) Then
                If MNS(rsML.Fields("Purge").Value) = "Yes" Then
                    cmbPurge.Text = "Yes"
                Else
                    cmbPurge.Text = "No"
                End If
            Else 'must be in VMList
                txtOFN.Text = MNS(rsVM.Fields("OriginalFormNameNumber").Value)
                cmbFT.Text = MNS(rsVM.Fields("Type").Value)
                cmbVS.Text = MNS(rsVM.Fields("VirtualSide").Value)
                If MNS(rsVM.Fields("Duplex").Value) = "X" Then
                    cmbDuplex.Text = "Yes"
                Else
                    cmbDuplex.Text = "No"
                End If
                If MNS(rsVM.Fields("Purge").Value) = "Yes" Then
                    cmbPurge.Text = "Yes"
                Else
                    cmbPurge.Text = "No"
                End If
        End If
            Cancel = False
		End If
EventExitSub: 
		eventArgs.Cancel = Cancel
	End Sub
	
	Private Sub txtOFN_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtOFN.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		'Check the contents of the Form Number box before proceeding
		If Trim(txtOFN.Text) = "" Then
			MsgBox("The OriginalFormNameNumber is missing, please fill out", MsgBoxStyle.Critical, "Missing required field")
			Cancel = True
			GoTo EventExitSub
		End If
		
		txtOFN.Text = UCase(txtOFN.Text)
		Cancel = False
		
EventExitSub: 
		eventArgs.Cancel = Cancel
	End Sub
	Private Sub ClearFields()
		txtFN.Text = ""
		txtOFN.Text = ""
		cmbFT.SelectedIndex = 1
		cmbVS.SelectedIndex = 4
		cmbPurge.SelectedIndex = 1
		cmbDuplex.SelectedIndex = 1
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