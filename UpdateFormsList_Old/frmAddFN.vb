Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Data.SqlClient
Friend Class frmAddFN
    Inherits System.Windows.Forms.Form
    Private Sub cmbFT_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles cmbFT.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        cmbFT.Text = UCase(cmbFT.Text)
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
        Dim sDuplex, sVS, sOFN, sFN, sFT, sPurge, sCurrentUser As String
        Dim sDRPOutFile, sUser, sQuery As String
        Dim iDRPOut As Short
        Dim rsForms As DAO.Recordset

        sFN = Trim(txtFN.Text)
        sOFN = Trim(txtOFN.Text)
        'Check for blanks
        If sOFN = "" Then
            MsgBox("Original Form Name Number field cannot be blank")
            txtOFN.Focus()
            Exit Sub
        End If
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

        '======================Add new form info to db and drp================================
        'Master list Release db/table
        sQuery = "INSERT INTO NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, Duplex, Purge) "
        sQuery = sQuery & "VALUES('" & sUser & "', '" & sFN & "', '" & sOFN & "', '" & sFT & "', '" & sVS & "', '" & sDuplex & "', '" & sPurge & "')"
        db.Execute(sQuery)

        'Validate/Completion VMList db/table
        sQuery = "INSERT INTO VMList_NewMaster(Source, FormNameNumber, OriginalFormNameNumber, Type, VirtualSide, Duplex, Purge) "
        sQuery = sQuery & "VALUES('" & sUser & "', '" & sFN & "', '" & sOFN & "', '" & sFT & "', '" & sVS & "', '" & sDuplex & "', '" & sPurge & "')"
        db.Execute(sQuery)

        'For KTM  - update Dictionary file from DB Table - since it's for Validation use VMList for "OTHERxxxx"
        sQuery = "Select Distinct FormNameNumber From VMList_NewMaster Where FormNameNumber Is Not NULL Order By FormNameNumber"
        rsForms = db.OpenRecordset(sQuery, DAO.RecordsetTypeEnum.dbOpenDynaset)
        rsForms.MoveFirst()
        iDRPOut = FreeFile()
        sDRPOutFile = sDRPDir & "Tmp"
        FileOpen(iDRPOut, sDRPOutFile, OpenMode.Output)
        'Print FormName records from DB Table
        Do Until rsForms.EOF
            PrintLine(iDRPOut, rsForms.Fields("FormNameNumber").Value)
            rsForms.MoveNext()
        Loop
        rsForms.Close()
        FileClose()
        fs.MoveFile(sDRPDir, VB.Left(sDRPDir, Len(sDRPDir) - 4) & "_" & VB6.Format(Str(Today.ToOADate), "yymmdd") & "_" & VB6.Format(Str(TimeOfDay.ToOADate), "hhmmss") & ".txt")
        'Rename new Dictionary file to original Dictionary file name
        fs.MoveFile(sDRPOutFile, sDRPDir)
        MsgBox("New record added = success!")
        Call ClearFields()
    End Sub

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub
    Private Sub frmAddFN_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Me.txtFN.Focus()
    End Sub
    Private Sub frmAddFN_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
    Private Sub frmAddFN_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        rsVM.Close()
        rsML.Close()
    End Sub
    Private Sub frmAddFN_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

        Dim sTest, sFNFind, strNew As String
        Dim bFNNotFound, bSpecial As Boolean
        Dim i As Short

        'Check the db's to see if it already exists
        sFNFind = "FormNameNumber = '" & Trim(UCase(txtFN.Text)) & "'"

        rsML.MoveFirst()
        rsML.FindFirst(sFNFind)
        bFNNotFound = rsML.NoMatch

        If bFNNotFound = True Then 'If the 1st db is OK then check the other
            rsVM.MoveFirst()
            rsVM.FindFirst(sFNFind)
            bFNNotFound = rsVM.NoMatch
        End If

        If bFNNotFound = False Then
            MsgBox("That Form Number is already in the system." & vbCrLf & "If you want to modify a Form Number Record use the Modify Form Number Record button")
            txtFN.Text = ""
            Cancel = True
        Else
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
                MsgBox("Special characters and spaces not permitted in Form Numer" & vbCrLf & "They have been removed.", MsgBoxStyle.Critical, "Invalid characters detected!")
                txtFN.Text = strNew
            End If
            txtFN.Text = UCase(txtFN.Text)
            'Give Original FNN a head start
            txtOFN.Text = txtFN.Text
            Cancel = False
        End If

EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtOFN_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtOFN.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        'Check the contents of the Form Number box before proceeding
        If Trim(txtOFN.Text) = "" Then
            MsgBox("You must enter an Original Form Number before proceeding", MsgBoxStyle.Critical, "Missing required field")
            Cancel = True
            GoTo EventExitSub
        End If

        Dim sFNFind As String
        Dim bFNNotFound As Boolean

        'Check the db's to see if it already exists
        sFNFind = "OriginalFormNameNumber = '" & Trim(UCase(txtOFN.Text)) & "'"

        rsML.MoveFirst()
        rsML.FindFirst(sFNFind)
        bFNNotFound = rsML.NoMatch

        If bFNNotFound = True Then 'If the 1st db is OK then check the other
            rsVM.MoveFirst()
            rsVM.FindFirst(sFNFind)
            bFNNotFound = rsVM.NoMatch
        End If

        If bFNNotFound = False Then
            MsgBox("That Original Form Number is already in the system." & vbCrLf & "If you want to modify a Form Number Record use the Modify Form Number Record button")
            txtOFN.Text = ""
            Cancel = True
        Else
            txtOFN.Text = UCase(txtOFN.Text)
            Cancel = False
        End If

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
End Class