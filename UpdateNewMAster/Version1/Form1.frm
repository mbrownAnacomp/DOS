VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Upload New Forms List"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUploadVM 
      Caption         =   "Upload To VMList"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtMoved 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdSelectFile 
      Caption         =   "Browse"
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtUploadFile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   5895
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6720
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Upload to NewMaster"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblMoved 
      Caption         =   "Records moved:"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select Forms List Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUpload_Click()
    Dim oDB1 As Database, oDB2 As Database
    Dim oRS1 As Recordset, oRS2 As Recordset
    Dim sSource As String, sFormNameNumber As String, sOriginalFormNameNumber As String, sFormDescription As String
    Dim sType As String, sVirtualSide As String, sDuplex As String, sNOAC As String, sNOAEffFromDate As String
    Dim sNOAEffToDate As String, sPurge As String, sAlias As String
    Dim strConnect As String, lRecCount As Long
    Dim sActionQuery As String, sQuery As String
    lblMoved.Visible = True
    txtMoved.Visible = True
    'Set SQL connection
    strConnect = "ODBC;DSN=DoS;DATABASE=DoS;UID=sa;PWD=f!$tomcat" 'SQL connection
    'strConnect = "ODBC;DSN=DEVTracker;DATABASE=HRKofaxTracker;UID=hreesvcdevkofax;PWD=Wduw3bt?1234" 'SQL connection
    'strConnect = "ODBC;DSN=PRDTracker;DATABASE=HRKofaxTracker;UID=hreesvcprdkofax;PWD=P@ssword12345!" 'SQL connection
    Set oDB1 = DBEngine.Workspaces(0).OpenDatabase(Trim(txtUploadFile.Text))
    Set oDB2 = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Set oRS1 = oDB1.OpenRecordset("FormsList", dbOpenDynaset, dbSeeChanges)
    oRS1.MoveLast
    oRS1.MoveFirst
    Do Until oRS1.EOF
        If oRS1.Fields("NOACodes").Value = "0 NOA Codes" Then 'means it's not a NOAC form
            sSource = oRS1.Fields("FormOwner")
            sOriginalFormNameNumber = oRS1.Fields("FormName").Value
            sFormNameNumber = CreateShortName(sOriginalFormNameNumber)
            sFormDescription = CheckForTicks(oRS1.Fields("FormDescription").Value)
            sType = ParseType(oRS1.Fields("FormTypes").Value)
            sVirtualSide = oRS1.Fields("VirtualSide").Value
            sDuplex = oRS1.Fields("Duplex").Value
            sAlias = oRS1.Fields("FormAlias").Value
            sPurge = "No"
            sActionQuery = "INSERT INTO NewMaster (Source, FormNameNumber, OriginalFormNameNumber, [Form Description], Type, VirtualSide, Duplex, Purge, Alias) VALUES('" & sSource & "', '" & sFormNameNumber & "', '" & sOriginalFormNameNumber & "', '" & sFormDescription & "', '" & sType & "', '" & sVirtualSide & "', '" & sDuplex & "', '" & sPurge & "', '" & sAlias & "')"
            oDB2.Execute sActionQuery ', 64
            lRecCount = lRecCount + 1
            txtMoved.Text = Str(lRecCount)
            txtMoved.Refresh
        Else 'We have NOAC
            sOriginalFormNameNumber = oRS1.Fields("FormName").Value
            sQuery = "Select * From NOACodes Where FormName = '" & sOriginalFormNameNumber & "' Order By NOACode"
            Set oRS2 = oDB1.OpenRecordset(sQuery, dbOpenSnapshot, dbSQLPassThrough)
            If oRS2.RecordCount > 0 Then
                sSource = oRS1.Fields("FormOwner")
                sFormNameNumber = CreateShortName(sOriginalFormNameNumber)
                sFormDescription = CheckForTicks(oRS1.Fields("FormDescription").Value)
                sType = "NOA"
                'sVirtualSide = oRS1.Fields("VirtualSide").Value
                sDuplex = oRS1.Fields("Duplex").Value
                sAlias = oRS1.Fields("FormAlias").Value
                sPurge = "No"
                With oRS2
                    .MoveFirst
                    Do Until .EOF
                        sNOAC = .Fields("NOACode")
                        sNOAEffFromDate = .Fields("EffectiveDateBegin").Value
                        If sNOAEffFromDate <> "NULL" Then sNOAEffFromDate = Left(sNOAEffFromDate, 10)
                        sNOAEffToDate = .Fields("EffectiveDateEnd").Value
                        If sNOAEffToDate = "NULL" Then
                            sNOAEffToDate = "2099-12-31"
                        Else
                            sNOAEffToDate = Left(sNOAEffToDate, 10)
                        End If
                        sVirtualSide = .Fields("VirtualSide").Value
                        sActionQuery = "INSERT INTO NewMaster (Source, FormNameNumber, OriginalFormNameNumber, [Form Description], Type, VirtualSide, Duplex, NOAC, [NOA Eff From Date], [NOA Eff To Date], Purge, Alias) VALUES('" & sSource & "', '" & sFormNameNumber & "', '" & sOriginalFormNameNumber & "', '" & sFormDescription & "', '" & sType & "', '" & sVirtualSide & "', '" & sDuplex & "', '" & sNOAC & "', '" & sNOAEffFromDate & "', '" & sNOAEffToDate & "', '" & sPurge & "', '" & sAlias & "')"
                        oDB2.Execute sActionQuery ', 64
                        .MoveNext
                        lRecCount = lRecCount + 1
                        txtMoved.Text = Str(lRecCount)
                    Loop
                End With
            End If
        End If
        oRS1.MoveNext
    Loop
    oRS1.Close
    oRS2.Close
    oDB1.Close
    oDB2.Close
    MsgBox Str(lRecCount) & " Records Moved"
End Sub
Function CreateShortName(sLongName As String) As String
    Dim strNew As String, i As Integer
    For i = 1 To Len(sLongName)
        Select Case UCase(Mid(sLongName, i, 1))
            Case "A" To "Z", 0 To 9
                strNew = strNew & Mid(sLongName, i, 1)
        End Select
    Next
    CreateShortName = strNew
End Function
Function ParseType(sType As String) As String
    Dim iComma As Integer
    iComma = InStr(1, sType, ",")
    If iComma > 0 Then
        If Mid(sType, 1, 9) = "EXCEPTION" Then
            ParseType = Mid(sType, iComma + 1)
        Else
            ParseType = Mid(sType, 1, iComma - 1)
        End If
    Else
        ParseType = ""
    End If
End Function
Function CheckForTicks(sInput As String) As String
    Dim iTick As Integer
    iTick = InStr(1, sInput, "'")
    If iTick > 0 Then
        CheckForTicks = Left(sInput, iTick) & "'" & Mid(sInput, iTick + 1)
    Else
        CheckForTicks = sInput
    End If
End Function

Private Sub cmdSelectFile_Click()
    Dim Path As String
    Dim nTextLen As Integer
    Dim LastPos As Integer
    Dim TmpPos As Integer
    
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo Error_Dir
    ' Initialize the dialog to the file and path specified in the text box.
    ' If nothing is specified and no previous default directory has been
    ' set, default to C:\
    nTextLen = Len(txtUploadFile.Text)
    If nTextLen <> 0 Then
         CommonDialog1.InitDir = txtUploadFile.Text
    Else
        If CommonDialog1.InitDir = "" Then
            CommonDialog1.InitDir = "C:\"
        End If
    End If
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNExplorer + cdlOFNNoValidate + cdlOFNPathMustExist
    ' Set filters
    CommonDialog1.Filter = "Upload File (*.mdb)|*.mdb"
    ' Specify default filter
    CommonDialog1.FilterIndex = 1
    CommonDialog1.DialogTitle = "Select Upload DB to Process"
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file
    txtUploadFile.Text = CommonDialog1.FileName
     
    Exit Sub
  
Error_Dir:
  'User pressed the Cancel button
    Select Case Err.Number
        Case 32755
            ' Always reset the dialog for next time
            CommonDialog1.FileName = ""
            CommonDialog1.InitDir = ""
            Exit Sub
        Case Else
            MsgBox "Error Number : " & Err.Number & vbCrLf & Err.Description
    End Select
    
    ' Always reset the dialog for next time
    CommonDialog1.FileName = ""
    CommonDialog1.InitDir = ""
  
    Exit Sub
End Sub

Private Sub cmdUploadVM_Click()
    Dim oDB1 As Database, oDB2 As Database
    Dim oRS1 As Recordset, oRS2 As Recordset
    Dim sSource As String, sFormNameNumber As String, sOriginalFormNameNumber As String, sFormDescription As String
    Dim sType As String, sVirtualSide As String, sDuplex As String, sNOAC As String, sNOAEffFromDate As String
    Dim sNOAEffToDate As String, sPurge As String, sAlias As String
    Dim strConnect As String, lRecCount As Long
    Dim sActionQuery As String, sQuery As String
    lblMoved.Visible = True
    txtMoved.Visible = True
    'Set SQL connection
    strConnect = "ODBC;DSN=DoS;DATABASE=DoS;UID=sa;PWD=f!$tomcat" 'SQL connection
    'strConnect = "ODBC;DSN=DEVTracker;DATABASE=HRKofaxTracker;UID=hreesvcdevkofax;PWD=Wduw3bt?1234" 'SQL connection
    'strConnect = "ODBC;DSN=PRDTracker;DATABASE=HRKofaxTracker;UID=hreesvcprdkofax;PWD=P@ssword12345!" 'SQL connection
    Set oDB1 = DBEngine.Workspaces(0).OpenDatabase(Trim(txtUploadFile.Text))
    Set oDB2 = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Set oRS1 = oDB1.OpenRecordset("FormsList", dbOpenDynaset, dbSeeChanges)
    oRS1.MoveLast
    oRS1.MoveFirst
    Do Until oRS1.EOF
        If oRS1.Fields("NOACodes").Value = "0 NOA Codes" Then 'means it's not a NOAC form
            sSource = oRS1.Fields("FormOwner")
            sOriginalFormNameNumber = oRS1.Fields("FormName").Value
            If sOriginalFormNameNumber <> "OTHER" Then
                sFormNameNumber = CreateShortName(sOriginalFormNameNumber)
                sFormDescription = CheckForTicks(oRS1.Fields("FormDescription").Value)
                sType = ParseType(oRS1.Fields("FormTypes").Value)
                sVirtualSide = oRS1.Fields("VirtualSide").Value
                sDuplex = oRS1.Fields("Duplex").Value
                sAlias = oRS1.Fields("FormAlias").Value
                sPurge = "No"
                sActionQuery = "INSERT INTO VMList_NewMaster (Source, FormNameNumber, OriginalFormNameNumber, [Form Description], Type, VirtualSide, Duplex, Purge, Alias) VALUES('" & sSource & "', '" & sFormNameNumber & "', '" & sOriginalFormNameNumber & "', '" & sFormDescription & "', '" & sType & "', '" & sVirtualSide & "', '" & sDuplex & "', '" & sPurge & "', '" & sAlias & "')"
                oDB2.Execute sActionQuery ', 64
            Else
                sActionQuery = "INSERT INTO VMList_NewMaster (Source, FormNameNumber, OriginalFormNameNumber, [Form Description], Type, VirtualSide, Duplex, Purge, Alias) VALUES('ANACOMP', 'OTHERBENEFITS', 'BENEFITS', 'Permanent', 'OTHER', 'OTHER', 'Duplex', 'No', 'OTHER')"
                oDB2.Execute sActionQuery ', 64
                sActionQuery = "INSERT INTO VMList_NewMaster (Source, FormNameNumber, OriginalFormNameNumber, [Form Description], Type, VirtualSide, Duplex, Purge, Alias) VALUES('ANACOMP', 'OTHEREMPLOYEE', 'EMPLOYEE', 'Permanent', 'OTHER', 'OTHER', 'Duplex', 'No', 'OTHER')"
                oDB2.Execute sActionQuery ', 64
                sActionQuery = "INSERT INTO VMList_NewMaster (Source, FormNameNumber, OriginalFormNameNumber, [Form Description], Type, VirtualSide, Duplex, Purge, Alias) VALUES('ANACOMP', 'OTHERINVESTIGATIONSSECURITYCLEARANCE', 'INVESTIGATION', 'Permanent', 'OTHER', 'OTHER', 'Duplex', 'No', 'OTHER')"
                oDB2.Execute sActionQuery ', 64
                sActionQuery = "INSERT INTO VMList_NewMaster (Source, FormNameNumber, OriginalFormNameNumber, [Form Description], Type, VirtualSide, Duplex, Purge, Alias) VALUES('ANACOMP', 'OTHERPAYROLL', 'PAYROLL', 'Permanent', 'OTHER', 'OTHER', 'Duplex', 'No', 'OTHER')"
                oDB2.Execute sActionQuery ', 64
                sActionQuery = "INSERT INTO VMList_NewMaster (Source, FormNameNumber, OriginalFormNameNumber, [Form Description], Type, VirtualSide, Duplex, Purge, Alias) VALUES('ANACOMP', 'OTHERPERFORMANCEAPPRAISAL', 'PERFORMANCE APPRAISAL', 'Permanent', 'OTHER', 'OTHER', 'Duplex', 'No', 'OTHER')"
                oDB2.Execute sActionQuery ', 64
                sActionQuery = "INSERT INTO VMList_NewMaster (Source, FormNameNumber, OriginalFormNameNumber, [Form Description], Type, VirtualSide, Duplex, Purge, Alias) VALUES('ANACOMP', 'OTHERPERSONNELACTIONSUPPORTDOC', 'PERSONNEL ACTION/SUPPORT DOC', 'Permanent', 'OTHER', 'OTHER', 'Duplex', 'No', 'OTHER')"
                oDB2.Execute sActionQuery ', 64
                sActionQuery = "INSERT INTO VMList_NewMaster (Source, FormNameNumber, OriginalFormNameNumber, [Form Description], Type, VirtualSide, Duplex, Purge, Alias) VALUES('ANACOMP', 'OTHERPOSITION', 'POSITION', 'Permanent', 'OTHER', 'OTHER', 'Duplex', 'No', 'OTHER')"
                oDB2.Execute sActionQuery ', 64
                sActionQuery = "INSERT INTO VMList_NewMaster (Source, FormNameNumber, OriginalFormNameNumber, [Form Description], Type, VirtualSide, Duplex, Purge, Alias) VALUES('ANACOMP', 'OTHERTRAINING', 'TRAINING', 'Permanent', 'OTHER', 'OTHER', 'Duplex', 'No', 'OTHER')"
                oDB2.Execute sActionQuery ', 64
            End If
            lRecCount = lRecCount + 1
            txtMoved.Text = Str(lRecCount)
            txtMoved.Refresh
        Else 'We have NOAC
            sOriginalFormNameNumber = oRS1.Fields("FormName").Value
            sQuery = "Select * From NOACodes Where FormName = '" & sOriginalFormNameNumber & "' Order By NOACode"
            Set oRS2 = oDB1.OpenRecordset(sQuery, dbOpenSnapshot, dbSQLPassThrough)
            If oRS2.RecordCount > 0 Then
                sSource = oRS1.Fields("FormOwner")
                sFormNameNumber = CreateShortName(sOriginalFormNameNumber)
                sFormDescription = CheckForTicks(oRS1.Fields("FormDescription").Value)
                sType = "NOA"
                'sVirtualSide = oRS1.Fields("VirtualSide").Value
                sDuplex = oRS1.Fields("Duplex").Value
                sAlias = oRS1.Fields("FormAlias").Value
                sPurge = "No"
                With oRS2
                    .MoveFirst
                    Do Until .EOF
                        sNOAC = .Fields("NOACode")
                        sNOAEffFromDate = .Fields("EffectiveDateBegin").Value
                        If sNOAEffFromDate <> "NULL" Then sNOAEffFromDate = Left(sNOAEffFromDate, 10)
                        sNOAEffToDate = .Fields("EffectiveDateEnd").Value
                        If sNOAEffToDate = "NULL" Then
                            sNOAEffToDate = "2099-12-31"
                        Else
                            sNOAEffToDate = Left(sNOAEffToDate, 10)
                        End If
                        sVirtualSide = .Fields("VirtualSide").Value
                        sActionQuery = "INSERT INTO VMList_NewMaster (Source, FormNameNumber, OriginalFormNameNumber, [Form Description], Type, VirtualSide, Duplex, NOAC, [NOA Eff From Date], [NOA Eff To Date], Purge, Alias) VALUES('" & sSource & "', '" & sFormNameNumber & "', '" & sOriginalFormNameNumber & "', '" & sFormDescription & "', '" & sType & "', '" & sVirtualSide & "', '" & sDuplex & "', '" & sNOAC & "', '" & sNOAEffFromDate & "', '" & sNOAEffToDate & "', '" & sPurge & "', '" & sAlias & "')"
                        oDB2.Execute sActionQuery ', 64
                        .MoveNext
                        lRecCount = lRecCount + 1
                        txtMoved.Text = Str(lRecCount)
                    Loop
                End With
            End If
        End If
        oRS1.MoveNext
    Loop
    oRS1.Close
    oRS2.Close
    oDB1.Close
    oDB2.Close
    MsgBox Str(lRecCount) & " Records Moved"

End Sub
