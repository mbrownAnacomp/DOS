VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Move UniqueID's"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
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
      Caption         =   "Upload"
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
    Dim oRS1 As Recordset, sDT As String, sOP As String
    Dim strConnect As String, lRecCount As Long
    Dim sQuery As String
    lblMoved.Visible = True
    txtMoved.Visible = True
    'Set SQL connection
    strConnect = "ODBC;DSN=DEVTRACKER;DATABASE=DIA;UID=sa;PWD=f!$tomcat" 'SQL connection
    Set oDB1 = DBEngine.Workspaces(0).OpenDatabase(Trim(txtUploadFile.Text))
    Set oDB2 = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Set oRS1 = oDB1.OpenRecordset("DIA_UniqueID", dbOpenDynaset, dbSeeChanges)
    oRS1.MoveLast
    oRS1.MoveFirst
    Do Until oRS1.EOF
        sDT = oRS1.Fields("DateTimeCreated")
        sOP = oRS1.Fields("Operator")
        sOP = Left(sOP, Len(sOP) - 1)
        sQuery = "INSERT INTO DIA_UniqueID (DateTimeCreated, Operator) VALUES('" & sDT & "', '" & sOP & "')"
        oDB2.Execute sQuery ', 64
        lRecCount = lRecCount + 1
        txtMoved.Text = Str(lRecCount)
        txtMoved.Refresh
        oRS1.MoveNext
    Loop
    oRS1.Close
    oDB1.Close
    MsgBox Str(lRecCount) & " Records Moved"
End Sub

Function CreateShortName(sLongName As String) As String
    Dim strNew As String
    For i = 1 To Len(sLongName)
        Select Case UCase(Mid(sLongName, i, 1))
            Case "A" To "Z", 0 To 9
                strNew = strNew & Mid(sLongName, i, 1)
        End Select
    Next
    CreateShortName = strNew
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
