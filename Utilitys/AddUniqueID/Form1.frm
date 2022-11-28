VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Add UniqueID's"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAdded 
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
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAddRecords 
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
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add records"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblAdded 
      Caption         =   "Records added:"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter number of UniqueID records to add"
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

Private Sub cmdAdd_Click()
    Dim oDB1 As Database, x As Long
    Dim sDT As String, sOP As String
    Dim strConnect As String, lRecCount As Long
    Dim sQuery As String
    lblAdded.Visible = True
    txtAdded.Visible = True
    'Set SQL connection
    strConnect = "ODBC;DSN=DIA;DATABASE=DIA;UID=sa;PWD=f!$tomcat" 'SQL connection
    Set oDB1 = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    If IsNumeric(Trim(txtAddRecords.Text)) Then
        lRecCount = Val(Trim(txtAddRecords.Text))
    Else
        MsgBox "Enter a Numeric value for records to add"
        txtAddRecords.Text = ""
        Exit Sub
    End If
    For x = 1 To lRecCount
        sDT = Now()
        sOP = "Record added programatically"
        sQuery = "INSERT INTO DIA_UniqueID (DateTimeCreated, Operator) VALUES('" & sDT & "', '" & sOP & "')"
        oDB1.Execute sQuery ', 64
        txtAdded.Text = Str(x)
        txtAdded.Refresh
    Next
    oDB1.Close
    MsgBox Str(lRecCount) & " Records Added"
End Sub

