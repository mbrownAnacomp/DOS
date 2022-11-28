VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "#Ascent Capture - Release Setup Wait Dialog"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblWait 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   105
      TabIndex        =   0
      Top             =   360
      Width           =   4635
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*** module level error variables for holding returned error values
'*** in the events (should they occur)
Private m_lErrorVal As Long
Private m_strErrorDescription As String
Private m_strErrorSource As String

Private Sub Form_Activate()

    '*** After the wait for is displayed, we load our settings
    '*** and then hide the waiting form
    On Error GoTo FormActivate_Error
        
    Me.Refresh
        
    Call frmSetup.LoadFormSettings
        
    Me.Hide
    Unload Me
Exit Sub
    
FormActivate_Error:
    '*** if there is an error, save it in the variable to be raised again later
    m_lErrorVal = Err.Number
    m_strErrorDescription = Err.Description
    m_strErrorSource = Err.Source
    Me.Hide
    Unload Me
End Sub


'*************************************************************************
'*** Function:  LoadSettings
'*** Purpose:   Kicks off the waiting form by setting the mouse
'***            pointer and showing the form.  Raises errors which
'***            may occur and be returned in the events Form_Load and
'***            Form_Activate
'***
'*** Input:     none
'***
'*** Output:    Generates an error should one occur in any of the events of this form
'*************************************************************************
Public Sub LoadSettings(strCaption As String)
    
    m_lErrorVal = 0
    
    frmWait.Caption = strCaption
    lblWait.Caption = LoadResString(CAP_LOADING)
    Me.MousePointer = vbHourglass
    Load Me
    Me.Show vbModal
    Me.MousePointer = vbNormal
    
    If m_lErrorVal = 0 Then
        Exit Sub
    Else
        Err.Raise m_lErrorVal, m_strErrorSource, m_strErrorDescription
    End If
    
End Sub
