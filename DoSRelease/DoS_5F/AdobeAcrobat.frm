VERSION 5.00
Object = "{97962CD3-8A13-11D4-ACFE-005004186384}#1.1#0"; "CAPPdf4.ocx"
Begin VB.Form frmAdobeAcrobatSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#Adobe Acrobat Capture Settings"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdHelp 
      Caption         =   "#Help"
      Height          =   312
      Left            =   6600
      TabIndex        =   5
      Top             =   5960
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   312
      Left            =   5600
      TabIndex        =   4
      Top             =   5960
      Width           =   900
   End
   Begin VB.Frame frmAdvancedPDFSettings 
      Caption         =   "#Advanced PDF Settings"
      Height          =   1120
      Left            =   150
      TabIndex        =   1
      Top             =   4680
      Width           =   7370
      Begin VB.CheckBox chkDeleteOnHung 
         Caption         =   "#Delete Adobe Acrobat Capture document if a processing error occurs"
         Height          =   210
         Left            =   160
         TabIndex        =   3
         Top             =   720
         Width           =   6000
      End
      Begin VB.CheckBox chkWaitForStatus 
         Caption         =   "#Poll Adobe Acrobat Capture for status on submitted documents"
         Height          =   210
         Left            =   160
         TabIndex        =   2
         Top             =   360
         Width           =   6000
      End
   End
   Begin CAPPdfCtl4.CAPPdf3 pdfImageFormat3 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   7858
   End
End
Attribute VB_Name = "frmAdobeAcrobatSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================
' Form Level Variables
'=======================
Dim m_fDirty As Boolean
Dim m_bValidate As Boolean

'*************************************************
' Validate [Let Property]
'-------------------------------------------------
' Purpose:  The validate property indicates if
'           Adobe Acrobat is in validate mode
' Inputs:   NewStatus   Boolean indicating if
'                       in validate mode (TRUE)
'                       or not (FALSE)
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Property Let Validate(NewStatus As Boolean)
        m_bValidate = NewStatus
End Property

'*************************************************
' Dirty [Let Property]
'-------------------------------------------------
' Purpose:  The dirty property will set the
'           current status of the data.  If
'           the data is dirty, the Apply
'           button is enabled.
' Inputs:   NewStatus   Boolean indicating if
'                       data is dirty (TRUE)
'                       or clean (FALSE)
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Property Let Dirty(NewStatus As Boolean)
        m_fDirty = NewStatus
End Property

'*************************************************
' Dirty [Get Property]
'-------------------------------------------------
' Purpose:  The dirty property will return
'           the current status of the data.
' Inputs:   None
' Outputs:  TRUE if the data is dirty
'           FALSE if the data is clean
' Returns:  None
' Notes:    None
'*************************************************
Public Property Get Dirty() As Boolean
        Dirty = m_fDirty
End Property

'*************************************************
' chkWaitForStatus_Click
'-------------------------------------------------
' Purpose:  The user toggled whether or not to
'           wait for status on submitted
'           documents
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub chkWaitForStatus_Click()
    chkDeleteOnHung.Enabled = CBool(chkWaitForStatus.Value)
    Me.Dirty = True
End Sub

'*************************************************
' chkDeleteOnHung_Click
'-------------------------------------------------
' Purpose:  The user toggled whether or not to
'           delete Adobe Carobat Capture document
'           if any processing error occurs.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub chkDeleteOnHung_Click()
    Me.Dirty = True
End Sub

'*************************************************
' cmdHelp_Click
'-------------------------------------------------
' Purpose:  Display the help topic for the
'           Adobe Acrobat Capture Setup
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    This release script uses a proprietary
'           COM object for its help system.
'*************************************************
Private Sub cmdHelp_Click()
       Dim oKChmHlp As KChmHlp
       Dim bRetVal As Boolean
       Dim HelpFilePath As String

       Set oKChmHlp = New KChmHlp
       HelpFilePath = App.Path & "\" & App.HelpFile
       
       Call oKChmHlp.ShowHelp(ByVal HelpFilePath, CLng(ADOBE_ACROBAT_HELPID))
End Sub

'*************************************************
' Form_Activate
'-------------------------------------------------
' Purpose:  The dialog is active. If validate
'           flage is TRUE, validate the PDF
'           settings.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub Form_Activate()
    If m_bValidate = True Then
        pdfImageFormat3.SetValidateFocus
        Call MsgBox(pdfImageFormat3.ValidateError, _
                vbOKOnly + vbExclamation, _
                LoadResString(TITLE_DATAVERIFYFAIL))
    End If
End Sub
'*************************************************
' Form_Initialize
'-------------------------------------------------
' Purpose:  Initialize dialog properties
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub Form_Initialize()
    Me.Dirty = False
    Me.Validate = False
End Sub

'*************************************************
' pdfImageFormat3_Change
'-------------------------------------------------
' Purpose:  Some setting on the PDF3 control (tab)
'           has changed.  Mark the form dirty.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub pdfImageFormat3_Change()
    Me.Dirty = True
End Sub

'*************************************************
' cmdOK_Click
'-------------------------------------------------
' Purpose:  Ready to move on and dismiss the
'           Adobe Acrobat Capture Setup dialog
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub cmdOK_Click()
     Me.Hide
End Sub
