VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Menus"
   ClientHeight    =   1704
   ClientLeft      =   132
   ClientTop       =   588
   ClientWidth     =   2412
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1704
   ScaleWidth      =   2412
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Label Label1 
      Caption         =   $"Menu.frx":030A
      Height          =   1308
      Left            =   108
      TabIndex        =   0
      Top             =   144
      Width           =   2136
   End
   Begin VB.Menu mnuLinks 
      Caption         =   "Link Menu"
      Begin VB.Menu mnuCurrLink 
         Caption         =   "#Current Link"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "#Delete"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDocumentID 
         Caption         =   "#Document ID"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuIndexFields 
         Caption         =   "#Index Fields"
         Begin VB.Menu mnuIFields 
            Caption         =   "Index Field Name"
            Index           =   0
         End
      End
      Begin VB.Menu mnuBatchFields 
         Caption         =   "#Batch Fields"
         Begin VB.Menu mnuBFields 
            Caption         =   "Batch Field Name"
            Index           =   0
         End
      End
      Begin VB.Menu mnuAscentValues 
         Caption         =   "#Ascent Capture Values"
         Begin VB.Menu mnuACList 
            Caption         =   "Batch Description"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTextConst 
         Caption         =   "#Text Constant"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This string is written to the error log to identify
' in which source module the error occurred
Private Const M_MENUFORM = "Text Release Setup Menu"

Private LocForm As frmSetup

'*************************************************
' MyForm [Property Set]
'-------------------------------------------------
' Purpose:  Sets a reference to the form that
'           uses the popup menu for linking
'           index values.
' Inputs:   oForm   Form that will use the menu
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Property Set MyForm(oForm As Form)
    Set LocForm = oForm
End Property

'*************************************************
' BuildAscentMenu
'-------------------------------------------------
' Purpose:  This routine will add the list of
'           Ascent Capture Values into the
'           popup menu.
' Inputs:   oSetupData  ReleaseSetupData object
'                       passed to the script by
'                       Admin
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Sub BuildAscentMenu(oSetupData As ReleaseSetupData)
    Dim oBatchVar As Variant
    Dim i As Integer
    
    On Error GoTo BAM_LogAndPropError
        
        ' Add the Ascent Capture Values to the menu
130     If oSetupData.BatchVariableNames.Count > 0 Then
            i = 0
140         For Each oBatchVar In oSetupData.BatchVariableNames
                If (i <> 0) Then
150                 Load mnuACList(i)
                End If
160             mnuACList(i).Visible = True
170             mnuACList(i).Caption = oBatchVar
                i = i + 1
            Next
        Else
            mnuAscentValues.Visible = False
        End If
            
        Exit Sub
        
BAM_LogAndPropError:

        Call oError.LogTheError(Err, Err.Description, M_MENUFORM, Erl, True, False)
        
End Sub

'*************************************************
' BuildBatchMenu
'-------------------------------------------------
' Purpose:  This routine will add the list of
'           batch fields into the popup menu.
' Inputs:   oSetupData  ReleaseSetupData object
'                       passed to the script by
'                       Admin
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Sub BuildBatchMenu(oSetupData As ReleaseSetupData)
    Dim CurrField As AscentRelease.BatchField
    Dim i As Integer
    
    On Error GoTo BBM_LogAndPropError
        
        ' Add the Batch Fields to the menu
        If oSetupData.BatchFields.Count > 0 Then
            i = 0
            For Each CurrField In oSetupData.BatchFields
                If (i <> 0) Then
200                 Load mnuBFields(i)
                End If
210             mnuBFields(i).Visible = True
220             mnuBFields(i).Caption = CurrField.Name
                i = i + 1
            Next
        Else
            ' There are no Batch Fields so hide the submenu
            mnuBatchFields.Visible = False
        End If
    
        Exit Sub
        
BBM_LogAndPropError:

        Call oError.LogTheError(Err, Err.Description, M_MENUFORM, Erl, True, False)
        
End Sub

'*************************************************
' BuildIndexMenu
'-------------------------------------------------
' Purpose:  This routine will add the list of
'           Index Fields into the popup menu.
' Inputs:   oSetupData  ReleaseSetupData object
'                       passed to the script by
'                       Admin
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Sub BuildIndexMenu(oSetupData As ReleaseSetupData)
    Dim CurrField As AscentRelease.IndexField
    Dim i As Integer
        
    On Error GoTo BIM_LogAndPropError
        
        ' Add the Index Fields to the menu
        If oSetupData.IndexFields.Count > 0 Then
            i = 0
            For Each CurrField In oSetupData.IndexFields
                If (i <> 0) Then
250                Load mnuIFields(i)
                End If
260             mnuIFields(i).Visible = True
270             mnuIFields(i).Caption = CurrField.Name
                i = i + 1
            Next
        Else
            ' There are no index fields so hide the submenu
            mnuIndexFields.Visible = False
        End If
    
        Exit Sub

BIM_LogAndPropError:
    
        Call oError.LogTheError(Err, Err.Description, M_MENUFORM, Erl, True, False)
        
End Sub

'*************************************************
' Form_Load
'-------------------------------------------------
' Purpose:  This subroutine is called whenever
'           the form is loaded into memory.  It
'           loads the menu captions from the
'           resource file.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub Form_Load()

        On Error Resume Next
        
        ' The captions for the popup menu must be initialized
        ' from the resource strings when the form is first loaded
500     mnuCurrLink.Caption = LoadResString(MNU_CURRLINK)
510     mnuDelete.Caption = LoadResString(MNU_DELETE)
520     mnuDocumentID.Caption = LoadResString(MNU_DOCUMENTID)
530     mnuIndexFields.Caption = LoadResString(MNU_INDEXFIELDS)
540     mnuBatchFields.Caption = LoadResString(MNU_BATCHFIELDS)
550     mnuAscentValues.Caption = LoadResString(MNU_ASCENTVALUES)
560     mnuTextConst.Caption = LoadResString(MNU_TEXTCONST)

End Sub

'*************************************************
' mnuACList_Click
'-------------------------------------------------
' Purpose:  This event handler sets the globals
'           used to complete a data link.  The
'           caption of the menu item holds the
'           Ascent Capture Value name.
' Inputs:   Index   the selected menu item
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub mnuACList_Click(Index As Integer)
    LocForm.Dirty = True
    LocForm.gNewIndexType = KFX_REL_VARIABLE
    LocForm.gNewIndexData = mnuACList(Index).Caption
End Sub

'*************************************************
' mnuBFields_Click
'-------------------------------------------------
' Purpose:  This event handler sets the globals
'           used to complete a data link.  The
'           caption of the menu item holds the
'           Batch Field name.
' Inputs:   Index   the selected menu item
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub mnuBFields_Click(Index As Integer)
    LocForm.Dirty = True
    LocForm.gNewIndexType = KFX_REL_BATCHFIELD
    LocForm.gNewIndexData = mnuBFields(Index).Caption
End Sub

'*************************************************
' mnuDelete_Click
'-------------------------------------------------
' Purpose:  This event handler deletes the link
'           that is currently selected by faking
'           a click of the Delete button.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub mnuDelete_Click()
    LocForm.Dirty = True
    LocForm.gNewIndexType = DELETE_LINK
    LocForm.gNewIndexData = ""
    Call LocForm.cmdDeleteIndex_Click
End Sub

'*************************************************
' mnuDocumentID_Click
'-------------------------------------------------
' Purpose:  This event handler sets the globals
'           used to complete a data link.  This
'           Document ID menu is disabled in the
'           Text release script because it does
'           not require a Document ID link like
'           the Database release script.  The
'           Document ID is still available from
'           the Ascent Capture Values submenu.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub mnuDocumentID_Click()
    LocForm.Dirty = True
    LocForm.gNewIndexType = KFX_REL_DOCUMENTID
    LocForm.gNewIndexData = StripAmpersands(mnuDocumentID.Caption)
End Sub

'*************************************************
' mnuIFields_Click
'-------------------------------------------------
' Purpose:  This event handler sets the globals
'           used to complete a data link.  The
'           caption of the menu item holds the
'           Index Field name.
' Inputs:   Index   the selected menu item
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub mnuIFields_Click(Index As Integer)
    LocForm.Dirty = True
    LocForm.gNewIndexType = KFX_REL_INDEXFIELD
    LocForm.gNewIndexData = mnuIFields(Index).Caption
End Sub

'*************************************************
' mnuTextConst_Click
'-------------------------------------------------
' Purpose:  This event handler sets the globals
'           used to complete a data link.  This
'           selection will cause the textbox to
'           become unlocked and allow the user
'           to enter a Text Constant.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub mnuTextConst_Click()
    LocForm.Dirty = True
    LocForm.gNewIndexType = KFX_REL_TEXTCONSTANT
    LocForm.gNewIndexData = ""
End Sub
