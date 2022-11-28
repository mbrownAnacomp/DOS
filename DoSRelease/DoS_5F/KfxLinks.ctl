VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.UserControl KfxLinks 
   ClientHeight    =   2796
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7296
   LockControls    =   -1  'True
   ScaleHeight     =   2796
   ScaleWidth      =   7296
   Begin VB.CommandButton cmdMenu 
      Height          =   288
      Index           =   8
      Left            =   5400
      Picture         =   "KfxLinks.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton cmdMenu 
      Height          =   288
      Index           =   7
      Left            =   5400
      Picture         =   "KfxLinks.ctx":030A
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton cmdMenu 
      Height          =   288
      Index           =   6
      Left            =   5400
      Picture         =   "KfxLinks.ctx":0614
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton cmdMenu 
      Height          =   288
      Index           =   5
      Left            =   5400
      Picture         =   "KfxLinks.ctx":091E
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1476
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton cmdMenu 
      Height          =   288
      Index           =   4
      Left            =   5400
      Picture         =   "KfxLinks.ctx":0C28
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton cmdMenu 
      Height          =   288
      Index           =   3
      Left            =   5400
      Picture         =   "KfxLinks.ctx":0F32
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   860
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton cmdMenu 
      Height          =   288
      Index           =   2
      Left            =   5400
      Picture         =   "KfxLinks.ctx":123C
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton cmdMenu 
      Height          =   288
      Index           =   1
      Left            =   5400
      Picture         =   "KfxLinks.ctx":1546
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   300
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton cmdMenu 
      Height          =   288
      Index           =   0
      Left            =   5400
      Picture         =   "KfxLinks.ctx":1850
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton cmdDeleteAllIndex 
      Caption         =   "#Delete All"
      Height          =   312
      Left            =   6120
      TabIndex        =   30
      Top             =   840
      Width           =   900
   End
   Begin VB.CommandButton cmdAddIndex 
      Caption         =   "#Add"
      Height          =   312
      Left            =   6120
      TabIndex        =   28
      Top             =   0
      Width           =   900
   End
   Begin VB.CommandButton cmdDeleteIndex 
      Caption         =   "#Delete"
      Enabled         =   0   'False
      Height          =   312
      Left            =   6120
      TabIndex        =   29
      Top             =   420
      Width           =   900
   End
   Begin VB.TextBox txtIndexData 
      Height          =   288
      Index           =   4
      Left            =   1680
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   3990
   End
   Begin VB.TextBox txtIndexData 
      Height          =   288
      Index           =   3
      Left            =   1680
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   900
      Visible         =   0   'False
      Width           =   3990
   End
   Begin VB.TextBox txtIndexData 
      Height          =   288
      Index           =   2
      Left            =   1680
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   3990
   End
   Begin VB.TextBox txtIndexData 
      Height          =   288
      Index           =   1
      Left            =   1680
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   300
      Visible         =   0   'False
      Width           =   3990
   End
   Begin VB.TextBox txtIndexData 
      Height          =   288
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   3990
   End
   Begin VB.TextBox txtIndexData 
      Height          =   288
      Index           =   5
      Left            =   1692
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1500
      Visible         =   0   'False
      Width           =   3990
   End
   Begin VB.TextBox txtIndexData 
      Height          =   288
      Index           =   6
      Left            =   1680
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   3990
   End
   Begin VB.TextBox txtIndexData 
      Height          =   288
      Index           =   7
      Left            =   1680
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   3990
   End
   Begin VB.TextBox txtIndexData 
      Height          =   288
      Index           =   8
      Left            =   1680
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   3990
   End
   Begin VB.TextBox txtSequence 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   288
      Index           =   8
      Left            =   0
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "9"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox txtSequence 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   288
      Index           =   7
      Left            =   0
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "8"
      Top             =   2100
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox txtSequence 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   288
      Index           =   6
      Left            =   0
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "7"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox txtSequence 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   288
      Index           =   5
      Left            =   0
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "6"
      Top             =   1500
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox txtSequence 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   288
      Index           =   4
      Left            =   0
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "5"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.VScrollBar vsbIndex 
      Height          =   2688
      LargeChange     =   8
      Left            =   5652
      TabIndex        =   27
      Top             =   0
      Width           =   204
   End
   Begin VB.TextBox txtSequence 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   288
      Index           =   3
      Left            =   0
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "4"
      Top             =   900
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox txtSequence 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   288
      Index           =   2
      Left            =   0
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "3"
      Top             =   600
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox txtSequence 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   288
      Index           =   1
      Left            =   0
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "2"
      Top             =   300
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox txtSequence 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   288
      Index           =   0
      Left            =   0
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1620
   End
   Begin ComCtl2.UpDown updnIndex 
      Height          =   600
      Left            =   6120
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1260
      Width           =   192
      _ExtentX        =   275
      _ExtentY        =   1058
      _Version        =   327681
      Alignment       =   0
      OrigLeft        =   6330
      OrigTop         =   1575
      OrigRight       =   6525
      OrigBottom      =   2175
      Max             =   0
      Min             =   8
      Enabled         =   0   'False
   End
   Begin VB.Label lblMove 
      Caption         =   "#Move"
      Enabled         =   0   'False
      Height          =   240
      Left            =   6420
      TabIndex        =   32
      Top             =   1428
      Width           =   600
   End
   Begin VB.Menu mnuLinks 
      Caption         =   ""
      Enabled         =   0   'False
      Begin VB.Menu mnuCurrLink 
         Caption         =   "CurrentLink"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuUnlink 
         Caption         =   "&Unlink Field"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDocumentID 
         Caption         =   "&Document ID"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuIndexFields 
         Caption         =   "&Index Fields"
         Begin VB.Menu mnuIFields 
            Caption         =   "FieldName"
            Index           =   0
         End
      End
      Begin VB.Menu mnuBatchFields 
         Caption         =   "&Batch Fields"
         Begin VB.Menu mnuBFields 
            Caption         =   "FieldName"
            Index           =   0
         End
      End
      Begin VB.Menu mnuAscentValues 
         Caption         =   "&Ascent Capture Values"
         Begin VB.Menu mnuACList 
            Caption         =   "Batch Description"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTextConst 
         Caption         =   "&Text Constant"
      End
   End
End
Attribute VB_Name = "KfxLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Display state
Private m_nLinkTextWidth      ' Width of the textbox column in the control.
Private m_nLinkCaptionWidth   ' Width of the caption column in the control.
Private m_bShowButtons        ' Whether buttons are visible.
Private m_bEnabled As Boolean ' Whether control is enabled or grayed.

'=========================
' Links Global Constants
'=========================
Private Const LINK_BOX_SIZE = 9

Private Const LOCK_TEXT_BOX = 0
Private Const UNLOCK_TEXT_BOX = -1

Private Const CURRENT_LINK = -2
Private Const NO_SELECTION = -3
Private Const DELETE_LINK = -4

Const IN_PROGRESS = "IP"
Const BUTTON_CLICK = "BC"
Const SKIP_EVENT = "SKIP"
Dim m_IndexList() As KfxLink
Dim m_ILCount As Integer
Dim m_SelectedIndex As Integer

' -- Linking variables --
Dim m_NewIndexType As Integer
Dim m_NewIndexData As String

Const M_KFXLINK = "KfxLink.CTL"

Public Event Change()
Public Event Error(ErrNum As Long, _
                ErrMsg As String, _
                SourceFile As String, _
                LineNo As Integer, _
                ReRaise As Boolean, _
                Display As Boolean)
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&
Private Const MF_BYPOSITION = &H400&

' Resource string ID's.
Private Const MSG_DELETEINDEX = 3700
Private Const TITLE_KFXLINKSERROR = 3701
Private Const MSG_DELETEALLINDEX = 3702
Private Const CMD_ADDINDEX = 3703
Private Const CMD_DELETE = 3704
Private Const CMD_DELETEALL = 3705
Private Const CAP_MOVE = 3706

'*************************************************
' AddMenuColumns
'-------------------------------------------------
' Purpose:  This subroutinue breaks all submenus
'           off the main menu into columns.  The
'           OS used to do this for us but that
'           changed so we have to do it now.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub AddMenuColumns()
    Dim I As Integer
    Dim J As Integer
    Dim hmnuMainHandle As Long
    Dim hmnuSubHandle As Long
    Dim nRet As Long
    Dim nNum As Long
    Dim nMainMenuCount As Long
    Dim nSubMenuCount As Long
    Dim nIndex As Long
    Dim nSubMenuID As Long
    Dim sSubMenuCaption As String
    Dim nSubMenuCaptionLen As Long
    
    ' Compute how many menu items there can be on the screen
    nNum = Screen.Height / (TextHeight("W") + 65)
    ' Get the main menu handle
    hmnuMainHandle = GetMenu(hwnd)
    hmnuMainHandle = GetSubMenu(hmnuMainHandle, 0)
    nMainMenuCount = GetMenuItemCount(hmnuMainHandle)
    
    For J = 0 To nMainMenuCount - 1
        ' Get the sub menu handle
        hmnuSubHandle = GetSubMenu(hmnuMainHandle, J)
        If (hmnuSubHandle > 0) Then
            ' Find out how many items there are in the menu
            nSubMenuCount = GetMenuItemCount(hmnuSubHandle)
            ' Loop through the menu and place column breaks at the correct spots
            For I = 1 To nSubMenuCount / nNum
                nIndex = I * nNum
                ' Get the menu caption
                sSubMenuCaption = Space(80)
                nSubMenuCaptionLen = GetMenuString(hmnuSubHandle, _
                                                    nIndex, _
                                                    sSubMenuCaption, _
                                                    Len(sSubMenuCaption), _
                                                    MF_BYPOSITION)
                sSubMenuCaption = left(sSubMenuCaption, nSubMenuCaptionLen)
                nSubMenuID = GetMenuItemID(hmnuSubHandle, nIndex)
                ' Set the column break
                nRet = ModifyMenu(hmnuSubHandle, _
                                    nIndex, _
                                    MF_BYPOSITION Or MF_MENUBARBREAK, _
                                    nSubMenuID, _
                                    sSubMenuCaption)
            Next
        End If
    Next
End Sub

'*************************************************
' ShowScrollBar [Property Get]
'-------------------------------------------------
' Purpose:  Determine whether link scrollbar is
'           visible or not.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Property Get ShowScrollBar() As Boolean
    ShowScrollBar = vsbIndex.Visible
End Property

'*************************************************
' ShowScrollBar [Property Let]
'-------------------------------------------------
' Purpose:  Show or hide the link scrollbar
' Inputs:   bShow   Show=True; Hide=False
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Property Let ShowScrollBar(bShow As Boolean)
    vsbIndex.Visible = bShow
End Property

'*************************************************
' ShowButtons [Property Get]
'-------------------------------------------------
' Purpose:  Determine whether or not the buttons
'           on the link control are visible
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Property Get ShowButtons() As Boolean
    ShowButtons = m_bShowButtons
End Property

'*************************************************
' ShowButtons [Property Let]
'-------------------------------------------------
' Purpose:  Show or hide buttons on the control
' Inputs:   bShow   Show=True; Hide=False
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Property Let ShowButtons(bShow As Boolean)
    m_bShowButtons = bShow
    cmdDeleteAllIndex.Visible = bShow
    cmdDeleteIndex.Visible = bShow
    cmdAddIndex.Visible = bShow
    updnIndex.Visible = bShow
    lblMove.Visible = bShow
    
    ' When there are no buttons, then deleting is impossible.
    ' We only allow "Unlink" menu item in that case.
    If bShow Then
        mnuDelete.Visible = True
        mnuUnlink.Visible = False
    Else
        mnuUnlink.Visible = True
        mnuDelete.Visible = False
    End If
    
End Property

'*************************************************
' Add
'-------------------------------------------------
' Purpose:  Add an Index Value link to the array
'           of data values and give it focus.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Sub Add(ByVal strSource As String, ByVal eType As AscentRelease.KfxLinkSourceType, ByVal strDest As String, ByVal strCaption As String)
    
    On Error GoTo AddIndex_Failure
    
        RaiseEvent Change
    
        ' Add a blank Index Value at the end of the list
400     ReDim Preserve m_IndexList(m_ILCount)
        m_ILCount = m_ILCount + 1
        Set m_IndexList(m_ILCount - 1) = New KfxLink

        ' Populate the data in the Index Value
410     With m_IndexList(m_ILCount - 1)
            .Destination = strDest
            .Source = strSource
            .SourceType = eType
            .Caption = strCaption
        End With
    
        ' Show the new index value and give it the focus
        If m_ILCount > LINK_BOX_SIZE Then
420         Call DisplayIndexValues(m_ILCount - LINK_BOX_SIZE)
            If txtIndexData(LINK_BOX_SIZE - 1).Enabled And txtIndexData(m_ILCount - 1).Visible Then
430             txtIndexData(LINK_BOX_SIZE - 1).SetFocus
            End If
        Else
440         Call DisplayIndexValues(0)
450         If txtIndexData(m_ILCount - 1).Enabled And txtIndexData(m_ILCount - 1).Visible Then
                txtIndexData(m_ILCount - 1).SetFocus
            End If
        End If
        
        Exit Sub

AddIndex_Failure:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, False, True)

End Sub

'*************************************************
' mnuUnlink_Click
'-------------------------------------------------
' Purpose:  This event handler removes the link
'           between the Ascent Index Value and
'           the destination previously specified
'           in the storage settings.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub mnuUnlink_Click()
    m_NewIndexType = KFX_REL_UNDEFINED_LINK
    m_NewIndexData = ""
End Sub

'*************************************************
' txtIndexData_GotFocus
'-------------------------------------------------
' Purpose:  When an Index Value gets focus, we
'           display the menu button for that row
'           and highlight the selected row. Also,
'           certain controls are only enabled when
'           an Index Value is currently selected.
' Inputs:   Index   the control array index of the
'                   selected Index Value
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub txtIndexData_GotFocus(Index As Integer)
    Dim I As Integer

        On Error GoTo IndexDataGF_Error

2800    If cmdMenu(Index).Visible <> True Then
            ' Show the Menu button
2810        With txtIndexData(Index)
2820            .Width = .Width - cmdMenu(Index).Width
            End With
2830        cmdMenu(Index).Visible = True
2840        cmdMenu(Index).Tag = ""

            ' Highlight the selected row
2850        With txtIndexData(Index)
                .BackColor = vbHighlight
                .ForeColor = vbHighlightText
            End With
2860        With txtSequence(Index)
                .BackColor = vbHighlight
                .ForeColor = vbHighlightText
            End With
        End If
        
        ' The control with focus serves as the
        ' TabStop for the entire control array
        For I = 0 To LINK_BOX_SIZE - 1
            If Index = I Then
                txtIndexData(I).TabStop = True
            Else
                txtIndexData(I).TabStop = False
            End If
        Next I
        
        ' Enable controls
        cmdDeleteIndex.Enabled = True
        updnIndex.Enabled = True
        lblMove.Enabled = True
        m_SelectedIndex = Index + vsbIndex.Value
2870    Call SetMoveControl
        
        Exit Sub

IndexDataGF_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, False, True)

End Sub

'*************************************************
' txtIndexData_KeyDown
'-------------------------------------------------
' Purpose:  Process keystrokes while an Index
'           Value has focus.  This allows the
'           user to move between Index Values
'           with the keyboard or to press the
'           Enter key when done editting a
'           Text Constant.
' Inputs:   Index   control array index of the
'                   Index Value with focus
'           KeyCode the key that was pressed
'           Shift   flags for Alt, Shift, Ctrl
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub txtIndexData_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
        On Error GoTo IndexDataKD_Error
        
        If txtIndexData(Index).Locked = False Then
            Select Case KeyCode
                Case vbKeyReturn
2910                txtIndexData(Index).Tag = BUTTON_CLICK
2920                txtIndexData_LostFocus (Index)
            End Select
        Else
            Select Case KeyCode
                Case vbKeySpace
2930                cmdMenu_Click (Index)
                Case vbKeyUp, vbKeyLeft
                    If Index <> 0 Then
2940                    txtIndexData(Index - 1).SetFocus
                    ElseIf (vsbIndex.Visible) Then
                        With vsbIndex
                            If .Value <> .Min Then
                                .Value = .Value - 1
                            End If
                        End With
                    End If
                Case vbKeyDown, vbKeyRight
                    If (Shift = vbAltMask) And (KeyCode = vbKeyDown) Then
                        ' Alt + Down Arrow pops up the linking menu
2950                    cmdMenu_Click (Index)
                    ElseIf Index <> LINK_BOX_SIZE - 1 And Index < m_ILCount - 1 Then
2960                    txtIndexData(Index + 1).SetFocus
                    ElseIf (vsbIndex.Visible) Then
                        With vsbIndex
                            If .Value <> .Max Then
                                .Value = .Value + 1
                            End If
                        End With
                    End If
            
            End Select
        End If

        Exit Sub

IndexDataKD_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, False, True)

End Sub

'*************************************************
' txtIndexData_LostFocus
'-------------------------------------------------
' Purpose:  Cleans up the selected Index Value
'           when it loses focus.  The includes
'           removing the highlighting, disabling
'           controls that are only valid when an
'           Index Value is selected, and adding
'           quotes around a Text Constant if the
'           user was in the middle of editting.
' Inputs:   Index   control array index of the
'                   Index Value that had focus
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub txtIndexData_LostFocus(Index As Integer)

        On Error GoTo IndexDataLF_Error

        With txtIndexData(Index)
            If .Locked = False Then
                If Trim$(.Text) <> "" Then
3010                m_IndexList(vsbIndex.Value + Index).Source = Trim$(.Text)
                    .Text = """" & Trim(.Text) & """"
                Else
3020                m_IndexList(vsbIndex.Value + Index).Source = ""
3030                m_IndexList(vsbIndex.Value + Index).SourceType = NO_LINK
                End If
                .Locked = True
                .MousePointer = vbArrow
                .BackColor = vbHighlight
                .ForeColor = vbHighlightText
            End If
        End With
        
        If txtIndexData(Index).Tag <> BUTTON_CLICK Then
        
3040        If cmdMenu(Index).Visible Then
                With txtIndexData(Index)
3050                .Width = .Width + cmdMenu(Index).Width
                End With
3060            cmdMenu(Index).Visible = False
3070            cmdMenu(Index).Tag = ""
            End If
        
            With txtIndexData(Index)
                .BackColor = vbWindowBackground
                .ForeColor = vbWindowText
            End With
            With txtSequence(Index)
                .BackColor = vbButtonFace
                .ForeColor = vbButtonText
            End With
        
            ' Disable controls
            cmdDeleteIndex.Enabled = False
            updnIndex.Enabled = False
            lblMove.Enabled = False
        Else
            txtIndexData(Index).Tag = ""
        End If
        
        Exit Sub

IndexDataLF_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, False, True)

End Sub

'*************************************************
' cmdAddIndex_Click
'-------------------------------------------------
' Purpose:  Add a blank (unlinked) Index Value
'           to the end of the list and place
'           focus on the control.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub cmdAddIndex_Click()
    
    Me.Add "", NO_LINK, CStr(Me.LinkCount + 1), CStr(Me.LinkCount + 1)
    
End Sub

'*************************************************
' cmdDeleteIndex_Click
'-------------------------------------------------
' Purpose:  Ask the user if it is OK and then
'           delete the selected Index Value.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    VB bahavior causes this function to
'           get called recursively so we set a
'           local InProgress flag to skip it.
'*************************************************
Private Sub cmdDeleteIndex_Click()
    Dim Result As Integer

        On Error GoTo DeleteIndex_Error
        
        ' Verify that the user REALLY
        ' wants to delete the selected index value
        Result = MsgBox(LoadResString(MSG_DELETEINDEX), _
                        vbExclamation + vbYesNo, _
                        LoadResString(TITLE_KFXLINKSERROR))
                        
        If (Result = vbYes) Then
            ' Go ahead and delete it
550         DeleteIndex (m_SelectedIndex)
            RaiseEvent Change
        
            ' If the last index value was deleted then
            ' m_SelectedIndex is now invalid and the
            ' previous index value is selected
            If m_SelectedIndex = m_ILCount Then
                m_SelectedIndex = m_SelectedIndex - 1
            End If
        
            ' Display the modified list
560         Call DisplayIndexValues(vsbIndex.Value)
        End If
    
        ' Set focus back to the textbox
        If m_ILCount > 0 Then
570         txtIndexData(m_SelectedIndex - vsbIndex).SetFocus
        End If
    
        Exit Sub

DeleteIndex_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, False, True)

End Sub

'*************************************************
' cmdDeleteIndex_GotFocus
'-------------------------------------------------
' Purpose:  When the user clicks the Delete
'           button, the Index Value textbox
'           LostFocus event fires which in
'           turn disables the button. Therefore
'           the cmdDeleteIndex_Click event does
'           not occur and we must call it here.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    This is a side effect of the fact
'           that the Delete button is only
'           enabled when an Index Value is
'           selected (has focus).
'*************************************************
Private Sub cmdDeleteIndex_GotFocus()

        On Error GoTo Unexpected_Error
        
        Call cmdDeleteIndex_Click
        
        Exit Sub

Unexpected_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, False, True)

End Sub

'*************************************************
' cmdMenu_Click
'-------------------------------------------------
' Purpose:  This routine begins the sequence of
'           events that cause the popup menu to
'           appear just below the link box.
' Inputs:   Index   control array index
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub cmdMenu_Click(Index As Integer)
    Static InProgress As Boolean
    
        On Error GoTo MenuClick_Error
        
        RaiseEvent Change
    
        ' InProgress used to keep this code from being
        ' re-entrant.  The buttons tag value can be "IP",
        ' "SKIP", or empty.  The value is set to skip if the
        ' user clicks the button while the menu is still up.
        ' This will keep the menu from popping back up.  The
        ' "IP" value is set if we are In Progress and is
        ' used by the MouseDown event.
700     If (InProgress = False And cmdMenu(Index).Tag <> SKIP_EVENT) Then
            ' Set our in progress flags.
            InProgress = True
710         cmdMenu(Index).Tag = IN_PROGRESS
            ' Move the focus back to the link box
720         txtIndexData(Index).SetFocus
            ' Start the indexing code.  This routine will
            ' not return until the popup menu goes away.
730         Call DoTheLink(txtIndexData(Index), Index + vsbIndex.Value)
            ' Allow events to fire.  If the user clicked on the
            ' menu button to drop the menu down, this will cause
            ' this routine to be fired re-entrantly.
            DoEvents
            ' Clear the in-progress flags
            InProgress = False
            ' If the tag is set to SKIP, let the other instance
            ' of this event clear the tag.
740         If (cmdMenu(Index).Tag <> SKIP_EVENT) Then
750             cmdMenu(Index).Tag = ""
            End If
        Else
            ' Clear the tag and set focus back on the
            ' link box.
760         cmdMenu(Index).Tag = ""
770         txtIndexData(Index).SetFocus
        End If
        
        Exit Sub

MenuClick_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, False, True)

End Sub

'*************************************************
' cmdMenu_MouseDown
'-------------------------------------------------
' Purpose:  This event handler has two purposes.
'           It checks to see if the popup menu
'           is up to keep the click event from
'           re-displaying the menu again.  It
'           also keeps the link boxes SetFocus
'           event from rerunning when the focus
'           is returning from the menu button.
' Inputs:   Index   control array index
'           Button  which mouse button was down
'           Shift   flag for Ctrl, Alt, Shift
'           x       horizontal position
'           y       vertical position
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub cmdMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
        
        On Error GoTo Menu_MouseDown_Error
        
        ' If the menu button's click event is in progress,
        ' tell the event that the next time it's fired to
        ' skip the event handling code
780     If cmdMenu(Index).Tag = IN_PROGRESS Then
790         cmdMenu(Index).Tag = SKIP_EVENT
        End If
        
        ' Tell the link textbox that the command
        ' button is returning the focus.
800     txtIndexData(Index).Tag = BUTTON_CLICK
        
        Exit Sub

Menu_MouseDown_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, False, True)

End Sub

'*************************************************
' DisplayIndexValues
'-------------------------------------------------
' Purpose:  This routine displays Index Values
'           and their associated link in the link
'           box.  The caller specifies where in
'           the links array to start displaying.
' Inputs:   Index   specifies the starting point
'                   in the links array
' Outputs:  None
' Returns:  None
' Notes:    Text Constants are enclosed in quotes
'           when they are not edittable and all
'           Ascent Capture system-provided values
'           are displayed in bold.
'*************************************************
Private Sub DisplayIndexValues(Index As Integer)
    Dim I As Integer
    
        On Error GoTo DisplayIndexValues_Error
        
        ' Fill the link box display
        For I = 0 To LINK_BOX_SIZE - 1
            ' If our index is within the link array,
            ' display the array values
            If (I + Index) < m_ILCount Then
                ' Make the column name text box visible
                ' and fill it with the column name
1000            txtSequence(I).Visible = True
1010            txtSequence(I).Text = m_IndexList(I + Index).Caption
                                                       
                ' Display the link data box and fill it with
                ' the appropriate value
1020            With txtIndexData(I)
                    .Visible = True
1030                Select Case m_IndexList(I + Index).SourceType
                        Case NO_LINK
                            .FontBold = False
                            .Text = ""
                        Case KFX_REL_TEXTCONSTANT
                            .FontBold = False
1040                        .Text = """" & m_IndexList(I + Index).Source & """"
                        Case KFX_REL_VARIABLE
                            .FontBold = True
1050                        .Text = "{" & m_IndexList(I + Index).Source & "}"
                        Case KFX_REL_INDEXFIELD
                            .FontBold = True
1060                        .Text = m_IndexList(I + Index).Source
                        Case KFX_REL_BATCHFIELD
                            .FontBold = True
1065                        .Text = "{$" & m_IndexList(I + Index).Source & "}"
                        Case KFX_REL_DOCUMENTID
                            .FontBold = True
                            .Text = m_IndexList(I + Index).Source
                    End Select
                End With
            Else
                ' Otherwise hide the sequence number
                ' and the data box
1070            txtSequence(I).Visible = False
1080            txtSequence(I).Text = ""
                
1090            txtIndexData(I).Visible = False
1100            txtIndexData(I).Text = ""
            End If
        Next I
        
        ' Enable or Disable the <Delete All> button
        cmdDeleteAllIndex.Enabled = (m_ILCount > 0)
        
        ' Update the scrollbar to
        ' represent the display
1110    SetScrollBar (Index)
        
        Exit Sub

DisplayIndexValues_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, True, False)

End Sub

'*************************************************
' cmdDeleteAllIndex_Click
'-------------------------------------------------
' Purpose:  Ask the user if it is OK and then
'           delete all defined Index Values.
'           Mark the data dirty if the Index
'           Values are deleted.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub cmdDeleteAllIndex_Click()
    Dim Result As Integer

        On Error GoTo DeleteAll_Error
        
        ' Verify that the user REALLY
        ' wants to delete all index values
        Result = MsgBox(LoadResString(MSG_DELETEALLINDEX), _
                        vbExclamation + vbYesNo, _
                        LoadResString(TITLE_KFXLINKSERROR))
                        
        If (Result = vbYes) Then
            ' Go ahead and delete them
500         Call DeleteAllIndex
            RaiseEvent Change
    
            ' Display the empty list
510         Call DisplayIndexValues(0)
        End If

        Exit Sub

DeleteAll_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, False, True)

End Sub

'*************************************************
' DoTheLink
'-------------------------------------------------
' Purpose:  This routine performs the steps to
'           allow a user to select a link value.
' Inputs:   LBox    selected text box
'           Index   array index of the link
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub DoTheLink(LBox As TextBox, Index As Integer)
    Dim BoxLeft As Integer
    Dim BoxTop As Integer
    Dim Results As Integer
        
        On Error GoTo DoTheLink_Error
        
        ' Initialize the return value to use
        ' no selection.  Handles when user
        ' clicks away from the menu without making
        ' a menu selection.
        m_NewIndexType = NO_SELECTION
        
        ' Pop up the menu just below the textbox
1150    Call GetBoxPosition(LBox, BoxLeft, BoxTop)
1160    Call PopupMenu(mnuLinks, 0, BoxLeft, BoxTop)
        
        ' After the menu returns, establish the new link
        If m_NewIndexType <> DELETE_LINK And _
           m_NewIndexType <> NO_SELECTION Then
1170        Results = EstablishLink(Index)
            If Results = UNLOCK_TEXT_BOX Then
                With LBox
1180                .Text = m_IndexList(Index).Source
                    .Locked = False
                    .MousePointer = vbDefault
                    .BackColor = vbWindowBackground
                    .ForeColor = vbWindowText
                    .FontBold = False
                End With
            Else
1190            Call DisplayIndexValues(vsbIndex.Value)
            End If
        End If
        
        Exit Sub

DoTheLink_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, True, False)

End Sub

'*************************************************
' UserControl_Initialize
'-------------------------------------------------
' Purpose:  Initializes the constituent controls
'           when this control is first loaded.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub UserControl_Initialize()
    
    ' Store startup size for later resizing.
    m_nLinkCaptionWidth = txtSequence(0).Width
    m_nLinkTextWidth = txtIndexData(0).Width
    m_bShowButtons = True
    
    ' Internationalize control UI from resource file.
    InternationalizeUI
        
End Sub

'*************************************************
' InternationalizeUI
'-------------------------------------------------
' Purpose:  Internationalize the constituent
'           controls on the user interface from
'           resource file strings.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub InternationalizeUI()

1480    cmdAddIndex.Caption = LoadResString(CMD_ADDINDEX)
1485    cmdDeleteIndex.Caption = LoadResString(CMD_DELETE)
1490    cmdDeleteAllIndex.Caption = LoadResString(CMD_DELETEALL)
1495    lblMove.Caption = LoadResString(CAP_MOVE)

End Sub

'*************************************************
' DeleteAllIndex
'-------------------------------------------------
' Purpose:  This routine deletes all Index Values
'           from the list of links.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub DeleteAllIndex()
    Dim I As Integer
    
    ' Discard all Index Values
    m_ILCount = 0
    ReDim m_IndexList(0)

End Sub

'*************************************************
' DeleteIndex
'-------------------------------------------------
' Purpose:  This routine deletes the specified
'           Index Value from the list of links
' Inputs:   Index   index into the links array
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub DeleteIndex(Index As Integer)
    Dim I As Integer
    
        On Error GoTo DeleteIndex_Error
        
        ' If more than one index value in the list
        ' and not the last index value in the list,
        ' shift remaining values up one position
        If (m_ILCount > 1) And (Index <> m_ILCount - 1) Then
            For I = Index To m_ILCount - 2
950             m_IndexList(I).Source = m_IndexList(I + 1).Source
960            m_IndexList(I).SourceType = m_IndexList(I + 1).SourceType
            Next
        End If
        
        ' Decrement the number of index values
        m_ILCount = m_ILCount - 1
        
        ' Resize the list of index values
        If m_ILCount > 0 Then
970         ReDim Preserve m_IndexList(m_ILCount - 1)
        Else
            ReDim m_IndexList(0)
        End If
        
        Exit Sub

DeleteIndex_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, True, False)

End Sub

'*************************************************
' EstablishLink
'-------------------------------------------------
' Purpose:  This routine will build the link
'           between a destination value and
'           the selected link data
' Inputs:   Index   index to the links array
' Outputs:  None
' Returns:  Lock Status of the text box
' Notes:    None
'*************************************************
Private Function EstablishLink(Index As Integer) As Integer

        ' Default to locking the text box on the
        ' conclusion of this routine.  Only the
        ' selection of a text constant link will
        ' unlock the text box.
        EstablishLink = LOCK_TEXT_BOX
    
        ' Depending upon the type,
        ' handle source data value.
        Select Case m_NewIndexType
            Case KFX_REL_TEXTCONSTANT
                ' If the current source type is not
                ' Text Constant, initialize the text
1300            If m_IndexList(Index).SourceType <> KFX_REL_TEXTCONSTANT Then
1310                m_IndexList(Index).Source = ""
                End If
                ' Unlock the text box and wait until
                ' the user has entered the text string
                EstablishLink = UNLOCK_TEXT_BOX
            
            Case KFX_REL_BATCHFIELD, _
                        KFX_REL_INDEXFIELD, _
                        KFX_REL_VARIABLE, _
                        KFX_REL_DOCUMENTID
                ' For the above just store the
                ' field name as the source
1320            m_IndexList(Index).Source = m_NewIndexData
        End Select
        
        ' Set the new link type
1330    m_IndexList(Index).SourceType = m_NewIndexType
        
        Exit Function

EstablishLink_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, True, False)

End Function

'*************************************************
' InitializeIndexValues
'-------------------------------------------------
' Purpose:  This routine will initialize the
'           Index Values with all Batch Fields
'           and Index Fields
' Inputs:   oSetupData  ReleaseSetupData object
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub InitializeIndexValues(oSetupData As ReleaseSetupData)
    Dim BField As AscentRelease.BatchField
    Dim IField As AscentRelease.IndexField
    Dim I As Integer
    
        On Error GoTo IIV_LogAndPropError
            
        I = 0
    
        ' Add each batch field to the list of index values
        For Each BField In oSetupData.BatchFields
            Me.Add BField.Name, KFX_REL_BATCHFIELD, I, I
            I = I + 1
        Next
        
        ' Add each index field to the list of index values
        For Each IField In oSetupData.IndexFields
            Me.Add IField.Name, KFX_REL_INDEXFIELD, I, I
            I = I + 1
        Next
        
        Me.Refresh
        
        Exit Sub
        
IIV_LogAndPropError:

    RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, True, False)
        
End Sub

'***********************************************
' LinksExist
'-----------------------------------------------
' Purpose:  This routine will search all of the
'           link entries to see if any links
'           exist.  If so, return true
' Inputs:   None
' Outputs:  None
' Returns:  True/False
' Notes:    None
'***********************************************
Private Function LinksExist() As Boolean
    Dim I As Integer

        On Error GoTo LinksExist_Error

        ' Loop through each entry and look for
        ' any sourcetype value other than NO_LINK
        For I = 0 To m_ILCount - 1
2030        If m_IndexList(I).SourceType <> NO_LINK Then
                LinksExist = True
                Exit Function
            End If
        Next I
        
        Exit Function

LinksExist_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, True, False)

End Function

'*************************************************
' Init
'-------------------------------------------------
' Purpose:  Initialize the control. Builds the
'           menus with valid fieldnames from the
'           Ascent database and loads previous
'           Index Value links.
' Inputs:   oSetupData  ReleaseSetupData object
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Sub Init(SetupData As ReleaseSetupData)

1600    Call InitializeIndexValues(SetupData)

2640    Call BuildLinkingMenu(SetupData)
2650    Call AddMenuColumns


End Sub

'*************************************************
' SetMoveControl
'-------------------------------------------------
' Purpose:  The updnIndex control is used to track
'           which Index Value currently has focus
'           as well as whether an Index Value has
'           moved to the top/bottom of the list.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub SetMoveControl()

        On Error GoTo SMC_Error
        
        With updnIndex
            ' Recalculate the range
2550        .Min = m_ILCount - 1
2560        .Max = 0
2570        .Value = m_SelectedIndex
        End With
        
        Exit Sub

SMC_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, True, False)

End Sub

'*************************************************
' SetScrollBar
'-------------------------------------------------
' Purpose:  If there are more Index Values than
'           the text boxes can display then the
'           scroll bar is made visible and the
'           scroll range is set.  The scroll bar
'           may be set to a new position.
' Inputs:   Position    the new scroll bar value
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub SetScrollBar(Position As Integer)
    Dim Overflow As Integer
    
        On Error GoTo SSB_Error
        
        With vsbIndex
            ' Calculate the new range for the scroll bar
            Overflow = m_ILCount - LINK_BOX_SIZE
            If Overflow > 0 Then
                .Visible = True
2580            .Max = Overflow
            Else
                .Visible = False
2590            .Max = 0
            End If
            
            ' Set the new position if in valid range
            If (Position >= .Min) And (Position <= .Max) Then
2600            .Value = Position
            End If
        End With
        
        Exit Sub

SSB_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, True, False)

End Sub

'*************************************************
' GetBoxPosition
'-------------------------------------------------
' Purpose:  This routine will find a control's
'           absolute X and Y coordinates on the
'           usercontrol.  It takes into account
'           all containers.
' Inputs:   TheControl      specified control
' Outputs:  Left            x coordinate
'           Top             y coordinate
' Returns:  None
' Notes:    This routine must be called ByRef.
'           The left and top values will not be
'           returned if it is called ByVal.
'*************************************************
Private Sub GetBoxPosition(TheControl As Control, left As Integer, top As Integer)
    Dim TheParent As Object
    
    On Error GoTo GBP_Error

        ' Initialize variables and get the
        ' controls position within it's
        ' current container
2000    Set TheParent = TheControl
        left = TheControl.left
        top = TheControl.top
            
        ' Now set where in the text box the upper left
        ' corner of the menu will appear.  Currently set
        ' to appear left aligned and just below the text box.
        top = top + (TheControl.Height)
        left = left
        
        Exit Sub

GBP_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, True, False)

End Sub

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
Private Sub BuildAscentMenu(oSetupData As ReleaseSetupData)
    Dim oBatchVar As Variant
    Dim I As Integer
    
    On Error GoTo BAM_LogAndPropError
        
        ' Add the Ascent Capture values to the menu
        I = 0
        For Each oBatchVar In oSetupData.BatchVariableNames
            If (I <> 0) Then
150             Load mnuACList(I)
            End If
160         mnuACList(I).Visible = True
170         mnuACList(I).Caption = oBatchVar
            I = I + 1
        Next
    
        Exit Sub
        
BAM_LogAndPropError:

        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, True, False)
        
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
Private Sub BuildBatchMenu(oSetupData As ReleaseSetupData)
    Dim CurrField As AscentRelease.BatchField
    Dim I As Integer
    
    On Error GoTo BBM_LogAndPropError
        
        ' Add the Batch Fields to the menu
        If oSetupData.BatchFields.Count > 0 Then
            I = 0
            ' Populate each element with the data
            For Each CurrField In oSetupData.BatchFields
                If (I <> 0) Then
200                 Load mnuBFields(I)
                End If
210             mnuBFields(I).Visible = True
220             mnuBFields(I).Caption = CurrField.Name
                I = I + 1
            Next
        Else
            ' There are no Batch Fields so remove the menu option
            mnuBatchFields.Visible = False
        End If
    
        Exit Sub
        
BBM_LogAndPropError:

        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, True, False)
        
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
    Dim I As Integer
        
    On Error GoTo BIM_LogAndPropError
        
        ' Add the Index Fields to our menu.
        If oSetupData.IndexFields.Count > 0 Then
            I = 0
            ' Populate each menu item
            For Each CurrField In oSetupData.IndexFields
                If (I <> 0) Then
250                Load mnuIFields(I)
                End If
260             mnuIFields(I).Visible = True
270             mnuIFields(I).Caption = CurrField.Name
                I = I + 1
            Next
        Else
            ' There are no Index Fields so remove the menu option
            mnuIndexFields.Visible = False
        End If
    
        Exit Sub

BIM_LogAndPropError:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, True, False)
        
End Sub

'*************************************************
' BuildLinkingMenu
'-------------------------------------------------
' Purpose:  This routine will build the linking
'           popup menu used by the link box.
'           It will add entries for each of the
'           Index Fields, Batch Fields, and
'           Ascent Capture Values.
' Inputs:   oSetupData  ReleaseSetupData object
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Sub BuildLinkingMenu(oSetupData As ReleaseSetupData)

    ' Build the three variable linking
    ' menu lists.  This only needs
    ' to be done once at the start.
    On Error GoTo BLM_LogAndPropError

300     Call BuildAscentMenu(oSetupData)
310     Call BuildBatchMenu(oSetupData)
320     Call BuildIndexMenu(oSetupData)

        Exit Sub

BLM_LogAndPropError:

    RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, True, False)

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
    m_NewIndexType = KFX_REL_VARIABLE
    m_NewIndexData = mnuACList(Index).Caption
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
    m_NewIndexType = KFX_REL_BATCHFIELD
    m_NewIndexData = mnuBFields(Index).Caption
End Sub

'*************************************************
' mnuCurrIndex_Click
'-------------------------------------------------
' Purpose:  This event handler sets the globals
'           used to complete a data link.  The
'           caption of the menu item holds the
'           current Index Value.
' Inputs:   Index   the selected menu item
' Outputs:  None
' Returns:  None
' Notes:    We currently have this menu option
'           disabled.  It was intended to allow
'           the user a way to retain the current
'           data link after displying the popup
'           menu.  This is not really necessary
'           since the user can press the ESC key
'           to dismiss the menu.  It is left here
'           for end users that might wish to use
'           this functionality by enabling the
'           mnuCurrIndex menu.
'*************************************************
Private Sub mnuCurrIndex_Click()
    m_NewIndexType = CURRENT_LINK
    m_NewIndexData = ""
End Sub

'*************************************************
' mnuDelete_Click
'-------------------------------------------------
' Purpose:  This event handler deletes the link
'           that is currently selected.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub mnuDelete_Click()
    m_NewIndexType = DELETE_LINK
    m_NewIndexData = ""
    cmdDeleteIndex_Click
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
    m_NewIndexType = KFX_REL_DOCUMENTID
    m_NewIndexData = StripAmpersands(mnuDocumentID.Caption)
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
    m_NewIndexType = KFX_REL_INDEXFIELD
    m_NewIndexData = mnuIFields(Index).Caption
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
    m_NewIndexType = KFX_REL_TEXTCONSTANT
    m_NewIndexData = ""
End Sub

'*************************************************
' MoveIndex
'-------------------------------------------------
' Purpose:  This routine moves the Index Value
'           up or down one position in the array
' Inputs:   Direction (UP_ONE or DOWN_ONE)
' Outputs:  None
' Returns:  None
' Notes:    UP and DOWN represent the visual
'           appearance in the list to the user
'           so UP actually moves the link to a
'           lower index while DOWN moves it to
'           a higher index.
'*************************************************
Private Sub MoveIndex(Direction As Integer)
    Dim tmpIndex As New KfxLink
    Dim I As Integer

        On Error GoTo MoveIndex_Error

        I = m_SelectedIndex
    
2200    tmpIndex = m_IndexList(I)
        Select Case Direction
            Case UP_ONE
                ' Make sure we're not already at the start of list
                If I > 0 Then
2210                m_IndexList(I).Source = m_IndexList(I - 1).Source
2220                m_IndexList(I).SourceType = m_IndexList(I - 1).SourceType
2230                m_IndexList(I - 1).Source = tmpIndex.Source
2240                m_IndexList(I - 1).SourceType = tmpIndex.SourceType
                    If I = vsbIndex.Value Then
                        ' Scroll the list
                        vsbIndex.Value = vsbIndex.Value - 1
                        m_SelectedIndex = m_SelectedIndex - 1
                    Else
2250                    Call DisplayIndexValues(vsbIndex.Value)
2260                    txtIndexData(I - vsbIndex.Value - 1).SetFocus
                    End If
                End If
    
            Case DOWN_ONE
                ' Make sure we're not already at the end of list
                If I < (m_ILCount - 1) Then
                    ' Swap the two items in the list
2270                m_IndexList(I).Source = m_IndexList(I + 1).Source
2280                m_IndexList(I).SourceType = m_IndexList(I + 1).SourceType
2290                m_IndexList(I + 1).Source = tmpIndex.Source
2300                m_IndexList(I + 1).SourceType = tmpIndex.SourceType
                    If I = vsbIndex.Value + LINK_BOX_SIZE - 1 Then
                        ' Scroll the list
                        vsbIndex.Value = vsbIndex.Value + 1
                        m_SelectedIndex = m_SelectedIndex + 1
                    Else
2310                    Call DisplayIndexValues(vsbIndex.Value)
2320                    txtIndexData(I - vsbIndex.Value + 1).SetFocus
                    End If
                End If
        End Select
        
        Exit Sub

MoveIndex_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, True, False)

End Sub

'*************************************************
' Enabled [Property Let]
'-------------------------------------------------
' Purpose:  Sets the grayed state of the control.
' Inputs:   nEnabled    True/False determines if
'                       the constituent controls
'                       are enabled (True) or
'                       disabled (False)
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Property Let Enabled(bEnabled As Boolean)

    Dim cc As Control
    For Each cc In Controls
        If Not TypeOf cc Is Menu Then
            cc.Enabled = bEnabled
            
            ' See Microsoft Article ID: Q185880
            ' Fixes accelerator keys
            On Error Resume Next
            cc.Caption = cc.Caption
        End If
    Next cc
    On Error GoTo 0
    m_bEnabled = bEnabled
    
End Property

'*************************************************
' Enabled [Property Get]
'-------------------------------------------------
' Purpose:  Gets the grayed state of the control.
' Inputs:   None
' Outputs:  None
' Returns:  True or False indicating whether the
'           constituent controls are enabled (True)
'           or disabled (False).
' Notes:    None
'*************************************************
Public Property Get Enabled() As Boolean
    Enabled = m_bEnabled
End Property

'*************************************************
' txtSequence_GotFocus
'-------------------------------------------------
' Purpose:  The user never really wants focus on
'           left side of the link box, so ship
'           focus to the active (right) side.
' Inputs:   Index   Index Value that got focus
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub txtSequence_GotFocus(Index As Integer)
    If txtIndexData(Index).Visible Then
        txtIndexData(Index).SetFocus
    End If
End Sub

'*************************************************
' updnIndex_DownClick
'-------------------------------------------------
' Purpose:  Moves the selected Index Value down
'           one position in the list.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub updnIndex_DownClick()

        On Error GoTo IDC_Error

        RaiseEvent Change
3100    Call MoveIndex(DOWN_ONE)
        Exit Sub

IDC_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, False, True)

End Sub

'*************************************************
' updnIndex_UpClick
'-------------------------------------------------
' Purpose:  Moves the selected Index Value up one
'           position in the list.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub updnIndex_UpClick()

        On Error GoTo IUC_Error

        RaiseEvent Change
3120    Call MoveIndex(UP_ONE)
        Exit Sub

IUC_Error:
    
        RaiseEvent Error(Err, Err.Description, M_KFXLINK, Erl, False, True)

End Sub

'*************************************************
' UserControl_Resize
'-------------------------------------------------
' Purpose:  Forces the control to be repainted
'           when it is resized.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub UserControl_Resize()
    Refresh
End Sub

'*************************************************
' UserControl_Show
'-------------------------------------------------
' Purpose:  Make all constituent controls visible
'           at design time.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub UserControl_Show()
    If Ambient.UserMode = False Then
        Dim cc As Control
        For Each cc In Controls
            If Not TypeOf cc Is Menu Then
                cc.Visible = True
            End If
        Next cc
    End If
    
End Sub

'*************************************************
' vsbIndex_Change
'-------------------------------------------------
' Purpose:  This routine handles the scrolling of
'           the Index Values.  A static variable
'           keeps it from being re-entrant. It
'           also handles when a user scrolls while
'           in the middle of entering a Text
'           Constant.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub vsbIndex_Change()
    Static LinksInProgress As Boolean   ' Are we currently processing a change event?
    Static LastScrollValue As Integer   ' The value after the last scroll event
    Dim CurrentScrollValue As Integer
    Dim I As Integer
    
        On Error GoTo Scroll_Error
        
        ' Before starting, check to see if this routine
        ' is currently running.  If so, skip all the code.
        If Not LinksInProgress Then
            ' Enter Critical Region
            LinksInProgress = True
        
            ' Check each of the text boxes to see if
            ' any are currently editing a text constant.  If
            ' so, go back to the previous scroll setting
            ' and allow the text box to finish before scrolling
            For I = 0 To LINK_BOX_SIZE - 1
3400            If txtIndexData(I).Locked = False Then
                    CurrentScrollValue = vsbIndex.Value
3410                vsbIndex.Value = LastScrollValue
3420                Call txtIndexData_LostFocus(I)
3430                vsbIndex.Value = CurrentScrollValue
                End If
            Next I
        
            ' Now display the new group of indexes
3450        Call DisplayIndexValues(vsbIndex.Value)
            
            ' Keep a record of the last scroll setting
            LastScrollValue = vsbIndex.Value
            
            ' Exit Critical Region
            LinksInProgress = False
        End If
        
        Exit Sub

Scroll_Error:
    ' Abort the change
    LinksInProgress = False
    ' Reassert the error
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Sub

'*************************************************
' RestoreDefaults
'-------------------------------------------------
' Purpose:  Restore control to default state.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Sub RestoreDefaults()
    Stop
    Refresh
End Sub

'*************************************************
' LinkCaptionWidth [Property Get]
'-------------------------------------------------
' Purpose:  Gets the width of the Caption column
'           textbox in the control.
' Inputs:   None
' Outputs:  None
' Returns:  Width of the Caption column textbox
' Notes:    None
'*************************************************
Public Property Get LinkCaptionWidth() As Integer
    LinkCaptionWidth = m_nLinkCaptionWidth
End Property

'*************************************************
' LinkCaptionWidth [Property Let]
'-------------------------------------------------
' Purpose:  Set the desired width of the Caption
'           column textbox in the control.
' Inputs:   nWidth    specified width of textbox
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Property Let LinkCaptionWidth(ByVal nWidth As Integer)
    m_nLinkCaptionWidth = nWidth
    Refresh
End Property

'*************************************************
' LinkTextWidth [Property Get]
'-------------------------------------------------
' Purpose:  Gets the width of Index Value column
'           textbox in the control.
' Inputs:   None
' Outputs:  None
' Returns:  Width of the Index Value column textbox
' Notes:    None
'*************************************************
Public Property Get LinkTextWidth() As Integer
    LinkTextWidth = m_nLinkTextWidth
End Property

'*************************************************
' LinkTextWidth [Property Let]
'-------------------------------------------------
' Purpose:  Sets the desired width of the Index
'           Value column textbox in the control.
' Inputs:   nWidth    specified width of textbox
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Property Let LinkTextWidth(ByVal nWidth As Integer)
    m_nLinkTextWidth = nWidth
    Refresh
End Property

'*************************************************
' Link [Property Get]
'-------------------------------------------------
' Purpose:  Gets information about one link in
'           the list.
' Inputs:   nIdx    index of the link in the list
' Outputs:  None
' Returns:  The requested link object.
' Notes:    None
'*************************************************
Public Property Get Link(ByVal nIdx As Integer) As KfxLink
    Set Link = m_IndexList(nIdx)
End Property

'*************************************************
' Refresh
'-------------------------------------------------
' Purpose:  Repaint the control using current
'           properties.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Sub Refresh()
    Dim ii As Integer
    
    For ii = 0 To LINK_BOX_SIZE - 1
        txtSequence(ii).Width = m_nLinkCaptionWidth
        txtIndexData(ii).left = txtSequence(ii).left + txtSequence(ii).Width + 60
        txtIndexData(ii).Width = m_nLinkTextWidth
        cmdMenu(ii).left = txtIndexData(ii).left + txtIndexData(ii).Width - cmdMenu(ii).Width
    Next ii
    
    Call DisplayIndexValues(0)
            
End Sub

'*************************************************
' LinkCount [Property Get]
'-------------------------------------------------
' Purpose:  Gets the number of links in the
'           control.
' Inputs:   None
' Outputs:  None
' Returns:  Number of links
' Notes:    None
'*************************************************
Public Property Get LinkCount() As Integer
    LinkCount = m_ILCount
End Property

'*************************************************
' RemoveAll
'-------------------------------------------------
' Purpose:  Remove all the links from the control.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Sub RemoveAll()

400     ReDim m_IndexList(0)
        m_ILCount = 0
        Me.Refresh
    
End Sub

'*************************************************
' EnsureVisible
'-------------------------------------------------
' Purpose:  Makes sure the specified index item
'           in the list is visible within the
'           link box and sets focus to it.
' Inputs:   nLinkIndex  index into link array
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Sub EnsureVisible(ByVal nLinkIdx As Integer)

        If m_ILCount - nLinkIdx > LINK_BOX_SIZE Then
3230        Call DisplayIndexValues(nLinkIdx)
3240        txtIndexData(nLinkIdx).SetFocus
        ElseIf m_ILCount > LINK_BOX_SIZE Then
3250        Call DisplayIndexValues(m_ILCount - LINK_BOX_SIZE)
3260        txtIndexData(nLinkIdx - (m_ILCount - LINK_BOX_SIZE)).SetFocus
        Else
3270        Call DisplayIndexValues(0)
3280        txtIndexData(nLinkIdx).SetFocus
        End If

End Sub

'*************************************************
' cmdMenu_LostFocus
'-------------------------------------------------
' Purpose:  The menu button returns focus to
'           the link box after setting its own
'           Tag property to IN_PROGRESS in the
'           Click event. In that instance we do
'           nothing.  If the menu button loses
'           focus with a different Tag value,
'           then the Click event did not occur.
'           The user must have moved the mouse
'           off the menu button before doing a
'           Mouse Up and clicked a different
'           focus.  We therefore need to clean
'           up the link box so it no longer
'           looks like it has focus.
' Inputs:   Index
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub cmdMenu_LostFocus(Index As Integer)
        If cmdMenu(Index).Tag <> IN_PROGRESS Then
3350        txtIndexData(Index).Tag = ""
3360        Call txtIndexData_LostFocus(Index)
        End If
End Sub

