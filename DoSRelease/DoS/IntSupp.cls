VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CInternationalSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'==============================
' Constants for the API calls
'==============================
Private Const LF_FACESIZE& = 32
Private Const SYSTEM_FONT& = 13
Private Const HKEY_CURRENT_USER = &H80000001
Private Const KEY_QUERY_VALUE& = &H1
Private Const DEFAULT_GUI_FONT& = 17

Private Const LANG_BULGARIAN = &H2
Private Const LANG_CHINESE = &H4
Private Const LANG_CROATIAN = &H1A
Private Const LANG_CZECH = &H5
Private Const LANG_DANISH = &H6
Private Const LANG_DUTCH = &H13
Private Const LANG_ENGLISH = &H9
Private Const LANG_FINNISH = &HB
Private Const LANG_FRENCH = &HC
Private Const LANG_GERMAN = &H7
Private Const LANG_GREEK = &H8
Private Const LANG_HUNGARIAN = &HE
Private Const LANG_ICELANDIC = &HF
Private Const LANG_ITALIAN = &H10
Private Const LANG_JAPANESE = &H11
Private Const LANG_KOREAN = &H12
Private Const LANG_NEUTRAL = &H0
Private Const LANG_NORWEGIAN = &H14
Private Const LANG_POLISH = &H15
Private Const LANG_PORTUGUESE = &H16
Private Const LANG_ROMANIAN = &H18
Private Const LANG_RUSSIAN = &H19
Private Const LANG_SLOVAK = &H1B
Private Const LANG_SLOVENIAN = &H24
Private Const LANG_SPANISH = &HA
Private Const LANG_SWEDISH = &H1D
Private Const LANG_TURKISH = &H1F

'=========================
' Data Type Declarations
'=========================
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE - 1) As Byte
End Type

'===========================
' Windows API Declarations
'===========================
Private Declare Function SelectObject& Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long)
Private Declare Function GetStockObject& Lib "gdi32" (ByVal nIndex As Long)
Private Declare Function lstrlen& Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String)
Private Declare Function GetTextFace& Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String)
Private Declare Function ReleaseDC& Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long)
Private Declare Function GetSystemDefaultLangID% Lib "kernel32" ()

'==================
' Local Variables
'==================
Private sNewFont As String

'*************************************************
' Class_Initialize
'-------------------------------------------------
' Purpose:  Initialize local variables
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub Class_Initialize()

    sNewFont = ""
    
End Sub

'*************************************************
' FixAllFonts
'-------------------------------------------------
' Purpose:  This routine "fixes" the fonts on a
'           form and its controls to the correct
'           system default.  VB, unfortunately,
'           doesn't do this by default which may
'           cause some systems (like Japanese)
'           to fail.
' Inputs:   frmX    form to be "fixed"
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Sub FixAllFonts(frmX As Form)
    Dim hDesktop As Long
    Dim hSysFont As Long
    Dim lOldFont As Long
    Dim lRet As Long
    Dim ctl As Control
    
    If Len(sNewFont) = 0 Then
        ' Size the string for the API call.
        sNewFont = Space(80)
        
        ' Get a handle to the system's font or the default UI font.
        hSysFont = GetStockObject(DEFAULT_GUI_FONT)
        
        ' Set my window to the correct default font.
        lOldFont = SelectObject(frmX.hdc, hSysFont)
        lRet = GetTextFace(frmX.hdc, Len(sNewFont), sNewFont)
        
        sNewFont = left$(sNewFont, lstrlen(sNewFont))
    
        ' Release the object.
        lRet = SelectObject(frmX.hdc, lOldFont)
        lRet = ReleaseDC(frmX.hwnd, frmX.hdc)
    End If
    
    ' For speed, we only fix up the fonts if they are different than the default
    ' VB setting of "‚l‚r ‚oƒSƒVƒbƒN"
    If sNewFont <> "‚l‚r ‚oƒSƒVƒbƒN" Then
        ' Update all the fonts.
        frmX.Font.Name = sNewFont
        On Error Resume Next
        For Each ctl In frmX.Controls
            ctl.Font.Name = sNewFont
        Next
    End If
    
    On Error GoTo 0
    
End Sub

