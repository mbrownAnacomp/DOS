Attribute VB_Name = "ResConstants"
Option Explicit

'==========================
' String Table References
'==========================
' -- Tab Captions --
Public Const CAP_TAB_INDEX = 3
Public Const CAP_TAB_DOCUMENT = 4
Public Const CAP_TAB_IMAGEFORMAT = 5

' -- Frame and Label Captions --
Public Const CAP_FILENAME = 9
Public Const CAP_IDXFRAME = 10
Public Const CAP_SEQUENCE = 11
Public Const CAP_MOVE = 22

' -- Command Button Captions --
Public Const CMD_ADDINDEX = 56
Public Const CMD_DELETE = 57
Public Const CMD_DELETEALL = 58

' -- Popup Links Menu --
Public Const MNU_DELETE = 71

'================================
' Error Values And Message Refs
'================================
' -- Error Values --
Public Const BASE_ERR = 32000
Public Const ERR_FAILEDTOLOADFORM = BASE_ERR + 4

' -- Message Box Titles --
Public Const TITLE_FORM = 950
Public Const TITLE_RSETUPERROR = 951
Public Const TITLE_RELEASEERROR = 952
Public Const TITLE_TEXTRSETUP = 953
Public Const TITLE_DIRSELDIALOG = 954
Public Const TITLE_SELECTIMGDIR = 955
Public Const TITLE_SELECTOCRDIR = 956
Public Const TITLE_ERRORMSGBOX = 958
Public Const TITLE_SELECTPDFDIR = 963
Public Const TITLE_ADOBEACROBAT = 964


' -- Error Message Resource References --
Public Const MSG_DELETEINDEX = 1000
Public Const MSG_DELETEALLINDEX = 1001
Public Const MSG_READONLYFILE = 1011
Public Const MSG_INVALIDFILE = 1012
Public Const MSG_INVALIDPATH = 1013
Public Const MSG_COULDNOTOPENFILE = 1014
Public Const MSG_FILEREQUIRED = 1015
Public Const MSG_NOINDEXVALUES = 1016
Public Const MSG_BLANKINDEXVALUE = 1017

'=============================
' Keys for Custom Properties
'=============================
Public Const KEY_ASCIIFILE = "ASCII File Name"
Public Const KEY_ADOBE_DELETE_HUNG = "Adobe Delete Hung"
Public Const KEY_ADOBE_WAIT_FOR_STATUS = "Adobe Wait For Status"
Public Const KEY_DISABLE_IMAGE_EXPORT = "DisableImageExport"
Public Const KEY_DISABLE_TEXT_EXPORT = "DisableTextExport"

'=========================
' Links Global Constants
'=========================
Public Const LINK_BOX_SIZE = 9

Public Const LOCK_TEXT_BOX = 0
Public Const UNLOCK_TEXT_BOX = -1

Public Const NO_LINK = -1 'KFX_REL_UNDEFINED
Public Const CURRENT_LINK = -2
Public Const NO_SELECTION = -3
Public Const DELETE_LINK = -4

'=============
' Data Types
'=============
Type T_Link
    Destination As String
    SourceType As Integer
    Source As String
End Type

Type T_Value
    DataType As Integer
    Destination As String
    SourceType As Integer
    SourceName As String
    Value As String
End Type

'===================
' Help Context IDs
'===================
Public Const TABS_FIRST_HELPID = &H26201
Public Const IMAGE_BROWSER_HELPID = &H26210
Public Const OCR_BROWSER_HELPID = &H26211
Public Const ADOBE_ACROBAT_HELPID = &H26213

'==========================
' SSTab Control Constants
'==========================
Public Const FIRST_TAB = 0
Public Const INDEX_TAB = 0
Public Const DOCUMENT_TAB = 1
Public Const IMAGE_TAB = 2
Public Const ODBC_TAB = 3
Public Const LAST_TAB = 3

'===========================
' UpDown Control Constants
'===========================
Public Const DOWN_ONE = -1
Public Const UP_ONE = 1



