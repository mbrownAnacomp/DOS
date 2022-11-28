Attribute VB_Name = "modRelCommon"
Option Explicit

' Shared resource constants
Public Const CAP_BATCH_CLASS = 1
Public Const CAP_DOC_CLASS = 2
Public Const CAP_LOADING = 6
Public Const CAP_INDEXVALUE = 12
Public Const CAP_IMGFRAME = 13
Public Const CAP_SKIPFIRST = 14
Public Const CAP_IMGTYPE = 16
Public Const CAP_RELDIRNAME = 17
Public Const CAP_OCRFRAME = 20
Public Const CAP_OCRRELDIR = 21
Public Const CAP_DIRECTORY = 28
Public Const CAP_DRIVE = 29
Public Const CAP_RELEASEFILESAS = 30
Public Const CAP_ALLFILES = 31
Public Const CAP_PDFFRAME = 34
Public Const CAP_PDFRELDIR = 35
Public Const CAP_PDFFORMAT_JPEG = 41
Public Const CAP_PDFFORMAT_MTIFF = 42
Public Const CAP_PDFFORMAT_PCX = 43
Public Const CAP_PDFFORMAT_TIFF = 44
Public Const CAP_FRA_ADV_PDF_SETTINGS = 80
Public Const CAP_FRA_REL_RELATED = 82
Public Const CAP_CHK_WAIT_STATUS = 88
Public Const CAP_CHK_DEL_HUNG = 90
Public Const CAP_NAME = 91
Public Const CAP_RELEASE_IMAGE_FILES = 92
Public Const CAP_RELEASE_FULL_TEXT = 93
Public Const CAP_RELEASE_KOFAX_PDF_FILES = 94
                        
Public Const MSG_ANERROR = 7001
Public Const MSG_HAPPENEDLOGGING = 7002
Public Const MSG_INTERNALERROR = 7003
Public Const MSG_NOVALIDATIONFAILURE = 7004
Public Const MSG_MISSINGIMAGETYPE = 7005
Public Const MSG_FATALERROR = 7006

' -- Command Buttons --
Public Const CMD_BROWSE = 50
Public Const CMD_APPLY = 51
Public Const CMD_OK = 52
Public Const CMD_CANCEL = 53
Public Const CMD_HELP = 54
Public Const CMD_BROWSE2 = 55
Public Const CMD_BROWSE3 = 59
Public Const CMD_SETTINGS = 60

' -- Popup Links Menu --
Public Const MNU_CURRLINK = 70
Public Const MNU_DOCUMENTID = 72
Public Const MNU_INDEXFIELDS = 73
Public Const MNU_BATCHFIELDS = 74
Public Const MNU_ASCENTVALUES = 75
Public Const MNU_TEXTCONST = 76
Public Const MNU_UNLINK = 77

' -- Message Box Titles --
Public Const TITLE_DATAVERIFYFAIL = 958
Public Const TITLE_SAVESETTINGS = 959
Public Const TITLE_DIRNOTEXIST = 960
Public Const TITLE_QUEUE = 961
Public Const TITLE_DATAVERIFY = 962

' -- Error Messages --
Public Const MSG_DISCARDTEXTCONST = 1002
Public Const MSG_BADDESTINATION = 1003
Public Const MSG_FAILEDTOLOADFORM = 1005
Public Const MSG_UNCSHARENOTEXIST = 1018
Public Const MSG_SAVESETTINGS = 1019
Public Const MSG_BADDRIVELETTER = 1020
Public Const MSG_ASKTOCREATEDIR = 1021
Public Const MSG_DRIVENOTAVAILABLE = 1022
Public Const MSG_NOTALLINDEXUSED = 1024
Public Const MSG_NOTALLBATCHUSED = 1025
Public Const MSG_NOIMAGEDIRECTORY = 1026
Public Const MSG_BADIMAGEDIRECTORY = 1027
Public Const MSG_BADOCRDIRECTORY = 1028
Public Const MSG_NOOCRDIRECTORY = 1029
Public Const MSG_OCRFILESDISCARDED = 1030
Public Const MSG_USEDEFAULT = 1031
Public Const MSG_ANDMORE = 1032
Public Const MSG_CANNOTFIND = 1033
Public Const MSG_FILE = 1034
Public Const MSG_HELP = 1035
Public Const MSG_HELPFILE = 1036
Public Const MSG_HELPFILES = 1037
Public Const MSG_FINDYOURSELF = 1038
Public Const MSG_HELPFILENOTEXIST = 1039
Public Const MSG_NEEDTOREINSTALL = 1040
Public Const MSG_INVALIDHELPFILE = 1041
Public Const MSG_BADKFXPDFDIRECTORY = 1042
Public Const MSG_NOKFXPDFDIRECTORY = 1043
Public Const MSG_PDFSETUPFAILED1 = 2000
Public Const MSG_PDFSETUPFAILED2 = 2001
Public Const MSG_PDFPUBLISHFAILED1 = 2002
Public Const MSG_PDFPUBLISHFAILED2 = 2003
Public Const MSG_UNCOMPRESSEDTIFF1 = 2004
Public Const MSG_UNCOMPRESSEDTIFF2 = 2005
Public Const MSG_UNCOMPRESSEDTIFF3 = 2006
Public Const MSG_UNCOMPRESSEDTIFF4 = 2007
Public Const MSG_UNCOMPRESSEDTIFF5 = 2008

'===========================
'DoS Index constants
'===========================
Public Const DOC_TYPE = "DocType"
Public Const SSN = "SSN"
Public Const EFF_DATE = "EffectiveDate"
Public Const NOACODE_1 = "NOACode1"
Public Const NOACODE_2 = "NOACode2"
Public Const DOC_TYPE_C = "DocTypeC"
Public Const SIDE = "Side"
Public Const EXCEPTION = "Exception"

' Index file name addon
Public indexFileName As String
Public firstDoc As Boolean
Public imageCount As Integer
Public docCount As Integer
Public totalImageCount As Integer
Public purgeImageCount As Integer

'For persisted Side filed
Public sSide As String

'for multi function db lookup
Public db As Database, rs As Recordset
'To set batch status and store document and page counts in Manifesting db
Public dbBatch As Database, rsBatch As Recordset
'For storage of rs for later purge logic
Public dbNew As Database, rsNew As Recordset
Public tdfNew As TableDef
'For renaming ssn database at script close
Public newDbLocation As String 'will modify at script close
