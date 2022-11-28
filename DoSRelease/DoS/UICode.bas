Attribute VB_Name = "UtilCode"
Option Explicit

Public oError As New TextError  ' Used to record all errors to the Capture error log

'============================================
' Menu Manipulation Functions and Constants
'============================================
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&
Private Const MF_BYPOSITION = &H400&

'==========================================
' Folder Browsing Functions and Constants
'==========================================
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFLags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

'*************************************************
' AddMenuColumns
'-------------------------------------------------
' Purpose:  This subroutinue breaks all submenus
'           off the main menu into columns.  The
'           OS used to do this for us but that
'           changed so we have to do it now.
' Inputs:   oForm   the form object that has the
'                   menu on it
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Sub AddMenuColumns(oForm As Form)
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
    nNum = Screen.Height / (oForm.TextHeight("W") + 65)
    
    ' Get the main menu handle
    hmnuMainHandle = GetMenu(oForm.hwnd)
    hmnuMainHandle = GetSubMenu(hmnuMainHandle, 0)
    nMainMenuCount = GetMenuItemCount(hmnuMainHandle)
    
    For J = 0 To nMainMenuCount - 1
        ' Get the submenu handle
        hmnuSubHandle = GetSubMenu(hmnuMainHandle, J)
        If (hmnuSubHandle > 0) Then
            ' Find out how many items are in the menu
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
                sSubMenuCaption = Left(sSubMenuCaption, nSubMenuCaptionLen)
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
' BrowseFolders
'-------------------------------------------------
' Purpose:  Displays a dialog with a tree view of
'           file system folders from which a user
'           may select a system folder.
' Inputs:   hWndParent  calling form
'           sPrompt     instructions to display
'                       above the tree view
'           ulFLags     flags to control the info
'                       displayed and/or returned
' Outputs:  None
' Returns:  The path of the selected directory
' Notes:    None
'*************************************************
Public Function BrowseFolders(hWndParent As Long, sPrompt As String, ulFLags As Long) As String
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo

    With tBrowseInfo
        .hWndOwner = hWndParent
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFLags = ulFLags
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        BrowseFolders = sBuffer
    Else
        BrowseFolders = ""
    End If

End Function

'*************************************************
' CreateDirectory
'-------------------------------------------------
' Purpose:  This routine will create the directory
'           path passed in.
' Inputs:   DirName     directory path to create
' Outputs:  None
' Returns:  True/False
' Notes:    This function is recursive.
'*************************************************
Function CreateDirectory(DirName As String) As Boolean
        Dim ParentDir As String
        Dim Results As Integer
        Dim sName As String
        
        ' Step 1 -
        ' Create the parent directory.  If we are working in the
        ' root directory, causes immediate failure.  We shouldn't
        ' be trying to create a new root directory ...
        ParentDir = RemoveLastPathSegment(DirName)
        If ParentDir = "" Then GoTo CantMakeDir
        
        ' Step 2 -
        ' Make sure that the parent directory is valid before
        ' attempting to create the current directory.  If
        ' the parent is not valid, attempt to create it.
        On Error GoTo DirDoesNotExist
        If (Left(ParentDir, 2) = "\\") Then
            sName = Dir(ParentDir, vbDirectory)
            If (sName = "") Then
                Call Err.Raise(76, "", "ERROR")
            End If
        Else
            ChDir ParentDir
        End If
        
        ' Step 3 -
        ' Attempt to create the directory.  If a failure
        ' occurs here it must be propogated back up the line.
        On Error GoTo CantMakeDir
        MkDir DirName
        
        CreateDirectory = True
        Exit Function
        
'----------------
' Error Handler
'----------------
DirDoesNotExist:
        ' If we get here we failed while attempting
        ' to change directories.  If we failed with
        ' the 'Path not found' error, then try to
        ' recursivly create the parent directory.
        If Err = 76 Then ' 76 = Path not found
            If CreateDirectory(ParentDir) Then Resume Next
        End If
        
        ' If we failed to make the directory, fall
        ' through into CantMakeDirectory
        
CantMakeDir:
        ' If we get here, we either hit an error that
        ' we couldn't handle or the recursive directory
        ' creation failed.
        CreateDirectory = False
        Exit Function

End Function

'*************************************************
' RemoveLastPathSegment
'-------------------------------------------------
' Purpose:  This function removes the last path
'           segment from the passed in directory.
'           If the passed in directory is just the
'           drive letter, an empty string will be
'           returned.
' Inputs:   DirName     directory path to modify
' Outputs:  None
' Returns:  The directory minus the last path seg
' Notes:    None
'*************************************************
Function RemoveLastPathSegment(ByVal DirName As String) As String
    Dim CurChar As Integer  ' The location of the last '\' character in the path
    Dim TmpChar As Integer

        ' Trim the passed in string to remove all white space
        DirName = Trim$(DirName)
        
        ' If there is a trailing '\' character, remove it
        ' before we chop off the last directory entry.
        If right$(DirName, 1) = "\" Then
            DirName = Left$(DirName, Len(DirName) - 1)
        End If
        
        ' Find the last "\" character in the path
        Do
            CurChar = TmpChar
            TmpChar = InStr(CurChar + 1, DirName, "\")
        Loop Until TmpChar = 0
        
        ' Handle a UNC
        If (Left(DirName, 2) = "\\") Then
            If (CurChar = InStr(3, DirName, "\")) Then
                CurChar = 0
            End If
        End If
        
        ' If we found at least one '\' in the path, then
        ' pass back the path up to that last '\'.  Otherwise
        ' pass back an empty string.
        If CurChar = 0 Then
            RemoveLastPathSegment = ""
        Else
            RemoveLastPathSegment = Left$(DirName, CurChar)
        End If
        
End Function

'*************************************************
' VerifyDirectoryName
'-------------------------------------------------
' Purpose:  Check the passed in directory, and if
'           it does not exist, allow the user the
'           option to create it.
' Inputs:   DirName     directory name to verify
' Outputs:  None
' Returns:  True/False if vlaid directory name
' Notes:    None
'*************************************************
Function VerifyDirectoryName(ByVal DirName As String) As Boolean
        Dim CurrentDirectory As String
        Dim Results As Integer
        Dim lError As Long

        VerifyDirectoryName = False
        
        If (DirName = "") Then
            Exit Function
        End If
        
        ' Windows 98 allows the path to include a space
        ' after the drive letter but before the directory.
        ' This appends the directory to the current working
        ' directory of the specified drive letter.  We don't
        ' allow this for consistency with Window 95 and NT.
        If (Mid$(DirName, 2, 2) = ": ") Then
            Exit Function
        End If
        
        ' Handle a UNC
        If (Left$(DirName, 2) = "\\") Then
            Dim sName As String

            ' The UNC path has to end with a \ at the end
            ' before passing it to Dir function in order to
            ' work properly in certain cases.
            ' This seems to occur only in Windows 2000.
            If (Right$(DirName, 1) <> "\") Then
                DirName = DirName & "\"
            End If
            
            ' This chunk of code handles bad share names for UNC.
            ' Ex: "\\server\share2\" where share2 does not exist.
            On Error Resume Next
            sName = Dir(DirName, vbDirectory)
            If Err.Number <> 0 Then
                lError = Err.Number
                On Error GoTo 0
                Err.Raise MSG_UNCSHARENOTEXIST, "UtilCode.VerifyDirectoryName", LoadResString(MSG_UNCSHARENOTEXIST)
            End If
            On Error GoTo 0
            
            If (sName = "") Then GoTo TryToCreate
            VerifyDirectoryName = True
            Exit Function
        End If
        
        ' Save the current directory and try to
        ' change directory to the one passed in.
        On Error GoTo InvalidDriveLetter
        CurrentDirectory = CurDir(DirName)
        
        ' Try to change directories to the passed in value.
        On Error GoTo InvalidDirectory
        ChDir DirName
        
        ' If we get here, the directory exists
        VerifyDirectoryName = True
        
ExitVerifyDirectoryName:
        ' Set the directory back to its value before we started.
        On Error GoTo InvalidDriveLetter
        
        If (CurrentDirectory <> "") Then
            ChDir CurrentDirectory
        End If
        
        Exit Function
        
'----------------
' Error Handlers
'----------------
InvalidDriveLetter:
        MsgBox LoadResString(MSG_BADDRIVELETTER), vbOKOnly + vbExclamation, LoadResString(TITLE_DIRNOTEXIST)
        VerifyDirectoryName = False
        Exit Function
        
InvalidDirectory:
        ' If we get here, the passed in directory does not
        ' yet exist.  Give the user the option to create
        ' the directory, but only if the error message is
        ' for an non-existent directory name.
        If Err = 76 Then    ' Error 76 = Path not found
            GoTo TryToCreate
        ElseIf Err <> 0 Then
            Resume ExitVerifyDirectoryName
        Else
            GoTo ExitVerifyDirectoryName
        End If

        
TryToCreate:
        Results = MsgBox(DirName & vbCr & LoadResString(MSG_ASKTOCREATEDIR), vbYesNo + vbQuestion, LoadResString(TITLE_DIRNOTEXIST))
        If Results = vbYes Then
            VerifyDirectoryName = CreateDirectory(DirName)
        End If
        If Err.Number = 0 Then
            GoTo ExitVerifyDirectoryName
        Else
            Resume ExitVerifyDirectoryName
        End If

End Function

'*************************************************
' VerifyFileName
'-------------------------------------------------
' Purpose:  Validates if a filename was provided
'           with a ".txt" extenstion and checks
'           the directory and prompts the user
'           to create it if it doesn't exist.
' Inputs:   FullPath    the full path of the file
'                       with drive and directory
' Outputs:  None
' Returns:  True/False
' Notes:    The file is not created if it doesn't
'           already exist.  It will be created at
'           Release time.
'*************************************************
Public Function VerifyFileName(ByVal FullPath As String) As Boolean
    Dim File As String
    Dim Path As String
    
        ' Default to not verified
        VerifyFileName = False
        
        ' Check if the directory exists
        ' and prompt to create it if not.
2010    Path = RemoveLastPathSegment(FullPath)
2020    If VerifyDirectoryName(Path) = False Then
            Exit Function
        End If
        
        ' Create/open the file (only real way to see if filename is valid)
        VerifyFileName = Touch(FullPath)
        
End Function

'*************************************************
' StripAmpersands
'-------------------------------------------------
' Purpose:  Removes all ampersands from a string
'           passed by the caller.
' Inputs:   sString     the string to purge of
'                       ampersands
' Outputs:  None
' Returns:  The string without ampersands
' Notes:    Some of the popup linking menu items
'           contain ampersands as accelerator
'           keys.  When we use the menu caption as
'           an Index Value, we need to strip out
'           the ampersand.
'*************************************************
Function StripAmpersands(ByVal sString As String) As String
  Dim I As Integer
  Dim sTmp As String
  Dim sChar As String
  
  For I = 1 To Len(sString)
    sChar = Mid$(sString, I, 1)
    If sChar <> "&" Then
      sTmp = sTmp & sChar
    End If
  Next I
    
  StripAmpersands = sTmp

End Function

'********************************************************************************
' Routine:  Touch
' Purpose:  Verify a filename is valid
' Input     strFilePath - Full path of filename to be created
' Returns:  True if the file was opened successfully
'********************************************************************************
Private Function Touch(ByVal strFilePath As String) As Boolean
    On Error GoTo trap
    
    Dim hFile As Integer
    
    hFile = FreeFile
    Open strFilePath For Append As hFile
    Close hFile
    
    Touch = True
    
    Exit Function
    
trap:
    Close hFile
End Function
