VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KfxReleaseSetupScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ReleaseSetupData object is set by the Release
' Setup Controller.  This object is used during
' the document type setup process.  It will contain
' all of the information and interfaces you need to
' define a document type's release process.
Public SetupData As ReleaseSetupData

Private Const M_RSETUP = "Text Release Setup Script"

'*************************************************
' CloseScript
'-------------------------------------------------
' Purpose:  Script release point.  Perform any
'           necessary cleanup such as releasing
'           resources, etc.
' Inputs:   None
' Outputs:  None
' Returns:  One of the following:
'             KFX_REL_SUCCESS, KFX_REL_ERROR,
'             KFX_REL_FATALERROR, KFX_REL_REINIT
'             KFX_REL_DOCCLASSERROR,
' Notes:    Called by Release Setup Controller
'           once just before the script object
'           is released.
'*************************************************
Public Function CloseScript() As KfxReturnValue

100     Set oError.DataObject = Nothing

        CloseScript = KFX_REL_SUCCESS

End Function

'*************************************************
' ActionEvent
'-------------------------------------------------
' Purpose:  This method allows the setup script
'           to respond to various events in the
'           Administration module.  The script
'           has the opportunity to make any
'           necessary changes to the release
'           settings in the ReleaseSetupData
'           object or any other external data
'           source.
' Inputs:   intActionID     ID of the event
'           strData1        Action parameter 1
'           strData2        Action parameter 2
' Outputs:  None
' Returns:  One of the following:
'             KFX_REL_SUCCESS, KFX_REL_ERROR,
'             or KFX_REL_UNSUPPORTED
' Notes:    Refer to the documentation for a list
'           of actions and associated parameters.
'*************************************************
Public Function ActionEvent(intActionID As KfxActionValue, _
                            strData1 As String, _
                            strData2 As String) As KfxReturnValue
    Dim strMsgHeader As String
    Dim nMaxLen As Integer
    Dim nBatchLen As Integer
    Dim nDocLen As Integer
    Dim nNameLen As Integer
    Dim oLink As AscentRelease.Link
    Static bShowUI As Boolean
    
        On Error GoTo Err_ActionEvent
    
        ActionEvent = KFX_REL_SUCCESS

        If (Not SetupData.New) Then
            Select Case intActionID
                
                Case KFX_REL_START
                    ' A new series of Action Events is
                    ' starting so initialize any flags
190                 bShowUI = False
                
                Case KFX_REL_INDEXFIELD_DELETE
                    ' Delete any links to this Index Field
200                 ActionEvent = RemoveTheLink(strData1, SetupData, KFX_REL_INDEXFIELD)
                
                Case KFX_REL_INDEXFIELD_INSERT
                    ' Set the flag to display the UI so the user
                    ' can create a link to the new Index Field
210                 bShowUI = True
                
                Case KFX_REL_INDEXFIELD_RENAME
                    ' If this Index Field is used in a link,
                    ' change the link to reflect the new name.
220                 For Each oLink In SetupData.Links
                        If (oLink.Source = strData1 And oLink.SourceType = KFX_REL_INDEXFIELD) Then
                            oLink.Source = strData2
                        End If
                    Next
                
                Case KFX_REL_BATCHFIELD_DELETE
                    ' Delete any links to this Batch Field
230                 ActionEvent = RemoveTheLink(strData1, SetupData, KFX_REL_BATCHFIELD)
                
                Case KFX_REL_BATCHFIELD_INSERT
                    ' Set the flag to display the UI so the user
                    ' can create a link to the new Batch Field
240                 bShowUI = True
                
                Case KFX_REL_BATCHFIELD_RENAME
                    ' If this Batch Field is used in a link,
                    ' change the link to reflect the new name.
250                 For Each oLink In SetupData.Links
                        If (oLink.Source = strData1 And oLink.SourceType = KFX_REL_BATCHFIELD) Then
                            oLink.Source = strData2
                        End If
                    Next
                
                Case KFX_REL_RELEASESETUP_DELETE
                    ' Nothing to do
                    
                Case KFX_REL_IMPORT
260                 ActionEvent = RunUI

                Case KFX_REL_BATCHCLASS_RENAME
                    ' Nothing to do
                    
                Case KFX_REL_DOCCLASS_RENAME
                    ' Nothing to do
                    
                Case KFX_REL_PUBLISH_CHECK
                
                    ' Calculate string length for padding purpose
                    nBatchLen = Len(LoadResString(CAP_BATCH_CLASS))
                    nDocLen = Len(LoadResString(CAP_DOC_CLASS))
                    nNameLen = Len(LoadResString(CAP_NAME))
                    nMaxLen = nBatchLen
                    nMaxLen = IIf(nDocLen > nMaxLen, nDocLen, nMaxLen)
                    nMaxLen = IIf(nNameLen > nMaxLen, nNameLen, nMaxLen)
                
                    ' Pad each string to the max length with spaces and append a tab character since padding
                    ' with spaces don't quite line up.  This ensures that all strings have equal length, and
                    ' works with localized strings.
                    strMsgHeader = LoadResString(CAP_BATCH_CLASS) & Space(nMaxLen - nBatchLen) & vbTab & SetupData.BatchClassName & vbCrLf & _
                        LoadResString(CAP_DOC_CLASS) & Space(nMaxLen - nDocLen) & vbTab & SetupData.DocClassName & vbCrLf & _
                        StripAmpersands(LoadResString(CAP_NAME)) & Space(nMaxLen - nNameLen) & vbTab & SetupData.Name & vbCrLf & vbCrLf

                    Select Case SetupData.ImageType
                        Case CAP_FORMAT_PDF
                            ' If PDF 2.01 is detected, return error.
                            Call MsgBox(strMsgHeader & LoadResString(MSG_PDFPUBLISHFAILED1) & vbCrLf & _
                                LoadResString(MSG_PDFPUBLISHFAILED2), _
                                vbOKOnly + vbExclamation, LoadResString(TITLE_DATAVERIFYFAIL))
                            ActionEvent = KFX_REL_ERROR
                            Exit Function
                        
                        Case CAP_FORMAT_MTIFF_RAW, CAP_FORMAT_TIFF_RAW
                            ' Give warning about 'Multipage TIFF - Uncompressed'
                            ' and 'TIFF - Uncompressed" formats which might not
                            ' work properly.
                            Call MsgBox(strMsgHeader & LoadResString(MSG_UNCOMPRESSEDTIFF1) & vbCrLf & _
                                LoadResString(MSG_UNCOMPRESSEDTIFF2) & vbCrLf & _
                                LoadResString(MSG_UNCOMPRESSEDTIFF3) & vbCrLf & _
                                LoadResString(MSG_UNCOMPRESSEDTIFF4) & vbCrLf & _
                                LoadResString(MSG_UNCOMPRESSEDTIFF5), _
                                vbOKOnly + vbQuestion, LoadResString(TITLE_DATAVERIFYFAIL))
                            ActionEvent = KFX_REL_SUCCESS
                            Exit Function
                    End Select
                    
                Case KFX_REL_UPGRADE
                    ' The default release scripts do not
                    ' support the UPGRADE action at this time
                    ActionEvent = KFX_REL_UNSUPPORTED
                    
                Case KFX_REL_END
                    ' Check if the flag was set to display the
                    ' UI by any Action Events in the series.
                    If bShowUI = True Then
                        bShowUI = False
                        ActionEvent = RunUI
                    End If
                    
                Case Else
                    ActionEvent = KFX_REL_UNSUPPORTED
                    
            End Select
        
            'Save our changes.
270         Call SetupData.Apply
        End If
    
Exit_ActionEvent:
    
        Exit Function

Err_ActionEvent:
    
        ActionEvent = KFX_REL_ERROR
        SetupData.LogError Err, 0, 0, Err.Description, M_RSETUP, Erl
        Resume Exit_ActionEvent

End Function

'*************************************************
' OpenScript
'-------------------------------------------------
' Purpose:  Script initialization point.  Perform
'           any necessary initialization such as
'           logging in to a remote data source,
'           allocating resources, etc.
' Inputs:   None
' Outputs:  None
' Returns:  One of the following:
'             KFX_REL_SUCCESS, KFX_REL_ERROR,
'             KFX_REL_FATALERROR, KFX_REL_REINIT
'             KFX_REL_DOCCLASSERROR,
' Notes:    Called by the Release Controller
'           once when the script object is loaded
'           and before a call to RunUI() or
'           ActionEvent() is made.
'*************************************************
Public Function OpenScript() As KfxReturnValue
            
800     Set oError.DataObject = SetupData
810     oError.TitleString = LoadResString(TITLE_RSETUPERROR)
    
        OpenScript = KFX_REL_SUCCESS

End Function

'*************************************************
' RunUI
'-------------------------------------------------
' Purpose:  User interface display point.  This
'           method is called by the Release Setup
'           Controller to display the setup form
'           specific to this script.
' Inputs:   None
' Outputs:  None
' Returns:  One of the following:
'             KFX_REL_SUCCESS, KFX_REL_ERROR,
'             or KFX_REL_STOPPED
' Notes:    Called by Release Setup Controller
'           when the Administration module asks
'           to run the script and whenever a
'           Batch Field or Index Field is inserted.
'*************************************************
Public Function RunUI() As KfxReturnValue
    Dim bResult As Boolean

        On Error GoTo Err_RunUI

1000    Load frmSetup
1010    bResult = frmSetup.ShowForm(SetupData)
        If (bResult) Then
            RunUI = KFX_REL_SUCCESS
        Else
            RunUI = KFX_REL_STOPPED
        End If

Exit_RunUI:
    
        Unload frmSetup
        Exit Function

Err_RunUI:
        RunUI = KFX_REL_ERROR
        Call SetupData.LogError(Err, 0, 0, Err.Description, M_RSETUP, Erl)
    
        Resume Exit_RunUI

End Function

'*************************************************
' RemoveTheLink
'-------------------------------------------------
' Purpose:  Deletes links from the release script
'           when the associated Index Field or
'           Batch Field is removed from the
'           Document Class or Batch Class
'           respectively in the Administration
'           module.
' Returns:  KFX_REL_SUCCESS or KFX_REL_ERROR
' Inputs:   strData1    the link source name
'           SetupData   the SetupData object
'           nLinkType   the link source type
' Notes:    Called by ActionEvent() when the
'           action specified is either
'           KFX_REL_INDEXFIELD_DELETE or
'           KFX_REL_BATCHFIELD_DELETE.
'
'           *** IMPORTANT ***
'           The Text Release script stores the
'           sequence in the Destination member
'           of each link.  Deleting a link means
'           that each subsequent link must be
'           modified to indicate its new sequence.
'
'           Multiple links may be deleted at once
'           if a Batch Field or Index Field is
'           used multiple times.  To resequence
'           the links that we wish to save, we
'           copy them to a temporary array,
'           delete the entire collection, then
'           re-add the saved links to the
'           collection with their new sequence.
'*************************************************
Private Function RemoveTheLink(strData1 As String, SetupData As ReleaseSetupData, nLinkType As KfxLinkSourceType) As KfxReturnValue
    Dim oLink As AscentRelease.Link
    Dim SavedLinks() As T_Link
    Dim nCount As Integer
    Dim I As Integer
    Dim nSequence As Integer
    Dim nNonPDFLinksCount As Integer
    Dim jj As Integer
        
        On Error GoTo Err_RemoveTheLink
        
        RemoveTheLink = KFX_REL_SUCCESS
        nNonPDFLinksCount = 0
        
        ' Resize the temporary array of saved links
        nCount = SetupData.Links.Count
        If nCount > 0 Then
2000        ReDim SavedLinks(nCount)
        Else
            Exit Function
        End If
        
        ' Loop through all of the links in the collection and determine which ones will be saved
        ' If set to NO_LINK, add it to output list, but later it will be dropped
        For Each oLink In SetupData.Links
            If Left$(oLink.Destination, 3) <> "PDF" Then
                nNonPDFLinksCount = nNonPDFLinksCount + 1
                I = Val(oLink.Destination)
                If (oLink.Source = strData1 And oLink.SourceType = nLinkType) Then
                    ' We are deleting this link
2010                SavedLinks(I).Destination = oLink.Destination
2020                SavedLinks(I).SourceType = NO_LINK
2030                SavedLinks(I).Source = ""
                Else
                    ' We are saving this link
2050                SavedLinks(I).Destination = oLink.Destination
2060                SavedLinks(I).SourceType = oLink.SourceType
2070                SavedLinks(I).Source = oLink.Source
                End If
            End If
        Next oLink

        ' Begin adding PDF indexes after last non-PDF index
        jj = nNonPDFLinksCount - 1

        ' Do the same thing for PDF links
        For Each oLink In SetupData.Links
            If Left$(oLink.Destination, 3) = "PDF" Then
                jj = jj + 1
                If (oLink.Source = strData1 And oLink.SourceType = nLinkType) Then
                    ' Can't delete PDF links, so make them UNDEFINED.
                    SavedLinks(jj).Destination = oLink.Destination
                    SavedLinks(jj).SourceType = KFX_REL_UNDEFINED_LINK
                    SavedLinks(jj).Source = ""
                Else
                    SavedLinks(jj).Destination = oLink.Destination
                    SavedLinks(jj).SourceType = oLink.SourceType
                    SavedLinks(jj).Source = oLink.Source
                End If
            End If
        Next oLink

        ' Delete the entire collection
2080    Call SetupData.Links.RemoveAll

        ' Add all the links back into the collection
        nSequence = 0
        For I = 0 To nCount - 1
            If SavedLinks(I).SourceType <> NO_LINK Then
                ' Destination for PDF links is a name like "PDF Author Header Link"
                If Left$(SavedLinks(I).Destination, 3) = "PDF" Then
2090                Call SetupData.Links.Add(SavedLinks(I).Source, _
                        SavedLinks(I).SourceType, _
                        SavedLinks(I).Destination)
                Else
                    ' Destination for non-PDF links is a number (order matters)
                    Call SetupData.Links.Add(SavedLinks(I).Source, _
                        SavedLinks(I).SourceType, _
                        Format(nSequence))
                        nSequence = nSequence + 1
                End If
            End If
        Next

Exit_RemoveTheLink:
        Exit Function

Err_RemoveTheLink:
        SetupData.LogError Err, 0, 0, Err.Description, M_RSETUP, Erl
        RemoveTheLink = KFX_REL_ERROR
        Resume Exit_RemoveTheLink

End Function
