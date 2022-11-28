VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KfxReleaseScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ReleaseData object is set by the release controller.
' This object is to be used during the document release
' process as it will contain the document data and the
' external data source information defined during the
' setup process.
Public DocumentData As ReleaseData
Public sReleaseDirectory As String

Private oTextFile As New ASCIITextFile

Private Const M_RELEASE = "DoS Release with XML indexing"
Private bOpen As Boolean

Public strOutPdf As String

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
' Notes:    Called by the Release Controller
'           once just before the script object
'           is released.
'*************************************************
Public Function CloseScript() As KfxReturnValue
        'Dim hFileO As Integer
        'Dim outLog As String
        Dim doneNewDb, sQuery As String
        Dim noFindLoc As Integer
        Dim sBatchName As String
        On Error GoTo Err_CloseScript
        
        ' Assume success for now
        CloseScript = KFX_REL_SUCCESS
        'rs.Close
        rsNew.Close
        dbNew.Close
        'get rid of "tmp" and hence make dB available to external process
        noFindLoc = InStr(newDbLocation, "nofindtmp")
        doneNewDb = left(newDbLocation, noFindLoc - 1)
        doneNewDb = doneNewDb & Mid(newDbLocation, noFindLoc + 9)
        If Dir(doneNewDb) <> "" Then Kill doneNewDb 'no duplicates allowed
        Name newDbLocation As doneNewDb
        'Update batch database
        sBatchName = Trim(DocumentData.BatchName)
        'Check to see if this batch has a valid name
        sQuery = "Select * From Batch Where BatchName ='" & sBatchName & "'"
        Set rsBatch = db.OpenRecordset(sQuery, dbOpenDynaset, 512) '512 = dbSeeChanges
        If rsBatch.RecordCount > 0 Then
            sQuery = "UPDATE Batch Set BatchStatus = 5, ScanDateTime ='" & Now() & "', Pages =" & imageCount _
            & ", Documents =" & docCount & " Where BatchName ='" & sBatchName & "'"
            db.Execute sQuery, 64
            'rsBatch.Edit
            'rsBatch.Fields("BatchStatus") = 5 'Completed Release
            'rsBatch.Fields("ReleaseDateTime") = Now()
            'rsBatch.Fields("Pages") = imageCount
            'rsBatch.Fields("Documents") = docCount
            'rsBatch.Update
        Else 'Give them somthing
            sQuery = "INSERT INTO Batch (BatchStatus, ReleaseDateTime, Pages, Documents) "
            sQuery = sQuery & "VALUES(5, '" & Now() & "', " & imageCount & ", " & docCount & ")"
            db.Execute sQuery, 64
            'With rsBatch
            '    .AddNew
            '        ![BatchName] = sBatchName
            '        ![BatchStatus] = 99
            '        ![ReleaseDateTime] = Now()
            '        ![Pages] = imageCount
            '        ![Documents] = docCount
            '    .Update
            'End With
        End If
        rsBatch.Close
        db.Close
        If (Not bOpen) Then
            GoTo Exit_CloseScript
        End If
Exit_CloseScript:
100     Set oError.DataObject = Nothing
        bOpen = False

        Exit Function

Err_CloseScript:
        Call oError.LogTheError(Err, Err.Description, M_RELEASE + ":" + Err.Source, Erl, False, False)
        CloseScript = KFX_REL_ERROR
        Resume Exit_CloseScript

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
'           once when the script object is loaded.
'*************************************************
Public Function OpenScript() As KfxReturnValue
        Dim sDSN, sDB, sUID, sPWD, sWinAuth, sQuery, sConnect As String
        On Error GoTo Err_OpenScript
        'Get values for ODBC connection stored during setup
        sDSN = DocumentData.CustomProperties("ODBC DSN").Value
        sDB = DocumentData.CustomProperties("ODBC DB").Value
        sUID = DocumentData.CustomProperties("ODBC UID").Value
        sPWD = DocumentData.CustomProperties("ODBC PWD").Value
        sWinAuth = DocumentData.CustomProperties("WinAuth").Value
        'Open DoS Forms List database for various fucntions
        sConnect = "ODBC;DSN=" & sDSN & ";DATABASE=" & sDB
        If sWinAuth = "False" Then sConnect = sConnect & ";UID=" & sUID & ";PWD=" & sPWD
        Set db = OpenDatabase("", False, False, sConnect)
        Dim fs As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
        If Not fs.FolderExists(DocumentData.KofaxPDFPath & "\Databases") Then
            fs.CreateFolder DocumentData.KofaxPDFPath & "\Databases"
        End If
        Set fs = Nothing
        'Create new database for this batch only - used for purge logic
        'prefix with "tmp" to prevent premature opening my external process
        newDbLocation = DocumentData.KofaxPDFPath & "\Databases\" & "nofindtmp" & Trim(DocumentData.BatchName) & ".mdb"
        If Dir(newDbLocation) <> "" Then Kill newDbLocation
        Set dbNew = CreateDatabase(newDbLocation, dbLangGeneral)
        Set tdfNew = dbNew.CreateTableDef("FolderList")
        With tdfNew
            ' Create fields and append them to the new TableDef object.
            .Fields.Append .CreateField("SSN", dbText)
            .Fields.Append .CreateField("FormNumber", dbText)
            .Fields.Append .CreateField("OriginalFormNumber", dbText)
            .Fields.Append .CreateField("EffDate", dbText)
            .Fields.Append .CreateField("NOA1", dbText)
            .Fields.Append .CreateField("NOA2", dbText)
            .Fields.Append .CreateField("SourceSide", dbText)
            .Fields.Append .CreateField("FormType", dbText)
            .Fields.Append .CreateField("DeliverySide", dbText)
            .Fields.Append .CreateField("Duplex", dbBoolean)
            .Fields.Append .CreateField("Purge", dbBoolean)
            .Fields.Append .CreateField("Pages", dbText)
            .Fields.Append .CreateField("PDFFileName", dbText)
        End With
        dbNew.TableDefs.Append tdfNew
        Set rsNew = dbNew.OpenRecordset("FolderList", dbOpenTable)

200     Set oError.DataObject = DocumentData
        oError.TitleString = LoadResString(TITLE_RELEASEERROR)

        OpenScript = KFX_REL_SUCCESS
        bOpen = True
        firstDoc = True
        
Exit_OpenScript:
        Exit Function

Err_OpenScript:
        Call DocumentData.LogError(Err, 0, 0, Err.Description, M_RELEASE + ":" + Err.Source, Erl)
        OpenScript = KFX_REL_ERROR
        Resume Exit_OpenScript

End Function

'*************************************************
' ReleaseDoc
'-------------------------------------------------
' Purpose:  Document release point.  Use the
'           ReleaseData object to release the
'           current document's data to the
'           external data repository.
' Inputs:   None
' Outputs:  None
' Returns:  One of the following:
'             KFX_REL_SUCCESS, KFX_REL_ERROR,
'             KFX_REL_FATALERROR, KFX_REL_REINIT
'             KFX_REL_DOCCLASSERROR,
' Notes:    Called by the Release Controller once
'           for each document to be released.
'*************************************************
Public Function ReleaseDoc() As KfxReturnValue

        Dim bReleaseImageFiles As Boolean
        Dim bReleaseOCRFullText As Boolean
        Dim bReleasedToDirectory As Boolean
        Dim bReleaseKofaxPDF As Boolean
        bReleasedToDirectory = False

        Dim nextPDF As String
        Dim hFile1 As Integer
        Dim hFile2 As Integer
        Dim fChunk As String
        
        ' Don't release if OpenScript() hasn't been called first
        If (Not bOpen) Then
            ReleaseDoc = KFX_REL_ERROR
            GoTo Exit_ReleaseDoc
        End If

        oTextFile.ErrorLineNum = 0

        ' Release the images (operates differently if PDF selected)
        Dim oPdfRel As New clsCAPPdfRelease ' CAPTools PDF release object
        Dim strImageFilePath As String
        
        On Error GoTo Handler_SetReleaseImageFilesDefault
        bReleaseImageFiles = Not CBool(DocumentData.CustomProperties.Item(KEY_DISABLE_IMAGE_EXPORT).Value)
        
        On Error GoTo Err_ReleaseDoc
        
        If bReleaseImageFiles Then
            If oPdfRel.ReleaseEnabled(DocumentData) Then
                ' If PDF enabled, let it release the images.
                strImageFilePath = oPdfRel.ReleaseImages(DocumentData)
                
                ' If the document contains non image files, release them
                If DocumentData.ImageFiles.ContainsNonImageFile Then ReleaseNonImageFiles
            Else
                On Error GoTo Handler_ReleaseImages
                ' Let Ascent release the images.
                Call DocumentData.ImageFiles.Copy
                strImageFilePath = DocumentData.ImageFiles.ReleasedDirectory
                bReleasedToDirectory = DocumentData.ImageFiles.ReleasedToDirectory
            End If
        End If
        
        On Error GoTo Handler_SetReleaseOCRFullTextDefault
        bReleaseOCRFullText = Not CBool(DocumentData.CustomProperties.Item(KEY_DISABLE_TEXT_EXPORT).Value)
        
        On Error GoTo Err_ReleaseDoc
        
        ' Release the Full Text OCR file.
        If bReleaseOCRFullText Then
            If (DocumentData.TextFilePath <> "") Then
                Call DocumentData.TextFiles.Copy
            End If
        End If
        
        On Error GoTo Handler_SetReleaseKofaxPDFDefault
        bReleaseKofaxPDF = DocumentData.KofaxPDFReleaseScriptEnabled
        
        On Error GoTo Err_ReleaseDoc
        
        ' Release the Kofax PDF file.
        If bReleaseKofaxPDF Then
            If (DocumentData.KofaxPDFPath <> "") Then
                If firstDoc Then
                    Dim oFS As Object
                    Set oFS = CreateObject("Scripting.FileSystemObject")
                    If Not oFS.FolderExists(DocumentData.KofaxPDFPath & "\" & Trim(DocumentData.BatchName)) Then
                        oFS.CreateFolder (DocumentData.KofaxPDFPath & "\" & Trim(DocumentData.BatchName))
                        sReleaseDirectory = DocumentData.KofaxPDFPath & "\" & Trim(DocumentData.BatchName)
                    Else 'Added to deal with same batch being released twice
                        oFS.DeleteFile (DocumentData.KofaxPDFPath & "\" & Trim(DocumentData.BatchName) & "\*.*") 'delete all old files
                        sReleaseDirectory = DocumentData.KofaxPDFPath & "\" & Trim(DocumentData.BatchName)
                    End If
                End If
                Call DocumentData.CopyKofaxPDFFileToPath(sReleaseDirectory)
            End If

            ' If Kofax PDF is checked and the Ascent release image is not checked,
            ' use the Kofax PDF image path as the release image path.
            If (Not bReleaseImageFiles) And Len(strImageFilePath) = 0 Then
                ' Format the pdf file name
                Dim strPDFFileName As String
                strPDFFileName = Hex$(DocumentData.UniqueDocumentID)
                strPDFFileName = String$(8 - Len(strPDFFileName), "0") & strPDFFileName & ".pdf"
                If right(sReleaseDirectory, 1) <> "\" Then sReleaseDirectory = sReleaseDirectory & "\"
                strImageFilePath = sReleaseDirectory & strPDFFileName
            End If
        End If
        docCount = docCount + 1
        imageCount = imageCount + DocumentData.ImageFiles.Count
        ' Release the index information to the text file.
1040    'Modified to use PDF file name only in index (docHarbor request)
        'Call oTextFile.ReleaseIndexes(DocumentData, strImageFilePath, sReleaseDirectory)
        Call oTextFile.ReleaseIndexes(DocumentData, strPDFFileName, sReleaseDirectory)
 
        ReleaseDoc = KFX_REL_SUCCESS

Exit_ReleaseDoc:
        Exit Function
        
Handler_SetReleaseImageFilesDefault:
        bReleaseImageFiles = True
        Resume Next

Handler_SetReleaseOCRFullTextDefault:
        bReleaseOCRFullText = True
        Resume Next
        
Handler_SetReleaseKofaxPDFDefault:
        bReleaseKofaxPDF = True
        Resume Next
        
Handler_ReleaseImages:
        '*** If error occurs while copying the image file, the file or folder
        '*** has been created, but these variables have yet not been set.
        strImageFilePath = DocumentData.ImageFiles.ReleasedDirectory
        bReleasedToDirectory = DocumentData.ImageFiles.ReleasedToDirectory

Err_ReleaseDoc:
        Dim iLine As Integer
        Dim nErrNum As Long
        Dim sErrDesc As String
        
        '*** Remove Image file or directory
        If strImageFilePath <> "" Then
            Dim oFileSystemObject
            Set oFileSystemObject = CreateObject("Scripting.FileSystemObject")
            
            If bReleasedToDirectory Then
                If oFileSystemObject.FolderExists(strImageFilePath) Then
                    Call oFileSystemObject.DeleteFolder(strImageFilePath)
                End If
            Else
                If oFileSystemObject.FileExists(strImageFilePath) Then
                    Call oFileSystemObject.DeleteFile(strImageFilePath)
                End If
            End If
            Set oFileSystemObject = Nothing
        End If
        
        ReleaseDoc = KFX_REL_ERROR
        
        If (oTextFile.ErrorLineNum <> 0) Then
            iLine = oTextFile.ErrorLineNum
        Else
            iLine = Erl
        End If
        nErrNum = Err
        sErrDesc = Err.Description
        
        Call oError.LogTheError(nErrNum, sErrDesc, M_RELEASE + ":" + Err.Source, iLine, False, False)
                
        ' Mark the document in error
        Call DocumentData.SendMessage(sErrDesc, nErrNum, KFX_REL_DOC_ERROR)
        
        Resume Exit_ReleaseDoc

End Function

'*************************************************
' Class_Initialize
'-------------------------------------------------
' Purpose:  VB calls this event when an object
'           of this class is instantiated.  We
'           simply initialize the variables.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub Class_Initialize()
    bOpen = False
End Sub

'*************************************************
' ReleaseNonImageFiles
'-------------------------------------------------
' Purpose:  Release Non Image Files to image file path.
'           This is needed when releasing to PDF image format
'           and eDocuments are part of the same document.
'           Since the PDF Releaser doesn't support eDocuments,
'           these have to be handled separately
'
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub ReleaseNonImageFiles()
On Error GoTo ReleaseNonImageFiles_Error

    Dim fso As Object               '*** For file system object
    Dim oFile As Object             '*** For ImageFiles to be deleted
    Dim strMultipageFileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    '*** Use the MTIFF_G4 format to make the deletion of images easy
    Call DocumentData.ImageFiles.Copy(DocumentData.ImageFilePath, CAP_FORMAT_MTIFF_G4)

    '*** Delete the image file (if any). If there are no images within this
    '*** document, then the deletion will fail. Ignore the error.
On Error Resume Next
    strMultipageFileName = PadFileName(Hex(DocumentData.UniqueDocumentID)) + ".TIF"
    '*** The following code gets the name for the multipage file, when the image type
    '*** is set to PDF
    Set oFile = fso.GetFile(DocumentData.ImageFilePath + "\" + strMultipageFileName)
    
    '*** Delete the file
    oFile.Delete (True)
    
Exit Sub

ReleaseNonImageFiles_Error:
    Call DocumentData.LogError(Err, 0, 0, Err.Description, M_RELEASE + ":" + Err.Source, Erl)
        
End Sub

'*************************************************
' PadFileName
'-------------------------------------------------
' Purpose:  Builds a string from strFileName padded with zero's
'
' Inputs:   strFileName - the string to pad with zero's
' Outputs:  None
' Returns:  strFileName prepended with zero's to make it's length
'           a total of eight characters e.g. "F" becomes "0000000F"
'
' Notes:    strFileName is the prefix file name without extension
'*************************************************
Public Function PadFileName(ByVal strFileName As String) As String
    
    Dim strPad As String
    Dim nRetVal As Integer
        
    For nRetVal = Len(strFileName) To 7
        strPad = strPad & "0"
    Next
    
    PadFileName = strPad & strFileName

End Function


