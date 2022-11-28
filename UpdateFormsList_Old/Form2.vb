Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmOPMXMLDF
    Inherits System.Windows.Forms.Form
    'Setup Global database
    Public dbOPM As DAO.Database
    Public rs As DAO.Recordset
    Public fs As Scripting.FileSystemObject
    ' Setup Recordset Variables
    Public rsPurge As DAO.Recordset
    Public rsData As DAO.Recordset
    Public rsOPM As DAO.Recordset
    Public dbNew As DAO.Database
    Public rsNew As DAO.Recordset

    ' Setup Global filenames
    Public strXml As String
    Public strBatchName As String
    Public strOrigBatchName As String
    Public strDenverZip As String
    Public strHerndonZip As String
    Public strTransID As String
    Public strDataName As String
    Public strZipFile As String
    Public strBackupDateDir As String
    Public strBackupDateData As String

    'Setup Global directories
    Public strFolderDir As String
    ' Count variables and temp dir variable
    Public strOPMDir As String
    Public strOPMXMLDir As String
    Public strProcessStartDate As String
    Public lngPurge As Integer
    Public lngPurgeDocs As Integer
    Public lngPurgePages As Integer
    Public lngPages As Integer
    Public lngTotalPages As Integer
    Public lngCount As Integer
    Public lngDocs As Integer
    Public lngBar As Integer
    Public strBatVol As String
    Public CSVOut As Short
    Public strCSVFile As String
    Public strLogFile As String
    Public ProbOut As Short
    Public blnStatus As Boolean
    Public blnValidDate As Boolean
    Public StrStatusInfo As String

    Function CheckBatchName(ByRef strOrig As String, ByRef strOp As String, ByRef strD As String) As Boolean
        Dim strProbWrite As String
        Dim strProbRec As String
        Dim strProbStat As String

        On Error GoTo CheckBatchName_Error

        '**** Check to see if the len of the database is equal to 19, if it does not
        '**** Move database to problem folder database directory and the folder to the
        '**** problem folder directory and also log the error in the error log
        If (Len(strOrigBatchName) = 19) And (IsNumeric(Mid(strOrigBatchName, 4, 15))) Then
            CheckBatchName = True
        Else
            'String written to problem text file
            strProbWrite = "Folder name is not a valid length or format : " & strOrigBatchName & " - " & Now & vbCrLf

            'String written to status label
            strProbStat = strOrigBatchName & " - Folder name not a valid length or format "

            'String written to problem table
            strProbRec = strOrigBatchName & " - Error in Batch Name : Folder name not a valid length or format "

            HandleProblems(strOrig, strD, strOp, strProbStat, strProbWrite, strProbRec, False, strProblemFolderDir, True, False)

            CheckBatchName = False
        End If

        Exit Function

CheckBatchName_Error:

        Select Case Err.Number

            Case Else

                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description & vbCrLf & StrStatusInfo & vbCrLf & "Current Batch in process : " & strOrigBatchName)
                'Close #ProbOut
                Exit Function
        End Select

    End Function

    Function CheckEmptyDatabase(ByRef strOrig As String, ByRef strB As String, ByRef strD As String) As Boolean
        Dim flsTest As Scripting.Files
        Dim fldTest As Scripting.Folder
        Dim lngFileCnt As Integer
        Dim lngRecCnt As Integer
        Dim strProbRec As String
        Dim strProbWrite As String
        Dim strProbStat As String

        On Error GoTo CheckEmptyDatabase_Error

        CheckEmptyDatabase = True

        fldTest = fs.GetFolder(txtDir(0).Text & strOrig)
        flsTest = fldTest.Files
        lngFileCnt = flsTest.Count
        fldTest = Nothing
        flsTest = Nothing

        'Checking for empty database table
        If rsOPM.EOF Then

            'String written to problem text file
            strProbWrite = "Database table does not contain any records : " & strOrigBatchName & " - " & Now & vbCrLf

            'String written to status label
            strProbStat = strOrigBatchName & " - Database table does not contain " & " any records"

            'String written to problem table
            strProbRec = strOrigBatchName & " - Folderlist table in " & strD & " database is empty"

            HandleProblems(strOrig, strD, strOPMDir, strProbStat, strProbWrite, strProbRec, True, strProblemFolderDir, True, True)

            CheckEmptyDatabase = False
            'blnStatus = False
            Exit Function
        End If

        If CheckEmptyDatabase = True Then
            'Database does contain records, get the total count
            rsOPM.MoveLast()
            lngRecCnt = rsOPM.RecordCount

            If lngFileCnt <> lngRecCnt Then

                'String written to problem text file
                strProbWrite = "Record count in database table does not match file count in folder : " & strOrigBatchName & " - ( Record Count = " & lngRecCnt & " ) and ( File Count in Folder = " & lngFileCnt & " ) :  " & Now & vbCrLf

                'String written to status label
                strProbStat = strOrigBatchName & " - Record count in database table does not match file count in folder " & " - ( Record Count = " & lngRecCnt & " ) and ( File Count in Folder = " & lngFileCnt & " )"

                'String written to problem table
                strProbRec = strOrigBatchName & " : " & " - Record count in database table does not match file count in folder " & " - ( Record Count = " & lngRecCnt & " ) and ( File Count in Folder = " & lngFileCnt & " )"

                HandleProblems(strOrig, strD, strOPMDir, strProbStat, strProbWrite, strProbRec, True, strProblemFolderDir, True, True)

                CheckEmptyDatabase = False
            ElseIf rsOPM.RecordCount <= 2 Then
                'Check to see if the table has 2 or less records - if so record as an error

                'String written to problem text file
                strProbWrite = "Database table contains 2 or less records : " & strOrigBatchName & " - " & Now & vbCrLf

                'String written to status label
                strProbStat = strOrigBatchName & " - Database table contains 2 or less records "

                'String written to problem table
                strProbRec = strOrigBatchName & " - Folderlist table in " & strD & " database contains 2 or less records"

                HandleProblems(strOrig, strD, strOPMDir, strProbStat, strProbWrite, strProbRec, True, strProblemFolderDir, True, True)

                CheckEmptyDatabase = False

            Else
                CheckEmptyDatabase = True
            End If
        End If

        Exit Function

CheckEmptyDatabase_Error:

        Select Case Err.Number

            Case Else

                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description & vbCrLf & StrStatusInfo & vbCrLf & "Current Batch in process : " & strOrigBatchName)
                'Close #ProbOut
                Exit Function
        End Select

    End Function

    Sub HandleProblems(ByRef strOrigProb As String, ByRef strDataProb As String, ByRef strBatchProb As String, ByRef strStatProb As String, ByRef strWriteProb As String, ByRef strRecProb As String, ByRef blnData As Boolean, ByRef strProbOutDir As String, ByRef blnDeleteFolder As Boolean, ByRef blnDeleteFile As Boolean)

        On Error GoTo HandleProblems_Error
        Dim sQuery As String
        If fs.FolderExists(strProbOutDir & "Databases") Then

            If fs.FileExists(strBackupDateData) Then
                'copy backup database to the problem folder database subdirectory
                fs.CopyFile(strBackupDateData, strProbOutDir & "Databases\" & strOrigProb & ".mdb")
                'Check is required file is present to delete
            End If
            If blnDeleteFile = True Then
                'copy database to the problem folder database subdirectory
                'fs.CopyFile strDataProb, strProbOutDir & "Databases\" & strOrigProb & ".mdb"
                If blnData = True Then
                    rsOPM.Close()
                    dbOPM.Close()
                End If
                fs.DeleteFile(strDataProb)
            End If

            If fs.FolderExists(strBackupDateDir) Then
                'Copy Backup folder to the Problem Folder
                fs.CopyFolder(strBackupDateDir, strProbOutDir)
                'Check if required folder is present to process folder
            End If

            If blnDeleteFolder = True Then
                'Copy folder to the Problem Folder
                'fs.CopyFolder strBatchProb, strProbOutDir
                fs.DeleteFolder(strBatchProb)
            End If
            'Write out error to log file
            WriteLine(ProbOut, strWriteProb)

            StrStatusInfo = strStatProb

        Else
            fs.CreateFolder(strProbOutDir & "Databases")
            'copy database to the problem folder database subdirectory
            fs.CopyFile(strDataProb, strProbOutDir & "Databases\" & strOrigProb & ".mdb")
            If blnData = True Then
                rsOPM.Close()
                dbOPM.Close()
            End If
            fs.DeleteFile(strDataProb)
            'Check if reuired to process folder
            If blnDeleteFolder = True Then
                'Copy folder to the Problem Folder
                fs.CopyFolder(strBatchProb, strProbOutDir)
                fs.DeleteFolder(strBatchProb)
            End If

            'Write out error to log file
            WriteLine(ProbOut, strWriteProb)

            'StrStatusInfo = strOrigBatchName & " - Database table contains 2 or less records "
            StrStatusInfo = strStatProb

        End If
        'Create record in problem database of error
        sQuery = "Insert Into Problems(BatchName, ProblemDate, ProbDescrip) Values('" & strOrigProb & "', '" & Now & "', '" & strRecProb & "')"
        db.Execute(sQuery, 64)

HandleProblems_Exit:
        Exit Sub

HandleProblems_Error:

        Select Case Err.Number

            Case Else

                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description & vbCrLf & "Current Batch in process : " & strOrigBatchName)
                'Close #ProbOut
                'GoTo HandleProblems_Exit
                Resume

        End Select

    End Sub

    Function CheckDuplicateFolders(ByRef strOrig As String, ByRef strB As String, ByRef strD As String, ByRef strS As String) As Boolean
        Dim strProbWrite, sQuery As String, rsDupCheck As DAO.Recordset
        Dim strProbRec As String
        Dim strProbStat As String
        Dim strOp As String
        On Error GoTo CheckDuplicateFolders_Error
        sQuery = "Select * From OPMStatus Where BatchVol='" & strS & "'"
        rsDupCheck = db.OpenRecordset(sQuery, DAO.RecordsetTypeEnum.dbOpenDynaset, 512)
        'rsOPM = dbOPM.OpenRecordset(strSQL, DAO.RecordsetTypeEnum.dbOpenDynaset)
        If rsDupCheck.RecordCount > 0 Then

            'String written to problem text file
            strProbWrite = "Duplicate Database found - Moved database and folder to duplicate folder : " & strOrigBatchName & " - " & Now & vbCrLf

            'String written to status label
            strProbStat = strOrigBatchName & " - Duplicate Database found "

            'String written to problem table
            strProbRec = strOrigBatchName & " - Attempt to add duplicate folder to table"

            HandleProblems(strOrig, strD, strOPMDir, strProbStat, strProbWrite, strProbRec, True, strDuplicateDatabaseDir, True, True)

            CheckDuplicateFolders = False
        Else
            CheckDuplicateFolders = True
        End If

        Exit Function

CheckDuplicateFolders_Error:

        Select Case Err.Number

            Case Else

                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description & vbCrLf & StrStatusInfo & vbCrLf & "Current Batch in process : " & strOrigBatchName)
                'Close #ProbOut
                Exit Function
                'Resume

        End Select

    End Function
    Function IsValidDate(ByVal EffDate As Object) As Boolean

        On Error GoTo IsValidDate_Error

        'Completion Field Validation Function
        'Return 1 if validation succeeds, or 0 if validation fails.
        'If validation fails, set the variable 'message' to display a message in the statusbar.
        Dim testVal3, testVal1, testval, testVal2, testVal4 As Object
        testVal1 = Nothing
        testval = EffDate
        testVal2 = DateValue("1/1/1901")
        testVal3 = DateValue("12/31/2099")
        testVal4 = DateValue("1/1/1901")
        'Check value
        If IsDate(testval) Then testVal1 = DateValue(testval)
        If Not IsDate(testval) Then
            IsValidDate = False
        ElseIf testVal1 = testVal4 Then
            IsValidDate = 1
        ElseIf testVal1 < testVal2 Then
            IsValidDate = False
        ElseIf testVal1 > testVal3 Then
            IsValidDate = False
        ElseIf Trim(testval) = "" Then
            IsValidDate = False
        Else
            IsValidDate = True
        End If

        Exit Function

IsValidDate_Error:

        Select Case Err.Number

            Case Else

                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description & vbCrLf & StrStatusInfo & vbCrLf & "Current Batch in process : " & strOrigBatchName)

                Exit Function
        End Select

    End Function

    Sub CheckDatabaseDir()
        Dim fld As Scripting.Folder
        Dim fl As Scripting.File
        Dim fls As Scripting.Files
        Dim strFldDir As String
        Dim strBatchIter As String
        Dim OPMDir As String
        Dim NewOPMDir As String
        Dim strSSN As String
        Dim strSQL As String
        Dim strForm As String
        Dim intAd332 As Short
        Dim intOF8 As Short
        Dim strData As String
        Dim strNewData As String
        Dim blnProcDir As Boolean
        Dim strDateTest As String
        Dim strNewDate As String
        Dim strThirdDate As String
        Dim blnFirstPass As Boolean
        Dim blnMerge As Boolean
        Dim blnForiegn As Boolean
        Dim strBack As String
        Dim strTemp As String
        Dim strInfoText As String
        Dim strS As String
        Dim strW As String
        Dim strR As String
        Dim blnD As Boolean
        Dim blnF As Boolean
        Dim blnDF As Boolean

        On Error GoTo CheckDatabaseDir_Error

        blnProcDir = False
        blnValidDate = True 'Don't ask...
        'Added afetr changing timer minutes to seconds caused errors - WAJ
        strProcessStartDate = CStr(Now)
        '**************** Setup error log file ********************************
        If Not fs.FolderExists(strProblemFolderDir) Then
            fs.CreateFolder(strProblemFolderDir)
        End If
        strLogFile = strProblemFolderDir & "ErrorLog.txt"
        ProbOut = FreeFile() + 1

        'Open log file to append the error
        FileOpen(ProbOut, strLogFile, OpenMode.Append)
        '**********************************************************************

        strFldDir = txtDir(0).Text & "databases"
        fld = fs.GetFolder(strFldDir)
        fls = fld.Files

        prgBar1.Visible = True

        Dim sTempSSN, sNewPurgeFileName As String
        Dim iSSNVol As Short
        Dim bSSNOK As Boolean
        For Each fl In fls 'begin main batch processing loop
            If Mid(fl.Name, 1, 9) <> "nofindtmp" Then

                txtInfo.Text = "Begin Processing New Database : " & fl.Name
                txtInfo.Refresh()
                strInfoText = txtInfo.Text
                blnProcDir = True
                ' Reset SSN count variables
                lngPurge = 0
                lngPurgePages = 0
                lngPages = 0
                lngTotalPages = 0
                lngPurgeDocs = 0
                lngDocs = 0
                intAd332 = 0
                lngBar = 0
                intOF8 = 0
                prgBar1.Minimum = 0
                prgBar1.Value = 0
                blnMerge = False
                blnForiegn = False
                blnStatus = True

                prgBar1.Refresh()

                Label1.Text = "Start Processing Database ******* " & fl.Name & "  *******"
                Label1.Refresh()

                '*******************   Setup variable contents **************************

                'Rename Batch name to use for directory and also for database name later
                strOrigBatchName = fs.GetBaseName(fl.Name)
                'Set current database name and also new database name
                strData = strFldDir & "\" & fs.GetBaseName(fl.Name) & ".mdb"
                strDataName = strFldDir & "\ANA" & fs.GetBaseName(fl.Name) & ".mdb"
                strOPMDir = txtDir(0).Text & fs.GetBaseName(fl.Name) 'not sure why these 2 vars are both needed. See above
                OPMDir = txtDir(0).Text & fs.GetBaseName(fl.Name)

                '********  Copy OPM batch folder and database to backup directory  ********
                strBackupDateDir = strOPMBackup & VB6.Format(Now, "mm-dd-yyyy") & "\" & strOrigBatchName
                strBackupDateData = strOPMBackup & VB6.Format(Now, "mm-dd-yyyy") & "\Databases\" & strOrigBatchName & ".mdb"
                'Copy OPM SSN folder to backup directory
                If fs.FolderExists(OPMDir) Then
                    If fs.FolderExists(strOPMBackup & VB6.Format(Now, "mm-dd-yyyy") & "\") Then
                        fs.CopyFolder(OPMDir, strOPMBackup & VB6.Format(Now, "mm-dd-yyyy") & "\")
                    Else
                        fs.CreateFolder(strOPMBackup & VB6.Format(Now, "mm-dd-yyyy") & "\")
                        fs.CopyFolder(OPMDir, strOPMBackup & VB6.Format(Now, "mm-dd-yyyy") & "\")
                    End If
                Else
                    blnF = False
                    If fs.FileExists(strData) Then
                        blnD = False
                        blnDF = True
                        'Copy database to backup directory if checkbox is checked
                        If fs.FolderExists(strOPMBackup & VB6.Format(Now, "mm-dd-yyyy") & "\" & "Databases") Then
                            fs.CopyFile((strData), strOPMBackup & VB6.Format(Now, "mm-dd-yyyy") & "\" & "databases\" & strOrigBatchName & ".mdb")
                        Else
                            fs.CreateFolder(strOPMBackup & VB6.Format(Now, "mm-dd-yyyy") & "\" & "Databases\")
                            fs.CopyFile((strData), strOPMBackup & VB6.Format(Now, "mm-dd-yyyy") & "\" & "databases\" & strOrigBatchName & ".mdb")
                        End If
                    Else
                        blnD = False
                        blnDF = False
                    End If

                    strS = strOrigBatchName & " Folder does not exist - Begining of CheckBaseDir sub "
                    strW = strOrigBatchName & " Folder does not exist  - Begining of CheckBaseDir sub"
                    strR = strOrigBatchName & " Folder does not exist - Begining of CheckBaseDir sub"
                    HandleProblems(strOrigBatchName, strData, OPMDir, strS, strW, strR, blnD, strProblemFolderDir, blnF, blnDF)

                    blnStatus = False
                End If

                'Check to make sure Data file exists
                If blnStatus = True Then
                    If fs.FileExists(strData) Then
                        'Copy database to backup directory if checkbox is checked
                        If fs.FolderExists(strOPMBackup & VB6.Format(Now, "mm-dd-yyyy") & "\" & "Databases") Then
                            fs.CopyFile((strData), strOPMBackup & VB6.Format(Now, "mm-dd-yyyy") & "\" & "databases\" & strOrigBatchName & ".mdb")
                        Else
                            fs.CreateFolder(strOPMBackup & VB6.Format(Now, "mm-dd-yyyy") & "\" & "Databases\")
                            fs.CopyFile((strData), strOPMBackup & VB6.Format(Now, "mm-dd-yyyy") & "\" & "databases\" & strOrigBatchName & ".mdb")
                        End If
                    Else
                        blnD = False
                        blnDF = False
                        If fs.FolderExists(OPMDir) Then
                            blnF = True
                        Else
                            blnF = False
                        End If
                        strS = strOrigBatchName & " Folder does not exists - Begining of CheckDatabaseDir sub "
                        strW = strOrigBatchName & " Folder does not exists  - Begining of CheckDatabaseDir sub"
                        strR = strOrigBatchName & " Folder does not exists - Begining of CheckDatabaseDir sub"
                        HandleProblems(strOrigBatchName, strData, OPMDir, strS, strW, strR, blnD, strProblemFolderDir, blnF, blnDF)
                        blnStatus = False
                    End If
                End If

                ' ******** Check for valid batchname  : True status indicates success **********
                If blnStatus = True Then
                    blnStatus = CheckBatchName(strOrigBatchName, OPMDir, strData)
                End If

                'Finish checking for missing folders and invalid format folder names
                If blnStatus = True Then
                    ' No errors found so far continue processing
                    dbOPM = DAODBEngine_definst.OpenDatabase(strData)
                    NewField(dbOPM)
                    rsOPM = dbOPM.OpenRecordset("folderlist", DAO.RecordsetTypeEnum.dbOpenDynaset)

                    ' ******** Check for for empty database  : True status indicates success **********
                    blnStatus = CheckEmptyDatabase(strOrigBatchName, strOrigBatchName, strData)
                    '*******************************************************************************
                End If

                If blnStatus = True Then
                    rsOPM.MoveLast()
                    prgBar1.Maximum = rsOPM.RecordCount
                    rsOPM.MoveFirst()

                End If

                ' ******* Finish checking for errors - Start processing contents of table *********
                If blnStatus = True Then
                    Label1.Text = "Error Checking Complete - Processing Database " & strData & "  : Folderlist Table"
                    Label1.Refresh()

                    rsOPM.MoveFirst()
                    'Search Purge table to determine if record should be marked as Purged
                    With rsOPM
                        Do While Not .EOF
                            strForm = .Fields("formnumber").Value
                            If Not IsDBNull(.Fields("SSN").Value) Then
                                sTempSSN = Trim(.Fields("SSN").Value)
                            Else
                                sTempSSN = ""
                            End If
                            'Presumably a begin or side BC Sheet
                            If sTempSSN = "" Then sTempSSN = strOrigBatchName
                            If (.Fields("purge").Value = True) Then
                                'Update Purge status flag
                                .Edit()
                                .Fields("purgestatus").Value = True
                                .Update()
                                'Move purged pdf file to purged folder directory. Create subdirectory
                                ' if directory does not exist
                                If Not fs.FolderExists(strPurgeFolderDir & "\" & strOrigBatchName) Then fs.CreateFolder(strPurgeFolderDir & "\" & strOrigBatchName)
                                'Create SSN based filename
                                sNewPurgeFileName = strPurgeFolderDir & "\" & strOrigBatchName & "\" & sTempSSN & "_" & "_" & fs.GetFileName(.Fields("pdffilename").Value)
                                fs.CopyFile(strOPMDir & "\" & fs.GetFileName(.Fields("pdffilename").Value), sNewPurgeFileName)
                                fs.DeleteFile(strOPMDir & "\" & fs.GetFileName(.Fields("pdffilename").Value))
                            End If
                            lngBar = lngBar + 1
                            prgBar1.Value = lngBar
                            .MoveNext()
                        Loop
                    End With
                    rsOPM.Close()

                    StrStatusInfo = strDataName & " - Finish Phase 1 - Purge Check"

                    dbOPM.Close()

                    'rename database to new name ( before move database to complete directory )
                    Rename(strData, strDataName)

                    StrStatusInfo = strDataName & " - Renaming database with ANA prefix"

                    'Check database and folder and remove XML files and also additional copies of
                    'PDF files
                    CleanupOPMFolder()

                    'Process database and create XML
                    If blnStatus = True Then
                        ProcessDir()
                        StrStatusInfo = strDataName & " - Finish Processdir Sub"
                    End If

                    If blnStatus = True Then
                        'Check if the effective date was a valid date
                        '*** If the effective date is a valid date continue to zip files and create CSV File
                        '**
                        '*** If the effective date is not valid then move the folder adn database to the Probdate directory
                        If blnValidDate = True Then

                            StrStatusInfo = strDataName & " - Finish Copying Completed folder to CompletedFolder Directory and deleting original folder"

                            'Move completed folder to Completed floders directory
                            fs.CopyFolder(strOPMDir, strCompletedFoldersDir & "\")
                            'Move completed database to the Completed Database directory
                            fs.CopyFile(strDataName, strCompleteDatabaseDir & "\")
                            fs.DeleteFile((strDataName))
                            fs.DeleteFolder((strOPMDir))

                            StrStatusInfo = strDataName & " - Finish Copying Completed database to CompletedDatabase Directory and deleting original database"

                            txtInfo.Text = strInfoText & vbCrLf & "Finish Processing Database : " & strDataName
                            txtInfo.Refresh()
                            If fls.Count = 0 Then Timer1.Enabled = True
                            lngCount = 0

                            StrStatusInfo = strDataName & " - Finish Updating status database"

                        Else
                            ' Copy folder from backup directory to the ProbDate directory due to bad Effective Date
                            fs.CopyFolder(strBackupDateDir, strProbDateDir)
                            fs.DeleteFolder((strOPMDir))
                            StrStatusInfo = strDataName & " - Finish Copying folder to ProbDate Directory and deleting original folder"
                            'Copy database from backup directory to the Probdate directory due to bad Effective Date
                            'Create database directory if it does not exist
                            If Not fs.FolderExists(strProbDateDir & "Databases") Then
                                fs.CreateFolder(strProbDateDir & "Databases")
                                fs.CopyFile(strBackupDateData, strProbDateDir & "Databases\")
                            Else
                                fs.CopyFile(strBackupDateData, strProbDateDir & "Databases\")
                            End If
                            fs.DeleteFile((strDataName))

                            StrStatusInfo = strDataName & " - Finish Copying database to ProbDate Directory and deleting original database"

                            txtInfo.Text = "Finish Processing Database and moving files to ProbDate directory due to bad Effective Date : " & strDataName
                            txtInfo.Refresh()
                            Timer1.Enabled = True
                            lngCount = 0
                        End If
                    End If
                End If
            End If
        Next fl

        txtInfo.Text = "No databases found to process....."
        txtInfo.Refresh()
        Timer1.Enabled = True
        lngCount = 0
        prgBar1.Visible = False
        blnProcDir = False

        FileClose(ProbOut)
        Exit Sub

CheckDatabaseDir_Error:

        Select Case Err.Number

            Case Else

                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description & vbCrLf & StrStatusInfo & vbCrLf & "Current Batch in process : " & strOrigBatchName)
                'Close #ProbOut
                Exit Sub
                'Resume
        End Select

    End Sub

    'Cleanup folder subroutine checks database removes files that were added unintentionally
    Sub CleanupOPMFolder()
        Dim strSQL As String
        Dim fldTest As Scripting.Folder
        Dim flTest As Scripting.File
        Dim flsTest As Scripting.Files
        Dim strNewTest As String
        Dim strSearch As String
        Dim lngFiles As Integer
        Dim lngRecs As Integer
        Dim strProbWrite As String
        Dim strProbRec As String
        Dim strProbStat As String

        On Error GoTo CleanupOPMFolder_Error

        'Search folder for any old XML file(s) and delete file(s) if found
        If fs.FileExists(strOPMDir & "\*.xml") Then
            fs.DeleteFile((strOPMDir & "\*.xml"))
        End If

        strSQL = "Select * from folderlist "
        strSQL = strSQL & "Where Purgestatus = FALSE"

        If fs.FileExists(strDataName) Then
            'Open database for OPF
            dbOPM = DAODBEngine_definst.OpenDatabase(strDataName)
            rsOPM = dbOPM.OpenRecordset(strSQL, DAO.RecordsetTypeEnum.dbOpenDynaset)
        End If

        fldTest = fs.GetFolder(strOPMDir)
        flsTest = fldTest.Files

        prgBar1.Visible = True

        With rsOPM
            For Each flTest In flsTest
                strNewTest = fs.GetParentFolderName(.Fields("pdffilename").Value) & "\" & flTest.Name
                strSearch = strOPMDir & "\" & flTest.Name
                rsOPM.FindFirst("PdfFilename = '" & strNewTest & "'")
                If .NoMatch Then
                    fs.DeleteFile((strSearch))
                End If
            Next flTest
        End With

        'Check to see after removal if all of the records were deleted (EOF Condition)
        If Not rsOPM.EOF Then
            'Move to last record and get record count
            rsOPM.MoveLast()
            lngRecs = rsOPM.RecordCount
            'Go back to first record for processing next step
            rsOPM.MoveFirst()
            'Check to make sure all records match folder contents after removal Since not at EOF
            With rsOPM
                Do While Not .EOF
                    'strNewTest = fs.GetParentFolderName(!pdffilename) & "\" & fs.GetBaseName(!pdfiflename)
                    strSearch = strOPMDir & "\" & fs.GetBaseName(.Fields("pdffilename").Value) & ".pdf"

                    'If any database pdfname does not match folder content declare an error and
                    ' move folder abd database to problem folder
                    If Not fs.FileExists(strSearch) Then
                        blnStatus = False
                        'String written to problem text file
                        strProbWrite = "Database record does not match any file in directory : " & strOrigBatchName & " - " & Now & vbCrLf

                        'String written to status label
                        strProbStat = strOrigBatchName & " - Database record does not match any file in directory " & " - PDF Filename = " & fs.GetBaseName(.Fields("pdffilename").Value)

                        'String written to problem table
                        strProbRec = strOrigBatchName & " - Database record does not match any file in directory" & " - PDF Filename = " & fs.GetBaseName(.Fields("pdffilename").Value)

                        HandleProblems(strOrigBatchName, strDataName, strOPMDir, strProbStat, strProbWrite, strProbRec, True, strProblemFolderDir, True, True)

                        Exit Do
                    End If

                    .MoveNext()
                Loop
            End With
        Else
            'No records in database, abort processing zip file
            blnStatus = False

            'String written to problem text file
            strProbWrite = "Database table does not contain any records : " & strOrigBatchName & " - " & Now & vbCrLf

            'String written to status label
            strProbStat = strOrigBatchName & " - Database table does not contain " & " any records"

            'String written to problem table
            strProbRec = strOrigBatchName & " - Folderlist table in " & strDataName & " database is empty"

            HandleProblems(strOrigBatchName, strDataName, strOPMDir, strProbStat, strProbWrite, strProbRec, True, strProblemFolderDir, True, True)

        End If

        If blnStatus = True Then
            'Get file count
            lngFiles = flsTest.Count

            If lngRecs <> lngFiles Then
                blnStatus = False
                'String written to problem text file
                strProbWrite = "Database record count does not match files in directory : " & strOrigBatchName & " - " & Now & vbCrLf

                'String written to status label
                strProbStat = strOrigBatchName & " - Database record count does not match files in directory"

                'String written to problem table
                strProbRec = strOrigBatchName & " - Database record count does not match files in directory"

                HandleProblems(strOrigBatchName, strDataName, strOPMDir, strProbStat, strProbWrite, strProbRec, True, strProblemFolderDir, True, True)

            End If

        End If

        fldTest = Nothing
        flsTest = Nothing

        If blnStatus = True Then
            rsOPM.Close()
            dbOPM.Close()
        End If
        Exit Sub

CleanupOPMFolder_Error:

        Select Case Err.Number

            Case Else

                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description & vbCrLf & StrStatusInfo & vbCrLf & "Current Batch in process : " & strOrigBatchName)

                Exit Sub
                'Resume
        End Select

    End Sub

    Sub RenamePDFFiles(ByRef sNewSSN As String)
        'Rename PDFs to proper xml format, add up docs and pages for this ssn
        Dim strNewTest As String
        Dim strSearch As String
        Dim intXML, iFileCount As Short
        'UPGRADE_WARNING: Arrays in structure rsNewSSN may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        Dim sNewSSNQuery, sNewPDFPath As String
        Dim rsNewSSN As DAO.Recordset

        On Error GoTo RenamePDFFiles_Error

        strOPMXMLDir = txtDir(0).Text & strBatchName & "_0000000001" 'loose paper mod
        fs.CreateFolder(strOPMXMLDir)

        sNewSSNQuery = "Select * From FolderList "
        sNewSSNQuery = sNewSSNQuery & "Where SSN = '" & sNewSSN & "' "
        rsNewSSN = dbNew.OpenRecordset(sNewSSNQuery, DAO.RecordsetTypeEnum.dbOpenDynaset)
        lngPurge = 0
        lngPurgeDocs = 0
        lngPages = 0
        lngDocs = 0
        If rsNewSSN.RecordCount > 0 Then
            iFileCount = 0
            rsNewSSN.MoveLast()
            rsNewSSN.MoveFirst()
            Do Until rsNewSSN.EOF
                If rsNewSSN.Fields("PurgeStatus").Value = True Then
                    'Add to purge counts
                    lngPurge = lngPurge + Val(rsNewSSN.Fields("Pages").Value)
                    lngPurgeDocs = lngPurgeDocs + 1
                    lngTotalPages = lngTotalPages + lngPurge
                Else
                    lngBar = lngBar + 1 'prgbar1.max = non purge only
                    prgBar1.Value = lngBar
                    lngPages = lngPages + Val(rsNewSSN.Fields("Pages").Value)
                    lngTotalPages = lngTotalPages + lngPages
                    lngDocs = lngDocs + 1
                    iFileCount = iFileCount + 1
                    strSearch = strOPMDir & "\" & fs.GetFileName(rsNewSSN.Fields("PDFFileName").Value)
                    strNewTest = strOPMXMLDir & "\" & strBatchName & "_0000000001_" & VB6.Format(iFileCount, "00000000") & ".pdf"
                    fs.CopyFile(strSearch, strNewTest)
                    'Modify input table with new file name for correct sort later
                    sNewPDFPath = fs.GetParentFolderName(rsNewSSN.Fields("PDFFileName").Value) & "\" & strBatchName & "_0000000001_" & VB6.Format(iFileCount, "00000000") & ".pdf"
                    rsNewSSN.Edit()
                    rsNewSSN.Fields("PDFFileName").Value = sNewPDFPath
                    rsNewSSN.Update()
                End If
                rsNewSSN.MoveNext()
            Loop
        End If
        rsNewSSN.Close()
        Exit Sub

RenamePDFFiles_Error:

        Select Case Err.Number

            Case Else

                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description & vbCrLf & StrStatusInfo & vbCrLf & "Current Batch in process : " & strOrigBatchName)

                Exit Sub
        End Select

    End Sub

    'Sub ProcessDir(strDir As String, strNewBatch As String, strBN As String)
    Sub ProcessDir()
        Dim x As Short
        Dim xmlOut As Short
        Dim strFormNumber, sQuery As String
        Dim strNOA2, strFormType, strNOA1, strDeliverySide As String
        Dim strDuplex, xmlFileName As String
        Dim strZipFileName As String
        Dim strSQL As String
        Dim bBail As Boolean
        Dim intValidDate As Short
        Dim blnRecDate As Boolean
        Dim strSearch As String
        Dim strProbWrite As String
        Dim strProbRec As String
        Dim strProbStat As String
        Dim rsPurgeQuery As DAO.Recordset
        Dim strType, sSSNQuery As String
        Dim sTemp As String
        'Dim o As New cCreateGUID '2005/12/18 Use class
        Dim sOldSSN As String 'For loose paper
        Dim iFileCount As Short

        On Error GoTo Command1_Error

        strSQL = "Select * from folderlist "
        strSQL = strSQL & "Where Purgestatus = FALSE "
        strSQL = strSQL & "ORDER By SSN, PDFFileName" 'added for loose paper

        intValidDate = 0
        If fs.FileExists(strDataName) Then
            'Open database for OPF
            dbNew = DAODBEngine_definst.OpenDatabase(strDataName)
            rsNew = dbNew.OpenRecordset(strSQL, DAO.RecordsetTypeEnum.dbOpenDynaset)
        End If

        xmlOut = FreeFile() + 3
        rsNew.MoveLast()
        rsNew.MoveFirst()
        prgBar1.Maximum = rsNew.RecordCount
        prgBar1.Value = 0
        lngBar = 0
        With rsNew
            'Loop through database
            Do While Not .EOF
                sOldSSN = Trim(rsNew.Fields("SSN").Value)
                'Moved inside loop for loose paper
                strBatchName = GetBatchName(Trim(rsNew.Fields("SSN").Value))
                strZipFile = strBatchName & "_0000000001.zip"
                strXml = strBatchName & "_0000000001.xml"
                RenamePDFFiles(Trim(.Fields("SSN").Value)) 'strOPMXMLDir created and defined here
                xmlFileName = strOPMXMLDir & "\" & strXml
                FileOpen(xmlOut, xmlFileName, OpenMode.Output)
                PrintLine(xmlOut, "<?xml version=""1.0""?>")
                PrintLine(xmlOut, "<Batch xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" Status=""Delivery"">")
                PrintLine(xmlOut, "   <Name>" & strBatchName & "</Name>")
                PrintLine(xmlOut, "   <SSN>" & rsNew.Fields("SSN").Value & "</SSN>")
                PrintLine(xmlOut, "   <DeliveryID>ST-ST00-" & VB6.Format(Now, "yyyymmdd") & "</DeliveryID>")
                PrintLine(xmlOut, "   <DocumentList>")
                iFileCount = 0
                Do Until sOldSSN <> Trim(rsNew.Fields("SSN").Value)
                    iFileCount = iFileCount + 1
                    strSearch = strOPMXMLDir & "\" & strBatchName & "_0000000001_" & VB6.Format(iFileCount, "00000000") & ".pdf"

                    'If any database pdfname does not match folder content declare an error and
                    ' move folder abd database to problem folder
                    bBail = False
                    If Not fs.FileExists(strSearch) Then
                        blnStatus = False
                        bBail = True
                        'String written to problem text file
                        strProbWrite = "PROCESSDIR () - Database record does not match any file in directory : " & strOrigBatchName & " - " & Now & vbCrLf

                        'String written to status label
                        strProbStat = strOrigBatchName & " PROCESSDIR () - Database record does not match any file in directory " & " - PDF Filename = " & fs.GetBaseName(.Fields("pdffilename").Value)

                        'String written to problem table
                        strProbRec = strOrigBatchName & " PROCESSDIR () - Database record does not match any file in directory" & " - PDF Filename = " & fs.GetBaseName(.Fields("pdffilename").Value)

                        HandleProblems(strOrigBatchName, strDataName, strOPMXMLDir, strProbStat, strProbWrite, strProbRec, True, strProblemFolderDir, True, True)

                        Exit Do
                    End If
                    If bBail Then Exit Do 'No PDF found
                    sTemp = System.Guid.NewGuid.ToString()
                    'sTemp = Mid(sTemp, 2, 36)

                    'update DocGuid field with newly create DocGuid
                    .Edit()
                    .Fields("docguid").Value = sTemp
                    .Update()
                    PrintLine(xmlOut, "      <Document Type=""INS"">")
                    PrintLine(xmlOut, "         <DocGUID>" & sTemp & "</DocGUID>")
                    PrintLine(xmlOut, "         <Path>" & "\" & fs.GetBaseName(strOPMXMLDir) & "\" & strBatchName & "_0000000001_" & VB6.Format(iFileCount, "00000000") & ".pdf" & "</Path>")

                    'check id effective date is valid and record in problem table if date is invalid
                    blnRecDate = IsValidDate(.Fields("EffDate").Value)
                    If blnRecDate = False Then
                        'Create record in problem database of error
                        sQuery = "Insert Into Problems(BatchName, ProblemDate, ProbDescrip, PDFFileName) Values('" & strOrigBatchName & "', '" & Now & "', '" & "Problem Date" & "', '" & .Fields("pdffilename").Value & "')"
                        db.Execute(sQuery, 64)

                    End If

                    'Form Number
                    If (Not IsDBNull(.Fields("formnumber").Value)) And (Not IsDBNull(.Fields("originalformnumber").Value)) Then
                        ' Check to see if the form number = 2809, it should be SF2809
                        '** so we correct the error here. If formnumber not a 2809, keep original form number
                        If (.Fields("formnumber").Value = "2809") Then
                            strFormNumber = "SF2809"
                        Else
                            strFormNumber = .Fields("originalformnumber").Value
                        End If
                        'Check for ampersands
                        strFormNumber = CheckFormType(strFormNumber)

                        PrintLine(xmlOut, "         <FormNumber>" & Trim(strFormNumber) & "</FormNumber>")
                    ElseIf (Not IsDBNull(.Fields("formnumber").Value)) And (IsDBNull(.Fields("originalformnumber").Value)) Then
                        ' Check to see if the form number = 2809, it should be SF2809
                        '** so we correct the error here. If formnumber not a 2809, keep original form number
                        If (.Fields("formnumber").Value = "2809") Then
                            strFormNumber = "SF2809"
                        Else
                            strFormNumber = .Fields("formnumber").Value
                        End If
                        PrintLine(xmlOut, "         <FormNumber>" & Trim(strFormNumber) & "</FormNumber>")
                    Else
                        'strFormNumber = ""
                        PrintLine(xmlOut, "         <FormNumber />")
                    End If

                    'Form Type
                    If Len(Trim(.Fields("formtype").Value)) <> 0 Then
                        strFormType = .Fields("formtype").Value
                        strType = CheckFormType(strFormType) 'Check for & and /
                        PrintLine(xmlOut, "         <FormType>" & strType & "</FormType>")
                    Else
                        'strFormType = ""
                        PrintLine(xmlOut, "         <FormType />")
                    End If

                    If Len(Trim(.Fields("EffDate").Value)) <> 0 Then
                        PrintLine(xmlOut, "         <EffDate>" & .Fields("EffDate").Value & "</EffDate>")
                    Else
                        PrintLine(xmlOut, "         <EffDate />")
                    End If

                    'Merge Field
                    If (Len(strOrigBatchName) = 12) And (VB.Right(strOrigBatchName, 1) = "M") Then
                        PrintLine(xmlOut, "         <Merge>T</Merge>")
                    Else
                        PrintLine(xmlOut, "         <Merge>F</Merge>")
                    End If

                    'NOA1
                    If Len(Trim(.Fields("noa1").Value)) <> 0 Then
                        strNOA1 = Trim(.Fields("noa1").Value)
                    Else
                        strNOA1 = ""
                    End If
                    If strNOA1 <> "" Then
                        PrintLine(xmlOut, "         <NOA1>" & Trim(strNOA1) & "</NOA1>")
                    Else
                        PrintLine(xmlOut, "         <NOA1 />")
                    End If

                    'NOA2
                    If Len(Trim(.Fields("noa2").Value)) <> 0 Then
                        strNOA2 = Trim(.Fields("noa2").Value)
                    Else
                        strNOA2 = ""
                    End If
                    'NOA2 Check
                    If strNOA2 <> "" Then
                        PrintLine(xmlOut, "         <NOA2>" & strNOA2 & "</NOA2>")
                    Else
                        PrintLine(xmlOut, "         <NOA2 />")
                    End If

                    'SourceSide
                    If Len(Trim(rsNew.Fields("sourceside").Value)) <> 0 Then
                        PrintLine(xmlOut, "         <SourceSide>" & .Fields("sourceside").Value & "</SourceSide>")
                    Else
                        PrintLine(xmlOut, "         <SourceSide />")
                    End If

                    'DeliverySide Check
                    If Len(Trim(.Fields("deliveryside").Value)) <> 0 Then
                        strDeliverySide = .Fields("deliveryside").Value
                        PrintLine(xmlOut, "         <DeliverySide>" & Trim(strDeliverySide) & "</DeliverySide>")
                    Else
                        strDeliverySide = ""
                        PrintLine(xmlOut, "         <DeliverySide />")
                    End If

                    'Duplex Check
                    If Len(Trim(.Fields("Duplex").Value)) <> 0 Then
                        If .Fields("Duplex").Value = False Then
                            strDuplex = "F"
                            PrintLine(xmlOut, "         <Duplex>" & Trim(strDuplex) & "</Duplex>")
                        ElseIf .Fields("Duplex").Value = True Then
                            strDuplex = "T"
                            PrintLine(xmlOut, "         <Duplex>" & Trim(strDuplex) & "</Duplex>")
                        End If
                    Else
                        PrintLine(xmlOut, "         <Duplex />")
                    End If
                    'Print #xmlOut, "         <Duplex>" & Trim(strDuplex) & "</Duplex>"
                    PrintLine(xmlOut, "         <ErrCode />")
                    PrintLine(xmlOut, "         <ImageCount>" & .Fields("pages").Value & "</ImageCount>")
                    PrintLine(xmlOut, "      </Document>")
                    sOldSSN = .Fields("SSN").Value
                    'Add record for each document to Reconcilliation table
                    AddReconRecord(strBatchName & "_0000000001_" & VB6.Format(iFileCount, "00000000") & ".pdf", rsNew.Fields("ssn").Value, rsNew.Fields("OriginalFormNumber").Value, Val(rsNew.Fields("pages").Value), rsNew.Fields("DocGuid").Value)
                    .MoveNext()
                    If .EOF Then Exit Do
                Loop
                'Finish xml
                PrintLine(xmlOut, "   </DocumentList>")
                PrintLine(xmlOut, "</Batch>")
                FileClose(xmlOut)
                'Now zip it up
                Label1.Text = "XML Creation Completed : Begin Zip File Creation"
                Label1.Refresh()
                ZipFiles(strOPMXMLDir)
                StrStatusInfo = strDataName & " - Finish Zip File Routine"

                ' Update OPMStatus table with Purge and Pages totals
                strBatVol = Mid(strBatchName, 13, 6)
                sQuery = "INSERT INTO OPMStatus(BatchName, SSN, BatchVol, ProcessStartDate, ProcessEndDate, TotalPurgePages, TotalDeliveredPages, TotalPages, TotalDocs, TotalPurgeDocs, TotalDeliveredDocs, VolNo) "
                sQuery = sQuery & "VALUES('" & strBatchName & "', '" & Mid(strBatchName, 4, 9) & "', '" & strOrigBatchName & "', '" & strProcessStartDate & "', '" & Now & "', " & lngPurge & ", " & lngPages & ", " & lngTotalPages & ", " & lngDocs & ", " & lngPurgeDocs & ", " & lngDocs - lngPurgeDocs & ", '" & strBatVol & "')"
                db.Execute(sQuery, 64)

                'Delete ssn based xml and folder
                fs.DeleteFolder((strOPMXMLDir))
            Loop
        End With

        'If the effective date is not valid move the folder and database into the ProbDate
        'directory

        If blnValidDate = True Then
            rsNew.Close()
            dbNew.Close()
        Else
            rsNew.Close()
            dbNew.Close()
            'Finish Proceesing files but effective date is invalid
            Label1.Text = "XML Creation Completed but effective date is invalid : Zip File Creation is cancelled"
            Label1.Refresh()
        End If

        Exit Sub

Command1_Error:
        Select Case Err.Number
            Case Else
                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description)
                rsNew.Close()
                dbNew.Close()
                Exit Sub
        End Select
    End Sub

    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click

        Me.Close()

    End Sub

    Private Sub cmdDir_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDir.Click
        Dim MyResult As System.Windows.Forms.DialogResult

        ' Set CancelError is True
        'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
        On Error GoTo Error_Dir

        CommonDialog1Open.InitialDirectory = "C:\"
        CommonDialog1Open.FileName = "*.*"
        ' Set flags
        CommonDialog1Open.ShowReadOnly = False
        CommonDialog1Open.ValidateNames = False
        CommonDialog1Open.CheckPathExists = True
        ' Set filters
        CommonDialog1Open.Filter = "All Files (*.*)|*.*|Text Files" & "(*.txt)|*.txt"
        ' Specify default filter
        CommonDialog1Open.FilterIndex = 1
        CommonDialog1Open.Title = "Select Data File to Modify"
        ' Display the Open dialog box
        CommonDialog1Open.ShowDialog()
        If MyResult = DialogResult.Cancel Then GoTo Error_Dir
        ' Display name of selected file
        txtDir(0).Text = fs.GetParentFolderName(CommonDialog1Open.FileName)
Error_Dir:

    End Sub

    Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click

        On Error GoTo cmdSave_Click_Error

        ' Check directory structure. Create directories if they are not available
        If Not fs.FolderExists(txtWorkDir.Text & "\TempZip") Then
            fs.CreateFolder(txtWorkDir.Text & "\TempZip")
            strDenverZipDir = txtWorkDir.Text & "\TempZip"
        End If

        If Not fs.FolderExists(txtWorkDir.Text & "\Zip") Then
            fs.CreateFolder(txtWorkDir.Text & "\Zip")
            strHerndonZipDir = txtWorkDir.Text & "\Zip"
        End If

        If Not fs.FolderExists(txtWorkDir.Text & "\CompletedFolders") Then
            fs.CreateFolder(txtWorkDir.Text & "\CompletedFolders")
            strCompletedFoldersDir = txtWorkDir.Text & "\CompletedFolders"
        End If

        If Not fs.FolderExists(txtWorkDir.Text & "\CompletedDatabases") Then
            fs.CreateFolder(txtWorkDir.Text & "\CompletedDatabases")
            strCompleteDatabaseDir = txtWorkDir.Text & "\CompletedDatabases"
        End If

        If Not fs.FolderExists(txtWorkDir.Text & "\PurgedFolders") Then
            fs.CreateFolder(txtWorkDir.Text & "\PurgedFolders")
            strPurgeFolderDir = txtWorkDir.Text & "\PurgedFolders"
        End If

        'Return to "frmIndx" and update the global parms using the values in the "frmCntls".
        ' Hide controls screen and show main screen back.
        Call UpdateHdrParmsFromCntlsFrmDF()
        Call WriteIniFile(IniFileNameDF)

        Exit Sub

cmdSave_Click_Error:

        Select Case Err.Number

            Case Else

                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description & vbCrLf & StrStatusInfo & vbCrLf & "Current Batch in process : " & strOrigBatchName)

                Exit Sub
        End Select

    End Sub

    Private Sub cmdStartDoc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdStartDoc.Click

        On Error GoTo cmdStartDoc_Click_Error

        If cmdStartDoc.Text = "START" Then
            txtInfo.Text = ""
            txtInfo.BackColor = System.Drawing.ColorTranslator.FromOle(&HFF00)
            cmdStartDoc.Text = "STOP"
            cmdStartDoc.Refresh()
            lngCount = 0
            Timer1.Enabled = True
        Else
            cmdStartDoc.Text = "START"
            txtInfo.BackColor = System.Drawing.ColorTranslator.FromOle(&HFF)
            lngCount = 0
            Timer1.Enabled = False
        End If

        Exit Sub

cmdStartDoc_Click_Error:

        Select Case Err.Number

            Case Else

                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description & vbCrLf & StrStatusInfo & vbCrLf & "Current Batch in process : " & strOrigBatchName)
                Exit Sub
        End Select

    End Sub

    Private Sub cmdWorkDir_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdWorkDir.Click
        Dim MyResult As System.Windows.Forms.DialogResult

        ' Set CancelError is True
        'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
        On Error GoTo Error_WorkDir


        CommonDialog1Open.InitialDirectory = "C:\"
        CommonDialog1Open.FileName = "*.*"
        ' Set flags
        CommonDialog1Open.ShowReadOnly = False
        CommonDialog1Open.ValidateNames = False
        CommonDialog1Open.CheckPathExists = True
        ' Set filters
        CommonDialog1Open.Filter = "All Files (*.*)|*.*|Text Files" & "(*.txt)|*.txt"
        ' Specify default filter
        CommonDialog1Open.FilterIndex = 1
        CommonDialog1Open.Title = "Select Data File to Modify"
        ' Display the Open dialog box
        CommonDialog1Open.ShowDialog()
        If MyResult = DialogResult.Cancel Then GoTo Error_WorkDir
        ' Display name of selected file
        txtWorkDir.Text = fs.GetParentFolderName(CommonDialog1Open.FileName)

Error_WorkDir:

    End Sub

    Private Sub ZipFiles(ByRef strFolder As String)
        Dim fld As Scripting.Folder

        On Error GoTo ZipFiles_Error

        ''Clears previous files from ZipTool
        'Zip1.FileStore.Clear()

        'Zip1.FileStore.Add(strFolder & "\*.pdf", , True)
        'Zip1.FileStore.Add(strFolder & "\*.xml", , True)
        ''Zip1.FileStore.Zip strHerndonZipDir & "\" & strZipFile

        ''MAB Change - Change to help stop the zip file corruption problem
        ''   Zipping the files into the DeneverZip directory first and then move
        ''   the files over to the Herndon directory may eliminate the possibility
        ''   of the the files being zipped and also the ingestion happening at the same time
        'Zip1.FileStore.Zip(strDenverZipDir & "\" & strZipFile)
        'If Not fs.FileExists(strHerndonZipDir & "\" & strZipFile) Then
        '    fs.MoveFile(strDenverZipDir & "\" & strZipFile, strHerndonZipDir & "\" & strZipFile)
        'Else
        '    fs.DeleteFile(strDenverZipDir & "\" & strZipFile)
        'End If

        'Label1.Text = "Zipping Files Completed : Zip Files Stored In - " & strHerndonZipDir
        'Label1.Refresh()
        'Exit Sub

ZipFiles_Error:

        Select Case Err.Number

            Case Else

                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description & vbCrLf & StrStatusInfo & vbCrLf & "Current Batch in process : " & strOrigBatchName)
                Exit Sub
        End Select

    End Sub

    Private Sub frmOPMXMLDF_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Set filesystem object to nothing to release object
        fs = Nothing
        frmStart.Enabled = True
    End Sub

    Private Sub frmOPMXMLDF_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim RetCode As Boolean = False
        On Error GoTo FormLoad_Error
        'set all constants
        Call SetConstants()
        Call CheckFileName(IniFileNameDF, RetCode)
        If Not RetCode Then
            'file does not exist
            MsgBox("The file DoSXMLDFGenerator.INI is missing in the folder where you're executing this program and is required. Please correct and run the program again.", MsgBoxStyle.Exclamation, "DoSXMLGenerator - DoSXMLGenerator.INI missing....")
            End
        End If
        Call ReadIniFile(IniFileNameDF)
        'Set filesystemobject
        fs = New Scripting.FileSystemObject

        prgBar1.Visible = False
        Label1.Visible = False

        If cmdStartDoc.Text = "START" Then
            txtInfo.BackColor = System.Drawing.ColorTranslator.FromOle(&HFF)
        End If

        Timer1.Interval = 1000 ' Set Timer interval.
        Timer1.Enabled = False

        Call UpdateFormVariablesFromGlobalDF()
        If cboData.Text = "" Then
            MsgBox("Please select directory poll time from Poll " & vbCrLf & "Directory control (value displayed in minutes) and save the configuration.")
        End If

         If Len(Trim(txtWorkDir.Text)) <> 0 Then
            ' Check directory structure. Create directories if they are not available
            If Not fs.FolderExists(txtWorkDir.Text & "TempZip") Then
                fs.CreateFolder(txtWorkDir.Text & "TempZip")
                strDenverZipDir = txtWorkDir.Text & "TempZip"
            Else
                strDenverZipDir = txtWorkDir.Text & "TempZip"
            End If

            If Not fs.FolderExists(txtWorkDir.Text & "CompletedFolders") Then
                fs.CreateFolder(txtWorkDir.Text & "CompletedFolders")
                strCompletedFoldersDir = txtWorkDir.Text & "CompletedFolders"
            Else
                strCompletedFoldersDir = txtWorkDir.Text & "CompletedFolders"
            End If

            If Not fs.FolderExists(txtWorkDir.Text & "CompletedDatabases") Then
                fs.CreateFolder(txtWorkDir.Text & "CompletedDatabases")
                strCompleteDatabaseDir = txtWorkDir.Text & "CompletedDatabases"
            Else
                strCompleteDatabaseDir = txtWorkDir.Text & "CompletedDatabases"
            End If

            If Not fs.FolderExists(txtWorkDir.Text & "PurgedFolders") Then
                fs.CreateFolder(txtWorkDir.Text & "PurgedFolders")
                strPurgeFolderDir = txtWorkDir.Text & "PurgedFolders"
            Else
                strPurgeFolderDir = txtWorkDir.Text & "PurgedFolders"
            End If

            If Not fs.FolderExists(txtWorkDir.Text & "DOSZIP") Then
                fs.CreateFolder(txtWorkDir.Text & "DOSZIP")
                strHerndonZipDir = txtWorkDir.Text & "DOSZIP"
            Else
                strHerndonZipDir = txtWorkDir.Text & "DOSZIP"

            End If

            If Not fs.FolderExists(txtWorkDir.Text & "Reporting") Then
                fs.CreateFolder(txtWorkDir.Text & "Reporting")
                strOPMReportDir = txtWorkDir.Text & "Reporting"
            Else
                strOPMReportDir = txtWorkDir.Text & "Reporting"
            End If

            If Not fs.FolderExists(txtWorkDir.Text & "Backup") Then
                fs.CreateFolder(txtWorkDir.Text & "Backup")
                strOPMBackup = txtWorkDir.Text & "Backup"
            Else
                strOPMBackup = txtWorkDir.Text & "Backup"
            End If

            If Not fs.FolderExists(txtWorkDir.Text & "ProblemFolder") Then
                fs.CreateFolder(txtWorkDir.Text & "ProblemFolder")
                strProblemFolderDir = txtWorkDir.Text & "ProblemFolder"
            Else
                strProblemFolderDir = txtWorkDir.Text & "ProblemFolder"
            End If

            If Not fs.FolderExists(txtWorkDir.Text & "ProblemDate") Then
                fs.CreateFolder(txtWorkDir.Text & "ProblemDate")
                strProbDateDir = txtWorkDir.Text & "ProblemDate"
            Else
                strProbDateDir = txtWorkDir.Text & "ProblemDate"
            End If

            If Not fs.FolderExists(txtWorkDir.Text & "DuplicateFolder") Then
                fs.CreateFolder(txtWorkDir.Text & "DuplicateFolder")
                strDuplicateDatabaseDir = txtWorkDir.Text & "DuplicateFolder"
            Else
                strDuplicateDatabaseDir = txtWorkDir.Text & "DuplicateFolder"
            End If

            Call WriteIniFile(IniFileNameDF)
        Else
            MsgBox(" Please select work directory to store files and then click Save Configuration")
        End If

FormLoad_Exit:
        Exit Sub

FormLoad_Error:

        Select Case Err.Number

            Case Else
                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description)
                GoTo FormLoad_Exit
                'Resume

        End Select
    End Sub

    Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick

        On Error GoTo Timer_Timer_Error

        ' Update time display.
        If lngCount < Val(cboData.Text) Then
            lngCount = lngCount + 1
            txtInfo.Text = "Time remaining to search directory : " & (Val(cboData.Text) - lngCount) + 1 & " seconds"
            txtInfo.Refresh()
        Else
            Timer1.Enabled = False
            Label1.Visible = True
            CheckDatabaseDir()
            prgBar1.Visible = False
            Label1.Visible = False
        End If

        Exit Sub

Timer_Timer_Error:

        Select Case Err.Number

            Case Else

                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description & vbCrLf & StrStatusInfo & vbCrLf & "Current Batch in process : " & strOrigBatchName)
                Exit Sub
        End Select
    End Sub

    Sub NewField(ByRef dbs As DAO.Database)
        Dim tdf As DAO.TableDef
        Dim fld As DAO.Field
        Dim blnCheck As Boolean
        Dim blnDocGuid As Boolean

        On Error GoTo NewField_Error

        ' Return reference to Employees table.
        tdf = dbs.TableDefs.Item("folderlist")
        ' Enumerate all fields in Fields collection of TableDef object.
        For Each fld In tdf.Fields
            'Check to see if Database Table has already been updated with PurgeStatus field
            If fld.Name = "PurgeStatus" Then
                blnCheck = True
                'Exit For
            End If

            If fld.Name = "DocGuid" Then
                blnDocGuid = True
                'Exit For
            End If

            If (blnCheck = True) And (blnDocGuid = True) Then
                Exit For
            End If
        Next fld

        If blnCheck = False Then
            ' Create new Field object.
            fld = tdf.CreateField("PurgeStatus")
            ' Set Type and Size properties of Field object.
            fld.Type = DAO.DataTypeEnum.dbBoolean
            ' Append field.
            tdf.Fields.Append(fld)
        End If

        If blnDocGuid = False Then
            ' Create new Field object.
            fld = tdf.CreateField("DocGuid")
            ' Set Type and Size properties of Field object.
            fld.Type = DAO.DataTypeEnum.dbText
            fld.Size = 40
            ' Append field.
            tdf.Fields.Append(fld)
        End If

        Exit Sub

NewField_Error:

        Select Case Err.Number

            Case Else

                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description & vbCrLf & StrStatusInfo & vbCrLf & "Current Batch in process : " & strOrigBatchName)
                Exit Sub
        End Select
    End Sub

    Function CheckFormType(ByRef strType As String) As String
        Dim strCheck() As String
        Dim i, ii As Short
        Dim strNew As String

        On Error GoTo CheckFormType_Error

        i = InStr(1, strType, "/")
        ii = InStr(1, strType, "&") 'Added an & check 1/28/08 - waj
        If i = 0 And ii = 0 Then
            CheckFormType = strType
        Else
            CheckFormType = "<![CDATA[" & Trim(strType) & "]]>"
        End If

CheckFormType_Exit:
        Exit Function

CheckFormType_Error:

        Select Case Err.Number

            Case Else
                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description, MsgBoxStyle.Exclamation, "Error Occurred in CheckFormType Function")
                GoTo CheckFormType_Exit
        End Select

    End Function
    Function GetBatchName(ByRef sInputSSN As String) As String
        'Creates a new batchname for a loose paper SSN with iteration if SSN has been scanned before
        Dim sBatchQuery As String
        Dim rsSSN As DAO.Recordset
        Dim iSSNCount As Short
        'Shut this down for a second
        sBatchQuery = "Select OPMStatus.BatchName, OPMStatus.SSN From OPMStatus "
        sBatchQuery = sBatchQuery & "Where OPMStatus.SSN = '" & sInputSSN & "' "
        sBatchQuery = sBatchQuery & "Order By OPMStatus.BatchName"
        rsSSN = db.OpenRecordset(sBatchQuery, DAO.RecordsetTypeEnum.dbOpenDynaset)
        If rsSSN.RecordCount > 0 Then
            rsSSN.MoveFirst()
            rsSSN.MoveLast() 'get the highest value iteration
            iSSNCount = Val(Mid(rsSSN.Fields("BatchName").Value, 13, 6))
            GetBatchName = "DOF" & sInputSSN & VB6.Format(iSSNCount + 1, "000000") & "M"
        Else 'First one
            GetBatchName = "DOF" & sInputSSN & "000000M"
        End If
        rsSSN.Close()
    End Function
    Sub AddReconRecord(ByVal BFN As String, ByVal SSN As String, ByVal FN As String, ByVal Pages As Int32, ByVal GUID As String)
        Dim sQuery, sFN As String
        On Error GoTo AddReconRecord_Error
        sFN = Replace(FN, "'", "''")
        sQuery = "INSERT INTO Reconcilliation(BatchFileName, SSN, FormName, Pages, GUID, GUIDDate) "
        sQuery = sQuery & "VALUES('" & BFN & "', '" & SSN & "', '" & sFN & "', " & Pages & ", '" & GUID & "', '" & Now() & "')"
        db.Execute(sQuery, 64) '64 = dbSQLPassThrough
AddReconRecord_Exit:
        Exit Sub
AddReconRecord_Error:
        Select Case Err.Number
            Case Else
                MsgBox("Error Number : " & Err.Number & vbCrLf & Err.Description & vbCrLf & StrStatusInfo & vbCrLf & "Current Batch in process : " & strOrigBatchName)
                GoTo AddReconRecord_Exit
        End Select
    End Sub
End Class