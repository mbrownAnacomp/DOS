VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ASCIITextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const M_TEXTCLASS = "XML Class"
Const MAX_OPEN_ATTEMPTS = 10

'===================
' Global Variables
'===================
Public ErrorLineNum As Integer

'*************************************************
' ReleaseIndexes
'-------------------------------------------------
' Purpose:  This routine writes the batch class
'           name, doc class name, Index Values,
'           and image path to the text file.
' Inputs:   TheData           ReleaseData object
'           strImageFilePath  path to the images
' Outputs:  None
' Returns:  None
' Notes:    Because the Values collection may not
'           be in the correct order, we load the
'           data into an array where Destination
'           indicates the array index.  We then
'           output the Index Values in order.
'*************************************************
Sub ReleaseIndexes(TheData As ReleaseData, ByVal strImageFilePath, ByVal sReleasePath As String)
    Dim I As Integer
    Dim sIndexFile, sQuery As String
    Dim sOutput As String
    Dim sSource As String
    Dim sValue As String
    Dim sPath As String
    Dim hFile As Integer
    Dim nOpenAttempts As Integer
    Dim oValue As Value
    Dim SortedValues() As T_Value
    Dim nValues As Integer
    Dim sTemp As String
    'OPM Specific
    Dim rsSearch As String
    Dim sOriginalFormNameNumber As String
    Dim sSSN As String
    Dim sEffectiveDate As String
    Dim sNOACode1 As String
    Dim sNOACode2 As String
    Dim sDocTypeC As String 'not sure what this is for, Lili?
    Dim sVSide As String
    Dim sDuplex As String
    Dim sPurge As String
    Dim sType As String
    Dim sPurgeComment As String
    Dim sOtherType As String
    Dim sImageCount As String
    Dim rsNOACode1 As String
    Dim sException As String
    Dim sBarcode As String
    Dim sEffFrom As String, sEffTo As String
    Dim rsFind As String
    Dim sepPos As Integer, sResult As String
        
        On Error GoTo RI_LogAndPropError
        
        ' Resize the array of sorted values to
        ' the possible number in the collection
        If TheData.Values.Count > 0 Then
1050        ReDim SortedValues(TheData.Values.Count - 1)
        Else
1060        ReDim SortedValues(0)
        End If
        nValues = 0
            
        ' Loop through each Index Value ignoring the PDF values.
        For Each oValue In TheData.Values
            If left$(oValue.Destination, 3) <> "PDF" Then
                ' Destination is used as the index into the array
                I = Val(oValue.Destination)
1070            SortedValues(I).DataType = oValue.DataType
1080            SortedValues(I).Destination = oValue.Destination
1090            SortedValues(I).SourceType = oValue.SourceType
1100            SortedValues(I).SourceName = oValue.SourceName
1110            SortedValues(I).Value = oValue.Value

                ' Count the number of values actually kept
                nValues = nValues + 1
            End If
        Next oValue
        
        ' Resize the array of values to the number actually kept
        If nValues > 0 Then
            ReDim Preserve SortedValues(nValues - 1)
        End If

        I = 0
        While I < nValues
        
            ' Get the Index Value source
1120        Select Case SortedValues(I).SourceType


                Case KFX_REL_INDEXFIELD
1140                sSource = SortedValues(I).SourceName
                    If sSource = DOC_TYPE Then
                        rsSearch = Trim(SortedValues(I).Value)
                        If left(rsSearch, 5) = "OTHER" Then
                            sOtherType = Trim(Mid(rsSearch, 6))
                            rsSearch = "OTHER"
                        ElseIf UCase(left(rsSearch, 11)) = "BARCODESIDE" Then
                            sSide = Mid(rsSearch, 11)
                        ElseIf UCase(rsSearch) = "BARCODESSN" Then   '1st page of batch
                            sSide = "Value not set"
                        ElseIf UCase(left(rsSearch, 7)) = "BARCODE" And right(rsSearch, 3) <> "SSN" Then  'Old barcode side format
                            sSide = Mid(rsSearch, 8)
                        End If
                    ElseIf sSource = SSN Then
                        sSSN = Replace(Trim(SortedValues(I).Value), "-", "")
                        If sSSN = "" Then sSSN = "999999999" 'Barcode field I hope
                    ElseIf sSource = EFF_DATE Then
                        sEffectiveDate = Trim(SortedValues(I).Value)
                    ElseIf sSource = NOACODE_1 Then
                        sNOACode1 = Trim(SortedValues(I).Value)
                    ElseIf sSource = NOACODE_2 Then
                        sNOACode2 = Trim(SortedValues(I).Value)
                    'Eliminated as an index field because of KTM licensing
                    'ElseIf sSource = SIDE Then
                    '    sSide = Trim(SortedValues(I).Value)
                    '    If left(sSide, 7) = "Barcode" Then sSide = right(sSide, Len(sSide) - 7)
                    ElseIf sSource = EXCEPTION Then
                        sException = Trim(SortedValues(I).Value)
                    End If
    
                Case KFX_REL_TEXTCONSTANT
1160                'sSource = Chr$(34) & "[TEXT]" & Chr$(34)
                    
                Case NO_LINK, KFX_REL_DOCUMENTID
                    ' These should never happen since we force the user
                    ' to delete all "unlinked" index values in Release Setup
                    ' and because we don't display the Document ID in the
                    ' top level link menu.  If either of these facts changes,
                    ' then this area needs to be modified.
                    sSource = ""
                Case Else
                    sSource = ""
            End Select

            I = I + 1
        Wend
        If IsDate(sEffectiveDate) Then
            sEffectiveDate = Format(sEffectiveDate, "d-mmm-yyyy")
        Else
            sEffectiveDate = "1-Jan-1901"
        End If
        
        'Do database lookup
        'rsFind = "FormNameNumber = '" & rsSearch & "'"
        sQuery = "Select * From Newmaster Where FormNameNumber='" & rsSearch & "' Order By FormNameNumber, NOAC"
        Set rs = db.OpenRecordset(sQuery, dbOpenDynaset, 512) '512 = dbSeeChanges
        With rs
            If .RecordCount > 0 Then
                'Grab the unedited form number
                If IsNull(![OriginalFormNameNumber]) Then
                    sOriginalFormNameNumber = " "
                Else
                    sOriginalFormNameNumber = Trim(![OriginalFormNameNumber])
                End If
                Select Case rsSearch
                    Case "SF50"
                        'Catch Alpha noa codes and convert
                        If UCase(sNOACode1) = "BRC" Then
                            sNOACode1 = "982"
                        ElseIf UCase(sNOACode1) = "CUC" Then
                            sNOACode1 = "995"
                        ElseIf UCase(sNOACode1) = "CUD" Then
                            sNOACode1 = "996"
                        ElseIf UCase(sNOACode1) = "CUN" Then
                            sNOACode1 = "997"
                        ElseIf UCase(sNOACode1) = "M2B" Then
                            sNOACode1 = "999"
                        ElseIf UCase(sNOACode1) = "MSC" Then
                            sNOACode1 = "990"
                        ElseIf UCase(sNOACode1) = "RCD" Then
                            sNOACode1 = "979"
                        ElseIf UCase(sNOACode1) = "S1B" Then
                            sNOACode1 = "993"
                        ElseIf UCase(sNOACode1) = "TSP" Then
                            sNOACode1 = "998"
                        ElseIf UCase(sNOACode1) = "ZZZ" Then
                            sNOACode1 = "994"
                        End If
                        
                        'Not sure if there are alpha NOA2s but what the heck
                        If UCase(sNOACode2) = "BRC" Then
                            sNOACode2 = "982"
                        ElseIf UCase(sNOACode2) = "CUC" Then
                            sNOACode2 = "995"
                        ElseIf UCase(sNOACode2) = "CUD" Then
                            sNOACode2 = "996"
                        ElseIf UCase(sNOACode2) = "CUN" Then
                            sNOACode2 = "997"
                        ElseIf UCase(sNOACode2) = "M2B" Then
                            sNOACode2 = "999"
                        ElseIf UCase(sNOACode2) = "MSC" Then
                            sNOACode2 = "990"
                        ElseIf UCase(sNOACode2) = "RCD" Then
                            sNOACode2 = "979"
                        ElseIf UCase(sNOACode2) = "S1B" Then
                            sNOACode2 = "993"
                        ElseIf UCase(sNOACode2) = "TSP" Then
                            sNOACode2 = "998"
                        ElseIf UCase(sNOACode2) = "ZZZ" Then
                            sNOACode2 = "994"
                        End If
                        If sEffectiveDate = "" Or sEffectiveDate = "1-Jan-1901" Then
                            sType = "EXCEPTION"
                            sVSide = Trim(![VirtualSide])
                        Else
                            If sNOACode1 <> "" Then
                                sResult = checkForDuplicates(sEffectiveDate, sNOACode1, rsSearch)
                                sepPos = InStr(sResult, "|")
                                If sepPos > 0 Then
                                    sType = left(sResult, sepPos - 1)
                                    sVSide = Mid(sResult, sepPos + 1)
                                Else
                                    sType = Trim(sResult)
                                    sVSide = Trim(![VirtualSide])
                                End If
                            Else
                                sType = "EXCEPTION"
                                sVSide = Trim(![VirtualSide])
                            End If
                            'If the Effective date isn't an "Effective Date" then
                            If UCase(sException) = "TRUE" Then sType = "EXCEPTION"
                        End If
                        If IsNull(![Purge]) Then
                            sPurge = ""
                        Else
                            sPurge = Trim(![Purge])
                        End If
                    Case "SF50A"
                        If sEffectiveDate = "" Or sEffectiveDate = "1-Jan-1901" Then
                            sType = "EXCEPTION"
                            sVSide = Trim(![VirtualSide])
                        Else
                            If sNOACode1 <> "" Then
                                sResult = checkForDuplicates(sEffectiveDate, sNOACode1, rsSearch)
                                sepPos = InStr(sResult, "|")
                                If sepPos > 0 Then
                                    sType = left(sResult, sepPos - 1)
                                    sVSide = Mid(sResult, sepPos + 1)
                                Else
                                    sType = Trim(sResult)
                                    sVSide = Trim(![VirtualSide])
                                End If
                            Else
                                sType = "EXCEPTION"
                                sVSide = Trim(![VirtualSide])
                            End If
                        End If
                        If IsNull(![Purge]) Then
                            sPurge = ""
                        Else
                            sPurge = Trim(![Purge])
                        End If
                        If UCase(sException) = "TRUE" Then sType = "EXCEPTION"
                    Case "SF50B"
                        If sEffectiveDate = "" Or sEffectiveDate = "1-Jan-1901" Then
                            sType = "EXCEPTION"
                            sVSide = Trim(![VirtualSide])
                        Else
                            If sNOACode1 <> "" Then
                                sResult = checkForDuplicates(sEffectiveDate, sNOACode1, rsSearch)
                                sepPos = InStr(sResult, "|")
                                If sepPos > 0 Then
                                    sType = left(sResult, sepPos - 1)
                                    sVSide = Mid(sResult, sepPos + 1)
                                Else
                                    sType = Trim(sResult)
                                    sVSide = Trim(![VirtualSide])
                                End If
                            Else
                                sType = "EXCEPTION"
                                sVSide = Trim(![VirtualSide])
                            End If
                        End If
                        If IsNull(![Purge]) Then
                            sPurge = ""
                        Else
                            sPurge = Trim(![Purge])
                        End If
                        If UCase(sException) = "TRUE" Then sType = "EXCEPTION"
                    Case "SF52"
                        If sEffectiveDate = "" Or sEffectiveDate = "1-Jan-1901" Then
                            sType = "EXCEPTION"
                            sVSide = Trim(![VirtualSide])
                        Else
                            If sNOACode1 <> "" Then
                                sResult = checkForDuplicates(sEffectiveDate, sNOACode1, rsSearch)
                                sepPos = InStr(sResult, "|")
                                If sepPos > 0 Then
                                    sType = left(sResult, sepPos - 1)
                                    sVSide = Mid(sResult, sepPos + 1)
                                Else
                                    sType = Trim(sResult)
                                    sVSide = Trim(![VirtualSide])
                                End If
                            Else
                                sType = "EXCEPTION"
                                sVSide = Trim(![VirtualSide])
                            End If
                        End If
                        If IsNull(![Purge]) Then
                            sPurge = ""
                        Else
                            sPurge = Trim(![Purge])
                        End If
                        If UCase(sException) = "TRUE" Then sType = "EXCEPTION"
                    Case "PS50"
                        If sEffectiveDate = "" Or sEffectiveDate = "1-Jan-1901" Then
                            sType = "EXCEPTION"
                            sVSide = Trim(![VirtualSide])
                        Else
                            If sNOACode1 <> "" Then
                                sResult = checkForDuplicates(sEffectiveDate, sNOACode1, rsSearch)
                                sepPos = InStr(sResult, "|")
                                If sepPos > 0 Then
                                    sType = left(sResult, sepPos - 1)
                                    sVSide = Mid(sResult, sepPos + 1)
                                Else
                                    sType = Trim(sResult)
                                    sVSide = Trim(![VirtualSide])
                                End If
                            Else
                                sType = "EXCEPTION"
                                sVSide = Trim(![VirtualSide])
                            End If
                        End If
                        If IsNull(![Purge]) Then
                            sPurge = ""
                        Else
                            sPurge = Trim(![Purge])
                        End If
                        If UCase(sException) = "TRUE" Then sType = "EXCEPTION"
                    Case "OTHER"
                        sType = sOtherType
                        If IsNull(![VirtualSide]) Then
                            sVSide = " "
                        Else
                            sVSide = Trim(![VirtualSide])
                            If sVSide = "" Then sVSide = " "
                        End If
                        sPurge = ""
                        If UCase(sException) = "TRUE" Then sType = "EXCEPTION"
                    Case Else
                        If sEffectiveDate <> "" And sEffectiveDate <> "1-Jan-1901" Then
                            If IsNull(![Type]) Then
                                sType = " "
                            Else
                                sType = Trim(![Type])
                                If sType = "" Then sType = " "
                            End If
                        Else
                            sType = "EXCEPTION"
                        End If
                        If IsNull(![VirtualSide]) Then
                            sVSide = " "
                        Else
                            sVSide = Trim(![VirtualSide])
                            If sVSide = "" Then sVSide = " "
                        End If
                        If IsNull(![Purge]) Then
                            sPurge = ""
                        Else
                            sPurge = Trim(![Purge])
                        End If
                        If UCase(sException) = "TRUE" Then sType = "EXCEPTION"
                End Select
                If IsNull(![Duplex]) Then
                    sDuplex = "Simplex"
                ElseIf Trim(![Duplex]) = "" Then
                    sDuplex = "Simplex"
                Else
                    sDuplex = "Duplex"
                End If

                sImageCount = Format(Str(imageCount), "0000")
            Else
                'Somehow the FormNumber isn't in the database
                sOriginalFormNameNumber = rsSearch
                sType = "EXCEPTION"
                sImageCount = Format(Str(imageCount), "0000")
            End If
        End With
        rs.Close
        'DoS Specific PAYROLL Side rule
        If UCase(sSide) = "PAYROLL" And UCase(sVSide) <> "PERMANENT" Then sVSide = "Payroll"
        'DoS Specific Other rule
        If UCase(rsSearch) = "OTHER" And UCase(sSide) = "LEFT" Then sVSide = "Miscellaneous"
        ' Get the image file name
1180    sPath = strImageFilePath
        
        ' Now we are ready to write to the file.  Open the file
        ' for Append which will create it if it doesn't exist.
        ' Also place a Read/Write Lock so no other process can
        ' access the file while we have it open.
1190    hFile = FreeFile
        
        On Error GoTo RI_FileOpenError
        nOpenAttempts = 1
1195    If firstDoc Then 'Get date/time stamped file name on first document and hold it thereafter
            firstDoc = False
        End If
1200    'If Dir(indexFileName) <> "" Then Kill indexFileName 'no repeats allowed
        'Open indexFileName For Output As #hFile

        ' Write the output to the file and close the file
        On Error GoTo RI_LogAndPropError
1210    'Print #hFile, sOutput
        'Add info to batch databse table
        With rsNew
            .AddNew
                ![SSN] = left(sSSN, 9)
                ![FormNumber] = rsSearch
                If sOriginalFormNameNumber = "" Then sOriginalFormNameNumber = " "
                ![OriginalFormNumber] = sOriginalFormNameNumber
                If IsDate(sEffectiveDate) Then
                    ![effdate] = Format(sEffectiveDate, "yyyy-mm-dd")
                Else
                    ![effdate] = "1901-01-01"
                End If
                
                If sNOACode1 = "" Then
                    sNOACode1 = " "
                ElseIf IsNumeric(sNOACode1) Then
                    sNOACode1 = Trim(Str(Val(sNOACode1)))
                End If
                ![NOA1] = sNOACode1
                If sNOACode2 = "" Then
                    sNOACode2 = " "
                ElseIf IsNumeric(sNOACode2) Then
                    sNOACode2 = Trim(Str(Val(sNOACode2)))
                End If
                ![NOA2] = sNOACode2
                If sSide = "" Then sSide = " "
                ![SourceSide] = sSide
                If sType = "" Then sType = " "
                ![FormType] = sType
                If sVSide = "" Then sVSide = " "
                ![DeliverySide] = sVSide
                If sDuplex = "Duplex" Then
                    ![Duplex] = True
                Else
                    ![Duplex] = False
                End If
                If sPurge = "Yes" Then
                    ![Purge] = True
                Else
                    ![Purge] = False
                End If
                'If sImageCount = "" Then sImageCount = "0"
                '![Pages] = sImageCount
                ![Pages] = TheData.ImageFiles.Count
                If right(sReleasePath, 1) = "\" Then
                    ![PDFFileName] = sReleasePath & strImageFilePath 'Actually the PDF File Name
                Else
                    ![PDFFileName] = sReleasePath & "\" & strImageFilePath 'Actually the PDF file Name
                End If
            .Update
        End With
1220    Close #hFile

        Exit Sub

'---------------
' Error Handler
'---------------
RI_FileOpenError:
            
        ' If the path does not exist, don't try to re-open the file
        ' Append the missing path to the error message
        If Err = 76 Then
            Err.Description = "'" & sIndexFile & "' " & LoadResString(MSG_INVALIDPATH)
            GoTo RI_LogAndPropError
        End If
            
        ' The file may be locked by another process.
        ' Retry the defined number of times.
        If nOpenAttempts < MAX_OPEN_ATTEMPTS Then
            nOpenAttempts = nOpenAttempts + 1
            Resume
        Else
            Err.Description = Err.Description & ": " & LoadResString(MSG_COULDNOTOPENFILE) & " '" & sIndexFile & "'."
        End If
        
RI_LogAndPropError:

1230    If hFile <> 0 Then Close hFile
        
        ' Remember the line number of the error
        If (ErrorLineNum = 0) Then
            ErrorLineNum = Erl
        End If
        
        ' Let the caller deal with the error.
        Call Err.Raise(Err, M_TEXTCLASS + ".ReleaseIndexes:" + Err.Source, Err.Description)
End Sub

'=====================
'Some NOA codes are in the dB more than once in which case the effective date
'range needs to be check against the effective date
'=====================
Public Function checkForDuplicates(sEffDate As String, sNOA As String, sfNum As String) As String
    Dim dEffDate As Date, sNewNoa As String, sFormNumber As String, qRs As Recordset, sQuery As String, vSide As String
    sFormNumber = "'" & sfNum & "'"
    sNewNoa = "'" & sNOA & "'"
    dEffDate = CDate(sEffDate)
    sQuery = "SELECT [NewMaster].[Type], NewMaster.[VirtualSide] FROM NewMaster WHERE [NewMaster].[FormNameNumber]=" & sFormNumber
    sQuery = sQuery & " And [NewMaster].[NOAC]=" & sNewNoa
    sQuery = sQuery & " And [NewMaster].[NOA Eff From Date] <=#" & dEffDate & "#"
    sQuery = sQuery & " And [NewMaster].[NOA Eff To Date] >=#" & dEffDate & "#"
    Set qRs = db.OpenRecordset(sQuery, dbOpenDynaset)
    With qRs
        If qRs.RecordCount > 0 Then
            If Not IsNull(![Type]) Then
                checkForDuplicates = Trim(![Type]) & "|" & Trim(![VirtualSide])
            Else
                checkForDuplicates = "EXCEPTION"
            End If
        Else
            checkForDuplicates = "EXCEPTION"
        End If
    End With
    qRs.Close
End Function
