VERSION 5.00
Begin VB.UserControl ctlButtonCmd 
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   675
   ScaleHeight     =   705
   ScaleWidth      =   675
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "ctlButtonCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_oApp As AscentCaptureModule.Application
Public beginTime 'For capturing the elapsed time
Public sBatchName, sConnect, sDSN, sDB, sWinAuth, sUID, sPWD  As String
Public strDelType As String
Public oDB As Database, oRS As Recordset, sQuery As String

'************************************************
'************************************************
'*****  Function:   Property Set Application()
'*****  Purpose:    Sets the application object - Base of hierarchy
'*****  Inputs:     None.
'*****  Outputs:    None.
'*****  Errors:
'************************************************
'************************************************
Public Property Set Application(oApp As AscentCaptureModule.Application)
    
        On Error GoTo ApplicationError
        
        '***** Set the application object. *****
        Set m_oApp = oApp
        '***** Create the menu options. *****
        Call InitialMenus

Exit Property
ApplicationError:

        MsgBox "Application Set Error: "
        Call HandleError("Error creating application object", "Appplication Set")
End Property

'************************************************
'************************************************
'*****  Function:   ActionEvent()
'*****  Purpose:    Receive an event and proceed accordingly.
'*****  Inputs:     nActionNumber - number assigned to
'*****              event
'*****  Outputs:    vArgument - Menu item.
'*****              pnCancel - Response to the event.
'*****  Errors:
'************************************************
'************************************************
Public Function ActionEvent(ByVal nActionNumber As Integer, _
                            ByRef vArgument As Variant, _
                            ByRef pnCancel As Integer) As Integer
        
        Dim oWrkODBC As Workspace, sSP_BatchReadyForScan As String
        Dim intRet As Integer, strConnect As String, strRet As String, sElapsedTime As String
        Dim elapsedTime, hTime, mTime, sTime, iElapsedTime As Integer
        Dim bhTime, bmTime, bsTime, biElapsedTime As Integer, ciElapsedTime As Integer
        Dim sUserID As String, sSP_BatchScanCompleted As String
        Dim sMachineName As String, iResponse As Integer, elapsedTimeSuspended, timeTest, vElapsedTime
        Dim fso As FileSystemObject, sBasePath As String, iPaths As Integer, x As Integer
        Dim sSplitPath() As String, sQuery As String, iRet As Integer
        
        '***** Find out which event was performed. *****
        Select Case nActionNumber
            '*****************************
            '***** MENU CLICK EVENTS *****
            '*****************************
            '***** This is a menu event from the batch contents tree. *****
            Case AscentCaptureModule.KfxOcxEventMenuClicked
                Select Case vArgument
                End Select
            '************************
            '***** BATCH EVENTS *****
            '************************
            '***** When the batch opens, get the current info. *****
            
            Case AscentCaptureModule.KfxOcxEventBatchOpened
                If UCase(Left(m_oApp.ActiveBatch.ClassName, 3)) = "DOS" Then
                    Screen.MousePointer = vbHourglass
                    beginTime = Now
                    sBatchName = m_oApp.ActiveBatch.Name
                    sDSN = m_oApp.ActiveBatch.BatchFields("ODBC DSN").DefaultValue
                    sDB = m_oApp.ActiveBatch.BatchFields("ODBC DB").DefaultValue
                    sConnect = "ODBC;DSN=" & sDSN & ";DB=" & sDB
                    sWinAuth = m_oApp.ActiveBatch.BatchFields("WinAuth").DefaultValue
                    If sWinAuth = "False" Then
                        sUID = m_oApp.ActiveBatch.BatchFields("ODBC UID").DefaultValue
                        sPWD = m_oApp.ActiveBatch.BatchFields("ODBC PWD").DefaultValue
                        sConnect = sConnect & ";UID=" & sUID & ";PWD=" & sPWD
                    End If
                    Set oDB = OpenDatabase("", False, False, sConnect)
                    'Check to see if this batch has a valid name
                    sQuery = "Select * From Batch Where BatchName ='" & sBatchName & "'"
                    Set oRS = oDB.OpenRecordset(sQuery, dbOpenDynaset, 512)
                    If oRS.RecordCount > 0 Then 'Batchname exists - check status
                        If oRS.Fields("BatchStatus") = 1 _
                        Or oRS.Fields("BatchStatus") = 3 _
                        Or oRS.Fields("BatchStatus") = 4 Then '1=Ready to scan,3=suspended,4=scan complete
                            sUserID = "{Operator Name}"
                            m_oApp.TranslateAscentCaptureValues sUserID
                             'Master list Release db/table
                            sQuery = "UPDATE Batch Set BatchStatus = 2, ScanOperator ='" & sUserID & "' Where BatchName ='" & sBatchName & "'"
                            oDB.Execute sQuery, 64
                       
                        Else
                            MsgBox "This Batch is in the system but has an Incorrect status" & vbCrLf & _
                                   "Batch status =" & Str(oRS.Fields("BatchStatus")) & vbCrLf & _
                                   "This batch will be deleted" & vbCrLf & _
                                   "Check with your supervisor or try again", 16, "Batch status Error!"
                            strDelType = "Incorrect Status" 'Communicate with delete event
                            Timer1.Enabled = True 'Calls the batch close/deletion process
                        End If
                    Else
                        MsgBox "This is not a vaild Batchname" & vbCrLf & _
                               "This batch will be deleted" & vbCrLf & _
                               "Check with your supervisor or try again", 16, "Batch Name Error!"
                        strDelType = "Incorrect Batchname"
                        Timer1.Enabled = True 'Calls the batch close/deletion process
                    End If
                End If
            Case AscentCaptureModule.KfxOcxEventBatchClosing
                If Left(m_oApp.ActiveBatch.ClassName, 3) = "DoS" Then
                    'good to go
                    Screen.MousePointer = vbHourglass
                    sQuery = "UPDATE Batch Set BatchStatus = 4, ScanDateTime ='" & Now() & "' Where BatchName ='" & sBatchName & "'"
                    oDB.Execute sQuery, 64
                    oDB.Close
                    'Let the Batch close
                    Screen.MousePointer = 0
                End If
            Case AscentCaptureModule.KfxOcxEventBatchSuspending
                If Left(m_oApp.ActiveBatch.ClassName, 3) = "DoS" Then
                    elapsedTimeSuspended = Now - beginTime
                    m_oApp.ActiveBatch.CustomStorageString("SuspendTime") = Format(elapsedTimeSuspended, "hh:mm:ss")
                    sQuery = "UPDATE Batch Set BatchStatus = 3 Where BatchName ='" & sBatchName & "'"
                    oDB.Execute sQuery, 64
                End If
            Case AscentCaptureModule.KfxOcxEventBatchDeleting
                If Left(m_oApp.ActiveBatch.ClassName, 3) = "DoS" _
                And strDelType <> "Incorrect Batchname" And strDelType <> "Incorrect Status" _
                Then  'Don't change batch status for Incorrect Status or Batchname
                    sQuery = "UPDATE Batch Set BatchStatus = 1 Where BatchName ='" & sBatchName & "'"
                    oDB.Execute sQuery, 64
                End If
            Case AscentCaptureModule.KfxOcxEventBatchRejecting
                Call DestroyObjects
            '***************************
            '***** DOCUMENT EVENTS *****
            '***************************
            '***** When the document opens, get the current info. *****
            Case AscentCaptureModule.KfxOcxEventDocumentOpened
            Case AscentCaptureModule.KfxOcxEventDocumentValidating
            Case AscentCaptureModule.KfxOcxEventDocumentClosing
            '************************
            '***** FIELD EVENTS *****
            '************************
            Case AscentCaptureModule.KfxOcxEventFieldEntered
            Case AscentCaptureModule.KfxOcxEventFieldExiting
            '************************
            '***** MISC. EVENTS *****
            '************************
            Case AscentCaptureModule.KfxOcxEventDataEntryModeEntered
            Case AscentCaptureModule.KfxOcxEventScanBatchStarted
            Case AscentCaptureModule.KfxOcxEventScanStopped
            Case Else
        End Select
    
End Function
'************************************************
'************************************************
'*****  Function:   DestroyObjects()
'*****  Purpose:    Destroy any active object for the batch.
'*****  Inputs:     None.
'*****  Outputs:    None.
'*****  Errors:
'************************************************
'************************************************
Public Sub DestroyObjects()
        Set m_oApp = Nothing
End Sub

'************************************************
'************************************************
'*****  Function:   InitialMenus()
'*****  Purpose:    Create batch content menu items.
'*****  Inputs:     None.
'*****  Outputs:    None.
'*****  Errors:
'************************************************
'************************************************
Public Sub InitialMenus()

    '***** Display the panel. *****
    'm_oApp.ShowWindow (True)
End Sub
'************************************************
'************************************************
'*****  Function:   HandleError()
'*****  Purpose:    Try to display a message.  The error
'*****              is handled by the API.
'*****  Inputs:
'*****  Outputs:
'*****  Errors:
'************************************************
'************************************************
Public Sub HandleError(sMessage As String, sFunction)

            '***** Display message for err handling purpose.
        MsgBox sFunction & _
                vbCrLf & sMessage & _
                vbCrLf & "Error " & Err.Number & ", " & Err.Description & ", " & _
                vbCrLf & "Line: " & Erl, _
                vbExclamation, _
                Err.Source
        '***** Log the error to the AC error log.
        Call m_oApp.LogError(Err.Number, 0, 0, sMessage, Err.Source & "." & sFunction, Erl)
    
End Sub
Private Sub Timer1_Timer()
    Timer1.Enabled = False
    m_oApp.DeleteBatch
    SendKeys "%BN"
End Sub
Private Sub Pause(pauseTime As Variant)
    Dim Start
    Start = Timer
    Do While Timer < Start + pauseTime
        DoEvents
    Loop
End Sub
Private Sub UserControl_Terminate()
    Call DestroyObjects
End Sub
