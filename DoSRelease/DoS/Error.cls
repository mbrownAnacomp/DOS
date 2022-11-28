VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Force late bind as we won't know if this module
' is being used in Setup Release or Release
Public DataObject As Object

Public TitleString As String

'*************************************************
' LogTheError
'-------------------------------------------------
' Purpose:  This routine logs an error to the
'           Capture error log and may optionally
'           re-raise the error and/or display it
'           to the user.
' Inputs:   ErrNum      Error Number
'           ErrMsg      Error Message
'           SourceFile  The source module in which
'                       the error occurred
'           LineNo      Line Number
'           ReRaise     If true, will re-raise the
'                       error on the way out
'           Display     If true, will show the
'                       error in a message box
' Outputs:  None
' Returns:  None
' Notes:    The initialization routine for the
'           error handler must be performed
'           before calling this routine.
'*************************************************
Public Sub LogTheError(ErrNum As Long, _
                ErrMsg As String, _
                SourceFile As String, _
                LineNo As Integer, _
                ReRaise As Boolean, _
                Display As Boolean)
    On Error GoTo LE_Error
    
        ' First log the error through the
        ' setup or release data object
        If (Not DataObject Is Nothing) Then
1000        Call DataObject.LogError(ErrNum, 0, 0, ErrMsg, SourceFile, LineNo)
        End If
        
        ' If the caller wants a message box displayed for this
        ' error, pop it up with the standard title
        If (Display Or DataObject Is Nothing) Then
            Call MsgBox(ErrMsg & vbCr & "(#" & Format$(ErrNum) & ")", _
                        vbOKOnly + vbExclamation, TitleString)    ' DO NOT LOCALIZE
        End If
        
LogTheError_Exit:
        On Error GoTo 0
        ' If the user wants the error re-raised to propogate
        ' it up to the calling routine, raise the error again
        If ReRaise Then
            Call Err.Raise(ErrNum, SourceFile, ErrMsg)
        End If
        
        Exit Sub

LE_Error:
    Call MsgBox(LoadResString(MSG_ANERROR) + CStr(Err) + " - " + Err.Description + _
                LoadResString(MSG_HAPPENEDLOGGING) + vbCrLf + _
                ErrMsg & vbCr & "(#" & Format$(ErrNum) & ")", _
                vbOKOnly + vbExclamation, _
                TitleString)    ' DO NOT LOCALIZE
    Resume LogTheError_Exit
End Sub

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
    TitleString = LoadResString(TITLE_ERRORMSGBOX)
End Sub
