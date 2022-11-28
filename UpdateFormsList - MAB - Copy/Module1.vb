Option Strict Off
Option Explicit On
Module Module1
	
	'Global File Path Directories
	Public strPath As String
	Public strNewLoc As String
	Public strPoll As String
	Public intBackup As String
	'Folders directory locations
	Public strDenverZipDir As String
	Public strHerndonZipDir As String
	Public strOPMReportDir As String
	Public strCompletedFoldersDir As String
	Public strCompleteDatabaseDir As String
	Public strOPMBackup As String
	Public strPurgeFolderDir As String
	Public strProblemFolderDir As String
	Public strProbDateDir As String
	Public strDuplicateDatabaseDir As String
	
	
	
	'Filename variables
    Public IniFileName, IniFileNameDF As String
	'Declare INI file read and write functions
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
	
	Public Sub SetConstants()
		Dim CompileVersion As Object
		Dim CompileDate As Object
		Dim InstallPath As String
		'set constants
        CompileDate = "February 1, 2011"
        CompileVersion = "3.0"
		
		If Right(My.Application.Info.DirectoryPath, 1) <> "\" Then
			InstallPath = My.Application.Info.DirectoryPath & "\"
		Else
			InstallPath = My.Application.Info.DirectoryPath
		End If
        IniFileName = InstallPath & "DoSXMLGenerator.INI"
        IniFileNameDF = InstallPath & "DoSXMLDFGenerator.INI"
	End Sub
	
	Public Sub CheckFileName(ByRef InName As String, ByRef ReturnCode As Boolean)
		'check to see if a file exists
		'if the file does not exist then do.....
		'    if return code=true => display a message and set returncode to false
		'    if returncode=false => no message and set return code to false
		Dim name As String
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		name = Dir(InName, FileAttribute.Normal)
		Dim Response As Short
		If name = "" Then
			'file does not exist
			If ReturnCode Then
				Response = MsgBox("The file:" & vbCrLf & Chr(34) & InName & Chr(34) & vbCrLf & "does not exist!" & vbCrLf & "Check the name and try again!", MsgBoxStyle.Exclamation, "Bad File Name!")
			End If
			ReturnCode = False
		Else
			'file exists
			ReturnCode = True
		End If
	End Sub

	Public Sub ReadIniFile(ByRef InCfgFileName As String)
		'read parameters from an ini file.
		'-------------------
		'[Main]
		'
		'-------------------
		Dim SectionName As String
		Dim ParmName As String
		Dim lpDefault As String
		Dim ParmValue As String
		Dim ParmSize As Integer
		Dim RetCode As Integer
		
		ParmSize = 75
		SectionName = "Main"
		
		'Output Location
		lpDefault = ""
		ParmName = "OutPut_Location"
		ParmValue = Space(ParmSize)
		RetCode = GetPrivateProfileString(SectionName, ParmName, lpDefault, ParmValue, ParmSize, InCfgFileName)
		strPath = UCase(Trim(Left(ParmValue, RetCode)))
		If Right(strPath, 1) <> "\" And RetCode > 0 Then
			strPath = strPath & "\"
		End If
		
		'New file location
		lpDefault = ""
		ParmName = "Input_Location"
		ParmValue = Space(ParmSize)
		RetCode = GetPrivateProfileString(SectionName, ParmName, lpDefault, ParmValue, ParmSize, InCfgFileName)
		strNewLoc = UCase(Trim(Left(ParmValue, RetCode)))
		If Right(strNewLoc, 1) <> "\" And RetCode > 0 Then
			strNewLoc = strNewLoc & "\"
		End If
		
		'Diretory Poll Timer Value
		lpDefault = ""
		ParmName = "PollDir"
		ParmValue = Space(ParmSize)
		RetCode = GetPrivateProfileString(SectionName, ParmName, lpDefault, ParmValue, ParmSize, InCfgFileName)
		strPoll = Trim(Left(ParmValue, RetCode))
		
        'Backup Diretory Value
        lpDefault = ""
        ParmName = "Backup"
        ParmValue = Space(ParmSize)
        RetCode = GetPrivateProfileString(SectionName, ParmName, lpDefault, ParmValue, ParmSize, InCfgFileName)
        intBackup = Trim(Left(ParmValue, RetCode))
		
	End Sub
	
	Public Sub WriteIniFile(ByRef InCfgFileName As String)
		'write parameters to an ini file.
		'-------------------
		'[Main]
		'
		'-------------------
		Dim SectionName As String
		Dim ParmName As String
        Dim ParmValue As String
		Dim RetCode As Integer
		
		SectionName = "Main"
		
		ParmName = "OutPut_Location"
		ParmValue = strPath
		RetCode = WritePrivateProfileString(SectionName, ParmName, ParmValue, InCfgFileName)
		
		ParmName = "Input_Location"
		ParmValue = strNewLoc
		RetCode = WritePrivateProfileString(SectionName, ParmName, ParmValue, InCfgFileName)

		ParmName = "PollDir"
		ParmValue = strPoll
		RetCode = WritePrivateProfileString(SectionName, ParmName, ParmValue, InCfgFileName)
		
        ParmName = "Backup"
        ParmValue = intBackup
        RetCode = WritePrivateProfileString(SectionName, ParmName, ParmValue, InCfgFileName)
				
	End Sub
	
    Public Sub UpdateFormVariablesFromGlobal()

        'Update Work location
        frmOPMXML.txtWorkDir.Text = strPath
        'Update OPM Location
        frmOPMXML.txtDir(0).Text = strNewLoc
        'Update Directory default Poll timer
        frmOPMXML.cboData.Text = strPoll
        'Update Backup Directory setting
        If intBackup = "False" Then
            frmOPMXML.chkBackUp.CheckState = System.Windows.Forms.CheckState.Unchecked
        Else
            frmOPMXML.chkBackUp.CheckState = System.Windows.Forms.CheckState.Checked
        End If

    End Sub
    Public Sub UpdateFormVariablesFromGlobalDF()

        'Update Work location
        frmOPMXMLDF.txtWorkDir.Text = strPath
        'Update OPM Location
        frmOPMXMLDF.txtDir(0).Text = strNewLoc
        'Update Directory default Poll timer
        frmOPMXMLDF.cboData.Text = strPoll
        'Update Backup Directory setting
        If intBackup = "False" Then
            frmOPMXMLDF.chkBackUp.CheckState = System.Windows.Forms.CheckState.Unchecked
        Else
            frmOPMXMLDF.chkBackUp.CheckState = System.Windows.Forms.CheckState.Checked
        End If

    End Sub

	
	Public Sub UpdateHdrParmsFromCntlsFrm()
		'Update all header parameters from the values found in the "Cntls" form.
		
		If Trim(frmOPMXML.txtWorkDir.Text) <> "" Then
			If Right(frmOPMXML.txtWorkDir.Text, 1) = "\" Then
				strPath = frmOPMXML.txtWorkDir.Text
			Else
				strPath = frmOPMXML.txtWorkDir.Text & "\"
			End If
		Else
			MsgBox("Please Enter Anacomp Work Directory")
			Exit Sub
		End If
		
		'Update New File Directory
		If Trim(frmOPMXML.txtDir(0).Text) <> "" Then
			If Right(frmOPMXML.txtDir(0).Text, 1) = "\" Then
				strNewLoc = frmOPMXML.txtDir(0).Text
			Else
				strNewLoc = frmOPMXML.txtDir(0).Text & "\"
			End If
		Else
			MsgBox("Please Enter OPM Work Directory")
			Exit Sub
		End If
		
		'Update New File Directory
		If Trim(frmOPMXML.cboData.Text) <> "" Then
			strPoll = frmOPMXML.cboData.Text
		Else
            MsgBox("Please Select Directory Poll Time(seconds) from drop down control")
			Exit Sub
		End If
		
		'Set Database Backup variable setting
		If frmOPMXML.chkBackUp.CheckState = 0 Then
			intBackup = "False"
		Else
			intBackup = "True"
		End If
		
		
    End Sub
    Public Sub UpdateHdrParmsFromCntlsFrmDF()
        'Update all header parameters from the values found in the "Cntls" form.

        If Trim(frmOPMXMLDF.txtWorkDir.Text) <> "" Then
            If Right(frmOPMXMLDF.txtWorkDir.Text, 1) = "\" Then
                strPath = frmOPMXMLDF.txtWorkDir.Text
            Else
                strPath = frmOPMXMLDF.txtWorkDir.Text & "\"
            End If
        Else
            MsgBox("Please Enter Anacomp Work Directory")
            Exit Sub
        End If

        'Update New File Directory
        If Trim(frmOPMXMLDF.txtDir(0).Text) <> "" Then
            If Right(frmOPMXMLDF.txtDir(0).Text, 1) = "\" Then
                strNewLoc = frmOPMXMLDF.txtDir(0).Text
            Else
                strNewLoc = frmOPMXMLDF.txtDir(0).Text & "\"
            End If
        Else
            MsgBox("Please Enter OPM Work Directory")
            Exit Sub
        End If

        'Update New File Directory
        If Trim(frmOPMXMLDF.cboData.Text) <> "" Then
            strPoll = frmOPMXMLDF.cboData.Text
        Else
            MsgBox("Please Select Directory Poll Time(seconds) from drop down control")
            Exit Sub
        End If

        'Set Database Backup variable setting
        If frmOPMXMLDF.chkBackUp.CheckState = 0 Then
            intBackup = "False"
        Else
            intBackup = "True"
        End If


    End Sub
End Module