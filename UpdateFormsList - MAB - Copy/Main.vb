Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMain
	Inherits System.Windows.Forms.Form
	Private Sub cmdAddFN_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAddFN.Click
		frmAddFN.ShowDialog()
	End Sub
	Private Sub cmdAddNOA_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAddNOA.Click
		frmAddNOA.ShowDialog()
	End Sub
	Public Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub
	Private Sub cmdModifyNOA_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdModifyNOA.Click
		frmModNOA.ShowDialog()
	End Sub
	Public Sub mnuBackup_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuBackup.Click
        'How do I backup SQL DB?
        MsgBox("Master List and Validation Completion databases have been backed up")
    End Sub
	Private Sub cmdModifyFN_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdModifyFN.Click
		frmModFN.ShowDialog()
    End Sub
    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        frmStart.Enabled = True
    End Sub
    Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'Deal with ini file
        Dim iniFile As Short
        Dim bDRPSet, bNOASet, bPS50NOASet, bOtherNOASet As Boolean
        Dim varText As String
        Dim Slash, x As String
        iniFile = FreeFile()
        Slash = ""
        If VB.Right(My.Application.Info.DirectoryPath, 1) <> "\" Then Slash = "\"
        x = My.Application.Info.DirectoryPath & Slash & "UpdateFormsList.ini"
        If x <> "" Then
            FileOpen(iniFile, x, OpenMode.Input)
            Do While Not EOF(1)
                varText = LineInput(iniFile)
                If VB.Left(varText, 3) = "DRP" Then
                    sDRPDir = VB.Right(varText, Len(varText) - 14)
                ElseIf VB.Left(varText, 3) = "NOA" Then
                    sNOADir = VB.Right(varText, Len(varText) - 14)
                ElseIf VB.Left(varText, 4) = "PS50" Then
                    sPS50NOADir = VB.Right(varText, Len(varText) - 15)
                ElseIf VB.Left(varText, 5) = "Other" Then
                    sOtherNOADir = VB.Right(varText, Len(varText) - 16)
                End If
            Loop
            FileClose(iniFile)
            If sDRPDir = "" Then sDRPDir = "Forms Dictionary file not defined"
            If sNOADir = "" Then sNOADir = "SF50 NOA Code Dictionary file not defined"
            If sPS50NOADir = "" Then sPS50NOADir = "PS50 NOA Code Dictionary file not defined"
            If sOtherNOADir = "" Then sOtherNOADir = "All Other NOA Code Dictionary file not defined"
        Else
            sDRPDir = "Forms Dictionary file not defined"
            sNOADir = "SF50 NOA Code Dictionary file not defined"
            sPS50NOADir = "PS50 NOA Code Dictionary file not defined"
            sOtherNOADir = "All Other NOA Code Dictionary file not defined"
        End If
        bDRPSet = True
        bNOASet = True
        bPS50NOASet = True
        bOtherNOASet = True
        If sDRPDir = "Forms Dictionary file not defined" Then
            bDRPSet = False
        End If
        If sNOADir = "NOA Code Dictionary file not defined" Then
            bNOASet = False
        End If
        If sPS50NOADir = "PS50 NOA Code Dictionary file not defined" Then
            bPS50NOASet = False
        End If
        If sOtherNOADir = "All Other NOA Code Dictionary file not defined" Then
            bOtherNOASet = False
        End If
        If bDRPSet = False Then MsgBox("The Forms Dictionary file not defined. Use the File/Find Form Dictionary File menu to set", MsgBoxStyle.Critical, "Distionary file not set")
        If bNOASet = False Then MsgBox("The SF50 NOA Code Dictionary file not defined. Use the File/Find NOA Code Dictionary File menu to set", MsgBoxStyle.Critical, "Dictionary file not set")
        If bPS50NOASet = False Then MsgBox("The PS50 NOA Code Dictionary file not defined. Use the File/Find NOA Code Dictionary File menu to set", MsgBoxStyle.Critical, "Distionary file not set")
        If bOtherNOASet = False Then MsgBox("The All Other 50 NOA Code Dictionary file not defined. Use the File/Find NOA Code Dictionary File menu to set", MsgBoxStyle.Critical, "Dictionary file not set")
        'Don't let there be changes/additions if any db or drp is missing
        If bDRPSet = False Or bNOASet = False Or bPS50NOASet = False Or bOtherNOASet = False Then
            cmdModifyFN.Enabled = False
            cmdAddFN.Enabled = False
            cmdModifyNOA.Enabled = False
            cmdAddNOA.Enabled = False
        End If
        'Set filesystemobject
        fs = New Scripting.FileSystemObject
    End Sub
	Private Sub frmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'Save directory settings to ini like file
		Dim iniFile As Short
        If sDRPDir = "" Or UCase(VB.Right(sDRPDir, 4)) <> ".TXT" Then sDRPDir = "Forms Dictionary file not defined"
        If sNOADir = "" Or UCase(VB.Right(sNOADir, 4)) <> ".TXT" Then sNOADir = "SF50 NOA Code Dictionary file not defined"
        If sPS50NOADir = "" Or UCase(VB.Right(sPS50NOADir, 4)) <> ".TXT" Then sPS50NOADir = "PS50 NOA Code Dictionary file not defined"
        If sOtherNOADir = "" Or UCase(VB.Right(sOtherNOADir, 4)) <> ".TXT" Then sOtherNOADir = "All Other NOA Code Dictionary file not defined"
        iniFile = FreeFile()
        If VB.Right(My.Application.Info.DirectoryPath, 1) = "\" Then
            FileOpen(iniFile, My.Application.Info.DirectoryPath & "UpdateFormsList.ini", OpenMode.Output)
        Else
            FileOpen(iniFile, My.Application.Info.DirectoryPath & "\UpdateFormsList.ini", OpenMode.Output)
        End If
        PrintLine(iniFile, "DRP Directory=" & sDRPDir)
        PrintLine(iniFile, "NOA Directory=" & sNOADir)
        PrintLine(iniFile, "PS50 Directory=" & sPS50NOADir)
        PrintLine(iniFile, "Other Directory=" & sOtherNOADir)
        FileClose()
    End Sub
	Public Sub mnuAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAbout.Click
		frmAbout.ShowDialog()
	End Sub
    Public Sub mnuFormDict_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFormDict.Click
        'frmBrowseDR is legacy name - now browses for KTM Dictionary
        frmBrowseDR.ShowDialog()
    End Sub
    Private Sub mnuNOADict_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        frmBrowseNOA.ShowDialog()
    End Sub

    Private Sub mnuPS50_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPS50.Click
        frmBrowsePS50NOA.ShowDialog()
    End Sub

    Private Sub mnuAllOther_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAllOther.Click
        frmBrowseOtherNOA.ShowDialog()
    End Sub
End Class