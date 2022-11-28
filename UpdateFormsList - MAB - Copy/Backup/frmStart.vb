Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic

Friend Class frmStart
    Friend bODBCSet As Boolean
    Private Sub cmdUFL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUFL.Click
        Me.Enabled = False
        frmMain.Show()
    End Sub

    Private Sub cmdBC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBC.Click
        Me.Enabled = False
        frmOpenOption.Show()
    End Sub

    Private Sub frmStart_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'Save directory settings to ini like file
        Dim iniFile As Short
        If sODBC(0) = "" Then sMLDBDir = "ODBC Connection Not Defined"
        iniFile = FreeFile()
        If VB.Right(My.Application.Info.DirectoryPath, 1) = "\" Then
            FileOpen(iniFile, My.Application.Info.DirectoryPath & "Start.ini", OpenMode.Output)
        Else
            FileOpen(iniFile, My.Application.Info.DirectoryPath & "\" & "Start.ini", OpenMode.Output)
        End If
        PrintLine(iniFile, "ODBC DSN=" & sODBC(0))
        PrintLine(iniFile, "ODBC DB=" & sODBC(1))
        PrintLine(iniFile, "ODBC UID=" & sODBC(2))
        PrintLine(iniFile, "ODBC PWD=" & sODBC(3))
        PrintLine(iniFile, "WinAuth=" & IIf(bWinAuth = True, "True", "False"))
        FileClose()
        If bODBCSet Then db.Close()
    End Sub

    Private Sub frmStart_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Deal with ini file
        Dim iniFile As Short
        Dim varText, sConnect As String
        Dim Slash, x As String
        iniFile = FreeFile()
        Slash = ""
        If VB.Right(My.Application.Info.DirectoryPath, 1) <> "\" Then Slash = "\"
        x = My.Application.Info.DirectoryPath & Slash & "Start.ini"
        If x <> "" Then
            FileOpen(iniFile, x, OpenMode.Input)
            Do While Not EOF(1)
                varText = LineInput(iniFile)
                If VB.Left(varText, 8) = "ODBC DSN" Then
                    sODBC(0) = VB.Right(varText, Len(varText) - 9)
                ElseIf VB.Left(varText, 7) = "ODBC DB" Then
                    sODBC(1) = VB.Right(varText, Len(varText) - 8)
                ElseIf VB.Left(varText, 8) = "ODBC UID" Then
                    sODBC(2) = VB.Right(varText, Len(varText) - 9)
                ElseIf VB.Left(varText, 8) = "ODBC PWD" Then
                    sODBC(3) = VB.Right(varText, Len(varText) - 9)
                ElseIf VB.Left(varText, 7) = "WinAuth" Then
                    bWinAuth = IIf(VB.InStr(1, varText, "True") > 0, True, False)
                End If
            Loop
            FileClose(iniFile)
            If sODBC(0) = "" Then sODBC(0) = "ODBC Connection Not Defined"
        Else
            sODBC(0) = "ODBC Connection Not Defined"
        End If
        bODBCSet = True
        'Connect the db's if they have been defined
        If sODBC(0) <> "ODBC Connection Not Defined" Then
            Try
                sConnect = "ODBC;DSN=" & sODBC(0) & ";DATABASE=" & sODBC(1)
                If Not bWinAuth Then sConnect = sConnect & ";UID=" & sODBC(2) & ";PWD=" & sODBC(3)
                db = DAODBEngine_definst.OpenDatabase("", False, False, sConnect)
            Catch excGeneric As Exception
                MsgBox("ODBC Connection failed! Check ODBC settings and network access" & vbCrLf _
                       & "Error Message = " & excGeneric.Message, MsgBoxStyle.Critical)
                bODBCSet = False
            End Try
        Else
            bODBCSet = False
        End If
        If bODBCSet = False Then MsgBox("The ODBC Connection is not defined. Use the File/Define ODBC Connection menu to set", MsgBoxStyle.Critical, "ODBC Not Set")
        'Don't let there be changes/additions if any db or drp is missing
        If bODBCSet = False Then
            cmdBC.Enabled = False
            cmdUFL.Enabled = False
            cmdLPXML.Enabled = False
            cmdXML.Enabled = False
        End If
        'Set filesystemobject
        fs = New Scripting.FileSystemObject

    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
        End
    End Sub

    Private Sub ODBCConnectToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ODBCConnectToolStripMenuItem.Click
        Dim lRet As Short
        lRet = MsgBox("Do NOT us this dialog unless you are cetain you know the correct parameters" & vbCrLf _
                         & "Use the Cancel button if you are unsure", MsgBoxStyle.OkCancel)
        If lRet = MsgBoxResult.Cancel Then
            Exit Sub
        End If
        Me.Enabled = False
        frmODBC.Show()
    End Sub

    Private Sub cmdXML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdXML.Click
        Me.Enabled = False
        frmOPMXML.Show()
    End Sub

    Private Sub cmdLPXML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLPXML.Click
        Me.Enabled = False
        frmOPMXMLDF.Show()
    End Sub

    Private Sub mnuImportSSN_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuImportSSN.Click
        Me.Enabled = False
        frmSSNImport.Show()
    End Sub
    Private Sub mnuRecon_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRecon.Click
        Me.Enabled = False
        frmRecon.Show()
    End Sub
End Class