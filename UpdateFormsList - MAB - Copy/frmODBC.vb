Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmODBC
    Inherits System.Windows.Forms.Form
    Private Sub chkAuthentication_CheckedChanged(ByVal sender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkAuthentication.CheckedChanged
        If chkAuthentication.Checked Then
            txtPWD.Enabled = False
            txtUID.Enabled = False
        Else
            txtPWD.Enabled = True
            txtUID.Enabled = True
        End If
    End Sub
    Private Sub txtDB_Validating(ByVal sender As Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtDB.Validating
        Dim Cancel As Boolean = EventArgs.Cancel
        If Trim(txtDB.Text) = "" Then
            MsgBox("You must enter a value for the Database")
            Cancel = True
            GoTo EventExitSub
        Else
            Cancel = False
        End If
EventExitSub:
        EventArgs.Cancel = Cancel
    End Sub

    Private Sub txtDSN_Validating(ByVal sender As Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtDSN.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        If Trim(txtDSN.Text) = "" Then
            MsgBox("You must enter a value for the DSN")
            Cancel = True
            GoTo EventExitSub
        Else
            Cancel = False
        End If
EventExitSub:
        EventArgs.Cancel = Cancel
    End Sub

    Private Sub txtPWD_Validating(ByVal sender As Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtPWD.Validating
        Dim Cancel As Boolean = EventArgs.Cancel
        If Trim(txtPWD.Text) = "" Then
            MsgBox("If you are not using Windows Authentication, you must enter an SQL Password")
            Cancel = True
            GoTo EventExitSub
        Else
            Cancel = False
        End If
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub txtUID_Validating(ByVal sender As Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtUID.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        If Trim(txtUID.Text) = "" Then
            MsgBox("If you are not using Windows Authentication, you must enter an SQL User ID")
            Cancel = True
            GoTo EventExitSub
        Else
            Cancel = False
        End If
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.txtDB.CausesValidation = False
        Me.txtDSN.CausesValidation = False
        Me.txtUID.CausesValidation = False
        Me.txtPWD.CausesValidation = False
        Me.Close()
    End Sub

    Private Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim sTestConnect As String
        If Trim(Me.txtDSN.Text) = "" Then
            MsgBox("DSN cannot be blank", MsgBoxStyle.OkOnly, "Warning!")
            Me.txtDSN.Focus()
            Exit Sub
        End If
        If Trim(Me.txtDB.Text) = "" Then
            MsgBox("Datase field cannot be blank", MsgBoxStyle.OkOnly, "Warning!")
            Me.txtDB.Focus()
            Exit Sub
        End If
        If Trim(Me.txtUID.Text) = "" And chkAuthentication.Checked = False Then
            MsgBox("UID cannot be blank if you are not using Windows Authentication", MsgBoxStyle.OkOnly, "Warning!")
            Me.txtUID.Focus()
            Exit Sub
        End If
        If Trim(Me.txtPWD.Text) = "" And chkAuthentication.Checked = False Then
            MsgBox("Password cannot be blank if you are not using Windows Authentication", MsgBoxStyle.OkOnly, "Warning!")
            Me.txtPWD.Focus()
            Exit Sub
        End If
        'See if it works
        Dim dbTest As DAO.Database
        sTestConnect = "ODBC;DSN=" & Trim(txtDSN.Text) & ";DATABASE=" & Trim(txtDB.Text)
        If chkAuthentication.Checked = False Then
            sTestConnect = sTestConnect & ";UID=" & Trim(txtUID.Text) & ";PWD=" & Trim(txtPWD.Text)
        End If
        Try
            dbTest = DAODBEngine_definst.OpenDatabase("", 1, False, sTestConnect)
            dbTest.Close()
            'Must have worked - save it
            sODBC(0) = Trim(Me.txtDSN.Text)
            sODBC(1) = Trim(Me.txtDB.Text)
            If chkAuthentication.Checked = False Then
                sODBC(2) = Trim(txtUID.Text)
                sODBC(3) = Trim(txtPWD.Text)
            Else
                sODBC(2) = ""
                sODBC(3) = ""
            End If
            bWinAuth = IIf(chkAuthentication.Checked = True, True, False)
            MsgBox("Connection tested OK, new connection will be used next time this program is started")
            Me.Close()
        Catch excGeneric As Exception
            MsgBox("!Unable to connect via ODBC to the specified SQL database!" & vbCrLf & _
                   "connection failure = " & excGeneric.Message & vbCrLf & _
                   "Please re-enter ODBC values or use the Cancel button to exit without change", MsgBoxStyle.Critical)
            Me.txtDSN.Focus()
            Exit Sub
        End Try

    End Sub

    Private Sub frmODBC_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        frmStart.Enabled = True
    End Sub

    Private Sub frmODBC_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        KeyPreview = True

        Me.txtDB.CausesValidation = True
        Me.txtDSN.CausesValidation = True
        Me.txtUID.CausesValidation = True
        Me.txtPWD.CausesValidation = True

        Me.txtDSN.Text = sODBC(0)
        Me.txtDB.Text = sODBC(1)
        Me.txtUID.Text = sODBC(2)
        Me.txtPWD.Text = sODBC(3)
        If bWinAuth Then
            chkAuthentication.Checked = True
            txtPWD.Enabled = False
            txtUID.Enabled = False
        Else
            chkAuthentication.Checked = False
            txtPWD.Enabled = True
            txtUID.Enabled = True
        End If
        sHoldODBC(0) = sODBC(0)
        sHoldODBC(1) = sODBC(1)
        sHoldODBC(2) = sODBC(2)
        sHoldODBC(3) = sODBC(3)
        bHoldWinAuth = bWinAuth
    End Sub
End Class