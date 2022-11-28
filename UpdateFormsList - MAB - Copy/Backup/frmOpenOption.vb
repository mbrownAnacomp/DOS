Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmOpenOption
	Inherits System.Windows.Forms.Form
	Private Sub cmdBackfile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBackfile.Click
        frmDoSBarCode.Show()
		Me.Hide()
	End Sub
	Private Sub cmdLP_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLP.Click
		frmDayForward.Show()
		Me.Hide()
	End Sub
	
	Private Sub frmOpenOption_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
    End Sub

	Public Sub mnuAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAbout.Click
		frmAbout.ShowDialog()
	End Sub
	
    Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click
        'frmStart.Show()
        frmStart.Enabled = True
        Me.Close()
    End Sub
    Public Sub mnuDeleteBatch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDeleteBatch.Click
        frmDoSDeleteBatch.ShowDialog()
    End Sub

    Public Sub mnuRePrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRePrint.Click
        frmRePrint.ShowDialog()
    End Sub

    Public Sub mnuResetBatchStatus_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuResetBatchStatus.Click
        frmDoSResetBatchStatus.ShowDialog()
    End Sub
End Class
