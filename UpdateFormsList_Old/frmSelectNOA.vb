Option Strict Off
Option Explicit On
Friend Class frmSelectNOA
	Inherits System.Windows.Forms.Form
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        MSFlexGrid1.Row = MSFlexGrid1.RowSel
		MSFlexGrid1.Col = 1
		frmModNOA.txtFT.Text = MSFlexGrid1.Text
		MSFlexGrid1.Col = 2
		frmModNOA.txtFromDate.Text = MSFlexGrid1.Text
		MSFlexGrid1.Col = 3
		frmModNOA.txtToDate.Text = MSFlexGrid1.Text
        Me.Close()
	End Sub
	
	Private Sub frmSelectNOA_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim x As Integer, y As Int32
        MSFlexGrid1.set_ColWidth(0, 500)
		MSFlexGrid1.set_ColWidth(1, 9000)
		MSFlexGrid1.set_ColWidth(2, 1400)
        MSFlexGrid1.set_ColWidth(3, 1400)
        x = rsNOACheck.RecordCount
        MSFlexGrid1.Rows = x
        rsNOACheck.MoveFirst()
        For y = 0 To x - 1
            MSFlexGrid1.Col = 0
            MSFlexGrid1.Row = y
            MSFlexGrid1.Text = rsNOACheck.Fields("NOAC").Value
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Row = y
            MSFlexGrid1.Text = rsNOACheck.Fields("Type").Value
            MSFlexGrid1.Col = 2
            MSFlexGrid1.Row = y
            MSFlexGrid1.Text = rsNOACheck.Fields("NOA Eff From Date").Value
            MSFlexGrid1.Col = 3
            MSFlexGrid1.Row = y
            MSFlexGrid1.Text = rsNOACheck.Fields("NOA Eff To Date").Value
            rsNOACheck.MoveNext()
        Next
    End Sub
End Class