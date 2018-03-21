Public Class frmAPISPax
    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click

        DialogResult = DialogResult.OK
        Me.Close()

    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

        DialogResult = DialogResult.Cancel
        Me.Close()

    End Sub

    Private Sub frmAPISPax_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub EnableSelection()

    End Sub
End Class