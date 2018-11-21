Public Class frmUser
    Private WithEvents mobjDBUser As DBUser
    Public Sub New(ByVal pGDS As Utilities.EnumGDSCode, ByVal pPCC As String, pUserID As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        mobjDBUser = New DBUser(pGDS, pPCC, pUserID)
        DisplayUser()
        EnableOptions()

    End Sub

    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        Try
            mobjDBUser.Update()
            DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

        DialogResult = DialogResult.Cancel
        Me.Close()

    End Sub

    Private Sub txtUsername_TextChanged(sender As Object, e As EventArgs) Handles txtUsername.TextChanged
        mobjDBUser.Username = txtUsername.Text.Trim
    End Sub

    Private Sub txtEmail_TextChanged(sender As Object, e As EventArgs) Handles txtEmail.TextChanged
        mobjDBUser.Email = txtEmail.Text.Trim
    End Sub

    Private Sub txtQueue_TextChanged(sender As Object, e As EventArgs) Handles txtQueue.TextChanged
        mobjDBUser.QueueNumber = txtQueue.Text.Trim
    End Sub

    Private Sub txtOPQueue_TextChanged(sender As Object, e As EventArgs) Handles txtOPQueue.TextChanged
        mobjDBUser.OPQueue = txtOPQueue.Text.Trim
    End Sub
    Private Sub DisplayUser()
        If mobjDBUser.GDS = Utilities.EnumGDSCode.Amadeus Then
            lblGDS.Text = "Amadeus"
            lblQForTimeLimit.Text = "Queue for time limit (TK TL)"
            lblQForReminder.Text = "Queue for reminder (OP)"
            lblQHint.Text = "Enter queue numbers without Q (e.g. 72, 72C4)"
        ElseIf mobjDBUser.GDS = Utilities.EnumGDSCode.Galileo Then
            lblGDS.Text = "Galileo"
            lblQForTimeLimit.Text = "Queue for time limit (T.TAU)"
            lblQForReminder.Text = "Queue for reminder (RB.)"
            lblQHint.Text = "Enter queue numbers without Q (e.g. 72, 72*C4)"
        Else
            lblGDS.Text = "UNKNOWN"
        End If
        lblPCC.Text = mobjDBUser.PCC
        lblUser.Text = mobjDBUser.UserID

    End Sub
    Private Sub EnableOptions()
        With mobjDBUser
            cmdCancel.Enabled = True
            ColourTextBox(txtUsername, .isUserNameValid)
            ColourTextBox(txtEmail, .isEmailValid)
            ColourTextBox(txtQueue, .isQueueNumberValid)
            ColourTextBox(txtOPQueue, .isOPQueueValid)
            cmdSave.Enabled = .isValid
        End With
    End Sub
    Private Sub ColourTextBox(ByRef txtBox As TextBox, ByRef IsValid As Boolean)
        If IsValid Then
            txtBox.BackColor = Color.White
        Else
            txtBox.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub mobjDBUser_UserValid() Handles mobjDBUser.UserValid

        EnableOptions()

    End Sub
End Class