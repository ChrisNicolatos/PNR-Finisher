Imports System.ComponentModel

Public Class frmAPISPax
    Private WithEvents mobjPax As New ApisPaxItem
    Private mobjCountries As New ReferenceCountriesCollection
    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click

        mobjPax.Insert()
        DialogResult = DialogResult.OK
        Me.Close()

    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

        DialogResult = DialogResult.Cancel
        Me.Close()

    End Sub
    Private Sub frmAPISPax_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadLists()
        cmdSave.Enabled = mobjPax.IsValid
    End Sub
    Private Sub LoadLists()
        cmbGender.Items.Clear()
        txtNationality.AutoCompleteCustomSource.Clear()
        txtPassportIssuingCountry.AutoCompleteCustomSource.Clear()

        Dim pGender As New ReferenceGenderCollection
        For Each pItem As ReferenceItem In pGender.Values
            cmbGender.Items.Add(pItem)
        Next
        For Each pItem As ReferenceItem In mobjCountries.Values
            txtNationality.AutoCompleteCustomSource.Add(pItem.ToString)
            txtPassportIssuingCountry.AutoCompleteCustomSource.Add(pItem.ToString)
        Next
    End Sub
    Private Sub mobjPax_Valid(IsValid As Boolean) Handles mobjPax.Valid
        cmdSave.Enabled = IsValid
    End Sub
    Private Sub txtSurname_TextChanged(sender As Object, e As EventArgs) Handles txtSurname.TextChanged
        mobjPax.Surname = txtSurname.Text
    End Sub
    Private Sub txtFirstName_TextChanged(sender As Object, e As EventArgs) Handles txtFirstName.TextChanged
        mobjPax.FirstName = txtFirstName.Text
    End Sub
    Private Sub cmbGender_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbGender.SelectedIndexChanged
        If Not cmbGender.SelectedItem Is Nothing Then
            Dim pItem As ReferenceItem = cmbGender.SelectedItem
            mobjPax.Gender = pItem.Code
        End If
    End Sub
    Private Sub txtPassportNumber_TextChanged(sender As Object, e As EventArgs) Handles txtPassportNumber.TextChanged
        mobjPax.PassportNumber = txtPassportNumber.Text
    End Sub

    Private Sub txtDateOfBirth_TextChanged(sender As Object, e As EventArgs) Handles txtDateOfBirth.TextChanged
        Try
            If Not Date.TryParse(txtDateOfBirth.Text, mobjPax.BirthDate) Then
                mobjPax.BirthDate = Utilities.DateFromIATA(txtDateOfBirth.Text)
            End If
        Catch ex As Exception
            mobjPax.BirthDate = Date.MinValue
        End Try
    End Sub

    Private Sub txtPassportExpiryDate_TextChanged(sender As Object, e As EventArgs) Handles txtPassportExpiryDate.TextChanged
        Try
            If Not Date.TryParse(txtPassportExpiryDate.Text, mobjPax.ExpiryDate) Then
                mobjPax.ExpiryDate = Utilities.DateFromIATA(txtPassportExpiryDate.Text)
            End If
        Catch ex As Exception
            mobjPax.ExpiryDate = Date.MinValue
        End Try
    End Sub
    Private Sub txtSurname_LostFocus(sender As Object, e As EventArgs) Handles txtSurname.LostFocus
        txtSurname.Text = mobjPax.Surname
    End Sub

    Private Sub txtFirstName_LostFocus(sender As Object, e As EventArgs) Handles txtFirstName.LostFocus
        txtFirstName.Text = mobjPax.FirstName
    End Sub

    Private Sub txtDateOfBirth_LostFocus(sender As Object, e As EventArgs) Handles txtDateOfBirth.LostFocus
        txtDateOfBirth.Text = mobjPax.BirthDate
    End Sub

    Private Sub txtPassportNumber_LostFocus(sender As Object, e As EventArgs) Handles txtPassportNumber.LostFocus
        txtPassportNumber.Text = mobjPax.PassportNumber
    End Sub

    Private Sub txtPassportExpiryDate_LostFocus(sender As Object, e As EventArgs) Handles txtPassportExpiryDate.LostFocus
        txtPassportExpiryDate.Text = mobjPax.ExpiryDate
    End Sub
    Private Sub txtPassportIssuingCountry_TextChanged(sender As Object, e As EventArgs) Handles txtPassportIssuingCountry.TextChanged
        mobjPax.IssuingCountry = (txtPassportIssuingCountry.Text & "   ").Substring(0, 3)
    End Sub

    Private Sub txtPassportIssuingCountry_LostFocus(sender As Object, e As EventArgs) Handles txtPassportIssuingCountry.LostFocus
        txtPassportIssuingCountry.Text = mobjPax.IssuingCountry
    End Sub

    Private Sub txtNationality_TextChanged(sender As Object, e As EventArgs) Handles txtNationality.TextChanged
        mobjPax.Nationality = (txtNationality.Text & "   ").Substring(0, 3)
    End Sub

    Private Sub txtNationality_LostFocus(sender As Object, e As EventArgs) Handles txtNationality.LostFocus
        txtNationality.Text = mobjPax.Nationality
    End Sub
End Class