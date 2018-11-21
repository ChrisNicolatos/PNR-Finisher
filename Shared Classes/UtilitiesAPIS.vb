Friend Class UtilitiesAPIS

    Public Shared Sub APISPrepareGrid(ByRef dgvApis As DataGridView)
        dgvApis.Columns.Clear()
        dgvApis.Columns.Add("Id", "Id")
        dgvApis.Columns.Add("Surname", "Surname")
        dgvApis.Columns.Add("FirstName", "First Name")
        dgvApis.Columns.Add("IssuingCountry", "Issuing Country")
        dgvApis.Columns.Add("Passportnumber", "Passport number")
        dgvApis.Columns.Add("Nationality", "Nationality")
        dgvApis.Columns.Add("BirthDate", "Birth Date")
        dgvApis.Columns.Add("Gender", "Gender")
        dgvApis.Columns.Add("ExpiryDate", "Expiry Date")
    End Sub
    Public Shared Sub APISAddRow(ByRef dgvApis As DataGridView, ByVal PaxItem As ApisPaxItem)
        Dim dgvRow As New DataGridViewRow With {
            .DefaultCellStyle = dgvApis.RowsDefaultCellStyle
        }
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        dgvRow.Cells(0).Value = PaxItem.Id
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        dgvRow.Cells(1).Value = PaxItem.Surname
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        dgvRow.Cells(2).Value = APISModifyFirstName(PaxItem.FirstName)
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        dgvRow.Cells(3).Value = PaxItem.IssuingCountry ' Issuing Country
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        dgvRow.Cells(4).Value = PaxItem.PassportNumber ' Passport Number
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        dgvRow.Cells(5).Value = PaxItem.Nationality ' Nationality
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        dgvRow.Cells(6).Value = DateToIATA(PaxItem.BirthDate) ' Birth date
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        dgvRow.Cells(7).Value = PaxItem.Gender ' Gender
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        If PaxItem.ExpiryDate > Date.MinValue Then
            dgvRow.Cells(8).Value = DateToIATA(PaxItem.ExpiryDate) ' Expiry Date
        End If
        dgvApis.Rows.Add(dgvRow)
    End Sub
    Public Shared Sub APISAddRow(ByRef dgvApis As DataGridView, ByVal ElementNo As Short, ByVal LastName As String, ByVal Initial As String, ByVal IssuingCountry As String, ByVal PassportNumber As String, ByVal Nationality As String,
                          ByVal Birthdate As Date, ByVal Gender As String, ByVal Expirydate As Date)
        Dim dgvRow As New DataGridViewRow With {
            .DefaultCellStyle = dgvApis.RowsDefaultCellStyle
        }
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        dgvRow.Cells(0).Value = ElementNo
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        dgvRow.Cells(1).Value = LastName
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        dgvRow.Cells(2).Value = APISModifyFirstName(Initial)
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        dgvRow.Cells(3).Value = IssuingCountry ' Issuing Country
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        dgvRow.Cells(4).Value = PassportNumber ' Passport Number
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        dgvRow.Cells(5).Value = Nationality ' Nationality
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        If Birthdate = Date.MinValue Then
            dgvRow.Cells(6).Value = "" ' Birth date
        Else
            dgvRow.Cells(6).Value = DateToIATA(Birthdate) ' Birth date
        End If
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        dgvRow.Cells(7).Value = Gender ' Gender
        dgvRow.Cells.Add(New DataGridViewTextBoxCell)
        If Expirydate > Date.MinValue Then
            dgvRow.Cells(8).Value = DateToIATA(Expirydate) ' Expiry Date
        End If
        dgvApis.Rows.Add(dgvRow)
    End Sub
    Public Shared Function APISModifyFirstName(ByVal FirstName As String) As String
        Dim pSalutations As New ReferenceSalutationsCollection
        Dim pintFindPos As Integer
        FirstName = FirstName.Trim
        For Each pItem As ReferenceItem In pSalutations.Values
            pintFindPos = FirstName.IndexOf(pItem.Code)
            If pintFindPos > 0 And pintFindPos = FirstName.Length - pItem.Code.Length Then
                FirstName = FirstName.Substring(0, pintFindPos).Trim
            End If
        Next
        Return FirstName
    End Function

End Class
