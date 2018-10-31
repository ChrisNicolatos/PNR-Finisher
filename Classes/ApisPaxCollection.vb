Option Strict On
Option Explicit On
Public Class ApisPaxCollection
    Inherits Collections.Generic.Dictionary(Of Integer, ApisPaxItem)

    Public Sub Read(ByVal Surname As String, ByVal FirstName As String)
        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader
        Dim pobjItem As ApisPaxItem

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.StoredProcedure
            .CommandText = "PaxApisInformationAllSelect"
            .Parameters.Add("@ppSurname", SqlDbType.NVarChar, 30).Value = Surname
            .Parameters.Add("@ppFirstName", SqlDbType.NVarChar, 30).Value = FirstName
            pobjReader = .ExecuteReader
        End With

        Clear()

        With pobjReader
            Do While .Read
                pobjItem = New ApisPaxItem(CInt(.Item("ppId")), Surname, FirstName, CDate(If(IsDBNull(.Item("ppBirthdate")), Date.MinValue, .Item("ppBirthdate"))), CStr(.Item("ppGender")),
                                        CStr(.Item("ppDocIssuingCountry")), CStr(.Item("ppDocnumber")), CDate(If(IsDBNull(.Item("ppDocExpiryDate")), Date.MinValue, .Item("ppDocExpiryDate"))),
                                        CStr(.Item("ppNationality")))
                MyBase.Add(pobjItem.Id, pobjItem)
                'Exit Do
            Loop
            .Close()
        End With
        pobjConn.Close()
    End Sub
    Public Sub AddSSRDocsItem(ByVal Id As Integer, ByVal pSSRDocs As String)
        Dim pItem As New ApisPaxItem(Id, pSSRDocs)
        If pItem.Id > 0 AndAlso Not MyBase.ContainsKey(pItem.Id) Then
            MyBase.Add(pItem.Id, pItem)

        End If
    End Sub

End Class
