Public Class ReferenceCountriesCollection
    Inherits Collections.Generic.Dictionary(Of String, ReferenceItem)
    Public Sub New()

        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader
        Dim pobjClass As ReferenceItem

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = "SELECT countryISO3Code " &
                                "      ,countryName " &
                                " FROM AmadeusReports.dbo.zzCountries " &
                                " WHERE LEN(CountryCode) = 2 AND countryISO3Code IS NOT NULL " &
                                " ORDER BY countryname "
            pobjReader = .ExecuteReader
        End With

        With pobjReader
            Do While .Read
                pobjClass = New ReferenceItem
                pobjClass.SetValues(CStr(.Item("countryISO3Code")), CStr(.Item("countryName")))
                MyBase.Add(pobjClass.Code, pobjClass)
            Loop
            .Close()
        End With
        pobjConn.Close()
    End Sub

End Class
