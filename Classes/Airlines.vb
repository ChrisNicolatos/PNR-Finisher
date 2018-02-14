Option Strict Off
Option Explicit On
Friend Class Airlines

    Public ReadOnly Property AirlineName(ByVal AirlineCode As String) As String
        Get
            AirlineName = ReadAirline(AirlineCode)
        End Get
    End Property

    Private Function ReadAirline(ByVal airlineCode As String) As String

        Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = " SELECT airlineName " & _
                           " FROM [AmadeusReports].[dbo].[zzAirlines] " & _
                           " WHERE airlineCode2 = '" & airlineCode & "'"
            pobjReader = .ExecuteReader
        End With

        With pobjReader
            If .Read Then
                ReadAirline = .Item("airlineName")
            Else
                ReadAirline = airlineCode
            End If
            .Close()
        End With
        pobjConn.Close()

    End Function
End Class