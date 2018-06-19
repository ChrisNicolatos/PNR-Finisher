Option Strict Off
Option Explicit On

<CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1812:AvoidUninstantiatedInternalClasses")>
Friend Class Airlines

    Public Shared ReadOnly Property AirlineName(ByVal AirlineCode As String) As String
        Get
            AirlineName = ReadAirlineName(AirlineCode)
        End Get
    End Property
    Public Shared ReadOnly Property AirlineCode(ByVal AirlineNumber As String) As String
        Get
            AirlineCode = ReadAirlineCode(AirlineNumber)
        End Get
    End Property
    Private Shared Function ReadAirlineName(ByVal airlineCode As String) As String

        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .Parameters.Add("@AirlineCode", SqlDbType.NVarChar, 2).Value = airlineCode
            .CommandText = " SELECT airlineName " &
                           " FROM [AmadeusReports].[dbo].[zzAirlines] " &
                           " WHERE airlineCode2 = @AirlineCode"
            pobjReader = .ExecuteReader
        End With

        With pobjReader
            If .Read Then
                ReadAirlineName = .Item("airlineName")
            Else
                ReadAirlineName = airlineCode
            End If
            .Close()
        End With
        pobjConn.Close()

    End Function
    Private Shared Function ReadAirlineCode(ByVal airlineNumber As String) As String

        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .Parameters.Add("@AirlineNumber", SqlDbType.NVarChar, 3).Value = airlineNumber
            .CommandText = " SELECT airlineCode2 " &
                           " FROM [AmadeusReports].[dbo].[zzAirlines] " &
                           " WHERE airlineTktCode = @AirlineNumber"
            pobjReader = .ExecuteReader
        End With

        With pobjReader
            If .Read Then
                ReadAirlineCode = .Item("airlineCode2")
            Else
                ReadAirlineCode = airlineNumber
            End If
            .Close()
        End With
        pobjConn.Close()

    End Function
End Class