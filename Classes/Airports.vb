Option Strict Off
Option Explicit On
Friend Class Airports

    Private mCode As String = ""
    Private mCityName As String
    Private mCityAirportName As String
    Private mAirportShortname As String

    Public ReadOnly Property CityAirportName(ByVal CityCode As String) As String
        Get
            ReadCityName(CityCode)
            CityAirportName = mCityAirportName
        End Get
    End Property
    Public ReadOnly Property CityName(ByVal CityCode As String) As String
        Get
            ReadCityName(CityCode)
            CityName = mCityName
        End Get
    End Property
    Public ReadOnly Property AirportShortname(ByVal CityCode As String) As String
        Get
            ReadCityName(CityCode)
            AirportShortname = mAirportShortname
        End Get
    End Property
    Private Sub ReadCityName(ByVal cityCode As String)

        If cityCode <> mCode Then
            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .CommandText = " SELECT cityName, airportName, ISNULL(airportShortName, '') AS airportShortName " & _
                               " FROM [AmadeusReports].[dbo].[zzAirports] " & _
                               " LEFT JOIN AmadeusReports.dbo.zzCities " & _
                               " ON cityCode = airportCityCode_FK " & _
                               " WHERE airportCode = '" & cityCode & "'"
                pobjReader = .ExecuteReader
            End With

            With pobjReader
                If .Read Then
                    If .Item("cityName") = .Item("airportName") Then
                        mCityAirportName = .Item("cityName")
                    ElseIf .Item("airportName").ToString.StartsWith(.Item("cityName")) Then
                        mCityAirportName = .Item("airportName")
                    Else
                        mCityAirportName = .Item("cityName") & " " & .Item("airportName")
                    End If
                    mCityName = .Item("cityName")
                    mAirportShortname = .Item("airportShortName")
                Else
                    mCityAirportName = cityCode
                    mCityName = cityCode
                End If
                .Close()
            End With
            pobjConn.Close()
        End If
        mCode = cityCode

    End Sub

End Class