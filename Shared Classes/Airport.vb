Option Strict Off
Option Explicit On
Friend Class Airport

    Private Shared mCode As String = ""
    Private Shared mCityName As String
    Private Shared mCityAirportName As String
    Private Shared mAirportShortname As String
    Private Shared mCountryName As String

    Public Shared ReadOnly Property CityAirportName(ByVal CityCode As String) As String
        Get
            ReadCityName(CityCode)
            CityAirportName = mCityAirportName
        End Get
    End Property
    Public Shared ReadOnly Property CityName(ByVal CityCode As String) As String
        Get
            ReadCityName(CityCode)
            CityName = mCityName
        End Get
    End Property
    Public Shared ReadOnly Property AirportShortname(ByVal CityCode As String) As String
        Get
            ReadCityName(CityCode)
            AirportShortname = mAirportShortname
        End Get
    End Property
    Public Shared ReadOnly Property CountryName(ByVal CityCode As String) As String
        Get
            CountryName = mCountryName
        End Get
    End Property
    Private Shared Sub ReadCityName(ByVal airportCode As String)

        If airportCode <> mCode Then
            Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .Parameters.Add("@AirportCode", SqlDbType.NVarChar, 3).Value = airportCode
                .CommandText = " SELECT cityName, airportName, ISNULL(airportShortName, '') AS airportShortName , ISNULL(countryName, '') AS countryName " &
                               " FROM [AmadeusReports].[dbo].[zzAirports] " &
                               " LEFT JOIN AmadeusReports.dbo.zzCities " &
                               " ON cityCode = airportCityCode_FK " &
                               " LEFT JOIN AmadeusReports.dbo.zzCountries " &
                               " ON zzCities.cityCountryCode_FK = zzCountries.countryCode " &
                               " WHERE airportCode = @AirportCode"
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
                    mCountryName = .Item("countryName")
                Else
                    mCityAirportName = airportCode
                    mCityName = airportCode
                    mCountryName = ""
                End If
                .Close()
            End With
            pobjConn.Close()
        End If
        mCode = airportCode

    End Sub

End Class
