Public Class Config

    Private Structure ClassProps
        Dim PCCId As Integer
        Dim OfficeCityCode As String
        Dim CountryCode As String
        Dim OfficeName As String
        Dim CityName As String
        Dim Phone As String
        Dim AOHPhone As String
        Dim PCCBackOffice As Integer
        Dim AgentId As Integer
        Dim AgentQueue As String
        Dim AgentOPQueue As String
        Dim AgentName As String
        Dim AgentEmail As String
        Dim AirportName As Integer
        Dim AirlineLocator As Boolean
        Dim ClassOfService As Boolean
        Dim BanElectricalEquipment As Boolean
        Dim BrazilText As Boolean
        Dim USAText As Boolean
        Dim Tickets As Boolean
        Dim PaxSegPerTkt As Boolean
        Dim ShowStopovers As Boolean
        Dim ShowTerminal As Boolean
        Dim FlyingTime As Boolean
        Dim CostCentre As Boolean
        Dim Seating As Boolean
        Dim Vessel As Boolean
        Dim PlainFormat As Boolean
        Dim Administrator As String
    End Structure
    Private mudtProps As ClassProps
    Private mobjAmadeusUser As AmadeusUser
    Private mflgIsDirtyPCC As Boolean
    Private mflgIsDirtyUser As Boolean

    Dim mAmadeusReferences As New Collections.Generic.Dictionary(Of String, String)

    Public Sub New(mAmadeusUser As AmadeusUser)

        Try
            mobjAmadeusUser = mAmadeusUser

            mflgIsDirtyPCC = False
            mflgIsDirtyUser = False
            mAmadeusReferences.Clear()
            DBReadPCC()
            DBReadUser()
            DBReadReferences()

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Public ReadOnly Property Administrator As Boolean
        Get
            Administrator = mudtProps.Administrator
        End Get
    End Property
    Public ReadOnly Property IsDirty As Boolean
        Get
            IsDirty = mflgIsDirtyPCC Or mflgIsDirtyUser
        End Get
    End Property
    Public ReadOnly Property IsDirtyPCC As Boolean
        Get
            IsDirtyPCC = mflgIsDirtyPCC
        End Get
    End Property
    Public ReadOnly Property IsDirtyUser As Boolean
        Get
            IsDirtyUser = mflgIsDirtyUser
        End Get
    End Property
    Public Property AirlineLocator As Boolean
        Get
            AirlineLocator = mudtprops.AirlineLocator
        End Get
        Set(value As Boolean)
            If value <> mudtprops.AirlineLocator Then
                mflgIsDirtyUser = True
            End If
            mudtprops.AirlineLocator = value
        End Set
    End Property
    Public Property ClassOfService As Boolean
        Get
            ClassOfService = mudtprops.ClassOfService
        End Get
        Set(value As Boolean)
            If value <> mudtprops.ClassOfService Then
                mflgIsDirtyUser = True
            End If
            mudtprops.ClassOfService = value
        End Set
    End Property
    Public Property BanElectricalEquipment As Boolean
        Get
            BanElectricalEquipment = mudtprops.BanElectricalEquipment
        End Get
        Set(value As Boolean)
            If value <> mudtprops.BanElectricalEquipment Then
                mflgIsDirtyUser = True
            End If
            mudtprops.BanElectricalEquipment = value
        End Set
    End Property
    Public Property BrazilText As Boolean
        Get
            BrazilText = mudtprops.BrazilText
        End Get
        Set(value As Boolean)
            If value <> mudtprops.BrazilText Then
                mflgIsDirtyUser = True
            End If
            mudtprops.BrazilText = value
        End Set
    End Property
    Public Property USAText As Boolean
        Get
            USAText = mudtprops.USAText
        End Get
        Set(value As Boolean)
            If value <> mudtprops.USAText Then
                mflgIsDirtyUser = True
            End If
            mudtprops.USAText = value
        End Set
    End Property
    Public Property Tickets As Boolean
        Get
            Tickets = mudtprops.Tickets
        End Get
        Set(value As Boolean)
            If value <> mudtprops.Tickets Then
                mflgIsDirtyUser = True
            End If
            mudtprops.Tickets = value
        End Set
    End Property
    Public Property PaxSegPerTkt As Boolean
        Get
            PaxSegPerTkt = mudtprops.PaxSegPerTkt
        End Get
        Set(value As Boolean)
            If value <> mudtprops.PaxSegPerTkt Then
                mflgIsDirtyUser = True
            End If
            mudtprops.PaxSegPerTkt = value
        End Set
    End Property
    Public Property ShowStopovers As Boolean
        Get
            ShowStopovers = mudtprops.ShowStopovers
        End Get
        Set(value As Boolean)
            If value <> mudtprops.ShowStopovers Then
                mflgIsDirtyUser = True
            End If
            mudtprops.ShowStopovers = value
        End Set
    End Property
    Public Property ShowTerminal As Boolean
        Get
            ShowTerminal = mudtprops.ShowTerminal
        End Get
        Set(value As Boolean)
            If value <> mudtprops.ShowTerminal Then
                mflgIsDirtyUser = True
            End If
            mudtprops.ShowTerminal = value
        End Set
    End Property
    Public Property FlyingTime As Boolean
        Get
            FlyingTime = mudtprops.FlyingTime
        End Get
        Set(value As Boolean)
            If value <> mudtprops.FlyingTime Then
                mflgIsDirtyUser = True
            End If
            mudtprops.FlyingTime = value
        End Set
    End Property
    Public Property CostCentre As Boolean
        Get
            CostCentre = mudtprops.CostCentre
        End Get
        Set(value As Boolean)
            If value <> mudtprops.CostCentre Then
                mflgIsDirtyUser = True
            End If
            mudtprops.CostCentre = value
        End Set
    End Property
    Public Property Seating As Boolean
        Get
            Seating = mudtprops.Seating
        End Get
        Set(value As Boolean)
            If value <> mudtprops.Seating Then
                mflgIsDirtyUser = True
            End If
            mudtprops.Seating = value
        End Set
    End Property
    Public Property Vessel As Boolean
        Get
            Vessel = mudtprops.Vessel
        End Get
        Set(value As Boolean)
            If value <> mudtprops.Vessel Then
                mflgIsDirtyUser = True
            End If
            mudtprops.Vessel = value
        End Set
    End Property
    Public Property PlainFormat As Boolean
        Get
            PlainFormat = mudtprops.PlainFormat
        End Get
        Set(value As Boolean)
            If value <> mudtprops.PlainFormat Then
                mflgIsDirtyUser = True
            End If
            mudtprops.PlainFormat = value
        End Set
    End Property
    Public Property OfficeCityCode As String
        Get
            OfficeCityCode = mudtprops.OfficeCityCode
        End Get
        Set(value As String)
            If value <> mudtprops.OfficeCityCode Then
                mflgIsDirtyPCC = True
            End If
            mudtprops.OfficeCityCode = value
        End Set
    End Property
    Public Property CountryCode As String
        Get
            CountryCode = mudtprops.CountryCode
        End Get
        Set(value As String)
            If value <> mudtprops.CountryCode Then
                mflgIsDirtyPCC = True
            End If
            mudtprops.CountryCode = value
        End Set
    End Property
    Public Property OfficeName As String
        Get
            OfficeName = mudtprops.OfficeName
        End Get
        Set(value As String)
            If value <> mudtprops.OfficeName Then
                mflgIsDirtyPCC = True
            End If
            mudtprops.OfficeName = value
        End Set
    End Property
    Public Property CityName As String
        Get
            CityName = mudtprops.CityName
        End Get
        Set(value As String)
            If value <> mudtprops.CityName Then
                mflgIsDirtyPCC = True
            End If
            mudtprops.CityName = value
        End Set
    End Property
    Public Property Phone As String
        Get
            Phone = mudtprops.Phone
        End Get
        Set(value As String)
            If value <> mudtprops.Phone Then
                mflgIsDirtyPCC = True
            End If
            mudtprops.Phone = value
        End Set
    End Property
    Public Property AOHPhone As String
        Get
            AOHPhone = mudtprops.AOHPhone
        End Get
        Set(value As String)
            If value <> mudtprops.AOHPhone Then
                mflgIsDirtyPCC = True
            End If
            mudtprops.AOHPhone = value
        End Set
    End Property
    Public Property PCCBackOffice As Integer
        Get
            PCCBackOffice = mudtProps.PCCBackOffice
        End Get
        Set(value As Integer)
            mudtProps.PCCBackOffice = value
        End Set
    End Property
    Public Property AgentQueue As String
        Get
            AgentQueue = mudtprops.AgentQueue
        End Get
        Set(value As String)
            If value <> mudtprops.AgentQueue Then
                mflgIsDirtyUser = True
            End If
            mudtprops.AgentQueue = value
        End Set
    End Property
    Public Property AgentOPQueue As String
        Get
            AgentOPQueue = mudtprops.AgentOPQueue
        End Get
        Set(value As String)
            If value <> mudtprops.AgentOPQueue Then
                mflgIsDirtyUser = True
            End If
            mudtprops.AgentOPQueue = value
        End Set
    End Property
    Public Property AgentName As String
        Get
            AgentName = mudtprops.AgentName
        End Get
        Set(value As String)
            If value <> mudtprops.AgentName Then
                mflgIsDirtyUser = True
            End If
            mudtprops.AgentName = value
        End Set
    End Property
    Public Property AgentEmail As String
        Get
            AgentEmail = mudtprops.AgentEmail
        End Get
        Set(value As String)
            If value <> mudtprops.AgentEmail Then
                mflgIsDirtyUser = True
            End If
            mudtprops.AgentEmail = value
        End Set
    End Property
    Public Property AirportName As Integer
        Get
            AirportName = mudtprops.AirportName
        End Get
        Set(value As Integer)
            If value <> mudtprops.AirportName Then
                mflgIsDirtyUser = True
            End If
            mudtprops.AirportName = value
        End Set
    End Property

    Public ReadOnly Property AmadeusPCC As String
        Get
            AmadeusPCC = mobjAmadeusUser.PCC
        End Get
    End Property
    Public ReadOnly Property AmadeusUser As String
        Get
            AmadeusUser = mobjAmadeusUser.User
        End Get
    End Property

    Private Sub DBReadReferences()

        Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = " SELECT pfrID " &
                           " ,pfrKey" &
                           " ,pfrValue " &
                           " FROM [AmadeusReports].[dbo].[PNRFinisherGDS_BOReferences] " &
                           " WHERE pfrGDS_fkey = 1 AND pfrBO_fkey = " & mudtProps.PCCBackOffice
            pobjReader = .ExecuteReader
        End With
        With pobjReader
            While pobjReader.Read
                mAmadeusReferences.Add(.Item("pfrKey"), .Item("pfrValue"))
            End While
            .Close()
        End With
        pobjConn.Close()

    End Sub

    Private Sub DBReadPCC()

        Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = " SELECT pfpId " &
                           " ,pfpOfficeCityCode " &
                           " ,pfpCountryCode " &
                           " ,pfpOfficeName " &
                           " ,pfpCityName " &
                           " ,pfpOfficePhone " &
                           " ,pfpAOHPhone " &
                           " ,pfpBO_fkey " &
                           " FROM [AmadeusReports].[dbo].[PNRFinisherPCC] " &
                           " WHERE pfpPCC = '" & mobjAmadeusUser.PCC & "'"

            pobjReader = .ExecuteReader
        End With
        With pobjReader
            If pobjReader.Read Then
                mudtProps.PCCId = .Item("pfpId")
                mudtProps.OfficeCityCode = .Item("pfpOfficeCityCode")
                mudtProps.CountryCode = .Item("pfpCountryCode")
                mudtProps.OfficeName = .Item("pfpOfficeName")
                mudtProps.CityName = .Item("pfpCityName")
                mudtProps.Phone = .Item("pfpOfficePhone")
                mudtProps.AOHPhone = .Item("pfpAOHPhone")
                mudtProps.PCCBackOffice = .Item("pfpBO_fkey")
            Else
                mudtProps.PCCId = 0
                mudtProps.OfficeCityCode = ""
                mudtProps.CountryCode = ""
                mudtProps.OfficeName = ""
                mudtProps.CityName = ""
                mudtProps.Phone = ""
                mudtProps.AOHPhone = ""
                mudtProps.PCCBackOffice = 0
            End If
            .Close()
        End With
        pobjConn.Close()

    End Sub

    Private Sub DBUpdatePCC()

        If mudtProps.PCCId > 0 Then
            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .CommandText = " UPDATE [AmadeusReports].[dbo].[PNRFinisherPCC]" & _
                               "  SET pfpOfficeCityCode ='" & mudtProps.OfficeCityCode & "'" & _
                               " ,pfpCountryCode ='" & mudtProps.CountryCode & "'" & _
                               " ,pfpOfficeName ='" & mudtProps.OfficeName & "'" & _
                               " ,pfpCityName ='" & mudtProps.CityName & "'" & _
                               " ,pfpOfficePhone ='" & mudtProps.Phone & "'" & _
                               " ,pfpAOHPhone ='" & mudtProps.AOHPhone & "'" & _
                               " WHERE pfpId = " & mudtProps.PCCId
            End With
        Else
            Throw New Exception("Cannot update PCC")
        End If
    End Sub
    Private Sub DBUpdateUser()

        If mudtProps.AgentId > 0 Then
            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .CommandText = " UPDATE [AmadeusReports].[dbo].[PNRFinisherUsers]" & _
                               "  SET [pfAgentQueue] ='" & mudtProps.AgentQueue & "'" & _
                               "     ,[pfAgentOPQueue] ='" & mudtProps.AgentOPQueue & "'" & _
                               "     ,[pfAgentName] ='" & mudtProps.AgentName & "'" & _
                               "     ,[pfAgentEmail] ='" & mudtProps.AgentEmail & "'" & _
                               "     ,[pfAirportName] =" & mudtProps.AirportName & _
                               "     ,[pfAirlineLocator] =" & If(mudtProps.AirlineLocator, 1, 0) & _
                               "     ,[pfClassOfService] =" & If(mudtProps.ClassOfService, 1, 0) & _
                               "     ,[pfBanElectricalEquipment] =" & If(mudtProps.BanElectricalEquipment, 1, 0) & _
                               "     ,[pfBrazilText] =" & If(mudtProps.BrazilText, 1, 0) & _
                               "     ,[pfUSAText] =" & If(mudtProps.USAText, 1, 0) & _
                               "     ,[pfTickets] =" & If(mudtProps.Tickets, 1, 0) & _
                               "     ,[pfPaxSegPerTkt] =" & If(mudtProps.PaxSegPerTkt, 1, 0) & _
                               "     ,[pfShowStopovers] =" & If(mudtProps.ShowStopovers, 1, 0) & _
                               "     ,[pfShowTerminal] =" & If(mudtProps.ShowTerminal, 1, 0) & _
                               "     ,[pfFlyingTime] =" & If(mudtProps.FlyingTime, 1, 0) & _
                               "     ,[pfCostCentre] =" & If(mudtProps.CostCentre, 1, 0) & _
                               "     ,[pfSeating] =" & If(mudtProps.Seating, 1, 0) & _
                               "     ,[pfVessel] =" & If(mudtProps.Vessel, 1, 0) & _
                               "     ,[pfPlainFormat] =" & If(mudtProps.PlainFormat, 1, 0) & _
                               "   WHERE pfPCC = '" & mobjAmadeusUser.PCC & "' AND pfUser = '" & mobjAmadeusUser.User & "'"
                .ExecuteNonQuery()
            End With
            pobjConn.Close()
        Else
            Throw New Exception("Cannot update User")
        End If
    End Sub
    Private Sub DBReadUser()

        Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = " SELECT [pfID] " & _
                           "       ,[pfPCC] " & _
                           "       ,[pfUser] " & _
                           "       ,[pfAgentQueue] " & _
                           "       ,[pfAgentOPQueue] " & _
                           "       ,[pfAgentName] " & _
                           "       ,[pfAgentEmail] " & _
                           "       ,[pfAirportName] " & _
                           "       ,[pfAirlineLocator] " & _
                           "       ,[pfClassOfService] " & _
                           "       ,[pfBanElectricalEquipment] " & _
                           "       ,[pfBrazilText] " & _
                           "       ,[pfUSAText] " & _
                           "       ,[pfTickets] " & _
                           "       ,[pfPaxSegPerTkt] " & _
                           "       ,[pfShowStopovers] " & _
                           "       ,[pfShowTerminal] " & _
                           "       ,[pfFlyingTime] " & _
                           "       ,[pfCostCentre] " & _
                           "       ,[pfSeating] " & _
                           "       ,[pfVessel] " & _
                           "       ,[pfPlainFormat] " & _
                           "       ,[pfAdministrator] " & _
                           "   FROM [AmadeusReports].[dbo].[PNRFinisherUsers] " & _
                           "   WHERE pfPCC = '" & mobjAmadeusUser.PCC & "' AND pfUser = '" & mobjAmadeusUser.User & "'"

            pobjReader = .ExecuteReader
        End With
        With pobjReader
            If pobjReader.Read Then
                mudtProps.AgentId = .Item("pfID")
                mudtProps.AgentQueue = .Item("pfAgentQueue")
                mudtProps.AgentOPQueue = .Item("pfAgentOPQueue")
                mudtProps.AgentName = .Item("pfAgentName")
                mudtProps.AgentEmail = .Item("pfAgentEmail")
                mudtProps.AirportName = .Item("pfAirportName")
                mudtProps.AirlineLocator = .Item("pfAirlineLocator")
                mudtProps.ClassOfService = .Item("pfClassOfService")
                mudtProps.BanElectricalEquipment = .Item("pfBanElectricalEquipment")
                mudtProps.BrazilText = .Item("pfBrazilText")
                mudtProps.USAText = .Item("pfUSAText")
                mudtProps.Tickets = .Item("pfTickets")
                mudtProps.PaxSegPerTkt = .Item("pfPaxSegPerTkt")
                mudtProps.ShowStopovers = .Item("pfShowStopovers")
                mudtProps.ShowTerminal = .Item("pfShowTerminal")
                mudtProps.FlyingTime = .Item("pfFlyingTime")
                mudtProps.CostCentre = .Item("pfCostCentre")
                mudtProps.Seating = .Item("pfSeating")
                mudtProps.Vessel = .Item("pfVessel")
                mudtProps.PlainFormat = .Item("pfPlainFormat")
                mudtProps.Administrator = .Item("pfAdministrator")
            Else
                mudtProps.AgentId = 0
                mudtProps.AgentQueue = ""
                mudtProps.AgentOPQueue = ""
                mudtProps.AgentName = ""
                mudtProps.AgentEmail = ""
                mudtProps.AirportName = 0
                mudtProps.AirlineLocator = False
                mudtProps.ClassOfService = False
                mudtProps.BanElectricalEquipment = False
                mudtProps.BrazilText = False
                mudtProps.USAText = False
                mudtProps.Tickets = False
                mudtProps.PaxSegPerTkt = False
                mudtProps.ShowStopovers = False
                mudtProps.ShowTerminal = False
                mudtProps.FlyingTime = False
                mudtProps.CostCentre = False
                mudtProps.Seating = False
                mudtProps.Vessel = False
                mudtProps.PlainFormat = False
                mudtProps.Administrator = False

            End If
            .Close()
        End With
        pobjConn.Close()

    End Sub

    Public Sub Save()

        Try
            If mflgIsDirtyPCC Then
                DBUpdatePCC()
            End If
            If mflgIsDirtyUser Then
                DBUpdateUser()
            End If

        Catch ex As Exception
            Throw New Exception("Config.Save()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Public ReadOnly Property AmadeusValue(ByVal Key As String) As String
        Get

            ' "CountryCode"    ' %MID%
            ' "OfficePCC"      ' %PCC%
            ' "AgentQueue"     ' %AGENTQ%
            ' "AgentOPQueue"   ' %AGENTOPQ%
            ' "AgentName"      ' %AGENTNAME%
            ' "AgentEmail"     ' %AGENTEMAIL%
            ' "OfficeCityCode" ' %CITYCODE%
            ' "AOHPhone"       ' %AOHP%
            ' "Phone"          ' %PHONE%
            ' "AgentID"        ' %AgentID%
            ' "CityName"       ' %CITYNAME%

            Try

                Dim TempVal As String = ""


                TempVal = mAmadeusReferences.Item(Key)

                If TempVal.IndexOf("%") >= 0 Then
                    TempVal = ReplaceReference(TempVal, "%PCC%", mobjAmadeusUser.PCC)
                    TempVal = ReplaceReference(TempVal, "%AgentID%", mobjAmadeusUser.User)

                    TempVal = ReplaceReference(TempVal, "%MID%", mudtProps.CountryCode)
                    TempVal = ReplaceReference(TempVal, "%AGENTQ%", mudtProps.AgentQueue)
                    TempVal = ReplaceReference(TempVal, "%AGENTOPQ%", mudtProps.AgentOPQueue)
                    TempVal = ReplaceReference(TempVal, "%AGENTNAME%", mudtProps.AgentName)
                    TempVal = ReplaceReference(TempVal, "%AGENTEMAIL%", mudtProps.AgentEmail)
                    TempVal = ReplaceReference(TempVal, "%CITYCODE%", mudtProps.OfficeCityCode)
                    TempVal = ReplaceReference(TempVal, "%CITYNAME%", mudtProps.CityName)
                    TempVal = ReplaceReference(TempVal, "%AOHP%", mudtProps.AOHPhone)
                    TempVal = ReplaceReference(TempVal, "%PHONE%", mudtProps.Phone)
                    TempVal = ReplaceReference(TempVal, "%OFFICENAME%", mudtProps.OfficeName)
                End If
                AmadeusValue = TempVal

            Catch ex As Exception

                Throw New Exception("Key:" & Key & " not found in the collection")

            End Try
        End Get

    End Property

    Public ReadOnly Property isValid As Boolean
        Get
            With mudtProps
                isValid = mobjAmadeusUser.PCC <> "" And
                          mobjAmadeusUser.User <> "" And
                          .AgentQueue <> "" And
                          .AgentOPQueue <> "" And
                          .CountryCode <> "" And
                          .AgentName <> "" And
                          .AgentEmail <> "" And
                          .OfficeCityCode <> "" And
                          .CityName <> "" And
                          .OfficeName <> "" And
                          .AOHPhone <> "" And
                          .PCCBackOffice <> 0 And
                          .Phone <> ""
            End With
        End Get
    End Property

    Private Function ReplaceReference(ByVal InputValue As String, ByVal RefKey As String, ByVal RefValue As String)
        If InputValue.IndexOf(RefKey) >= 0 Then
            ReplaceReference = InputValue.Replace(RefKey, RefValue)
        Else
            ReplaceReference = InputValue
        End If
    End Function
End Class
