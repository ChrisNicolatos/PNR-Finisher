Option Strict Off
Option Explicit On
Public Class Config
    Public Enum OPTItinFormat
        ItnFormatDefault = 0
        ItnFormatPlain = 1
        ItnFormatSeaChefs = 2
        ItnSeaChefsWithCode = 3
        ItnFormatEuronav = 4
    End Enum
    Public Enum GDSCode
        GDSisUnknown = 0
        GDSisAmadeus = 1
        GDSisGalileo = 2
    End Enum
    Private Structure ClassProps
        Dim PCCId As Integer
        Dim OfficeCityCode As String
        Dim CountryCode As String
        Dim OfficeName As String
        Dim CityName As String
        Dim Phone As String
        Dim AOHPhone As String
        Dim PCCBackOffice As Integer
        Dim pCCDBDataSource As String
        Dim pCCDBInitialCatalog As String
        Dim pCCDBUserId As String
        Dim pCCDBUserPassword As String
        Dim pCCIATANumber As String
        Dim PCCFormalOfficeName As String

        Dim AgentId As Integer
        Dim AgentQueue As String
        Dim AgentOPQueue As String
        Dim AgentName As String
        Dim AgentEmail As String
        Dim AirportName As Integer
        Dim ShowAirlineLocator As Boolean
        Dim ShowClassOfService As Boolean
        Dim ShowBanElectricalEquipment As Boolean
        Dim ShowBrazilText As Boolean
        Dim ShowUSAText As Boolean
        Dim ShowTickets As Boolean
        Dim ShowPaxSegPerTkt As Boolean
        Dim ShowStopovers As Boolean
        Dim ShowTerminal As Boolean
        Dim ShowFlyingTime As Boolean
        Dim ShowCostCentre As Boolean
        Dim ShowSeating As Boolean
        Dim ShowVessel As Boolean
        Dim Administrator As String
        Dim FormatStyle As Integer
        Dim OSMVesselGroup As Integer
        Dim OSMLoGPerPax As Boolean
        Dim OSMLoGOnsigner As Boolean
        Dim OSMLoGPath As String

    End Structure

    Private mudtProps As ClassProps
    Private mobjGDSUser As GDSUser
    Private mflgIsDirtyPCC As Boolean
    Private mflgIsDirtyUser As Boolean


    Dim mGDSReferences As New Config_GDSReferences.Collection

    Public Sub New(mGDSUser As GDSUser)

        Try
            mobjGDSUser = mGDSUser

            mflgIsDirtyPCC = False
            mflgIsDirtyUser = False
            mGDSReferences.Clear()
            DBReadPCC()
            If PCCId = 0 Then
                Throw New Exception("You are signed in to Amadeus PCC : " & mobjGDSUser.PCC & vbCrLf & "This PCC is not registered in the PNR FInisher" & vbCrLf & "Please jump to your own PCC and restart the program")
            End If
            DBReadUser()
            If AgentID = 0 Then
                Throw New Exception("You are signed in to Amadeus PCC : " & mobjGDSUser.PCC & " as user : " & mobjGDSUser.User & vbCrLf & "This user is not registered in the PNR FInisher" & vbCrLf & "Please jump to your own PCC and restart the program")
            End If
            mGDSReferences.Read(PCCBackOffice, mobjGDSUser.GDSCode)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Public ReadOnly Property Administrator As Boolean
        Get
            Administrator = mudtProps.Administrator
        End Get
    End Property
    Public Property FormatStyle As OPTItinFormat
        Get
            FormatStyle = mudtProps.FormatStyle
        End Get
        Set(value As OPTItinFormat)
            If value <> mudtProps.FormatStyle Then
                mflgIsDirtyUser = True
            End If
            mudtProps.FormatStyle = value
        End Set
    End Property
    Public Property OSMVesselGroup As Integer
        Get
            OSMVesselGroup = mudtProps.OSMVesselGroup
        End Get
        Set(value As Integer)
            If value <> mudtProps.OSMVesselGroup Then
                mflgIsDirtyUser = True
            End If
            mudtProps.OSMVesselGroup = value
        End Set
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
    Public Property ShowAirlineLocator As Boolean
        Get
            ShowAirlineLocator = mudtProps.ShowAirlineLocator
        End Get
        Set(value As Boolean)
            If value <> mudtProps.ShowAirlineLocator Then
                mflgIsDirtyUser = True
            End If
            mudtProps.ShowAirlineLocator = value
        End Set
    End Property
    Public Property ShowClassOfService As Boolean
        Get
            ShowClassOfService = mudtProps.ShowClassOfService
        End Get
        Set(value As Boolean)
            If value <> mudtProps.ShowClassOfService Then
                mflgIsDirtyUser = True
            End If
            mudtProps.ShowClassOfService = value
        End Set
    End Property
    Public Property ShowBanElectricalEquipment As Boolean
        Get
            ShowBanElectricalEquipment = mudtProps.ShowBanElectricalEquipment
        End Get
        Set(value As Boolean)
            If value <> mudtProps.ShowBanElectricalEquipment Then
                mflgIsDirtyUser = True
            End If
            mudtProps.ShowBanElectricalEquipment = value
        End Set
    End Property
    Public Property ShowBrazilText As Boolean
        Get
            ShowBrazilText = mudtProps.ShowBrazilText
        End Get
        Set(value As Boolean)
            If value <> mudtProps.ShowBrazilText Then
                mflgIsDirtyUser = True
            End If
            mudtProps.ShowBrazilText = value
        End Set
    End Property
    Public Property ShowUSAText As Boolean
        Get
            ShowUSAText = mudtProps.ShowUSAText
        End Get
        Set(value As Boolean)
            If value <> mudtProps.ShowUSAText Then
                mflgIsDirtyUser = True
            End If
            mudtProps.ShowUSAText = value
        End Set
    End Property
    Public Property ShowTickets As Boolean
        Get
            ShowTickets = mudtProps.ShowTickets
        End Get
        Set(value As Boolean)
            If value <> mudtProps.ShowTickets Then
                mflgIsDirtyUser = True
            End If
            mudtProps.ShowTickets = value
        End Set
    End Property
    Public Property ShowPaxSegPerTkt As Boolean
        Get
            ShowPaxSegPerTkt = mudtProps.ShowPaxSegPerTkt
        End Get
        Set(value As Boolean)
            If value <> mudtProps.ShowPaxSegPerTkt Then
                mflgIsDirtyUser = True
            End If
            mudtProps.ShowPaxSegPerTkt = value
        End Set
    End Property
    Public Property ShowStopovers As Boolean
        Get
            ShowStopovers = mudtProps.ShowStopovers
        End Get
        Set(value As Boolean)
            If value <> mudtProps.ShowStopovers Then
                mflgIsDirtyUser = True
            End If
            mudtProps.ShowStopovers = value
        End Set
    End Property
    Public Property ShowTerminal As Boolean
        Get
            ShowTerminal = mudtProps.ShowTerminal
        End Get
        Set(value As Boolean)
            If value <> mudtProps.ShowTerminal Then
                mflgIsDirtyUser = True
            End If
            mudtProps.ShowTerminal = value
        End Set
    End Property
    Public Property ShowFlyingTime As Boolean
        Get
            ShowFlyingTime = mudtProps.ShowFlyingTime
        End Get
        Set(value As Boolean)
            If value <> mudtProps.ShowFlyingTime Then
                mflgIsDirtyUser = True
            End If
            mudtProps.ShowFlyingTime = value
        End Set
    End Property
    Public Property ShowCostCentre As Boolean
        Get
            ShowCostCentre = mudtProps.ShowCostCentre
        End Get
        Set(value As Boolean)
            If value <> mudtProps.ShowCostCentre Then
                mflgIsDirtyUser = True
            End If
            mudtProps.ShowCostCentre = value
        End Set
    End Property
    Public Property ShowSeating As Boolean
        Get
            ShowSeating = mudtProps.ShowSeating
        End Get
        Set(value As Boolean)
            If value <> mudtProps.ShowSeating Then
                mflgIsDirtyUser = True
            End If
            mudtProps.ShowSeating = value
        End Set
    End Property
    Public Property ShowVessel As Boolean
        Get
            ShowVessel = mudtProps.ShowVessel
        End Get
        Set(value As Boolean)
            If value <> mudtProps.ShowVessel Then
                mflgIsDirtyUser = True
            End If
            mudtProps.ShowVessel = value
        End Set
    End Property
    'Public Property UsePlainFormat As Boolean
    '    Get
    '        UsePlainFormat = mudtProps.UsePlainFormat
    '    End Get
    '    Set(value As Boolean)
    '        If value <> mudtProps.UsePlainFormat Then
    '            mflgIsDirtyUser = True
    '        End If
    '        mudtProps.UsePlainFormat = value
    '    End Set
    'End Property
    Public Property OSMLoGPerPax As Boolean
        Get
            OSMLoGPerPax = mudtProps.OSMLoGPerPax
        End Get
        Set(value As Boolean)
            If value <> mudtProps.OSMLoGPerPax Then
                mflgIsDirtyUser = True
            End If
            mudtProps.OSMLoGPerPax = value
        End Set
    End Property
    Public Property OSMLoGOnSigner As Boolean
        Get
            OSMLoGOnSigner = mudtProps.OSMLoGOnsigner
        End Get
        Set(value As Boolean)
            If value <> mudtProps.OSMLoGOnsigner Then
                mflgIsDirtyUser = True
            End If
            mudtProps.OSMLoGOnsigner = value
        End Set
    End Property
    Public Property OSMLoGPath As String
        Get
            OSMLoGPath = mudtProps.OSMLoGPath
        End Get
        Set(value As String)
            If value <> mudtProps.OSMLoGPath Then
                mflgIsDirtyUser = True
            End If
            mudtProps.OSMLoGPath = value
        End Set
    End Property
    Public ReadOnly Property PCCId As Integer
        Get
            PCCId = mudtProps.PCCId
        End Get
    End Property
    Public Property OfficeCityCode As String
        Get
            OfficeCityCode = mudtProps.OfficeCityCode.ToUpper
        End Get
        Set(value As String)
            If value <> mudtProps.OfficeCityCode Then
                mflgIsDirtyPCC = True
            End If
            mudtProps.OfficeCityCode = value
        End Set
    End Property
    Public Property CountryCode As String
        Get
            CountryCode = mudtProps.CountryCode.ToUpper
        End Get
        Set(value As String)
            If value <> mudtProps.CountryCode Then
                mflgIsDirtyPCC = True
            End If
            mudtProps.CountryCode = value
        End Set
    End Property
    Public Property OfficeName As String
        Get
            OfficeName = mudtProps.OfficeName.ToUpper
        End Get
        Set(value As String)
            If value <> mudtProps.OfficeName Then
                mflgIsDirtyPCC = True
            End If
            mudtProps.OfficeName = value
        End Set
    End Property
    Public Property FormalOfficeName As String
        Get
            FormalOfficeName = mudtProps.PCCFormalOfficeName
        End Get
        Set(value As String)
            If value <> mudtProps.PCCFormalOfficeName Then
                mflgIsDirtyPCC = True
            End If
            mudtProps.PCCFormalOfficeName = value
        End Set
    End Property
    Public Property CityName As String
        Get
            CityName = mudtProps.CityName.ToUpper
        End Get
        Set(value As String)
            If value <> mudtProps.CityName Then
                mflgIsDirtyPCC = True
            End If
            mudtProps.CityName = value
        End Set
    End Property
    Public Property Phone As String
        Get
            Phone = mudtProps.Phone.ToUpper
        End Get
        Set(value As String)
            If value <> mudtProps.Phone Then
                mflgIsDirtyPCC = True
            End If
            mudtProps.Phone = value
        End Set
    End Property
    Public Property AOHPhone As String
        Get
            AOHPhone = mudtProps.AOHPhone.ToUpper
        End Get
        Set(value As String)
            If value <> mudtProps.AOHPhone Then
                mflgIsDirtyPCC = True
            End If
            mudtProps.AOHPhone = value
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
    Public Property PCCDBDataSource As String
        Get
            Return mudtProps.pCCDBDataSource
        End Get
        Set(value As String)
            mudtProps.pCCDBDataSource = value
        End Set
    End Property

    Public Property PCCDBInitialCatalog As String
        Get
            Return mudtProps.pCCDBInitialCatalog
        End Get
        Set(value As String)
            mudtProps.pCCDBInitialCatalog = value
        End Set
    End Property

    Public Property PCCDBUserId As String
        Get
            Return mudtProps.pCCDBUserId
        End Get
        Set(value As String)
            mudtProps.pCCDBUserId = value
        End Set
    End Property

    Public Property PCCDBUserPassword As String
        Get
            Return mudtProps.pCCDBUserPassword
        End Get
        Set(value As String)
            mudtProps.pCCDBUserPassword = value
        End Set
    End Property
    Public Property IATANumber As String
        Get
            IATANumber = mudtProps.pCCIATANumber
        End Get
        Set(value As String)
            mudtProps.pCCIATANumber = value
        End Set
    End Property
    Public ReadOnly Property AgentID As Integer
        Get
            AgentID = mudtProps.AgentId
        End Get
    End Property
    Public Property AgentQueue As String
        Get
            AgentQueue = mudtProps.AgentQueue.ToUpper
        End Get
        Set(value As String)
            If value <> mudtProps.AgentQueue Then
                mflgIsDirtyUser = True
            End If
            mudtProps.AgentQueue = value
        End Set
    End Property
    Public Property AgentOPQueue As String
        Get
            AgentOPQueue = mudtProps.AgentOPQueue.ToUpper
        End Get
        Set(value As String)
            If value <> mudtProps.AgentOPQueue Then
                mflgIsDirtyUser = True
            End If
            mudtProps.AgentOPQueue = value
        End Set
    End Property
    Public Property AgentName As String
        Get
            AgentName = mudtProps.AgentName.ToUpper
        End Get
        Set(value As String)
            If value <> mudtProps.AgentName Then
                mflgIsDirtyUser = True
            End If
            mudtProps.AgentName = value
        End Set
    End Property
    Public Property AgentEmail As String
        Get
            AgentEmail = mudtProps.AgentEmail.ToUpper
        End Get
        Set(value As String)
            If value <> mudtProps.AgentEmail Then
                mflgIsDirtyUser = True
            End If
            mudtProps.AgentEmail = value
        End Set
    End Property
    Public Property AirportName As Integer
        Get
            AirportName = mudtProps.AirportName
        End Get
        Set(value As Integer)
            If value <> mudtProps.AirportName Then
                mflgIsDirtyUser = True
            End If
            mudtProps.AirportName = value
        End Set
    End Property

    Public ReadOnly Property GDSPcc As String
        Get
            GDSPcc = mobjGDSUser.PCC.ToUpper
        End Get
    End Property
    Public ReadOnly Property GDSUser As String
        Get
            GDSUser = mobjGDSUser.User.ToUpper
        End Get
    End Property
    Private Sub DBReadPCC()

        Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .Parameters.Add("@PCC", SqlDbType.NVarChar, 9).Value = mobjGDSUser.PCC
            .CommandText = " SELECT pfpId " &
                           " ,pfpOfficeCityCode " &
                           " ,pfpCountryCode " &
                           " ,pfpOfficeName " &
                           " ,pfpCityName " &
                           " ,pfpOfficePhone " &
                           " ,pfpAOHPhone " &
                           " ,pfpBO_fkey " &
                           " ,pfpDBDataSource " &
                           " ,pfpDBInitialCatalog " &
                           " ,pfpDBUserId " &
                           " ,pfpDBUserPassword " &
                           " ,pfpIATANumber " &
                           " ,pfpFormalOfficeName " &
                           " FROM [AmadeusReports].[dbo].[PNRFinisherPCC] " &
                           " WHERE pfpPCC = @PCC"

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
                mudtProps.pCCDBDataSource = .Item("pfpDBDataSource")
                mudtProps.pCCDBInitialCatalog = .Item("pfpDBInitialCatalog")
                mudtProps.pCCDBUserId = .Item("pfpDBUserId")
                mudtProps.pCCDBUserPassword = .Item("pfpDBUserPassword")
                mudtProps.pCCIATANumber = .Item("pfpIATANumber")
                mudtProps.PCCFormalOfficeName = .Item("pfpFormalOfficeName")
            Else
                mudtProps.PCCId = 0
                mudtProps.OfficeCityCode = ""
                mudtProps.CountryCode = ""
                mudtProps.OfficeName = ""
                mudtProps.CityName = ""
                mudtProps.Phone = ""
                mudtProps.AOHPhone = ""
                mudtProps.PCCBackOffice = 0
                mudtProps.pCCDBDataSource = ""
                mudtProps.pCCDBInitialCatalog = ""
                mudtProps.pCCDBUserId = ""
                mudtProps.pCCDBUserPassword = ""
                mudtProps.pCCIATANumber = ""
                mudtProps.PCCFormalOfficeName = ""
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
                .Parameters.Add("@PCCId", SqlDbType.Int).Value = mudtProps.PCCId
                .Parameters.Add("@pfpOfficeCityCode", SqlDbType.NChar, 3).Value = OfficeCityCode
                .Parameters.Add("@pfpCountryCode", SqlDbType.NChar, 2).Value = CountryCode
                .Parameters.Add("@pfpOfficeName", SqlDbType.NVarChar, 254).Value = OfficeName
                .Parameters.Add("@pfpCityName", SqlDbType.NVarChar, 254).Value = CityName
                .Parameters.Add("@pfpOfficePhone", SqlDbType.NVarChar, 254).Value = Phone
                .Parameters.Add("@pfpAOHPhone", SqlDbType.NVarChar, 254).Value = AOHPhone
                .CommandText = " UPDATE [AmadeusReports].[dbo].[PNRFinisherPCC]" &
                               "  SET pfpOfficeCityCode =@pfpOfficeCityCode " &
                               " ,pfpCountryCode =@pfpCountryCode " &
                               " ,pfpOfficeName =@pfpOfficeName " &
                               " ,pfpCityName =@pfpCityName " &
                               " ,pfpOfficePhone =@pfpOfficePhone " &
                               " ,pfpAOHPhone =@pfpAOHPhone " &
                               " WHERE pfpId = @PCCId"
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
                .CommandText = " UPDATE AmadeusReports.dbo.PNRFinisherUsers" &
                               "  SET pfAgentQueue ='" & AgentQueue & "'" &
                               "     ,pfAgentOPQueue ='" & AgentOPQueue & "'" &
                               "     ,pfAgentName ='" & AgentName & "'" &
                               "     ,pfAgentEmail ='" & AgentEmail & "'" &
                               "     ,pfAirportName =" & AirportName &
                               "     ,pfAirlineLocator =" & If(ShowAirlineLocator, 1, 0) &
                               "     ,pfClassOfService =" & If(ShowClassOfService, 1, 0) &
                               "     ,pfBanElectricalEquipment =" & If(ShowBanElectricalEquipment, 1, 0) &
                               "     ,pfBrazilText =" & If(ShowBrazilText, 1, 0) &
                               "     ,pfUSAText =" & If(ShowUSAText, 1, 0) &
                               "     ,pfTickets =" & If(ShowTickets, 1, 0) &
                               "     ,pfPaxSegPerTkt =" & If(ShowPaxSegPerTkt, 1, 0) &
                               "     ,pfShowStopovers =" & If(ShowStopovers, 1, 0) &
                               "     ,pfShowTerminal =" & If(ShowTerminal, 1, 0) &
                               "     ,pfFlyingTime =" & If(ShowFlyingTime, 1, 0) &
                               "     ,pfCostCentre =" & If(ShowCostCentre, 1, 0) &
                               "     ,pfSeating =" & If(ShowSeating, 1, 0) &
                               "     ,pfVessel =" & If(ShowVessel, 1, 0) &
                               "     ,pfPlainFormat = 0" &
                               "     ,pfFormatStyle =" & FormatStyle &
                               "     ,pfOSMVesselGroup = " & OSMVesselGroup &
                               "     ,pfOSMLOGPerPax = " & If(OSMLoGPerPax, 1, 0) &
                               "     ,pfOSMLOGOnSigner = " & If(OSMLoGOnSigner, 1, 0) &
                               "     ,pfOSMLOGPath = '" & OSMLoGPath & "' " &
                               "   WHERE pfPCC = '" & mobjGDSUser.PCC & "' AND pfUser = '" & mobjGDSUser.User & "'"
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
            .CommandText = " SELECT [pfID] " &
                           "       ,[pfPCC] " &
                           "       ,[pfUser] " &
                           "       ,[pfAgentQueue] " &
                           "       ,[pfAgentOPQueue] " &
                           "       ,[pfAgentName] " &
                           "       ,[pfAgentEmail] " &
                           "       ,[pfAirportName] " &
                           "       ,[pfAirlineLocator] " &
                           "       ,[pfClassOfService] " &
                           "       ,[pfBanElectricalEquipment] " &
                           "       ,[pfBrazilText] " &
                           "       ,[pfUSAText] " &
                           "       ,[pfTickets] " &
                           "       ,[pfPaxSegPerTkt] " &
                           "       ,[pfShowStopovers] " &
                           "       ,[pfShowTerminal] " &
                           "       ,[pfFlyingTime] " &
                           "       ,[pfCostCentre] " &
                           "       ,[pfSeating] " &
                           "       ,[pfVessel] " &
                           "       ,[pfPlainFormat] " &
                           "       ,[pfAdministrator] " &
                           "       ,[pfFormatStyle] " &
                           "       ,ISNULL(pfOSMVesselGroup,0) AS pfOSMVesselGroup " &
                           "       ,ISNULL(pfOSMLOGPerPax,0) AS pfOSMLOGPerPax " &
                           "       ,ISNULL(pfOSMLOGOnSigner,0) AS pfOSMLOGOnSigner " &
                           "       ,ISNULL(pfOSMLOGPath,'') AS pfOSMLOGPath " &
                           "   FROM [AmadeusReports].[dbo].[PNRFinisherUsers] " &
                           "   WHERE pfPCC = '" & mobjGDSUser.PCC & "' AND pfUser = '" & mobjGDSUser.User & "'"
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
                mudtProps.ShowAirlineLocator = .Item("pfAirlineLocator")
                mudtProps.ShowClassOfService = .Item("pfClassOfService")
                mudtProps.ShowBanElectricalEquipment = .Item("pfBanElectricalEquipment")
                mudtProps.ShowBrazilText = .Item("pfBrazilText")
                mudtProps.ShowUSAText = .Item("pfUSAText")
                mudtProps.ShowTickets = .Item("pfTickets")
                mudtProps.ShowPaxSegPerTkt = .Item("pfPaxSegPerTkt")
                mudtProps.ShowStopovers = .Item("pfShowStopovers")
                mudtProps.ShowTerminal = .Item("pfShowTerminal")
                mudtProps.ShowFlyingTime = .Item("pfFlyingTime")
                mudtProps.ShowCostCentre = .Item("pfCostCentre")
                mudtProps.ShowSeating = .Item("pfSeating")
                mudtProps.ShowVessel = .Item("pfVessel")
                mudtProps.Administrator = .Item("pfAdministrator")
                mudtProps.FormatStyle = .Item("pfFormatStyle")
                mudtProps.OSMVesselGroup = .Item("pfOSMVesselGroup")
                mudtProps.OSMLoGPerPax = .Item("pfOSMLOGPerPax")
                mudtProps.OSMLoGOnsigner = .Item("pfOSMLOGOnSigner")
                mudtProps.OSMLoGPath = .Item("pfOSMLOGPath")
            Else
                mudtProps.AgentId = 0
                mudtProps.AgentQueue = ""
                mudtProps.AgentOPQueue = ""
                mudtProps.AgentName = ""
                mudtProps.AgentEmail = ""
                mudtProps.AirportName = 0
                mudtProps.ShowAirlineLocator = False
                mudtProps.ShowClassOfService = False
                mudtProps.ShowBanElectricalEquipment = False
                mudtProps.ShowBrazilText = False
                mudtProps.ShowUSAText = False
                mudtProps.ShowTickets = False
                mudtProps.ShowPaxSegPerTkt = False
                mudtProps.ShowStopovers = False
                mudtProps.ShowTerminal = False
                mudtProps.ShowFlyingTime = False
                mudtProps.ShowCostCentre = False
                mudtProps.ShowSeating = False
                mudtProps.ShowVessel = False
                mudtProps.Administrator = False
                mudtProps.FormatStyle = 0
                mudtProps.OSMVesselGroup = 0
                mudtProps.OSMLoGPerPax = False
                mudtProps.OSMLoGOnsigner = False
                mudtProps.OSMLoGPath = ""
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
    Public ReadOnly Property GDSValue(ByVal Key As String) As String
        Get
            Try
                GDSValue = ConvertGDSValue(mGDSReferences.Item(Key).Value)
            Catch ex As Exception
                Throw New Exception("Key:" & Key & " not found in the collection")
            End Try
        End Get

    End Property
    Public ReadOnly Property GDSElement(ByVal Key As String) As String
        Get
            Try
                GDSElement = ConvertGDSValue(mGDSReferences.Item(Key).Element)
            Catch ex As Exception
                Throw New Exception("Key:" & Key & " not found in the collection")
            End Try
        End Get

    End Property
    Public ReadOnly Property GDSReferenceID(ByVal Key As String) As String
        Get
            Try
                GDSReferenceID = ConvertGDSValue(mGDSReferences.Item(Key).RefId)
            Catch ex As Exception
                Throw New Exception("Key:" & Key & " not found in the collection")
            End Try
        End Get

    End Property
    Public ReadOnly Property GDSReferenceDetail(ByVal Key As String) As String
        Get
            Try
                GDSReferenceDetail = ConvertGDSValue(mGDSReferences.Item(Key).RefDetail)
            Catch ex As Exception
                Throw New Exception("Key:" & Key & " not found in the collection")
            End Try
        End Get

    End Property
    Public ReadOnly Property CloseOffValue(ByVal CloseOffEntry As String) As String
        Get
            CloseOffValue = ConvertGDSValue(CloseOffEntry)
        End Get
    End Property

    Public Function ConvertGDSValue(ByVal ValueToConvert As String) As String

        ' "CountryCode"           ' %MID%
        ' "OfficePCC"             ' %PCC%
        ' "AgentQueue"            ' %AGENTQ%
        ' "AgentOPQueue"          ' %AGENTOPQ%
        ' "AgentName"             ' %AGENTNAME%
        ' "AgentEmail"            ' %AGENTEMAIL%
        ' "OfficeCityCode"        ' %CITYCODE%
        ' "AOHPhone"              ' %AOHP%
        ' "Phone"                 ' %PHONE%
        ' "AgentID"               ' %AgentID%
        ' "CityName"              ' %CITYNAME%
        ' GalileoTrackingCode     ' %GALTRACK%

        ConvertGDSValue = ValueToConvert

        If ConvertGDSValue.IndexOf("%") >= 0 Then
            If AgentQueue.IndexOf("/") >= 0 Then
                ConvertGDSValue = ReplaceReference(ConvertGDSValue, "%PCC-AGENTQ%", "/" & AgentQueue)
                ConvertGDSValue = ReplaceReference(ConvertGDSValue, "%AGENTQ%", "/" & AgentQueue)
            Else
                ConvertGDSValue = ReplaceReference(ConvertGDSValue, "%PCC-AGENTQ%", mobjGDSUser.PCC & "/" & AgentQueue)
                ConvertGDSValue = ReplaceReference(ConvertGDSValue, "%AGENTQ%", AgentQueue)
            End If
            If AgentOPQueue.IndexOf("/") >= 0 Then
                ConvertGDSValue = ReplaceReference(ConvertGDSValue, "%AGENTOPQ%", "/" & AgentOPQueue)
            Else
                ConvertGDSValue = ReplaceReference(ConvertGDSValue, "%AGENTOPQ%", AgentOPQueue)
            End If
            Do While ConvertGDSValue.IndexOf("//") >= 0
                ConvertGDSValue.Replace("//", "/")
            Loop
            ConvertGDSValue = ReplaceReference(ConvertGDSValue, "%PCC%", mobjGDSUser.PCC)
            ConvertGDSValue = ReplaceReference(ConvertGDSValue, "%AgentID%", mobjGDSUser.User)

            ConvertGDSValue = ReplaceReference(ConvertGDSValue, "%MID%", CountryCode)
            ConvertGDSValue = ReplaceReference(ConvertGDSValue, "%AGENTNAME%", AgentName)
            ConvertGDSValue = ReplaceReference(ConvertGDSValue, "%AGENTEMAIL%", AgentEmail)
            ConvertGDSValue = ReplaceReference(ConvertGDSValue, "%CITYCODE%", OfficeCityCode)
            ConvertGDSValue = ReplaceReference(ConvertGDSValue, "%CITYNAME%", CityName)
            ConvertGDSValue = ReplaceReference(ConvertGDSValue, "%AOHP%", AOHPhone)
            ConvertGDSValue = ReplaceReference(ConvertGDSValue, "%PHONE%", Phone)
            ConvertGDSValue = ReplaceReference(ConvertGDSValue, "%OFFICENAME%", OfficeName)

        End If

    End Function
    Public ReadOnly Property isValid As Boolean
        Get
            With mudtProps
                isValid = mobjGDSUser.PCC <> "" And
                          mobjGDSUser.User <> "" And
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
                          .pCCDBDataSource <> "" And
                          .pCCDBInitialCatalog <> "" And
                          .pCCDBUserId <> "" And
                          .pCCDBUserPassword <> "" And
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
