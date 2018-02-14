﻿Option Strict Off
Option Explicit On
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
        Dim FormatStyle As Integer
        Dim OSMVesselGroup As Integer

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
            If PCCId = 0 Then
                Throw New Exception("You are signed in to Amadeus PCC : " & mobjAmadeusUser.PCC & vbCrLf & "This PCC is not registered in the PNR FInisher" & vbCrLf & "Please jump to your own PCC and restart the program")
            End If
            DBReadUser()
            If AgentID = 0 Then
                Throw New Exception("You are signed in to Amadeus PCC : " & mobjAmadeusUser.PCC & " as user : " & mobjAmadeusUser.User & vbCrLf & "This user is not registered in the PNR FInisher" & vbCrLf & "Please jump to your own PCC and restart the program")
            End If
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
    Public Property FormatStyle As Integer
        Get
            FormatStyle = mudtProps.FormatStyle
        End Get
        Set(value As Integer)
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
    Public Property AirlineLocator As Boolean
        Get
            AirlineLocator = mudtProps.AirlineLocator
        End Get
        Set(value As Boolean)
            If value <> mudtProps.AirlineLocator Then
                mflgIsDirtyUser = True
            End If
            mudtProps.AirlineLocator = value
        End Set
    End Property
    Public Property ClassOfService As Boolean
        Get
            ClassOfService = mudtProps.ClassOfService
        End Get
        Set(value As Boolean)
            If value <> mudtProps.ClassOfService Then
                mflgIsDirtyUser = True
            End If
            mudtProps.ClassOfService = value
        End Set
    End Property
    Public Property BanElectricalEquipment As Boolean
        Get
            BanElectricalEquipment = mudtProps.BanElectricalEquipment
        End Get
        Set(value As Boolean)
            If value <> mudtProps.BanElectricalEquipment Then
                mflgIsDirtyUser = True
            End If
            mudtProps.BanElectricalEquipment = value
        End Set
    End Property
    Public Property BrazilText As Boolean
        Get
            BrazilText = mudtProps.BrazilText
        End Get
        Set(value As Boolean)
            If value <> mudtProps.BrazilText Then
                mflgIsDirtyUser = True
            End If
            mudtProps.BrazilText = value
        End Set
    End Property
    Public Property USAText As Boolean
        Get
            USAText = mudtProps.USAText
        End Get
        Set(value As Boolean)
            If value <> mudtProps.USAText Then
                mflgIsDirtyUser = True
            End If
            mudtProps.USAText = value
        End Set
    End Property
    Public Property Tickets As Boolean
        Get
            Tickets = mudtProps.Tickets
        End Get
        Set(value As Boolean)
            If value <> mudtProps.Tickets Then
                mflgIsDirtyUser = True
            End If
            mudtProps.Tickets = value
        End Set
    End Property
    Public Property PaxSegPerTkt As Boolean
        Get
            PaxSegPerTkt = mudtProps.PaxSegPerTkt
        End Get
        Set(value As Boolean)
            If value <> mudtProps.PaxSegPerTkt Then
                mflgIsDirtyUser = True
            End If
            mudtProps.PaxSegPerTkt = value
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
    Public Property FlyingTime As Boolean
        Get
            FlyingTime = mudtProps.FlyingTime
        End Get
        Set(value As Boolean)
            If value <> mudtProps.FlyingTime Then
                mflgIsDirtyUser = True
            End If
            mudtProps.FlyingTime = value
        End Set
    End Property
    Public Property CostCentre As Boolean
        Get
            CostCentre = mudtProps.CostCentre
        End Get
        Set(value As Boolean)
            If value <> mudtProps.CostCentre Then
                mflgIsDirtyUser = True
            End If
            mudtProps.CostCentre = value
        End Set
    End Property
    Public Property Seating As Boolean
        Get
            Seating = mudtProps.Seating
        End Get
        Set(value As Boolean)
            If value <> mudtProps.Seating Then
                mflgIsDirtyUser = True
            End If
            mudtProps.Seating = value
        End Set
    End Property
    Public Property Vessel As Boolean
        Get
            Vessel = mudtProps.Vessel
        End Get
        Set(value As Boolean)
            If value <> mudtProps.Vessel Then
                mflgIsDirtyUser = True
            End If
            mudtProps.Vessel = value
        End Set
    End Property
    Public Property PlainFormat As Boolean
        Get
            PlainFormat = mudtProps.PlainFormat
        End Get
        Set(value As Boolean)
            If value <> mudtProps.PlainFormat Then
                mflgIsDirtyUser = True
            End If
            mudtProps.PlainFormat = value
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

    Public ReadOnly Property AmadeusPCC As String
        Get
            AmadeusPCC = mobjAmadeusUser.PCC.ToUpper
        End Get
    End Property
    Public ReadOnly Property AmadeusUser As String
        Get
            AmadeusUser = mobjAmadeusUser.User.ToUpper
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
                           " WHERE pfrGDS_fkey = 1 AND pfrBO_fkey = " & PCCBackOffice
            pobjReader = .ExecuteReader
        End With

        mAmadeusReferences.Clear()

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
                           " ,pfpDBDataSource " &
                           " ,pfpDBInitialCatalog " &
                           " ,pfpDBUserId " &
                           " ,pfpDBUserPassword " &
                           " ,pfpIATANumber " &
                           " ,pfpFormalOfficeName " &
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
                .CommandText = " UPDATE [AmadeusReports].[dbo].[PNRFinisherPCC]" &
                               "  SET pfpOfficeCityCode ='" & OfficeCityCode & "'" &
                               " ,pfpCountryCode ='" & CountryCode & "'" &
                               " ,pfpOfficeName ='" & OfficeName & "'" &
                               " ,pfpCityName ='" & CityName & "'" &
                               " ,pfpOfficePhone ='" & Phone & "'" &
                               " ,pfpAOHPhone ='" & AOHPhone & "'" &
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
                .CommandText = " UPDATE [AmadeusReports].[dbo].[PNRFinisherUsers]" &
                               "  SET [pfAgentQueue] ='" & AgentQueue & "'" &
                               "     ,[pfAgentOPQueue] ='" & AgentOPQueue & "'" &
                               "     ,[pfAgentName] ='" & AgentName & "'" &
                               "     ,[pfAgentEmail] ='" & AgentEmail & "'" &
                               "     ,[pfAirportName] =" & AirportName &
                               "     ,[pfAirlineLocator] =" & If(AirlineLocator, 1, 0) &
                               "     ,[pfClassOfService] =" & If(ClassOfService, 1, 0) &
                               "     ,[pfBanElectricalEquipment] =" & If(BanElectricalEquipment, 1, 0) &
                               "     ,[pfBrazilText] =" & If(BrazilText, 1, 0) &
                               "     ,[pfUSAText] =" & If(USAText, 1, 0) &
                               "     ,[pfTickets] =" & If(Tickets, 1, 0) &
                               "     ,[pfPaxSegPerTkt] =" & If(PaxSegPerTkt, 1, 0) &
                               "     ,[pfShowStopovers] =" & If(ShowStopovers, 1, 0) &
                               "     ,[pfShowTerminal] =" & If(ShowTerminal, 1, 0) &
                               "     ,[pfFlyingTime] =" & If(FlyingTime, 1, 0) &
                               "     ,[pfCostCentre] =" & If(CostCentre, 1, 0) &
                               "     ,[pfSeating] =" & If(Seating, 1, 0) &
                               "     ,[pfVessel] =" & If(Vessel, 1, 0) &
                               "     ,[pfPlainFormat] =" & If(PlainFormat, 1, 0) &
                               "     ,[pfFormatStyle] =" & FormatStyle &
                               "     ,pfOSMVesselGroup = " & OSMVesselGroup &
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
                           "   FROM [AmadeusReports].[dbo].[PNRFinisherUsers] " &
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
                mudtProps.FormatStyle = .Item("pfFormatStyle")
                mudtProps.OSMVesselGroup = .Item("pfOSMVesselGroup")
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
                mudtProps.FormatStyle = 0
                mudtProps.OSMVesselGroup = 0
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
            Try
                AmadeusValue = ConvertAmadeusValue(mAmadeusReferences.Item(Key))
            Catch ex As Exception
                Throw New Exception("Key:" & Key & " not found in the collection")
            End Try
        End Get

    End Property
    Public ReadOnly Property CloseOffValue(ByVal CloseOffEntry As String) As String
        Get
            CloseOffValue = ConvertAmadeusValue(CloseOffEntry)
        End Get
    End Property
    Public Function ConvertAmadeusValue(ByVal ValueToConvert As String) As String

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

        ConvertAmadeusValue = ValueToConvert

        If ConvertAmadeusValue.IndexOf("%") >= 0 Then
            If AgentQueue.IndexOf("/") >= 0 Then
                ConvertAmadeusValue = ReplaceReference(ConvertAmadeusValue, "%PCC-AGENTQ%", "/" & AgentQueue)
                ConvertAmadeusValue = ReplaceReference(ConvertAmadeusValue, "%AGENTQ%", "/" & AgentQueue)
            Else
                ConvertAmadeusValue = ReplaceReference(ConvertAmadeusValue, "%PCC-AGENTQ%", mobjAmadeusUser.PCC & "/" & AgentQueue)
                ConvertAmadeusValue = ReplaceReference(ConvertAmadeusValue, "%AGENTQ%", AgentQueue)
            End If
            If AgentOPQueue.IndexOf("/") >= 0 Then
                ConvertAmadeusValue = ReplaceReference(ConvertAmadeusValue, "%AGENTOPQ%", "/" & AgentOPQueue)
            Else
                ConvertAmadeusValue = ReplaceReference(ConvertAmadeusValue, "%AGENTOPQ%", AgentOPQueue)
            End If
            Do While ConvertAmadeusValue.IndexOf("//") >= 0
                ConvertAmadeusValue.Replace("//", "/")
            Loop
            ConvertAmadeusValue = ReplaceReference(ConvertAmadeusValue, "%PCC%", mobjAmadeusUser.PCC)
            ConvertAmadeusValue = ReplaceReference(ConvertAmadeusValue, "%AgentID%", mobjAmadeusUser.User)

            ConvertAmadeusValue = ReplaceReference(ConvertAmadeusValue, "%MID%", CountryCode)
            ConvertAmadeusValue = ReplaceReference(ConvertAmadeusValue, "%AGENTNAME%", AgentName)
            ConvertAmadeusValue = ReplaceReference(ConvertAmadeusValue, "%AGENTEMAIL%", AgentEmail)
            ConvertAmadeusValue = ReplaceReference(ConvertAmadeusValue, "%CITYCODE%", OfficeCityCode)
            ConvertAmadeusValue = ReplaceReference(ConvertAmadeusValue, "%CITYNAME%", CityName)
            ConvertAmadeusValue = ReplaceReference(ConvertAmadeusValue, "%AOHP%", AOHPhone)
            ConvertAmadeusValue = ReplaceReference(ConvertAmadeusValue, "%PHONE%", Phone)
            ConvertAmadeusValue = ReplaceReference(ConvertAmadeusValue, "%OFFICENAME%", OfficeName)
        End If

    End Function
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
