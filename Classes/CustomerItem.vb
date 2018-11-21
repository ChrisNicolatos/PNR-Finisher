Public Class CustomerItem
    Private Structure ClassProps
        Dim ID As Integer
        Dim Code As String
        Dim Name As String
        Dim Logo As String
        Dim EntityKindLT As Integer
        Dim HasVessels As Boolean
        Dim HasDepartments As Boolean
        Dim AlertForFinisher As String
        Dim AlertForDownsell As String
        Dim GalileoTrackingCode As String
    End Structure
    Private mudtProps As ClassProps
    Private mobjCustomProperties As New CustomPropertiesCollection
    Private mflgCustomProperties As Boolean
    Private mobjAlerts As New AlertsCollection

    Public Overrides Function ToString() As String

        Return Code & " " & Logo ' Name

    End Function

    Public ReadOnly Property ID() As Integer
        Get
            Return mudtProps.ID
        End Get
    End Property

    Public ReadOnly Property Code() As String
        Get
            Return mudtProps.Code
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return mudtProps.Name.ToUpper
        End Get
    End Property
    Public ReadOnly Property Logo As String
        Get
            If mudtProps.Logo Is Nothing Then
                Return ""
            Else
                Return mudtProps.Logo.ToUpper
            End If
        End Get
    End Property
    Public ReadOnly Property EntityKindLT() As Integer
        Get
            Return mudtProps.EntityKindLT
        End Get
    End Property

    Public ReadOnly Property HasVessels() As Boolean
        Get
            Return mudtProps.HasVessels
        End Get
    End Property

    Public ReadOnly Property HasDepartments() As Boolean
        Get
            Return mudtProps.HasDepartments
        End Get
    End Property
    Public ReadOnly Property AlertForFinisher As String
        Get
            Return mudtProps.AlertForFinisher
        End Get
    End Property
    Public ReadOnly Property AlertForDownsell As String
        Get
            If mudtProps.AlertForDownsell Is Nothing Then
                Return ""
            Else
                Return mudtProps.AlertForDownsell
            End If

        End Get
    End Property
    Public ReadOnly Property GalileoTrackingCode As String
        Get
            Return mudtProps.GalileoTrackingCode
        End Get
    End Property
    Public ReadOnly Property CustomerProperties As CustomPropertiesCollection
        Get
            If Not mflgCustomProperties Then
                mobjCustomProperties.Load(mudtProps.ID)
                mflgCustomProperties = True
            End If
            Return mobjCustomProperties
        End Get
    End Property

    Friend Sub SetValues(ByVal pID As Integer, ByVal pCode As String, ByVal pName As String, ByVal pLogo As String, ByVal pEntityKindLT As Integer, ByVal pAlertForFinisher As String, ByVal pAlertForDownsell As String, ByVal pGalileoTrackingCode As String)
        With mudtProps
            .ID = pID
            .Code = pCode
            .Name = pName
            .Logo = pLogo
            .EntityKindLT = pEntityKindLT
            .AlertForFinisher = pAlertForFinisher.Trim
            .AlertForDownsell = pAlertForDownsell.Trim
            .GalileoTrackingCode = pGalileoTrackingCode
            ' TFEntityKind (from DB table [TravelForceCosmos].[dbo].[LookupTable])
            ' 404 = Other
            ' 405 = Individual
            ' 406 = Corporate
            ' 526 = Shipping Co
            ' 527 = Travel Agency
            Select Case pEntityKindLT
                Case 526, 527
                    .HasDepartments = True
                    .HasVessels = True
                Case Else
                    .HasDepartments = False
                    .HasVessels = False
            End Select
            mflgCustomProperties = False
        End With
    End Sub
    Public Sub Load(ByVal pCode As String)

        mobjAlerts.Load()

        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringACC) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = PrepareClientSelectCommand(pCode)
            pobjReader = .ExecuteReader
        End With
        With pobjReader
            If pobjReader.Read Then
                SetValues(CInt(.Item("Id")), CStr(.Item("Code")), CStr(.Item("Name")), CStr(.Item("Logo")), CInt(.Item("TFEntityKindLT")), mobjAlerts.AlertForFinisher(MySettings.PCCBackOffice, CStr(.Item("Code"))), mobjAlerts.AlertForDownsell(MySettings.PCCBackOffice, CStr(.Item("Code"))), CStr(.Item("GalileoTrackingCode")))
                .Close()
            End If
        End With
        pobjConn.Close()

    End Sub

    Private Function PrepareClientSelectCommand(ByVal pCode As String) As String

        Select Case MySettings.PCCBackOffice
            Case 1 ' Travel Force
                PrepareClientSelectCommand = " SELECT TFEntities.Id " &
                           " ,TFEntities.Code" &
                           " ,TFEntities.Name " &
                           " ,TFEntities.Logo" &
                           " ,TFEntityCategories.TFEntityKindLT " &
                           " ,ISNULL(DealCodes.Code, '') AS GalileoTrackingCode " &
                           " FROM [TravelForceCosmos].[dbo].[TFEntities] " &
                           " LEFT JOIN [TravelForceCosmos].[dbo].[TFEntityCategories] " &
                           " ON TFEntities.CategoryID = TFEntityCategories.Id " &
                           " LEFT JOIN TravelForceCosmos.dbo.DealCodes " &
                           " ON DealCodes.ClientID=TFEntities.Id And DealCodes.AirlineID=3352 " &
                           " WHERE TFEntities.IsClient = 1  " &
                           " AND TFEntities.CanHaveCT = 1 " &
                           " AND TFEntities.IsActive = 1 " &
                           " AND TFEntities.Code = '" & pCode & "' "
            Case 2 ' Discovery
                PrepareClientSelectCommand = " Select [Account_Id] As Id " &
                                            " ,[Account_Abbriviation] AS Code " &
                                            " ,[Account_Name] AS Name " &
                                            " ,[Account_Name] AS Logo " &
                                            " ,526 AS TFEntityKindLT " &
                                            " ,'' AS GalileoTrackingCode " &
                                            " From [Disco_Instone_EU].[dbo].[Company] " &
                                            " Where Account_Abbriviation = '" & pCode & "' "
            Case Else
                PrepareClientSelectCommand = ""
        End Select
    End Function
End Class
