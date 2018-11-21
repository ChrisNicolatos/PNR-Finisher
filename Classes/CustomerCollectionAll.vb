Public Class CustomerCollectionAll

    Inherits Collections.Generic.Dictionary(Of Integer, CustomerItem)

    Dim mobjAlerts As AlertsCollection
    Dim mintPCCBackoffice As Integer
    Public ReadOnly Property PCCBackOffice As Integer
        Get
            PCCBackOffice = mintPCCBackoffice
        End Get
    End Property
    Public ReadOnly Property GetCustomerByCode(ByVal pCode As String) As CustomerItem
        Get
            GetCustomerByCode = New CustomerItem
            For Each pItem As CustomerItem In MyBase.Values
                If pItem.Code = pCode Then
                    GetCustomerByCode = pItem
                    Exit For
                End If
            Next
        End Get
    End Property
    Public Sub Load()

        Dim pCommandText As String

        Try
            mobjAlerts = New AlertsCollection()
            mobjAlerts.Load()

            pCommandText = PrepareClientSelectCommand()
            ReadCustomers(pCommandText)
        Catch ex As Exception
            Throw New Exception("CustomerCollectionAll.Load()" & vbCrLf & ex.Message)
        End Try

    End Sub
    Private Function PrepareClientSelectCommand() As String

        mintPCCBackoffice = MySettings.PCCBackOffice

        Select Case MySettings.PCCBackOffice
            Case 1 ' Travel Force
                PrepareClientSelectCommand = " SELECT TFEntities.Id " &
                           " ,TFEntities.Code" &
                           " ,TFEntities.Name " &
                           " ,TFEntities.Logo " &
                           " ,TFEntityCategories.TFEntityKindLT " &
                           " ,ISNULL(DealCodes.Code, '') AS GalileoTrackingCode " &
                           " FROM [TravelForceCosmos].[dbo].[TFEntities] " &
                           " LEFT JOIN [TravelForceCosmos].[dbo].[TFEntityCategories] " &
                           " ON TFEntities.CategoryID = TFEntityCategories.Id " &
                           " LEFT JOIN TravelForceCosmos.dbo.DealCodes " &
                           " ON DealCodes.ClientID=TFEntities.Id and DealCodes.AirlineID=3352 " &
                           " WHERE TFEntities.IsClient = 1  " &
                           " AND TFEntities.CanHaveCT = 1 " &
                           " AND TFEntities.IsActive = 1 " &
                           " ORDER BY TFEntities.Code "
            Case 2 ' Discovery
                PrepareClientSelectCommand = "SELECT Company.[Account_Id] AS Id " &
                                             " ,[Account_Abbriviation] AS Code " &
                                             " ,[Account_Name] AS Name " &
                                             " ,[Account_Name] AS Logo " &
                                             " ,526 AS TFEntityKindLT" &
                                             " ,'' AS GalileoTrackingCode " &
                                             " From [Disco_Instone_EU].[dbo].[Company] " &
                                             " Left Join Disco_Instone_EU.dbo.CompProfile " &
                                             " On Company.Account_Id = CompProfile.Account_Id " &
                                             " Where CompProfile.Branch = '" & MySettings.PCCBranchCode & "' " &
                                             " ORDER BY Account_Abbriviation "
            Case Else
                PrepareClientSelectCommand = ""
        End Select
    End Function
    Private Sub ReadCustomers(ByVal CommandText As String)

        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringACC) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader
        Dim pobjClass As CustomerItem

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = CommandText
            pobjReader = .ExecuteReader
        End With

        MyBase.Clear()

        With pobjReader
            Do While .Read
                pobjClass = New CustomerItem
                pobjClass.SetValues(CInt(.Item("Id")), CStr(.Item("Code")), CStr(.Item("Name")), CStr(.Item("Logo")), CInt(.Item("TFEntityKindLT")), mobjAlerts.AlertForFinisher(MySettings.PCCBackOffice, CStr(.Item("Code"))), mobjAlerts.AlertForDownsell(MySettings.PCCBackOffice, CStr(.Item("Code"))), CStr(.Item("GalileoTrackingCode")))
                MyBase.Add(pobjClass.ID, pobjClass)
            Loop
            .Close()
        End With
        pobjConn.Close()

    End Sub

End Class