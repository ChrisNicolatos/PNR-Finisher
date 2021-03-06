﻿Option Strict On
Option Explicit On
Namespace Customers
    'Public Class CustomerItem
    '    Private Structure ClassProps
    '        Dim ID As Integer
    '        Dim Code As String
    '        Dim Name As String
    '        Dim Logo As String
    '        Dim EntityKindLT As Integer
    '        Dim HasVessels As Boolean
    '        Dim HasDepartments As Boolean
    '        Dim Alert As String
    '        Dim GalileoTrackingCode As String
    '    End Structure
    '    Private mudtProps As ClassProps
    '    Private mobjCustomProperties As New CustomProperties.Collection
    '    Private mflgCustomProperties As Boolean
    '    Private mobjAlerts As New AlertsCollection

    '    Public Overrides Function ToString() As String

    '        Return Code & " " & Logo ' Name

    '    End Function

    '    Public ReadOnly Property ID() As Integer
    '        Get
    '            ID = mudtProps.ID
    '        End Get
    '    End Property

    '    Public ReadOnly Property Code() As String
    '        Get
    '            Code = mudtProps.Code
    '        End Get
    '    End Property

    '    Public ReadOnly Property Name() As String
    '        Get
    '            Name = mudtProps.Name.ToUpper
    '        End Get
    '    End Property
    '    Public ReadOnly Property Logo As String
    '        Get
    '            Return mudtProps.Logo.ToUpper
    '        End Get
    '    End Property
    '    Public ReadOnly Property EntityKindLT() As Integer
    '        Get
    '            EntityKindLT = mudtProps.EntityKindLT
    '        End Get
    '    End Property

    '    Public ReadOnly Property HasVessels() As Boolean
    '        Get
    '            HasVessels = mudtProps.HasVessels
    '        End Get
    '    End Property

    '    Public ReadOnly Property HasDepartments() As Boolean
    '        Get
    '            HasDepartments = mudtProps.HasDepartments
    '        End Get
    '    End Property
    '    Public ReadOnly Property Alert As String
    '        Get
    '            Alert = mudtProps.Alert
    '        End Get
    '    End Property
    '    Public ReadOnly Property GalileoTrackingCode As String
    '        Get
    '            GalileoTrackingCode = mudtProps.GalileoTrackingCode
    '        End Get
    '    End Property
    '    Public ReadOnly Property CustomerProperties As CustomProperties.Collection
    '        Get
    '            If Not mflgCustomProperties Then
    '                mobjCustomProperties.Load(mudtProps.ID)
    '                mflgCustomProperties = True
    '            End If
    '            CustomerProperties = mobjCustomProperties
    '        End Get
    '    End Property

    '    Friend Sub SetValues(ByVal pID As Integer, ByVal pCode As String, ByVal pName As String, ByVal pLogo As String, ByVal pEntityKindLT As Integer, ByVal pAlert As String, ByVal pGalileoTrackingCode As String)
    '        With mudtProps
    '            .ID = pID
    '            .Code = pCode
    '            .Name = pName
    '            .Logo = pLogo
    '            .EntityKindLT = pEntityKindLT
    '            .Alert = pAlert.Trim
    '            .GalileoTrackingCode = pGalileoTrackingCode
    '            ' TFEntityKind (from DB table [TravelForceCosmos].[dbo].[LookupTable])
    '            ' 404 = Other
    '            ' 405 = Individual
    '            ' 406 = Corporate
    '            ' 526 = Shipping Co
    '            ' 527 = Travel Agency
    '            Select Case pEntityKindLT
    '                Case 526, 527
    '                    .HasDepartments = True
    '                    .HasVessels = True
    '                Case Else
    '                    .HasDepartments = False
    '                    .HasVessels = False
    '            End Select
    '            mflgCustomProperties = False
    '        End With
    '    End Sub
    '    Public Sub Load(ByVal pCode As String)

    '        mobjAlerts.Load()

    '        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringACC) ' ActiveConnection)
    '        Dim pobjComm As New SqlClient.SqlCommand
    '        Dim pobjReader As SqlClient.SqlDataReader

    '        pobjConn.Open()
    '        pobjComm = pobjConn.CreateCommand

    '        With pobjComm
    '            .CommandType = CommandType.Text
    '            .CommandText = PrepareClientSelectCommand(pCode)
    '            pobjReader = .ExecuteReader
    '        End With
    '        With pobjReader
    '            If pobjReader.Read Then
    '                SetValues(CInt(.Item("Id")), CStr(.Item("Code")), CStr(.Item("Name")), CStr(.Item("Logo")), CInt(.Item("TFEntityKindLT")), mobjAlerts.Alert(MySettings.PCCBackOffice, CStr(.Item("Code"))), CStr(.Item("GalileoTrackingCode")))
    '                .Close()
    '            End If
    '        End With
    '        pobjConn.Close()

    '    End Sub

    '    Private Function PrepareClientSelectCommand(ByVal pCode As String) As String

    '        Select Case MySettings.PCCBackOffice
    '            Case 1 ' Travel Force
    '                PrepareClientSelectCommand = " SELECT TFEntities.Id " &
    '                           " ,TFEntities.Code" &
    '                           " ,TFEntities.Name " &
    '                           " ,TFEntities.Logo" &
    '                           " ,TFEntityCategories.TFEntityKindLT " &
    '                           " ,ISNULL(DealCodes.Code, '') AS GalileoTrackingCode " &
    '                           " FROM [TravelForceCosmos].[dbo].[TFEntities] " &
    '                           " LEFT JOIN [TravelForceCosmos].[dbo].[TFEntityCategories] " &
    '                           " ON TFEntities.CategoryID = TFEntityCategories.Id " &
    '                           " LEFT JOIN TravelForceCosmos.dbo.DealCodes " &
    '                           " ON DealCodes.ClientID=TFEntities.Id And DealCodes.AirlineID=3352 " &
    '                           " WHERE TFEntities.IsClient = 1  " &
    '                           " AND TFEntities.CanHaveCT = 1 " &
    '                           " AND TFEntities.IsActive = 1 " &
    '                           " AND TFEntities.Code = '" & pCode & "' "
    '            Case 2 ' Discovery
    '                PrepareClientSelectCommand = " Select [Account_Id] As Id " &
    '                                            " ,[Account_Abbriviation] AS Code " &
    '                                            " ,[Account_Name] AS Name " &
    '                                            " ,[Account_Name] AS Logo " &
    '                                            " ,526 AS TFEntityKindLT " &
    '                                            " ,'' AS GalileoTrackingCode " &
    '                                            " From [Disco_Instone_EU].[dbo].[Company] " &
    '                                            " Where Account_Abbriviation = '" & pCode & "' "
    '            Case Else
    '                PrepareClientSelectCommand = ""
    '        End Select
    '    End Function
    'End Class

    'Public Class CustomerCollection
    '    Inherits Collections.Generic.Dictionary(Of Integer, CustomerItem)
    '    Private mAllCustomer As New AllCustomer
    '    Public Sub Load(ByVal SearchString As String)

    '        Try
    '            If mAllCustomer.Count = 0 Or mAllCustomer.PCCBackOffice <> MySettings.PCCBackOffice Then
    '                Cursor.Current = Cursors.WaitCursor
    '                mAllCustomer.Load()
    '            End If

    '            MyBase.Clear()

    '            Dim pItem As CustomerItem

    '            For Each pItem In mAllCustomer.Values
    '                If pItem.Code.ToUpper.IndexOf(SearchString.ToUpper) >= 0 Or pItem.Name.ToUpper.IndexOf(SearchString.ToUpper) >= 0 Or pItem.Logo.ToUpper.IndexOf(SearchString.ToUpper) >= 0 Then
    '                    MyBase.Add(pItem.ID, pItem)
    '                End If
    '            Next

    '        Catch ex As Exception
    '            Throw New Exception("Customers.Load()" & vbCrLf & ex.Message)
    '        Finally
    '            Cursor.Current = Cursors.Default
    '        End Try
    '    End Sub


    'End Class
    'Public Class CustomerCollectionAll

    '    Inherits Collections.Generic.Dictionary(Of Integer, CustomerItem)

    '    Dim mobjAlerts As New AlertsCollection
    '    Dim mintPCCBackoffice As Integer
    '    Public ReadOnly Property PCCBackOffice As Integer
    '        Get
    '            PCCBackOffice = mintPCCBackoffice
    '        End Get
    '    End Property

    '    Public Sub Load()

    '        Dim pCommandText As String

    '        Try
    '            mobjAlerts.Load()

    '            pCommandText = PrepareClientSelectCommand()
    '            ReadCustomers(pCommandText)
    '        Catch ex As Exception
    '            Throw New Exception("Customers.Load()" & vbCrLf & ex.Message)
    '        End Try

    '    End Sub
    '    Private Function PrepareClientSelectCommand() As String

    '        mintPCCBackoffice = MySettings.PCCBackOffice

    '        Select Case MySettings.PCCBackOffice
    '            Case 1 ' Travel Force
    '                PrepareClientSelectCommand = " SELECT TFEntities.Id " &
    '                           " ,TFEntities.Code" &
    '                           " ,TFEntities.Name " &
    '                           " ,TFEntities.Logo " &
    '                           " ,TFEntityCategories.TFEntityKindLT " &
    '                           " ,ISNULL(DealCodes.Code, '') AS GalileoTrackingCode " &
    '                           " FROM [TravelForceCosmos].[dbo].[TFEntities] " &
    '                           " LEFT JOIN [TravelForceCosmos].[dbo].[TFEntityCategories] " &
    '                           " ON TFEntities.CategoryID = TFEntityCategories.Id " &
    '                           " LEFT JOIN TravelForceCosmos.dbo.DealCodes " &
    '                           " ON DealCodes.ClientID=TFEntities.Id and DealCodes.AirlineID=3352 " &
    '                           " WHERE TFEntities.IsClient = 1  " &
    '                           " AND TFEntities.CanHaveCT = 1 " &
    '                           " AND TFEntities.IsActive = 1 " &
    '                           " ORDER BY TFEntities.Code "
    '            Case 2 ' Discovery
    '                PrepareClientSelectCommand = "SELECT Company.[Account_Id] AS Id " &
    '                                             " ,[Account_Abbriviation] AS Code " &
    '                                             " ,[Account_Name] AS Name " &
    '                                             " ,[Account_Name] AS Logo " &
    '                                             " ,526 AS TFEntityKindLT" &
    '                                             " ,'' AS GalileoTrackingCode " &
    '                                             " From [Disco_Instone_EU].[dbo].[Company] " &
    '                                             " Left Join Disco_Instone_EU.dbo.CompProfile " &
    '                                             " On Company.Account_Id = CompProfile.Account_Id " &
    '                                             " Where CompProfile.Branch = 19 " &
    '                                             " ORDER BY Account_Abbriviation "
    '            Case Else
    '                PrepareClientSelectCommand = ""
    '        End Select
    '    End Function
    '    Private Sub ReadCustomers(ByVal CommandText As String)

    '        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringACC) ' ActiveConnection)
    '        Dim pobjComm As New SqlClient.SqlCommand
    '        Dim pobjReader As SqlClient.SqlDataReader
    '        Dim pobjClass As CustomerItem

    '        pobjConn.Open()
    '        pobjComm = pobjConn.CreateCommand

    '        With pobjComm
    '            .CommandType = CommandType.Text
    '            .CommandText = CommandText
    '            pobjReader = .ExecuteReader
    '        End With

    '        MyBase.Clear()

    '        With pobjReader
    '            Do While .Read
    '                pobjClass = New CustomerItem
    '                pobjClass.SetValues(CInt(.Item("Id")), CStr(.Item("Code")), CStr(.Item("Name")), CStr(.Item("Logo")), CInt(.Item("TFEntityKindLT")), mobjAlerts.Alert(MySettings.PCCBackOffice, CStr(.Item("Code"))), CStr(.Item("GalileoTrackingCode")))
    '                MyBase.Add(pobjClass.ID, pobjClass)
    '            Loop
    '            .Close()
    '        End With
    '        pobjConn.Close()

    '    End Sub

    'End Class

    'Public Class CustomerGroupItem
    '    Private Structure ClassProps
    '        Dim ID As Integer
    '        Dim Name As String
    '    End Structure
    '    Private mudtProps As ClassProps
    '    Public Overrides Function ToString() As String
    '        With mudtProps
    '            Return .Name
    '        End With
    '    End Function
    '    Public ReadOnly Property ID() As Integer
    '        Get
    '            ID = mudtProps.ID
    '        End Get
    '    End Property
    '    Public ReadOnly Property Name() As String
    '        Get
    '            Name = mudtProps.Name
    '        End Get
    '    End Property
    '    Friend Sub SetValues(ByVal pID As Integer, ByVal pName As String)
    '        With mudtProps
    '            .ID = pID
    '            .Name = pName
    '        End With
    '    End Sub
    'End Class
    'Public Class CustomerGroupCollectionAll
    '    Inherits Collections.Generic.Dictionary(Of Integer, CustomerGroupItem)
    '    Public Sub Load()

    '        Dim pCommandText As String

    '        Try
    '            pCommandText = " USE TravelForceCosmos " &
    '                           " SELECT Id " &
    '                           " ,Description " &
    '                           " FROM Tags " &
    '                           " WHERE TagGroupId = 146 " &
    '                           " ORDER BY Description "
    '            ReadCustomerGroups(pCommandText)
    '        Catch ex As Exception
    '            Throw New Exception("AllCustomerGroups.Load()" & vbCrLf & ex.Message)
    '        End Try

    '    End Sub
    '    Private Sub ReadCustomerGroups(ByVal CommandText As String)

    '        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringACC) ' ActiveConnection)
    '        Dim pobjComm As New SqlClient.SqlCommand
    '        Dim pobjReader As SqlClient.SqlDataReader
    '        Dim pobjClass As CustomerGroupItem

    '        pobjConn.Open()
    '        pobjComm = pobjConn.CreateCommand

    '        With pobjComm
    '            .CommandType = CommandType.Text
    '            .CommandText = CommandText
    '            pobjReader = .ExecuteReader
    '        End With

    '        With pobjReader
    '            Do While .Read
    '                pobjClass = New CustomerGroupItem
    '                pobjClass.SetValues(CInt(.Item("Id")), CStr(.Item("Description")))
    '                MyBase.Add(pobjClass.ID, pobjClass)
    '            Loop
    '            .Close()
    '        End With
    '        pobjConn.Close()

    '    End Sub
    'End Class
    'Public Class CustomerGroupCollection
    '    Inherits Collections.Generic.Dictionary(Of Integer, CustomerGroupItem)
    '    Private mAllCustomer As New CustomerGroupCollectionAll

    '    Public Sub Load(ByVal SearchString As String)

    '        Try
    '            If mAllCustomer.Count = 0 Then
    '                Cursor.Current = Cursors.WaitCursor
    '                mAllCustomer.Load()
    '            End If

    '            MyBase.Clear()

    '            Dim pItem As CustomerGroupItem

    '            For Each pItem In mAllCustomer.Values
    '                If pItem.Name.ToUpper.IndexOf(SearchString.ToUpper) >= 0 Then
    '                    MyBase.Add(pItem.ID, pItem)
    '                End If
    '            Next


    '        Catch ex As Exception
    '            Throw New Exception("CustomerGroupCollection.Load()" & vbCrLf & ex.Message)
    '        Finally
    '            Cursor.Current = Cursors.Default
    '        End Try
    '    End Sub

    'End Class

End Namespace
