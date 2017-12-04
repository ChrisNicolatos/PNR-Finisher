Namespace Customers

    Public Class CustomerItem
        Private Structure ClassProps
            Dim ID As Long
            Dim Code As String
            Dim Name As String
            Dim EntityKindLT As Long
            Dim HasVessels As Boolean
            Dim HasDepartments As Boolean
            Dim AllowNullMLEntity As Boolean
        End Structure
        Private mudtProps As ClassProps
        Private mobjCustomProperties As New CustomProperties.Collection
        Private mflgCustomProperties As Boolean

        Public Overrides Function ToString() As String
            With mudtProps
                Return .Code & " " & .Name
            End With
        End Function

        Public ReadOnly Property ID() As Long
            Get
                ID = mudtProps.ID
            End Get
        End Property

        Public ReadOnly Property Code() As String
            Get
                Code = mudtProps.Code
            End Get
        End Property

        Public ReadOnly Property Name() As String
            Get
                Name = mudtProps.Name
            End Get
        End Property

        Public ReadOnly Property EntityKindLT() As Long
            Get
                EntityKindLT = mudtProps.EntityKindLT
            End Get
        End Property

        Public ReadOnly Property HasVessels() As Boolean
            Get
                HasVessels = mudtProps.HasVessels
            End Get
        End Property

        Public ReadOnly Property HasDepartments() As Boolean
            Get
                HasDepartments = mudtProps.HasDepartments
            End Get
        End Property

        Public ReadOnly Property AllowNullMLEntity As Boolean
            Get
                AllowNullMLEntity = mudtProps.AllowNullMLEntity
            End Get
        End Property

        Public ReadOnly Property CustomerProperties As CustomProperties.Collection
            Get
                If Not mflgCustomProperties Then
                    mobjCustomProperties.Load(mudtProps.ID)
                    mflgCustomProperties = True
                End If
                CustomerProperties = mobjCustomProperties
            End Get
        End Property

        Friend Sub SetValues(ByVal pID As Long, ByVal pCode As String, ByVal pName As String, ByVal pEntityKindLT As Long, ByVal pAllowNullMLEntity As Boolean)
            With mudtProps
                .ID = pID
                .Code = pCode
                .Name = pName
                .EntityKindLT = pEntityKindLT
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
                .AllowNullMLEntity = pAllowNullMLEntity
                mflgCustomProperties = False
            End With
        End Sub

        ''' <summary>
        ''' Will load a specific customer's details from TfEntites
        ''' </summary>
        ''' <param name="pCode">Used to send the required customer code to the select statement.</param>
        ''' <remarks></remarks>
        Public Sub Load(ByVal pCode As String)
            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringACC) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .CommandText = " SELECT TFEntities.Id " & _
                               " ,TFEntities.Code" & _
                               " ,TFEntities.Name " & _
                               " ,TFEntityCategories.TFEntityKindLT " & _
                               " ,TFEntities.AllowNullMLEntity " & _
                               " FROM [TravelForceCosmos].[dbo].[TFEntities] " & _
                               " LEFT JOIN [TravelForceCosmos].[dbo].[TFEntityCategories] " & _
                               " ON TFEntities.CategoryID = TFEntityCategories.Id " & _
                               " WHERE TFEntities.IsClient = 1  " & _
                               " AND TFEntities.CanHaveCT = 1 " & _
                               " AND TFEntities.IsActive = 1 " & _
                               " AND TFEntities.Code = '" & pCode & "' " & _
                               " ORDER BY TFEntities.Code "

                pobjReader = .ExecuteReader
            End With
            With pobjReader
                If pobjReader.Read Then
                    SetValues(.Item("Id"), .Item("Code"), .Item("Name"), .Item("TFEntityKindLT"), .Item("AllowNullMLEntity"))
                    .Close()
                End If
            End With
            pobjConn.Close()

        End Sub
    End Class

    Public Class CustomerCollection
        Inherits Collections.Generic.Dictionary(Of String, CustomerItem)
        Private mAllCustomer As New AllCustomer


        Public Sub Load(ByVal SearchString As String)

            Try
                If mAllCustomer.Count = 0 Then
                    Cursor.Current = Cursors.WaitCursor
                    mAllCustomer.Load()
                End If

                MyBase.Clear()

                Dim pItem As CustomerItem

                For Each pItem In mAllCustomer.Values
                    If pItem.Code.ToUpper.IndexOf(SearchString.ToUpper) >= 0 Or pItem.Name.ToUpper.IndexOf(SearchString.ToUpper) >= 0 Then
                        MyBase.Add(pItem.ID, pItem)
                    End If
                Next

            Catch ex As Exception
                Throw New Exception("Customers.Load()" & vbCrLf & ex.Message)
            Finally
                Cursor.Current = Cursors.Default
            End Try
        End Sub

        'Private Sub ReadCustomers(ByVal CommandText As String)

        '    Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringACC) ' ActiveConnection)
        '    Dim pobjComm As New SqlClient.SqlCommand
        '    Dim pobjReader As SqlClient.SqlDataReader
        '    Dim pobjClass As CustomerItem

        '    pobjConn.Open()
        '    pobjComm = pobjConn.CreateCommand

        '    With pobjComm
        '        .CommandType = CommandType.Text
        '        .CommandText = CommandText
        '        pobjReader = .ExecuteReader
        '    End With

        '    With pobjReader
        '        Do While .Read
        '            pobjClass = New CustomerItem
        '            pobjClass.SetValues(.Item("Id"), .Item("Code"), .Item("Name"), .Item("TFEntityKindLT"), .Item("AllowNullMLEntity"))
        '            MyBase.Add(pobjClass.ID, pobjClass)
        '        Loop
        '        .Close()
        '    End With
        '    pobjConn.Close()

        'End Sub
    End Class

    Public Class AllCustomer
        Inherits Collections.Generic.Dictionary(Of String, CustomerItem)

        Public Sub Load()

            Dim pCommandText As String

            Try
                pCommandText = " SELECT TFEntities.Id " & _
                               " ,TFEntities.Code" & _
                               " ,TFEntities.Name " & _
                               " ,TFEntityCategories.TFEntityKindLT " & _
                               " ,TFEntities.AllowNullMLEntity " & _
                               " FROM [TravelForceCosmos].[dbo].[TFEntities] " & _
                               " LEFT JOIN [TravelForceCosmos].[dbo].[TFEntityCategories] " & _
                               " ON TFEntities.CategoryID = TFEntityCategories.Id " & _
                               " WHERE TFEntities.IsClient = 1  " & _
                               " AND TFEntities.CanHaveCT = 1 " & _
                               " AND TFEntities.IsActive = 1 " & _
                               " ORDER BY TFEntities.Code "
                ReadCustomers(pCommandText)
            Catch ex As Exception
                Throw New Exception("Customers.Load()" & vbCrLf & ex.Message)
            End Try

        End Sub
        Private Sub ReadCustomers(ByVal CommandText As String)

            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringACC) ' ActiveConnection)
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

            With pobjReader
                Do While .Read
                    pobjClass = New CustomerItem
                    pobjClass.SetValues(.Item("Id"), .Item("Code"), .Item("Name"), .Item("TFEntityKindLT"), .Item("AllowNullMLEntity"))
                    MyBase.Add(pobjClass.ID, pobjClass)
                Loop
                .Close()
            End With
            pobjConn.Close()

        End Sub
    End Class

    Public Class CustomerGroupItem
        Private Structure ClassProps
            Dim ID As Long
            Dim Name As String
        End Structure
        Private mudtProps As ClassProps
        Public Overrides Function ToString() As String
            With mudtProps
                Return .Name
            End With
        End Function
        Public ReadOnly Property ID() As Long
            Get
                ID = mudtProps.ID
            End Get
        End Property
        Public ReadOnly Property Name() As String
            Get
                Name = mudtProps.Name
            End Get
        End Property
        Friend Sub SetValues(ByVal pID As Long, ByVal pName As String)
            With mudtProps
                .ID = pID
                .Name = pName
            End With
        End Sub
    End Class
    Public Class AllCustomerGroups
        Inherits Collections.Generic.Dictionary(Of String, CustomerGroupItem)
        Public Sub Load()

            Dim pCommandText As String

            Try
                pCommandText = " USE TravelForceCosmos " & _
                               " SELECT Id " & _
                               " ,Description " & _
                               " FROM Tags " & _
                               " WHERE TagGroupId = 146 " & _
                               " ORDER BY Description "
                ReadCustomerGroups(pCommandText)
            Catch ex As Exception
                Throw New Exception("AllCustomerGroups.Load()" & vbCrLf & ex.Message)
            End Try

        End Sub
        Private Sub ReadCustomerGroups(ByVal CommandText As String)

            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringACC) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader
            Dim pobjClass As CustomerGroupItem

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .CommandText = CommandText
                pobjReader = .ExecuteReader
            End With

            With pobjReader
                Do While .Read
                    pobjClass = New CustomerGroupItem
                    pobjClass.SetValues(.Item("Id"), .Item("Description"))
                    MyBase.Add(pobjClass.ID, pobjClass)
                Loop
                .Close()
            End With
            pobjConn.Close()

        End Sub
    End Class
    Public Class CustomerGroupCollection
        Inherits Collections.Generic.Dictionary(Of String, CustomerGroupItem)
        Private mAllCustomer As New AllCustomerGroups
        
        Public Sub Load(ByVal SearchString As String)

            Try
                If mAllCustomer.Count = 0 Then
                    Cursor.Current = Cursors.WaitCursor
                    mAllCustomer.Load()
                End If

                MyBase.Clear()

                Dim pItem As CustomerGroupItem

                For Each pItem In mAllCustomer.Values
                    If pItem.Name.ToUpper.IndexOf(SearchString.ToUpper) >= 0 Then
                        MyBase.Add(pItem.ID, pItem)
                    End If
                Next


            Catch ex As Exception
                Throw New Exception("CustomerGroupCollection.Load()" & vbCrLf & ex.Message)
            Finally
                Cursor.Current = Cursors.Default
            End Try
        End Sub

        'Private Sub ReadCustomers(ByVal CommandText As String)

        '    Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringACC) ' ActiveConnection)
        '    Dim pobjComm As New SqlClient.SqlCommand
        '    Dim pobjReader As SqlClient.SqlDataReader
        '    Dim pobjClass As CustomerGroupItem



        '    pobjConn.Open()
        '    pobjComm = pobjConn.CreateCommand

        '    With pobjComm
        '        .CommandType = CommandType.Text
        '        .CommandText = CommandText
        '        pobjReader = .ExecuteReader
        '    End With

        '    With pobjReader
        '        Do While .Read
        '            pobjClass = New CustomerGroupItem
        '            pobjClass.SetValues(.Item("Id"), .Item("Description"))
        '            MyBase.Add(pobjClass.ID, pobjClass)
        '        Loop
        '        .Close()
        '    End With
        '    pobjConn.Close()

        'End Sub
    End Class

End Namespace
