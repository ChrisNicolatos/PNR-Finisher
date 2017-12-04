Namespace osmVessels
    Public Class VesselItem
        Private Structure ClassProps
            Dim Id As Integer
            Dim VesselName As String
            Dim VesselGroup As String
            Dim InUse As Boolean
            Dim isNew As Boolean
            Dim isValid As Boolean
        End Structure
        Dim mudtProps As ClassProps

        Public Sub New()
            With mudtProps
                .Id = 0
                .VesselName = ""
                .VesselGroup = ""
                .InUse = True
                .isNew = True
                CheckValid()
            End With
        End Sub
        Public Overrides Function ToString() As String

            ToString = mudtProps.VesselName

        End Function
        Public ReadOnly Property Id As Integer
            Get
                Id = mudtProps.Id
            End Get
        End Property
        Public Property VesselName As String
            Get
                VesselName = mudtProps.VesselName
            End Get
            Set(value As String)
                mudtProps.VesselName = value
                CheckValid()
            End Set
        End Property
        Public Property VesselGroup As String
            Get
                VesselGroup = mudtProps.VesselGroup
            End Get
            Set(value As String)
                mudtProps.VesselGroup = value
                CheckValid()
            End Set
        End Property
        Public Property InUse As Boolean
            Get
                InUse = mudtProps.InUse
            End Get
            Set(value As Boolean)
                mudtProps.InUse = value
            End Set
        End Property
        Public ReadOnly Property isValid As Boolean
            Get
                isValid = mudtProps.isValid
            End Get
        End Property

        Public Sub SetValues(ByVal pId As Integer, ByVal pVesselName As String, ByVal pVesselGroup As String, ByVal pInUse As Boolean)
            With mudtProps
                .Id = pId
                .VesselName = pVesselName
                .VesselGroup = pVesselGroup
                .InUse = pInUse
                .isNew = False
            End With
            CheckValid()
        End Sub
        Private Sub CheckValid()

            With mudtProps
                .isValid = (.VesselName.Trim <> "")
            End With

        End Sub
        Public Sub Update()

            Try
                If mudtProps.isValid Then

                    Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR) ' ActiveConnection)
                    Dim pobjComm As New SqlClient.SqlCommand

                    pobjConn.Open()
                    pobjComm = pobjConn.CreateCommand

                    With pobjComm
                        .CommandType = CommandType.Text
                        If mudtProps.isNew Then
                            .CommandText = "IF (SELECT COUNT(*) FROM [AmadeusReports].[dbo].[osmVessels] WHERE osmvVesselName = '" & mudtProps.VesselName & "') = 0 " & _
                                           " INSERT INTO AmadeusReports.dbo.osmVessels " & _
                                           " (osmvVesselName, osmvVesselGroup, osmvInUse) " & _
                                           " VALUES " & _
                                           " ( '" & mudtProps.VesselName & "', '" & mudtProps.VesselGroup & "', " & If(mudtProps.InUse, 1, 0) & ") " & _
                                           " select scope_identity() as Id"
                            Dim pTemp As Object = .ExecuteScalar
                            If IsDBNull(pTemp) Then
                                Throw New Exception("Vessel Already exists")
                            Else
                                mudtProps.Id = pTemp
                                mudtProps.isNew = False
                            End If
                        Else
                            .CommandText = "UPDATE AmadeusReports.dbo.osmVessels " & _
                                           " SET osmvVesselName = '" & mudtProps.VesselName & "', " & _
                                           "     osmvVesselGroup = '" & mudtProps.VesselGroup & "', " & _
                                           "     osmvInUse = " & If(mudtProps.InUse, 1, 0) & " " & _
                                           " WHERE osmvId = " & mudtProps.Id
                            .ExecuteNonQuery()
                        End If
                    End With
                Else
                    Throw New Exception("Vessel name invalid")
                End If
            Catch ex As Exception
                Throw New Exception("Update Vessel Error" & vbCrLf & ex.Message)
            End Try


        End Sub
    End Class

    Public Class VesselCollection
        Inherits Collections.Generic.Dictionary(Of String, VesselItem)

        Public Sub Load()

            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader
            Dim pobjClass As VesselItem

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .CommandText = "SELECT osmvID, osmvVesselName, ISNULL(osmvVesselGroup, '') AS osmvVesselGroup, ISNULL(osmvInUse, 0) AS osmvInUse FROM AmadeusReports.dbo.osmVessels ORDER BY osmvVesselName"
                pobjReader = .ExecuteReader
            End With

            With pobjReader
                Do While .Read
                    pobjClass = New VesselItem
                    pobjClass.SetValues(.Item("osmvId"), .Item("osmvVesselName"), .Item("osmvVesselGroup"), .Item("osmvInUse"))
                    MyBase.Add(pobjClass.Id, pobjClass)
                Loop
                .Close()
            End With
            pobjConn.Close()

        End Sub
    End Class

    Public Class emailItem

        Private Structure Classprops
            Dim Id As Integer
            Dim Vessel_FK As Integer
            Dim Name As String
            Dim Details As String
            Dim EmailType As String
            Dim Email As String
            Dim VesselName As String
            Dim isNew As Boolean
            Dim isValid As Boolean
        End Structure
        Dim mudtprops As Classprops

        Public Sub New()

            With mudtprops
                .Id = 0
                .Vessel_FK = 0
                .Name = ""
                .Details = ""
                .EmailType = ""
                .Email = ""
                .VesselName = ""
                .isNew = True
            End With

        End Sub
        Public Sub New(ByVal pType As String, Optional ByVal pVessel_FK As Integer = 0)

            With mudtprops
                .Id = 0
                .Vessel_FK = pVessel_FK
                .Name = ""
                .Details = ""
                .EmailType = pType
                .Email = ""
                .VesselName = ""
                .isNew = True
                CheckValid()
            End With

        End Sub
        Private Sub CheckValid()

            With mudtprops
                .isValid = False
                If (.Vessel_FK <> 0 Or (.EmailType = "AGENT" And .Vessel_FK = 0)) And .Name <> "" And .EmailType <> "" And .Email <> "" Then
                    .isValid = True
                End If
            End With
        End Sub
        Public Overrides Function ToString() As String

            ToString = Chr(34) & mudtprops.Name & Chr(34) & " <" & mudtprops.Email & ">"
            Dim pEmail() As String = mudtprops.Email.Split({";"}, StringSplitOptions.RemoveEmptyEntries)
            If IsArray(pEmail) Then
                Dim pString As String = ""
                For i As Integer = 0 To pEmail.GetUpperBound(0)

                    If pString <> "" Then
                        pString &= ";"
                    End If
                    pString &= Chr(34) & mudtprops.Name & Chr(34) & " <" & pEmail(i) & ">"
                Next
                ToString = pString
            End If

        End Function

        Public ReadOnly Property Id As Integer
            Get
                Id = mudtprops.Id
            End Get
        End Property
        Public ReadOnly Property Vessel_FK As Integer
            Get
                Vessel_FK = mudtprops.Vessel_FK
            End Get
        End Property
        Public Property Name As String
            Get
                Name = mudtprops.Name
            End Get
            Set(value As String)
                mudtprops.Name = value
                CheckValid()
            End Set
        End Property
        Public Property Details As String
            Get
                Details = mudtprops.Details
            End Get
            Set(value As String)
                mudtprops.Details = value
                CheckValid()
            End Set
        End Property
        Public ReadOnly Property EmailType As String
            Get
                EmailType = mudtprops.EmailType
            End Get
        End Property
        Public Property Email As String
            Get
                Email = mudtprops.Email
            End Get
            Set(value As String)
                mudtprops.Email = value
                CheckValid()
            End Set
        End Property
        Public ReadOnly Property VesselName As String
            Get
                VesselName = mudtprops.VesselName
            End Get
        End Property
        Public ReadOnly Property isValid As Boolean
            Get
                isValid = mudtprops.isValid
            End Get
        End Property
        Public ReadOnly Property isNew As Boolean
            Get
                isNew = mudtprops.isNew
            End Get
        End Property
        Public Sub SetValues(ByVal pId As Integer, ByVal pVessel_FK As Integer, ByVal pName As String, ByVal pDetails As String, ByVal pEmailType As String, ByVal pEmail As String, ByVal pVesselName As String)

            With mudtprops
                .Id = pId
                .Vessel_FK = pVessel_FK
                .Name = pName
                .Details = pDetails
                .EmailType = pEmailType
                .Email = pEmail
                .VesselName = pVesselName
                .isNew = False
                CheckValid()
            End With
        End Sub
        Public Sub Update()

            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                If mudtprops.isNew Then
                    If mudtprops.EmailType = "AGENT" Then
                        .CommandText = "INSERT INTO [AmadeusReports].[dbo].[osmEmailDetails] " & _
                                       " (osmeName, osmeDetails, osmeType, osmeEmail) " & _
                                       " VALUES " & _
                                       " ( '" & mudtprops.Name & "', '" & mudtprops.Details & "', '" & mudtprops.EmailType & "', '" & mudtprops.Email & "') " & _
                                       " select scope_identity() as Id"
                    Else
                        .CommandText = "INSERT INTO [AmadeusReports].[dbo].[osmEmailDetails] " & _
                                       " (osmeVessel_FK, osmeName, osmeDetails, osmeType, osmeEmail) " & _
                                       " VALUES " & _
                                       " ( " & mudtprops.Vessel_FK & ", '" & mudtprops.Name & "', '" & mudtprops.Details & "', '" & mudtprops.EmailType & "', '" & mudtprops.Email & "') " & _
                                       " select scope_identity() as Id"

                    End If
                    mudtprops.Id = .ExecuteScalar
                    mudtprops.isNew = False
                Else
                    .CommandText = "UPDATE AmadeusReports.dbo.osmEmailDetails " & _
                                   " SET osmeName = '" & mudtprops.Name & "', " & _
                                   "     osmeDetails = '" & mudtprops.Details & "', " & _
                                   "     osmeType    = '" & mudtprops.EmailType & "' ," & _
                                   "     osmeEmail   = '" & mudtprops.Email & "' " & _
                                   " WHERE osmeId = " & mudtprops.Id
                    .ExecuteNonQuery()
                End If
            End With
        End Sub
        Public Sub Delete()
            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                If Not mudtprops.isNew Then
                    .CommandText = "DELETE FROM AmadeusReports.dbo.osmEmailDetails " & _
                                   " WHERE osmeId = " & mudtprops.Id
                    .ExecuteNonQuery()
                End If
            End With
        End Sub
    End Class
    Public Class emailCollection
        Inherits Collections.Generic.Dictionary(Of String, emailItem)

        Public Sub Load(ByVal pVesselID As Integer)

            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader
            Dim pobjClass As emailItem

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .CommandText = "SELECT osmeId " & _
                               "      ,ISNULL(osmeVessel_FK,0) AS osmeVessel_FK " & _
                               "      ,ISNULL(osmeName, '') AS osmeName " & _
                               "      ,ISNULL(osmeDetails, '') AS osmeDetails " & _
                               "      ,ISNULL(osmeType, '') AS osmeType " & _
                               "      ,ISNULL(osmeEmail, '') AS osmeEmail " & _
                               "      ,ISNULL(osmvVesselName, '') AS osmvVesselName " & _
                               "  FROM AmadeusReports.dbo.osmEmailDetails " & _
                               " LEFT JOIN AmadeusReports.dbo.osmVessels " & _
                               "   ON osmeVessel_FK = osmVessels.osmvID " & _
                               "  WHERE ISNULL(osmeVessel_FK, " & pVesselID & ") = " & pVesselID & " " & _
                               "  ORDER BY osmeType, osmeName"
                pobjReader = .ExecuteReader
            End With

            MyBase.Clear()
            With pobjReader
                Do While .Read
                    pobjClass = New emailItem
                    pobjClass.SetValues(.Item("osmeId"), .Item("osmeVessel_FK"), .Item("osmeName"), .Item("osmeDetails"), .Item("osmeType"), .Item("osmeEmail"), .Item("osmvVesselName"))
                    MyBase.Add(pobjClass.Id, pobjClass)
                Loop
                .Close()
            End With
            pobjConn.Close()

        End Sub
        Public Sub Load()

            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader
            Dim pobjClass As emailItem

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .CommandText = "SELECT osmeId " & _
                               "      ,ISNULL(osmeVessel_FK,0) AS osmeVessel_FK " & _
                               "      ,ISNULL(osmeName, '') AS osmeName " & _
                               "      ,ISNULL(osmeDetails, '') AS osmeDetails " & _
                               "      ,ISNULL(osmeType, '') AS osmeType " & _
                               "      ,ISNULL(osmeEmail, '') AS osmeEmail " & _
                               "      ,ISNULL(osmvVesselName, '') AS osmvVesselName " & _
                               "  FROM AmadeusReports.dbo.osmEmailDetails " & _
                               " LEFT JOIN AmadeusReports.dbo.osmVessels " & _
                               "   ON osmeVessel_FK = osmVessels.osmvID " & _
                               "  WHERE osmeType = 'AGENT' " & _
                               "  ORDER BY  osmeName"
                pobjReader = .ExecuteReader
            End With

            MyBase.Clear()
            With pobjReader
                Do While .Read
                    pobjClass = New emailItem
                    pobjClass.SetValues(.Item("osmeId"), .Item("osmeVessel_FK"), .Item("osmeName"), .Item("osmeDetails"), .Item("osmeType"), .Item("osmeEmail"), .Item("osmvVesselName"))
                    MyBase.Add(pobjClass.Id, pobjClass)
                Loop
                .Close()
            End With
            pobjConn.Close()

        End Sub
    End Class

End Namespace
