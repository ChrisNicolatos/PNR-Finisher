Public Class OSMVesselItem
    Private Structure ClassProps
        Dim Id As Integer
        Dim VesselName As String
        Dim VesselGroup() As Integer
        Dim VesselgroupCount As Integer
        Dim InUse As Boolean
        Dim isNew As Boolean
        Dim isValid As Boolean
    End Structure
    Dim mudtProps As ClassProps
    Dim mobjVessel_VesselGroup As New OSMVessel_VesselGroupCollection

    Public Sub New()
        With mudtProps
            .Id = 0
            .VesselName = ""
            .VesselgroupCount = 0
            ReDim .VesselGroup(0)
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
    Public ReadOnly Property VesselGroupCount As Integer
        Get
            VesselGroupCount = mudtProps.VesselgroupCount
        End Get
    End Property
    Public ReadOnly Property VesselGroup(ByVal Index As Integer) As Integer
        Get
            If Index >= 0 And Index < mudtProps.VesselgroupCount Then
                VesselGroup = mudtProps.VesselGroup(Index)
            Else
                VesselGroup = 0
            End If
        End Get
    End Property
    Public ReadOnly Property Vessel_VesselGroup As OSMVessel_VesselGroupCollection
        Get
            If mobjVessel_VesselGroup.Count = 0 Then
                mobjVessel_VesselGroup.Load(Id)
            End If
            Vessel_VesselGroup = mobjVessel_VesselGroup
        End Get
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
    Public Sub SetValues(ByVal pId As Integer, ByVal pVesselName As String, ByVal pInUse As Boolean)
        With mudtProps
            .Id = pId
            .VesselName = pVesselName
            .InUse = pInUse
            .isNew = False
            'mobjVessel_VesselGroup.Load(Id)
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

                Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
                Dim pobjComm As New SqlClient.SqlCommand

                pobjConn.Open()
                pobjComm = pobjConn.CreateCommand

                With pobjComm
                    .CommandType = CommandType.Text
                    If mudtProps.isNew Then
                        .CommandText = "IF (SELECT COUNT(*) FROM [AmadeusReports].[dbo].[osmVessels] WHERE osmvVesselName = '" & mudtProps.VesselName & "') = 0 " &
                                       " INSERT INTO AmadeusReports.dbo.osmVessels " &
                                       " (osmvVesselName, osmvVesselGroup, osmvInUse) " &
                                       " VALUES " &
                                       " ( '" & mudtProps.VesselName & "', '', " & If(mudtProps.InUse, 1, 0) & ") " &
                                       " select scope_identity() as Id"
                        Dim pTemp As Integer = CInt(.ExecuteScalar)
                        If IsDBNull(pTemp) Then
                            Throw New Exception("Vessel Already exists")
                        Else
                            mudtProps.Id = pTemp
                            mudtProps.isNew = False
                        End If
                    Else
                        .CommandText = "UPDATE AmadeusReports.dbo.osmVessels " &
                                       " SET osmvVesselName = '" & mudtProps.VesselName & "', " &
                                       "     osmvVesselGroup = '', " &
                                       "     osmvInUse = " & If(mudtProps.InUse, 1, 0) & " " &
                                       " WHERE osmvId = " & mudtProps.Id
                        .ExecuteNonQuery()
                    End If
                End With
                mobjVessel_VesselGroup.Update()
            Else
                Throw New Exception("Vessel name invalid")
            End If
        Catch ex As Exception
            Throw New Exception("Update Vessel Error" & vbCrLf & ex.Message)
        End Try


    End Sub

End Class