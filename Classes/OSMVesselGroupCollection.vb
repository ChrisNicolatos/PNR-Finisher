Public Class OSMVesselGroupCollection
    Inherits Collections.Generic.Dictionary(Of Integer, OSMVesselGroupItem)

    Public Sub Load(ByVal pVesselGroupID As Integer)

        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader
        Dim pobjClass As OSMVesselGroupItem

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .Parameters.Add("@VesselgroupId", SqlDbType.BigInt).Value = pVesselGroupID
            .CommandText = "SELECT osmvrId " &
                           "      ,osmvrGroupName " &
                           "  FROM AmadeusReports.dbo.osmVesselGroup " &
                           "  WHERE osmvrId = @VesselgroupId " &
                           "  ORDER BY osmvrGroupName"
            pobjReader = .ExecuteReader
        End With

        MyBase.Clear()
        With pobjReader
            Do While .Read
                pobjClass = New OSMVesselGroupItem
                pobjClass.SetValues(CInt(.Item("osmvrId")), CStr(.Item("osmvrGroupName")))
                MyBase.Add(pobjClass.Id, pobjClass)
            Loop
            .Close()
        End With
        pobjConn.Close()

    End Sub
    Public Sub Load()

        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader
        Dim pobjClass As OSMVesselGroupItem

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = "SELECT osmvrId " &
                           "      ,osmvrGroupName " &
                           "  FROM AmadeusReports.dbo.osmVesselGroup " &
                           "  ORDER BY osmvrGroupName"
            pobjReader = .ExecuteReader
        End With

        MyBase.Clear()
        With pobjReader
            Do While .Read
                pobjClass = New OSMVesselGroupItem
                pobjClass.SetValues(CInt(.Item("osmvrId")), CStr(.Item("osmvrGroupName")))
                MyBase.Add(pobjClass.Id, pobjClass)
            Loop
            .Close()
        End With
        pobjConn.Close()

    End Sub
End Class