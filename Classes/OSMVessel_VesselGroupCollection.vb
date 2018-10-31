Public Class OSMVessel_VesselGroupCollection
    Inherits Collections.Generic.Dictionary(Of Integer, OSMVessel_VesselGroupItem)

    Public Sub Load(ByVal pVesselId As Integer)

        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader
        Dim pobjClass As OSMVessel_VesselGroupItem

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = " SELECT [osmvrId] " &
                           "       ,[osmvrGroupName] " &
                           " 	  ,ISNULL(osmvVesselName, '') AS osmvVesselName " &
                           " 	  ," & pVesselId & " AS osmvId " &
                           " 	  ,ISNULL((SELECT osmvId_fkey FROM osmVesselGroup_Vessels WHERE osmvrId = osmvrId_fkey AND osmvId_fkey=" & pVesselId & "),0) AS osmvId_fkey " &
                           "   FROM [AmadeusReports].[dbo].[osmVesselGroup] " &
                           "   LEFT JOIN osmVessels " &
                           "   ON osmvID = " & pVesselId

            pobjReader = .ExecuteReader
        End With

        MyBase.Clear()
        Dim pId As Integer = 0
        With pobjReader
            Do While .Read
                pobjClass = New OSMVessel_VesselGroupItem
                pId += 1
                pobjClass.SetValues(pId, CInt(.Item("osmvId")), CInt(.Item("osmvrId")), CStr(.Item("osmvVesselName")), CStr(.Item("osmvrGroupName")), CInt(.Item("osmvId_fkey")))
                MyBase.Add(pobjClass.Id, pobjClass)
            Loop
            .Close()
        End With
        pobjConn.Close()

    End Sub
    Public Sub Update()

        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjClass As OSMVessel_VesselGroupItem

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        For Each pobjClass In MyBase.Values
            pobjComm.CommandType = CommandType.Text
            If pobjClass.Exists Then
                pobjComm.CommandText = "IF (SELECT COUNT(*) FROM AmadeusReports.dbo.osmVesselGroup_Vessels WHERE osmvrId_fkey = " & pobjClass.VesselGroupId & " AND osmvId_fkey = " & pobjClass.VesselId & ")=0" &
                    "INSERT INTO AmadeusReports.dbo.osmVesselGroup_Vessels (osmvrId_fkey ,osmvId_fkey) VALUES (" & pobjClass.VesselGroupId & "," & pobjClass.VesselId & ")"
            Else
                pobjComm.CommandText = "DELETE FROM AmadeusReports.dbo.osmVesselGroup_Vessels WHERE osmvrId_fkey = " & pobjClass.VesselGroupId & " AND osmvId_fkey = " & pobjClass.VesselId
            End If
            pobjComm.ExecuteNonQuery()
        Next
    End Sub
End Class