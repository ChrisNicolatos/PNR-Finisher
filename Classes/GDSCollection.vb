Public Class GDSCollection
    Inherits Collections.Generic.Dictionary(Of Integer, GDSItem)
    Public Sub Load()
        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader
        Dim pobjClass As GDSItem
        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand
        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = " SELECT  pfrGDSId, pfrGDSName, pfrGDSCode " &
                           " FROM AmadeusReports.dbo.PNRFinisherGDS " &
                           " Order By pfrGDSId"
            pobjReader = .ExecuteReader
        End With
        MyBase.Clear()
        With pobjReader
            Do While .Read
                pobjClass = New GDSItem
                pobjClass.SetValues(CInt(.Item("pfrGDSId")), CStr(.Item("pfrGDSName")), CStr(.Item("pfrGDSCode")))
                MyBase.Add(pobjClass.Id, pobjClass)
            Loop
            .Close()
        End With
        pobjConn.Close()

    End Sub
End Class