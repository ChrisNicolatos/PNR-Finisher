Public Class BackOfficeCollection
    Inherits Collections.Generic.Dictionary(Of Integer, BackOfficeItem)

    Public Sub Load()

        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader
        Dim pobjClass As BackOfficeItem
        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = " SELECT  pfrBOId, pfrBOName " &
                           " FROM AmadeusReports.dbo.PNRFinisherBackOffice " &
                           " Order By pfrBOId"
            pobjReader = .ExecuteReader
        End With

        MyBase.Clear()

        With pobjReader
            Do While .Read
                pobjClass = New BackOfficeItem
                pobjClass.SetValues(CInt(.Item("pfrBOId")), CStr(.Item("pfrBOName")))
                MyBase.Add(pobjClass.Id, pobjClass)
            Loop
            .Close()
        End With
        pobjConn.Close()

    End Sub
End Class