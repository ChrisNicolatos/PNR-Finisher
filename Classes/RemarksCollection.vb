Public Class RemarksCollection
    Inherits Collections.Generic.Dictionary(Of Integer, RemarksItem)
    Public Sub Load()
        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader
        Dim pobjClass As RemarksItem
        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = " SELECT [pnrrID]
                            ,[pnrrTitle]
                            ,[pnrrRemark]
                            ,[pnrrInUse]
                            FROM [AmadeusReports].[dbo].[PNRFinisherRemarks]
                            ORDER BY pnrrID "
            pobjReader = .ExecuteReader
        End With

        With pobjReader
            Do While .Read
                pobjClass = New RemarksItem
                pobjClass.SetValues(CInt(.Item("pnrrID")), CStr(.Item("pnrrTitle")), CStr(.Item("pnrrRemark")), CBool(.Item("pnrrInUse")))
                MyBase.Add(pobjClass.Id, pobjClass)
            Loop
            .Close()
        End With
        pobjConn.Close()
    End Sub
End Class
