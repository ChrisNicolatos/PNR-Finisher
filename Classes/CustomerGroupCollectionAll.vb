Public Class CustomerGroupCollectionAll
    Inherits Collections.Generic.Dictionary(Of Integer, CustomerGroupItem)
    Public Sub Load()

        Dim pCommandText As String

        Try
            pCommandText = " USE TravelForceCosmos " &
                           " SELECT Id " &
                           " ,Description " &
                           " FROM Tags " &
                           " WHERE TagGroupId = 146 " &
                           " ORDER BY Description "
            ReadCustomerGroups(pCommandText)
        Catch ex As Exception
            Throw New Exception("AllCustomerGroups.Load()" & vbCrLf & ex.Message)
        End Try

    End Sub
    Private Sub ReadCustomerGroups(ByVal CommandText As String)

        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringACC) ' ActiveConnection)
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
                pobjClass.SetValues(CInt(.Item("Id")), CStr(.Item("Description")))
                MyBase.Add(pobjClass.ID, pobjClass)
            Loop
            .Close()
        End With
        pobjConn.Close()

    End Sub
End Class