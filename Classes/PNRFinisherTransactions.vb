Public Class PNRFinisherTransactions
    Public Sub UpdateTransactions(ByVal pPNR As String, ByVal pGDS As String, ByVal pPCC As String, ByVal pUserID As String, ByVal pTransactionDate As Date _
           , ByVal pPax As String, ByVal pSegs As String, ByVal pFares As String, ByVal pClientCode As String, ByVal pNewEntry As Boolean)
        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand
        With pobjComm
            .CommandType = CommandType.Text
            .Parameters.Add("@pflPNR", SqlDbType.NVarChar, 6).Value = pPNR
            .Parameters.Add("@pflGDS", SqlDbType.Char, 2).Value = pGDS
            .Parameters.Add("@pflPCC", SqlDbType.NVarChar, 20).Value = pPCC
            .Parameters.Add("@pflUserID", SqlDbType.NVarChar, 20).Value = pUserID
            .Parameters.Add("@pflTransactionDate", SqlDbType.DateTime).Value = pTransactionDate
            .Parameters.Add("@pflPax", SqlDbType.NVarChar).Value = pPax
            .Parameters.Add("@pflSegs", SqlDbType.NVarChar).Value = pSegs
            .Parameters.Add("@pflFares", SqlDbType.NVarChar).Value = pFares
            .Parameters.Add("@pflClientCode", SqlDbType.NVarChar, 20).Value = pClientCode
            .Parameters.Add("@pflNewEntry", SqlDbType.Bit).Value = pNewEntry

            .CommandText = "INSERT INTO [AmadeusReports].[dbo].[PNRFinisherTransactions]
           ([pflPNR]
           ,[pflGDS]
           ,[pflPCC]
           ,[pflUserID]
           ,[pflTransactionDate]
           ,[pflPax]
           ,[pflSegs]
           ,[pflFares]
           ,[pflClientCode]
           ,[pflNewEntry])
     VALUES
           (@pflPNR
           ,@pflGDS
           ,@pflPCC
           ,@pflUserID
           ,@pflTransactionDate
           ,@pflPax
           ,@pflSegs
           ,@pflFares
           ,@pflClientCode
           ,@pflNewEntry)"
            .ExecuteNonQuery()
        End With
        pobjConn.Close()
    End Sub
End Class
