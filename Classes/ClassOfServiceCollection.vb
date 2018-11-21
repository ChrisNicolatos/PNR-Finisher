Public Class ClassOfServiceCollection

    Private mClassCollection As New Collections.Generic.Dictionary(Of String, ClassOfServiceItem)

    Public Function Load(ByVal pCarrier As String, ByVal pOrigin As String, ByVal pDestination As String, ByVal pClassOfService As String) As ClassOfServiceItem

        Dim pItem As New ClassOfServiceItem
        pItem.SetValues(pCarrier, pOrigin, pDestination, pClassOfService, "", "")
        If mClassCollection.ContainsKey(pItem.Key) Then
            pItem = mClassCollection.Item(pItem.Key)

        Else
            Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringTWS_MIS) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .Parameters.Add("@Carrier", SqlDbType.Char, 2).Value = pCarrier
                .Parameters.Add("@ClassCode", SqlDbType.VarChar, 2).Value = pClassOfService
                .Parameters.Add("@OriginCode", SqlDbType.VarChar, 3).Value = pOrigin
                .Parameters.Add("@DestinationCode", SqlDbType.VarChar, 3).Value = pDestination

                .CommandText = "SELECT TOP 1 ISNULL(rtrim(cabinDescription), '') AS CabinDescription
		   , ISNULL(twref_Air_Classes.Description, '') AS classDescription
 from TWS_MIS.mis.twref_Air_Classes
 left join TWS_MIS.mis.twref_City_Code oCity
 on oCity.Code = @OriginCode
 left join TWS_MIS.mis.twref_City_Code dCity
 on dCity.Code = @DestinationCode
 left join TWS_MIS.mis.twref_Airline_Code
 on twref_Airline_Code.Code = @Carrier
 where @ClassCode = classcode 
      and (carrier = '*' or carrier = @Carrier)                     
	  and (dep = '*' or dep = @OriginCode or dep = '*' + oCity.CountryCode or dep = oCity.ContinentCode)                     
	  and (arr = '*' or arr = @DestinationCode or arr = '*' + dCity.CountryCode or arr = dCity.ContinentCode)
 order by carrier desc, dep desc, arr desc"
                pobjReader = .ExecuteReader
            End With

            With pobjReader
                If .Read Then
                    pItem.SetValues(pCarrier, pOrigin, pDestination, pClassOfService, CStr(.Item("CabinDescription")), CStr(.Item("classDescription")))
                    mClassCollection.Add(pItem.Key, pItem)
                End If
                .Close()
            End With
            pobjConn.Close()
        End If
        Return pItem

    End Function
End Class
