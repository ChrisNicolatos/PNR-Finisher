Public Class ConfigGDSReferenceCollection
    Inherits Collections.Generic.Dictionary(Of String, ConfigGDSReferenceItem)
    Public Sub Read(ByVal BackOffice As Integer, ByVal GDSCode As Utilities.EnumGDSCode)

        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .Parameters.Add("@PCCBackOffice", SqlDbType.BigInt).Value = BackOffice
            .Parameters.Add("@GDS", SqlDbType.BigInt).Value = GDSCode
            .CommandText = " SELECT pfrID " &
                           " , ISNULL(pfrKey,'') AS pfrKey " &
                           " , ISNULL(pfrValue,'') AS pfrValue " &
                           " , ISNULL(pfrGDS_fkey,0) AS pfrGDS_fkey " &
                           " , ISNULL(pfrBO_fkey,0) AS pfrBO_fkey " &
                           " , ISNULL(pfrGDSElement,'') AS pfrGDSElement " &
                           " , ISNULL(pfrReferenceIdentifier,'') AS pfrReferenceIdentifier " &
                           " , ISNULL(pfrReferenceDetail,'') AS pfrReferenceDetail " &
                           " FROM [AmadeusReports].[dbo].[PNRFinisherGDS_BOReferences] " &
                           " WHERE pfrGDS_fkey = @GDS AND pfrBO_fkey = @PCCBackOffice"
            pobjReader = .ExecuteReader
        End With

        MyBase.Clear()

        With pobjReader
            While pobjReader.Read
                Dim pItem As New ConfigGDSReferenceItem
                pItem.SetValues(CInt(.Item("pfrID")), CStr(.Item("pfrKey")), CStr(.Item("pfrValue")), CInt(.Item("pfrGDS_fkey")), CInt(.Item("pfrBO_fkey")), CStr(.Item("pfrGDSElement")), CStr(.Item("pfrReferenceIdentifier")), CStr(.Item("pfrReferenceDetail")))
                MyBase.Add(pItem.Key, pItem)
            End While
            .Close()
        End With
        pobjConn.Close()

    End Sub
End Class