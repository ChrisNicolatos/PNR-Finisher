Option Strict On
Option Explicit On
Public Class OSMAddressCollection
    Inherits System.Collections.Generic.Dictionary(Of Integer, OSMAddressItem)
    Public Sub Load()
        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader
        Dim pobjClass As OSMAddressItem

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand
        MyBase.Clear()
        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = "SELECT osmaId
                          ,osmaOffice
                          ,osmaName
                          ,osmaTitle
                          ,osmaCompanyName
                          ,osmaAddress
                          ,osmaPCArea
                          ,osmaCountry
                          ,osmaTelephone
                          ,ISNULL(osmaLogoImage, 0) AS osmaLogoImage
                          ,ISNULL(osmaSignatureImage, 0) AS osmaSignatureImage
                      FROM AmadeusReports.dbo.osmOfficeAddresses
                      ORDER BY osmaOffice"
            pobjReader = .ExecuteReader
        End With

        With pobjReader
            Do While .Read
                pobjClass = New OSMAddressItem
                pobjClass.SetValues(CInt(.Item("osmaId")), CStr(.Item("osmaOffice")), CStr(.Item("osmaName")), CStr(.Item("osmaTitle")) _
                                    , CStr(.Item("osmaCompanyName")), CStr(.Item("osmaAddress")), CStr(.Item("osmaPCArea")), CStr(.Item("osmaCountry")) _
                                    , CStr(.Item("osmaTelephone")), CInt(.Item("osmaLogoImage")), CInt(.Item("osmaSignatureImage")))
                MyBase.Add(pobjClass.Id, pobjClass)
            Loop
            .Close()
        End With
        pobjConn.Close()
    End Sub
End Class
