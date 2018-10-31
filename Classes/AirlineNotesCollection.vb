Option Strict On
Option Explicit On
Public Class AirlineNotesCollection
    Inherits System.Collections.Generic.Dictionary(Of Integer, AirlineNotesItem)

    Public Sub Load(ByVal pAirlineCode As String, ByVal GDSCode As Utilities.EnumGDSCode)

        Dim pCommandText As String
        If GDSCode = Utilities.EnumGDSCode.Amadeus Then
            pCommandText = "SELECT anID, " &
                            " anAirlineCode, " &
                            " anFlightType, " &
                            " ISNULL(anSeaman, 0) AS anSeaman, " &
                            " anSeqNo, " &
                            " anAmadeusElement AS GDSElement, " &
                            " anAmadeusText AS GDSText " &
                            " FROM AmadeusReports.dbo.AirlineNotes " &
                            " WHERE anAirlineCode = @AirlineCode " &
                            " ORDER BY anSeqNo"
        ElseIf GDSCode = Utilities.EnumGDSCode.Galileo Then
            pCommandText = "SELECT anID, " &
                            " anAirlineCode, " &
                            " anFlightType, " &
                            " ISNULL(anSeaman, 0) AS anSeaman, " &
                            " anSeqNo, " &
                            " '' AS GDSElement, " &
                            " anGalileoEntry AS GDSText " &
                            " FROM AmadeusReports.dbo.AirlineNotes " &
                            " WHERE anAirlineCode = @AirlineCode " &
                            " ORDER BY anSeqNo"
        Else
            Throw New Exception("AirlineNotes.Collection.Load()" & vbCrLf & "GDS is not selected")
        End If
        ReadFromDB(pCommandText, pAirlineCode)

    End Sub
    Private Sub ReadFromDB(ByVal CommandText As String, ByVal pAirlineCode As String)

        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader
        Dim pobjClass As AirlineNotesItem
        Dim pID As Integer = 0

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand
        MyBase.Clear()
        With pobjComm
            .CommandType = CommandType.Text
            .Parameters.Add("@AirlineCode", SqlDbType.NVarChar, 10).Value = pAirlineCode
            .CommandText = CommandText
            pobjReader = .ExecuteReader
        End With

        With pobjReader
            Do While .Read
                pID += 1
                pobjClass = New AirlineNotesItem
                pobjClass.SetValues(CInt(.Item("anID")), CStr(.Item("anAirlineCode")), CStr(.Item("anFlightType")), CBool(.Item("anSeaman")),
                                        CInt(.Item("anSeqNo")), CStr(.Item("GDSElement")), CStr(.Item("GDSText")))
                MyBase.Add(pID, pobjClass)
            Loop
            .Close()
        End With
        pobjConn.Close()

    End Sub

End Class
