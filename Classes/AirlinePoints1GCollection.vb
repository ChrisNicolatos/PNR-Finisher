Option Strict On
Option Explicit On
Public Class AirlinePoints1GCollection
    Inherits Dictionary(Of Integer, AirlinePointsItem)

    Public Sub Load(ByVal pCustID As Integer, ByVal pIATACode As String)
        Dim pCommandText As String
        pCommandText = "SELECT TFEntities.Id" &
                               "	   , TFEntities.Code " &
                               " 	   , TFEntities.Name " &
                               " 	   , Airlines.IATACode " &
                               " 	   , Airlines.AirlineName " &
                               " 	   , FrequentFlyerCards.Remarks " &
                               " FROM FrequentFlyerCards  " &
                               " 	LEFT OUTER JOIN TFEntities  " &
                               " 		ON FrequentFlyerCards.TFEntityID = TFEntities.Id  " &
                               " 	LEFT OUTER JOIN Airlines  " &
                               " 		ON FrequentFlyerCards.AirlineID = Airlines.Id " &
                               " WHERE (TFEntities.Id = " & pCustID & ")  " &
                               " 			AND (Airlines.IATACode = '" & pIATACode & "')"
        ReadFromDB(pCommandText)

    End Sub

    Public Sub Load()
        Dim pCommandText As String
        pCommandText = "SELECT TFEntities.Id" &
                               "	   , TFEntities.Code " &
                               " 	   , TFEntities.Name " &
                               " 	   , Airlines.IATACode " &
                               " 	   , Airlines.AirlineName " &
                               " 	   , FrequentFlyerCards.Remarks " &
                               " FROM FrequentFlyerCards  " &
                               " 	LEFT OUTER JOIN TFEntities  " &
                               " 		ON FrequentFlyerCards.TFEntityID = TFEntities.Id  " &
                               " 	LEFT OUTER JOIN Airlines  " &
                               " 		ON FrequentFlyerCards.AirlineID = Airlines.Id " &
                               " ORDER BY TFEntities.Code, Airlines.IATACode"

        ReadFromDB(pCommandText)

    End Sub

    Private Sub ReadFromDB(ByVal CommandText As String)

        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringACC) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader
        Dim pobjClass As AirlinePointsItem
        Dim pID As Integer = 0

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand
        MyBase.Clear()
        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = CommandText
            pobjReader = .ExecuteReader
        End With

        With pobjReader
            Do While .Read
                Dim p1GRemark As String = Get1GRemark(CStr(.Item("Remarks")))
                If p1GRemark <> "" Then
                    pID += 1
                    pobjClass = New AirlinePointsItem
                    pobjClass.SetValues(CInt(.Item("ID")), CStr(.Item("Code")), CStr(.Item("Name")), CStr(.Item("IATACode")),
                                        CStr(.Item("AirlineName")), p1GRemark)
                    MyBase.Add(pID, pobjClass)
                End If
            Loop
            .Close()
        End With
        pobjConn.Close()
    End Sub
    Private Function Get1GRemark(ByVal p1ARemark As String) As String
        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand
        MyBase.Clear()
        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = "SELECT ISNULL(ff1GRemark,'') AS ff1GRemark " &
                               " FROM AmadeusReports.dbo.FrequentFlyerCards_1G " &
                               " WHERE ffTFCRemark = '" & p1ARemark & "'"
            pobjReader = .ExecuteReader
        End With
        With pobjReader
            If .Read Then
                Get1GRemark = .Item("ff1GRemark").ToString.Trim
            Else
                Get1GRemark = ""
            End If
        End With
    End Function

End Class
