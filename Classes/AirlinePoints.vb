Namespace AirlinePoints

    Public Class Item
        Private Structure ClassProps
            Dim CustomerID As Integer
            Dim CustomerCode As String
            Dim CustomerName As String
            Dim AirlineCode As String
            Dim AirlineName As String
            Dim PointsCommand As String
        End Structure
        Private mudtProps As ClassProps

        Public ReadOnly Property CustomerID() As Long
            Get
                CustomerID = mudtProps.CustomerID
            End Get
        End Property
        Public ReadOnly Property CustomerCode() As String
            Get
                CustomerCode = mudtProps.CustomerCode
            End Get
        End Property
        Public ReadOnly Property CustomerName() As String
            Get
                CustomerName = mudtProps.CustomerName
            End Get
        End Property
        Public ReadOnly Property AirlineCode() As String
            Get
                AirlineCode = mudtProps.AirlineCode
            End Get
        End Property
        Public ReadOnly Property AirlineName() As String
            Get
                AirlineName = mudtProps.AirlineName
            End Get
        End Property
        Public ReadOnly Property PointsCommand() As String
            Get
                PointsCommand = mudtProps.PointsCommand
            End Get
        End Property

        Friend Sub SetValues(ByVal pCustID As Integer, ByVal pCustCode As String, ByVal pCustName As String, _
                             ByVal pAirlineCode As String, ByVal pAirlineName As String, ByVal pPointsCommand As String)
            With mudtProps
                .CustomerID = pCustID
                .CustomerCode = pCustCode
                .CustomerName = pCustName
                .AirlineCode = pAirlineCode
                .AirlineName = pAirlineName
                .PointsCommand = pPointsCommand
            End With
        End Sub
    End Class

    Public Class Collection
        Inherits System.Collections.Generic.Dictionary(Of Integer, Item)

        Public Sub Load(ByVal pCustID As Integer, ByVal pIATACode As String)

            Dim pCommandText As String
            pCommandText = "SELECT TravelForceCosmos.dbo.TFEntities.Id" & _
                               "	   , TravelForceCosmos.dbo.TFEntities.Code " & _
                               " 	   , TravelForceCosmos.dbo.TFEntities.Name " & _
                               " 	   , TravelForceCosmos.dbo.Airlines.IATACode " & _
                               " 	   , TravelForceCosmos.dbo.Airlines.AirlineName " & _
                               " 	   , TravelForceCosmos.dbo.FrequentFlyerCards.Remarks " & _
                               " FROM TravelForceCosmos.dbo.FrequentFlyerCards  " & _
                               " 	LEFT OUTER JOIN TravelForceCosmos.dbo.TFEntities  " & _
                               " 		ON TravelForceCosmos.dbo.FrequentFlyerCards.TFEntityID = TravelForceCosmos.dbo.TFEntities.Id  " & _
                               " 	LEFT OUTER JOIN TravelForceCosmos.dbo.Airlines  " & _
                               " 		ON TravelForceCosmos.dbo.FrequentFlyerCards.AirlineID = TravelForceCosmos.dbo.Airlines.Id " & _
                               " WHERE (TravelForceCosmos.dbo.TFEntities.Id = " & pCustID & ")  " & _
                               " 			AND (TravelForceCosmos.dbo.Airlines.IATACode = '" & pIATACode & "')"
            ReadFromDB(pCommandText)

        End Sub

        Public Sub Load()

            Dim pCommandText As String
            pCommandText = "SELECT TravelForceCosmos.dbo.TFEntities.Id" & _
                               "	   , TravelForceCosmos.dbo.TFEntities.Code " & _
                               " 	   , TravelForceCosmos.dbo.TFEntities.Name " & _
                               " 	   , TravelForceCosmos.dbo.Airlines.IATACode " & _
                               " 	   , TravelForceCosmos.dbo.Airlines.AirlineName " & _
                               " 	   , TravelForceCosmos.dbo.FrequentFlyerCards.Remarks " & _
                               " FROM TravelForceCosmos.dbo.FrequentFlyerCards  " & _
                               " 	LEFT OUTER JOIN TravelForceCosmos.dbo.TFEntities  " & _
                               " 		ON TravelForceCosmos.dbo.FrequentFlyerCards.TFEntityID = TravelForceCosmos.dbo.TFEntities.Id  " & _
                               " 	LEFT OUTER JOIN TravelForceCosmos.dbo.Airlines  " & _
                               " 		ON TravelForceCosmos.dbo.FrequentFlyerCards.AirlineID = TravelForceCosmos.dbo.Airlines.Id " & _
                               " ORDER BY TFEntities.Code, TravelForceCosmos.dbo.Airlines.IATACode"

            ReadFromDB(pCommandText)

        End Sub

        Private Sub ReadFromDB(ByVal CommandText As String)

            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringACC) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader
            Dim pobjClass As Item
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
                    pID += 1
                    pobjClass = New Item
                    pobjClass.SetValues(.Item("ID"), .Item("Code"), .Item("Name"), .Item("IATACode"), _
                                        .Item("AirlineName"), .Item("Remarks"))
                    MyBase.Add(pID, pobjClass)
                Loop
                .Close()
            End With
            pobjConn.Close()

        End Sub

    End Class

End Namespace
