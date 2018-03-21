Imports System.Data.SqlClient

Namespace AirlinePoints1G

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

        Friend Sub SetValues(ByVal pCustID As Integer, ByVal pCustCode As String, ByVal pCustName As String,
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
        Inherits Dictionary(Of Integer, Item)

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

            Dim pobjConn As New SqlConnection(ConnectionStringACC) ' ActiveConnection)
            Dim pobjComm As New SqlCommand
            Dim pobjReader As SqlDataReader
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
                    Dim p1GRemark As String = Get1GRemark(.Item("Remarks"))
                    If p1GRemark <> "" Then
                        pID += 1
                        pobjClass = New Item
                        pobjClass.SetValues(.Item("ID"), .Item("Code"), .Item("Name"), .Item("IATACode"),
                                        .Item("AirlineName"), p1GRemark)
                        MyBase.Add(pID, pobjClass)
                    End If
                Loop
                .Close()
            End With
            pobjConn.Close()
        End Sub
        Private Function Get1GRemark(ByVal p1ARemark As String) As String
            Dim pobjConn As New SqlConnection(ConnectionStringPNR) ' ActiveConnection)
            Dim pobjComm As New SqlCommand
            Dim pobjReader As SqlDataReader

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

End Namespace
