Namespace PriceLookup

    Public Class Item
        Private Structure ClassProps
            Dim Id As Integer
            Dim CustomerName As String
            Dim Airline As String
            Dim ClassOfService As String
            Dim TicketCount As Integer
            Dim AveragePrice As Decimal
            Dim CustomerNameNull As Boolean
            Dim AirlineNull As Boolean
            Dim ClassOfServiceNull As Boolean
        End Structure
        Private mudtProps As ClassProps
        Public ReadOnly Property Id As Integer
            Get
                Id = mudtProps.Id
            End Get
        End Property
        Public ReadOnly Property CustomerName As String
            Get
                CustomerName = mudtProps.CustomerName
            End Get
        End Property
        Public ReadOnly Property Airline As String
            Get
                Airline = mudtProps.Airline
            End Get
        End Property
        Public ReadOnly Property ClassOfService As String
            Get
                ClassOfService = mudtProps.ClassOfService
            End Get
        End Property
        Public ReadOnly Property CustomerNameNull As Boolean
            Get
                CustomerNameNull = mudtProps.CustomerNameNull
            End Get
        End Property
        Public ReadOnly Property AirlineNull As Boolean
            Get
                AirlineNull = mudtProps.AirlineNull
            End Get
        End Property
        Public ReadOnly Property ClassOfServiceNull As Boolean
            Get
                ClassOfServiceNull = mudtProps.ClassOfServiceNull
            End Get
        End Property
        Public ReadOnly Property TicketCount As Integer
            Get
                TicketCount = mudtProps.TicketCount
            End Get
        End Property
        Public ReadOnly Property AveragePrice As Decimal
            Get
                AveragePrice = mudtProps.AveragePrice
            End Get
        End Property
        Friend Sub SetValues(ByVal pId As Integer, ByVal pCustomerName As Object, ByVal pAirline As Object, ByVal pClassOfService As Object, ByVal pTicketCount As Object, ByVal pAveragePrice As Object)
            With mudtProps
                .Id = pId
                If IsDBNull(pCustomerName) Then
                    .CustomerNameNull = True
                    .CustomerName = ""
                Else
                    .CustomerNameNull = False
                    .CustomerName = pCustomerName
                End If
                If IsDBNull(pAirline) Then
                    .AirlineNull = True
                    .Airline = ""
                Else
                    .AirlineNull = False
                    .Airline = pAirline
                End If
                If IsDBNull(pClassOfService) Then
                    .ClassOfServiceNull = True
                    .ClassOfService = ""
                Else
                    .ClassOfServiceNull = False
                    .ClassOfService = pClassOfService
                End If
                If IsDBNull(pTicketCount) Then
                    .TicketCount = 0
                Else
                    .TicketCount = pTicketCount
                End If
                If IsDBNull(pAveragePrice) Then
                    .AveragePrice = 0
                Else
                    .AveragePrice = pAveragePrice
                End If
            End With
        End Sub
    End Class
    Public Class Collection
        Inherits Collections.Generic.Dictionary(Of Integer, Item)

        Private mTicketCount As Integer
        Private mAveragePrice As Decimal

        Private mFromDate As Date
        Private mOrigin As String
        Private mDestination As String
        Private mValuesLoaded As Boolean = False

        Public Sub SetValues(ByVal pFromDate As Date, ByVal pItinerary As String)
            pItinerary = pItinerary.Trim
            If pItinerary.Length >= 6 Then
                mOrigin = pItinerary.Substring(0, 3)
                mDestination = pItinerary.Substring(pItinerary.IndexOf(" ") - 3, 3)
                mFromDate = pFromDate
                mValuesLoaded = True
            End If
        End Sub
        Public Function Load() As Boolean

            mTicketCount = 0
            mAveragePrice = 0
            Load = False

            If mValuesLoaded Then
                If mOrigin <> mDestination Then
                    Load(mFromDate, mOrigin, mDestination)
                    Load = True
                End If
            End If

        End Function

        Private Sub Load(ByVal FromDate As Date, ByVal Origin As String, ByVal Destination As String)

            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringACC) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader
            Dim pobjClass As Item

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .CommandText = " SELECT TFEntities.Code + '/' + TFEntities.Name AS ClientName  " & _
                                " 	  ,Airlines.IATACode AS Airline  " & _
                                "       ,AirSegFrom.ActualClass AS Class  " & _
                                "   	  ,COUNT(*) AS CountOfTkts  " & _
                                " 	  ,AVG(-(  CTCost.FaceValue          + CTCost.FVVatAmount        + CTCost.FaceValueExtra  " & _
                                "                + CTCost.FVXVatAmount       + CTCost.Taxes              + CTCost.TAXVatAmount  " & _
                                " 			   + CTCost.TaxesExtra         + CTCost.TAXXVatAmount      + CTCost.DiscountAmount  " & _
                                " 			   + CTCost.DISCVatAmount      + CTCost.CommissionAmount   + CTCost.COMVatAmount    " & _
                                " 			   + CTCost.ServiceFeeAmount   + CTCost.SFVatAmount        + CTCost.ExtraChargeAmount1  " & _
                                " 			   + CTCost.ExtraChargeAmount2 + CTCost.ExtraChargeAmount3 + CTCost.CancellationFeeAmount  " & _
                                " 			   + CTCost.CFVatAmount   " & _
                                " 		)) AS AverageCostPrice  " & _
                                "   FROM CommercialTransactions  " & _
                                "   LEFT JOIN CommercialTransactionValues CTV  " & _
                                " 	LEFT JOIN TFEntities  " & _
                                " 		ON CTV.CommercialEntityID = TFEntities.Id  " & _
                                " 	ON CommercialTransactions.Id = CTV.CommercialTransactionID   " & _
                                " 	   AND IsCost = 0  " & _
                                "   LEFT JOIN AirTicketTransactions  " & _
                                " 	LEFT JOIN AirSegments AirSegFrom  " & _
                                " 		ON AirSegFrom.AirTicketTransactionID = AirTicketTransactions.Id   " & _
                                " 		   AND OriginalPosition = (SELECT MIN(OriginalPosition) FROM AirSegments WHERE AirSegments.AirTicketTransactionID = AirTicketTransactions.Id)  " & _
                                " 	LEFT JOIN AirSegments AirSegTo  " & _
                                " 		On AirSegTo.AirTicketTransactionID = AirTicketTransactions.Id  " & _
                                " 			AND AirSegTo.OriginalPosition = (SELECT MAX(OriginalPosition) FROM AirSegments WHERE AirSegments.AirTicketTransactionID = AirTicketTransactions.Id)  " & _
                                " 	LEFT JOIN AirTickets  " & _
                                " 		ON AirTicketTransactions.AirTicketID = AirTickets.Id  " & _
                                " 	ON AirTicketTransactions.CommercialTransactionID = CommercialTransactions.Id  " & _
                                "   LEFT JOIN Airports AirportFrom  " & _
                                " 	ON AirSegFrom.FromAirportID = AirportFrom.Id  " & _
                                "   LEFT JOIN Airports AirportTo  " & _
                                " 	ON AirSegTo.ToAirportID = AirportTo.Id  " & _
                                "   LEFT JOIN Airlines  " & _
                                " 	ON AirSegFrom.CarrierAirlineID = Airlines.Id  " & _
                                "   LEFT JOIN CommercialTransactionValues CTCost  " & _
                                " 	ON CTCost.CommercialTransactionID = CTV.CommercialTransactionID   " & _
                                " 			  AND CTCost.IsCost=1  " & _
                                "   WHERE ComTransactionTypeID = 1                   " & _
                                "         AND ActionTypeID = 335                     " & _
                                " 		AND AirTickets.IssueDate  >= '" & Format(FromDate, "yyyy-MM-dd") & "'     " & _
                                " 		AND AirportFrom.Abbreviation = '" & Origin & "'  " & _
                                " 		AND AirportTo.Abbreviation   = '" & Destination & "'  " & _
                                "  GROUP BY GROUPING SETS ( " & _
                                " 						 (TFEntities.Code + '/' + TFEntities.Name) " & _
                                " 						,(TFEntities.Code + '/' + TFEntities.Name, Airlines.IATACode) " & _
                                " 						,(Airlines.IATACode ,AirSegFrom.ActualClass) " & _
                                " 						,(Airlines.IATACode) " & _
                                " 						,() " & _
                                " 						)" & _
                                "   ORDER BY TFEntities.Code + '/' + TFEntities.Name  " & _
                                " 		   ,Airlines.IATACode  " & _
                                " 		   ,AirSegFrom.ActualClass  "
                pobjReader = .ExecuteReader
            End With

            Dim pId As Integer = 0
            mTicketCount = 0
            mAveragePrice = 0
            MyBase.Clear()

            With pobjReader
                Do While .Read
                    pId = pId + 1
                    pobjClass = New Item
                    pobjClass.SetValues(pId, .Item("ClientName"), .Item("Airline"), .Item("Class"), .Item("CountOfTkts"), .Item("AverageCostPrice"))
                    MyBase.Add(pobjClass.Id, pobjClass)
                    If pobjClass.CustomerNameNull And pobjClass.AirlineNull And pobjClass.ClassOfServiceNull Then
                        mTicketCount = pobjClass.TicketCount
                        mAveragePrice = pobjClass.AveragePrice
                    End If
                Loop
                .Close()
            End With
            pobjConn.Close()

        End Sub
        Public ReadOnly Property TicketCount As Integer
            Get
                TicketCount = mTicketCount
            End Get
        End Property
        Public ReadOnly Property AveragePrice As Decimal
            Get
                AveragePrice = mAveragePrice
            End Get
        End Property
        Public ReadOnly Property Itinerary As String
            Get
                Itinerary = mOrigin & "-" & mDestination
            End Get
        End Property
        Public ReadOnly Property FromDate As Date
            Get
                FromDate = mFromDate
            End Get
        End Property
    End Class
End Namespace
