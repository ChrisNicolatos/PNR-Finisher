Option Strict On
Option Explicit On
Public Class AveragePriceCollection
    Inherits Collections.Generic.Dictionary(Of Integer, AveragePriceItem)

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

        If MySettings.PCCBackOffice = 1 Then

            Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringACC) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader
            Dim pobjClass As AveragePriceItem

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .CommandText = " SELECT TFEntities.Code + '/' + TFEntities.Name AS ClientName  " &
                                " 	  ,Airlines.IATACode AS Airline  " &
                                "       ,AirSegFrom.ActualClass AS Class  " &
                                "   	  ,COUNT(*) AS CountOfTkts  " &
                                " 	  ,AVG(-(  CTCost.FaceValue          + CTCost.FVVatAmount        + CTCost.FaceValueExtra  " &
                                "                + CTCost.FVXVatAmount       + CTCost.Taxes              + CTCost.TAXVatAmount  " &
                                " 			   + CTCost.TaxesExtra         + CTCost.TAXXVatAmount      + CTCost.DiscountAmount  " &
                                " 			   + CTCost.DISCVatAmount      + CTCost.CommissionAmount   + CTCost.COMVatAmount    " &
                                " 			   + CTCost.ServiceFeeAmount   + CTCost.SFVatAmount        + CTCost.ExtraChargeAmount1  " &
                                " 			   + CTCost.ExtraChargeAmount2 + CTCost.ExtraChargeAmount3 + CTCost.CancellationFeeAmount  " &
                                " 			   + CTCost.CFVatAmount   " &
                                " 		)) AS AverageCostPrice  " &
                                "   FROM CommercialTransactions  " &
                                "   LEFT JOIN CommercialTransactionValues CTV  " &
                                " 	LEFT JOIN TFEntities  " &
                                " 		ON CTV.CommercialEntityID = TFEntities.Id  " &
                                " 	ON CommercialTransactions.Id = CTV.CommercialTransactionID   " &
                                " 	   AND IsCost = 0  " &
                                "   LEFT JOIN AirTicketTransactions  " &
                                " 	LEFT JOIN AirSegments AirSegFrom  " &
                                " 		ON AirSegFrom.AirTicketTransactionID = AirTicketTransactions.Id   " &
                                " 		   AND OriginalPosition = (SELECT MIN(OriginalPosition) FROM AirSegments WHERE AirSegments.AirTicketTransactionID = AirTicketTransactions.Id)  " &
                                " 	LEFT JOIN AirSegments AirSegTo  " &
                                " 		On AirSegTo.AirTicketTransactionID = AirTicketTransactions.Id  " &
                                " 			AND AirSegTo.OriginalPosition = (SELECT MAX(OriginalPosition) FROM AirSegments WHERE AirSegments.AirTicketTransactionID = AirTicketTransactions.Id)  " &
                                " 	LEFT JOIN AirTickets  " &
                                " 		ON AirTicketTransactions.AirTicketID = AirTickets.Id  " &
                                " 	ON AirTicketTransactions.CommercialTransactionID = CommercialTransactions.Id  " &
                                "   LEFT JOIN Airports AirportFrom  " &
                                " 	ON AirSegFrom.FromAirportID = AirportFrom.Id  " &
                                "   LEFT JOIN Airports AirportTo  " &
                                " 	ON AirSegTo.ToAirportID = AirportTo.Id  " &
                                "   LEFT JOIN Airlines  " &
                                " 	ON AirSegFrom.CarrierAirlineID = Airlines.Id  " &
                                "   LEFT JOIN CommercialTransactionValues CTCost  " &
                                " 	ON CTCost.CommercialTransactionID = CTV.CommercialTransactionID   " &
                                " 			  AND CTCost.IsCost=1  " &
                                "   WHERE ComTransactionTypeID = 1                   " &
                                "         AND ActionTypeID = 335                     " &
                                " 		AND AirTickets.IssueDate  >= '" & Format(FromDate, "yyyy-MM-dd") & "'     " &
                                " 		AND AirportFrom.Abbreviation = '" & Origin & "'  " &
                                " 		AND AirportTo.Abbreviation   = '" & Destination & "'  " &
                                "  GROUP BY GROUPING SETS ( " &
                                " 						 (TFEntities.Code + '/' + TFEntities.Name) " &
                                " 						,(TFEntities.Code + '/' + TFEntities.Name, Airlines.IATACode) " &
                                " 						,(Airlines.IATACode ,AirSegFrom.ActualClass) " &
                                " 						,(Airlines.IATACode) " &
                                " 						,() " &
                                " 						)" &
                                "   ORDER BY TFEntities.Code + '/' + TFEntities.Name  " &
                                " 		   ,Airlines.IATACode  " &
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
                    pobjClass = New AveragePriceItem
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

        End If

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
