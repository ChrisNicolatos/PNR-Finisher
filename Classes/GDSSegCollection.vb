Public Class GDSSegCollection
    Inherits Collections.Generic.Dictionary(Of String, GDSSegItem)

    Private mMaxAirportNameLength As Integer = 11
    Private mMaxCityNameLength As Integer = 11
    Private mMaxAirportShortNameLength As Integer = 11
    Private mAirlineAlert As String = ""
    Private mAirlineAlerts As New AlertsCollection
    Private mAmadeusQueue As String
    Private mGalileoQueue As String
    Public Overloads Sub Clear()
        MyBase.Clear()
        mMaxAirportNameLength = 11
        mMaxCityNameLength = 11
        mMaxAirportShortNameLength = 11
        mAirlineAlert = ""
        mAirlineAlerts = New AlertsCollection
        mAmadeusQueue = ""
        mGalileoQueue = ""
    End Sub
    Friend Function AddItem(ByVal pAirline As String, ByVal pBoardPoint As String, ByVal pClass As String, ByVal pDepartureDate As Date, ByVal pArrivalDate As Date, ByVal pElementNo As Short, ByVal pFlightNo As String, ByVal pOffPoint As String, ByVal pStatus As String, ByVal pDepartTime As Date, ByVal pArriveTime As Date, ByVal pEquipment As String, ByVal pText As String, ByVal SegDo As String) As GDSSegItem

        Dim pobjClass As GDSSegItem

        pobjClass = New GDSSegItem

        pobjClass.SetValues(pAirline, pBoardPoint, pClass, pDepartureDate, pArrivalDate, pElementNo, pFlightNo, pOffPoint, pStatus, pDepartTime, pArriveTime, pEquipment, pText, SegDo)
        MyBase.Add(Format(pElementNo), pobjClass)

        SetItinValues(pobjClass)

        Return pobjClass

    End Function
    Friend Function AddItem(ByVal pAirline As String, ByVal pBoardPoint As String, ByVal pClass As String, ByVal pDepartureDate As Date, ByVal pArrivalDate As Date, ByVal pElementNo As Short, ByVal pFlightNo As String, ByVal pOffPoint As String, ByVal pStatus As String, ByVal pDepartTime As Date, ByVal pArriveTime As Date, ByVal pEquipment As String, ByVal pVL() As String, ByVal pText As String, ByVal pOperatedBy As String, ByVal SVC() As String) As GDSSegItem

        Dim pobjClass As GDSSegItem

        pobjClass = New GDSSegItem

        pobjClass.SetValues(pAirline, pBoardPoint, pClass, pDepartureDate, pArrivalDate, pElementNo, pFlightNo, pOffPoint, pStatus, pDepartTime, pArriveTime, pEquipment, pVL, pText, pOperatedBy, SVC)
        MyBase.Add(Format(pElementNo), pobjClass)

        SetItinValues(pobjClass)

        Return pobjClass

    End Function
    Private Sub SetItinValues(ByVal pobjClass As GDSSegItem)

        mMaxAirportNameLength = Math.Max(pobjClass.BoardAirportName.Length, mMaxAirportNameLength)
        mMaxAirportNameLength = Math.Max(pobjClass.OffPointAirportName.Length, mMaxAirportNameLength)
        mMaxCityNameLength = Math.Max(pobjClass.BoardCityName.Length, mMaxCityNameLength)
        mMaxCityNameLength = Math.Max(pobjClass.OffPointCityName.Length, mMaxCityNameLength)
        mMaxAirportShortNameLength = 11 ' Math.Max(pobjClass.BoardAirportShortName.Length, mMaxAirportShortNameLength)
        mMaxAirportShortNameLength = 11 ' Math.Max(pobjClass.OffPointAirportShortName.Length, mMaxAirportShortNameLength)

        Dim pAirlineAlert As String = ""
        pAirlineAlert = mAirlineAlerts.AirlineAlert(pobjClass.Airline)
        If pAirlineAlert <> "" And mAirlineAlert.IndexOf(pAirlineAlert) = -1 Then
            mAirlineAlert &= pAirlineAlert & vbCrLf
        End If
        mAmadeusQueue = mAirlineAlerts.AmadeusQueue(pobjClass.Airline)
        mGalileoQueue = mAirlineAlerts.GalileoQueue(pobjClass.Airline)

    End Sub
    Friend ReadOnly Property MaxAirportNameLength As Integer
        Get
            MaxAirportNameLength = mMaxAirportNameLength
        End Get
    End Property
    Friend ReadOnly Property MaxCityNameLength As Integer
        Get
            MaxCityNameLength = mMaxCityNameLength
        End Get
    End Property
    Friend ReadOnly Property MaxAirportShortNameLength As Integer
        Get
            MaxAirportShortNameLength = mMaxAirportShortNameLength
        End Get
    End Property
    Friend ReadOnly Property Itinerary As String
        Get
            Dim PrevOff As String = ""
            Itinerary = ""
            For Each Seg As GDSSegItem In MyBase.Values
                With Seg
                    If PrevOff = .BoardPoint Then
                        Itinerary &= " " & .Airline & " " & .OffPoint
                    Else
                        If Itinerary <> "" Then
                            Itinerary &= " *** "
                        End If
                        Itinerary &= .BoardPoint & " " & .Airline & " " & .OffPoint
                    End If
                    PrevOff = .OffPoint
                End With
            Next
        End Get
    End Property
    Friend ReadOnly Property FullItinerary As String
        Get
            Dim pItin As String = ""
            Dim pClasses As String = ""
            Dim pDates As String = ""
            Dim pFlights As String = ""
            Dim PrevOff As String = ""

            For Each Seg As GDSSegItem In MyBase.Values
                With Seg
                    If PrevOff = .BoardPoint Then
                        pItin &= " " & .Airline & " " & .OffPoint
                    Else
                        If pItin <> "" Then
                            pItin &= " *** "
                        End If
                        pItin &= .BoardPoint & " " & .Airline & " " & .OffPoint
                    End If
                    PrevOff = .OffPoint
                    pClasses &= .ClassOfService
                    pDates &= .DepartureDateIATA & " "
                    pFlights &= .Airline & .FlightNo & " "
                End With
            Next
            FullItinerary = pItin.Trim & "|" & pClasses.Trim & "|" & pDates.Trim & "|" & pFlights.Trim
        End Get
    End Property
    Friend ReadOnly Property AirlineAlert As String
        Get
            Return mAirlineAlert
        End Get
    End Property
    Friend ReadOnly Property AmadeusQueue As String
        Get
            Return mAmadeusQueue
        End Get
    End Property
    Friend ReadOnly Property GalileoQueue As String
        Get
            Return mGalileoQueue
        End Get
    End Property
End Class