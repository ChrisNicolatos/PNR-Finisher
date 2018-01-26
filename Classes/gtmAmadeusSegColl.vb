Option Strict Off
Option Explicit On
Friend Class gtmAmadeusSegColl
    Inherits Collections.Generic.Dictionary(Of String, gtmAmadeusSeg)

    Private mMaxAirportNameLength As Integer = 11
    Private mMaxCityNameLength As Integer = 11
    Private mMaxAirportShortNameLength As Integer = 11
    Friend Sub AddItem(ByRef pAirline As String, ByRef pBoardPoint As String, ByRef pClass As String, ByRef pDepartureDate As Date, ByRef pArrivalDate As Date, ByRef pElementNo As Short, ByRef pFlightNo As String, ByRef pOffPoint As String, ByRef pStatus As String, ByRef pFareBase As String, ByRef pDepartTime As Date, ByRef pArriveTime As Date, ByRef pText As String, ByVal SegDo As String)

        Dim pobjClass As gtmAmadeusSeg

        pobjClass = New gtmAmadeusSeg

        pobjClass.SetValues(pAirline, pBoardPoint, pClass, pDepartureDate, pArrivalDate, pElementNo, pFlightNo, pOffPoint, pStatus, pFareBase, pDepartTime, pArriveTime, pText, SegDo)
        MyBase.Add(Format(pElementNo), pobjClass)

        If pobjClass.BoardAirportName.Length > mMaxAirportNameLength Then
            mMaxAirportNameLength = pobjClass.BoardAirportName.Length
        End If
        If pobjClass.OffPointAirportName.Length > mMaxAirportNameLength Then
            mMaxAirportNameLength = pobjClass.OffPointAirportName.Length
        End If
        If pobjClass.BoardCityName.Length > mMaxCityNameLength Then
            mMaxCityNameLength = pobjClass.BoardCityName.Length
        End If
        If pobjClass.OffPointCityName.Length > mMaxCityNameLength Then
            mMaxCityNameLength = pobjClass.OffPointCityName.Length
        End If
        If pobjClass.BoardAirportShortName.Length > mMaxAirportShortNameLength Then
            mMaxAirportShortNameLength = pobjClass.BoardAirportShortName.Length
        End If
        If pobjClass.OffPointAirportShortName.Length > mMaxAirportShortNameLength Then
            mMaxAirportShortNameLength = pobjClass.OffPointAirportShortName.Length
        End If

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
            For Each Seg As gtmAmadeusSeg In MyBase.Values
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

End Class