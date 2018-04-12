Option Strict Off
Option Explicit On
Namespace GDSSeg

    Friend Class GDSSegItem

        Public Event Valid(ByRef Status As Boolean)

        Private Structure ClassProps
            Dim ElementNo As Short
            Dim Airline As String
            Dim AirlineName As String
            Dim FlightNo As String
            Dim ClassOfService As String
            Dim DepartureDate As Date
            Dim DepartureDateIATA As String
            Dim ArrivalDate As Date
            Dim ArrivalDateIATA As String
            Dim BoardPoint As String
            Dim BoardAirportName As String
            Dim BoardCityName As String
            Dim BoardAirportShortName As String
            Dim BoardCountryName As String
            Dim OffPoint As String
            Dim OffPointAirportName As String
            Dim OffPointCityName As String
            Dim OffPointAirportShortName As String
            Dim offPointCountryName As String
            Dim DepartTime As Date
            Dim ArriveTime As Date
            Dim EstimatedFlyingTime As String
            Dim AirlineLocator As String
            Dim Text As String
            Dim Stopovers As String
            Dim DepartTerminal As String
            Dim ArriveTerminal As String
            Dim Status As String
        End Structure

        Private mudtProps As ClassProps
        Private mobjAirlineDate As New s1aAirlineDate.clsAirlineDate
        Public ReadOnly Property ElementNo As Short
            Get
                ElementNo = mudtProps.ElementNo
            End Get
        End Property
        Public ReadOnly Property Airline() As String
            Get

                Airline = mudtProps.Airline.Trim

            End Get
        End Property
        Public ReadOnly Property AirlineName As String
            Get
                AirlineName = mudtProps.AirlineName.Trim
            End Get
        End Property
        Public ReadOnly Property BoardPoint() As String
            Get

                BoardPoint = mudtProps.BoardPoint.Trim

            End Get
        End Property
        Public ReadOnly Property BoardAirportName() As String
            Get

                BoardAirportName = mudtProps.BoardAirportName.Trim

            End Get
        End Property
        Public ReadOnly Property BoardCityName As String
            Get
                BoardCityName = mudtProps.BoardCityName.Trim
            End Get
        End Property
        Public ReadOnly Property BoardAirportShortName As String
            Get
                If mudtProps.BoardAirportShortName = "" Then
                    BoardAirportShortName = mudtProps.BoardCityName.Trim
                Else
                    BoardAirportShortName = mudtProps.BoardAirportShortName.Trim
                End If
            End Get
        End Property
        Public ReadOnly Property BoardCountryName As String
            Get
                BoardCountryName = mudtProps.BoardCountryName
            End Get
        End Property
        Public ReadOnly Property OffPointAirportName() As String
            Get
                OffPointAirportName = mudtProps.OffPointAirportName.Trim
            End Get
        End Property
        Public ReadOnly Property OffPointCityName As String
            Get
                OffPointCityName = mudtProps.OffPointCityName.Trim
            End Get
        End Property
        Public ReadOnly Property OffPointAirportShortName As String
            Get
                If mudtProps.OffPointAirportShortName = "" Then
                    OffPointAirportShortName = mudtProps.OffPointCityName.Trim
                Else
                    OffPointAirportShortName = mudtProps.OffPointAirportShortName.Trim
                End If
            End Get
        End Property
        Public ReadOnly Property OffPointCountryName As String
            Get
                OffPointCountryName = mudtProps.offPointCountryName
            End Get
        End Property
        Public ReadOnly Property Status As String
            Get
                Status = mudtProps.Status.Trim
            End Get
        End Property
        Public ReadOnly Property ClassOfService() As String
            Get

                ClassOfService = mudtProps.ClassOfService.Trim

            End Get
        End Property
        Public ReadOnly Property DepartureDate() As Date
            Get

                DepartureDate = mudtProps.DepartureDate

            End Get
        End Property
        Public ReadOnly Property DepartureDateIATA As String
            Get
                DepartureDateIATA = mudtProps.DepartureDateIATA.Trim
            End Get
        End Property
        Public ReadOnly Property ArrivalDate As Date
            Get
                ArrivalDate = mudtProps.ArrivalDate
            End Get
        End Property
        Public ReadOnly Property ArrivalDateIATA As String
            Get
                ArrivalDateIATA = mudtProps.ArrivalDateIATA.Trim
            End Get
        End Property
        Public ReadOnly Property DepartureDay() As String

            Get
                DepartureDay = WeekDaySeg(mudtProps.DepartureDate)
            End Get

        End Property
        Public ReadOnly Property ArrivalDay As String
            Get
                ArrivalDay = WeekDaySeg(mudtProps.ArrivalDate)
            End Get
        End Property
        Private Function WeekDaySeg(ByVal InputDate As Date) As String

            Try
                Select Case Weekday(InputDate)
                    Case 1
                        WeekDaySeg = "Sunday"
                    Case 2
                        WeekDaySeg = "Monday"
                    Case 3
                        WeekDaySeg = "Tuesday"
                    Case 4
                        WeekDaySeg = "Wednesday"
                    Case 5
                        WeekDaySeg = "Thursday"
                    Case 6
                        WeekDaySeg = "Friday"
                    Case 7
                        WeekDaySeg = "Saturday"
                    Case Else
                        WeekDaySeg = ""
                End Select
            Catch ex As Exception
                WeekDaySeg = ""
            End Try

        End Function
        '   Public ReadOnly Property ElementNo() As Short
        '	Get

        '		ElementNo = mudtProps.ElementNo

        '	End Get
        'End Property
        Public ReadOnly Property FlightNo() As String
            Get

                FlightNo = Trim(mudtProps.FlightNo)

            End Get
        End Property
        Public ReadOnly Property OffPoint() As String
            Get

                OffPoint = Trim(mudtProps.OffPoint)

            End Get
        End Property
        'Public Property FareBase() As String
        '	Get

        '		FareBase = Trim(mudtProps.FareBase)

        '	End Get
        '	Set(ByVal Value As String)

        '		mudtProps.FareBase = Value

        '	End Set
        'End Property
        Public ReadOnly Property DepartTime() As Date
            Get

                DepartTime = mudtProps.DepartTime

            End Get
        End Property
        Public ReadOnly Property ArriveTime() As Date
            Get

                ArriveTime = mudtProps.ArriveTime

            End Get
        End Property
        Public ReadOnly Property EstimatedFlyingTime As String
            Get
                EstimatedFlyingTime = mudtProps.EstimatedFlyingTime
            End Get
        End Property
        Public ReadOnly Property AirlineLocator() As String
            Get

                AirlineLocator = mudtProps.AirlineLocator

            End Get
        End Property

        Public ReadOnly Property Text() As String
            Get

                Text = Trim(mudtProps.Text)

            End Get
        End Property

        Public ReadOnly Property Stopovers As String
            Get
                Stopovers = mudtProps.Stopovers
            End Get
        End Property

        'Public ReadOnly Property ArriveTerminal As String
        '    Get
        '        ArriveTerminal = mudtProps.ArriveTerminal
        '    End Get
        'End Property

        Public ReadOnly Property DepartTerminal As String
            Get
                DepartTerminal = mudtProps.DepartTerminal
            End Get
        End Property
        Public ReadOnly Property OperatedBy As String
            Get
                OperatedBy = ""
                For i = 81 To mudtProps.Text.Length Step 80
                    If (mudtProps.Text & StrDup(80, " ")).Substring(i, 80).IndexOf("OPERATED BY") >= 0 Then
                        If OperatedBy <> "" Then
                            OperatedBy &= vbCrLf
                        End If
                        OperatedBy &= (mudtProps.Text.Trim & StrDup(80, " ")).Substring(i, 80)
                    End If
                Next
            End Get
        End Property
        Friend Sub SetValues(ByVal pAirline As String, ByVal pBoardPoint As String, ByVal pClass As String, ByVal pDepartureDate As Date, ByVal pArrivalDate As Date, ByVal pElementNo As Short, ByVal pFlightNo As String, ByVal pOffPoint As String, ByVal pStatus As String, ByVal pDepartTime As Date, ByVal pArriveTime As Date, ByVal pVL() As String, ByVal pText As String, ByVal SVC As String())
            ' Galileo
            With mudtProps
                .ElementNo = pElementNo
                .Airline = pAirline
                .AirlineName = Airlines.AirlineName(.Airline)
                .FlightNo = pFlightNo
                .ClassOfService = pClass
                .DepartureDate = pDepartureDate
                .ArrivalDate = pArrivalDate
                .BoardPoint = pBoardPoint
                .BoardAirportName = Airport.CityAirportName(.BoardPoint)
                .BoardCityName = Airport.CityName(.BoardPoint)
                .BoardAirportShortName = Airport.AirportShortname(.BoardPoint)
                .BoardCountryName = Airport.CountryName(.BoardPoint)
                .OffPoint = pOffPoint
                .OffPointAirportName = Airport.CityAirportName(.OffPoint)
                .OffPointCityName = Airport.CityName(.OffPoint)
                .OffPointAirportShortName = Airport.AirportShortname(.OffPoint)
                .offPointCountryName = Airport.CountryName(.OffPoint)
                .Status = pStatus
                .DepartTime = pDepartTime
                .ArriveTime = pArriveTime
                .AirlineLocator = ""
                For iVL As Integer = 1 To pVL.GetUpperBound(0)
                    If pVL(iVL).Substring(5, 2) = .Airline Then
                        If pVL(iVL).IndexOf("/") > 0 Then
                            .AirlineLocator = pVL(iVL).Substring(5, pVL(iVL).IndexOf("/") - 5)
                        Else
                            .AirlineLocator = pVL(iVL).Substring(5)
                        End If
                    End If
                Next
                If .AirlineLocator = "" Then
                    For iVL As Integer = 1 To pVL.GetUpperBound(0)
                        If pVL(iVL).Substring(5, 2) = "1A" Then
                            If pVL(iVL).IndexOf("/") > 0 Then
                                .AirlineLocator = pVL(iVL).Substring(5, pVL(iVL).IndexOf("/") - 5)
                            Else
                                .AirlineLocator = pVL(iVL).Substring(5)
                            End If
                        End If
                    Next

                End If
                .Text = pText
                Try
                    mobjAirlineDate.IgnoreAmadeusRange = True
                    mobjAirlineDate.VBDate = .DepartureDate
                Catch ex As Exception
                    mobjAirlineDate.VBDate = DateAdd(DateInterval.Year, -1, .DepartureDate)
                End Try
                .DepartureDateIATA = mobjAirlineDate.IATA

                Try
                    mobjAirlineDate.IgnoreAmadeusRange = True
                    mobjAirlineDate.VBDate = .ArrivalDate
                Catch ex As Exception
                    mobjAirlineDate.VBDate = DateAdd(DateInterval.Year, -1, .ArrivalDate)
                End Try
                .ArrivalDateIATA = mobjAirlineDate.IATA
                mudtProps.Stopovers = ""
                mudtProps.ArriveTerminal = ""
                mudtProps.DepartTerminal = ""
                mudtProps.EstimatedFlyingTime = ""
                AnalyzeSVC1G(SVC)
            End With
        End Sub

        Friend Sub SetValues(ByVal pAirline As String, ByVal pBoardPoint As String, ByVal pClass As String, ByVal pDepartureDate As Date, ByVal pArrivalDate As Date, ByVal pElementNo As Short, ByVal pFlightNo As String, ByVal pOffPoint As String, ByVal pStatus As String, ByVal pDepartTime As Date, ByVal pArriveTime As Date, ByVal pText As String, ByVal SegDo As String)
            ' Amadeus
            With mudtProps
                .ElementNo = pElementNo
                .Airline = pAirline
                .AirlineName = Airlines.AirlineName(.Airline)
                .FlightNo = pFlightNo
                .ClassOfService = pClass
                .DepartureDate = pDepartureDate
                .ArrivalDate = pArrivalDate
                .BoardPoint = pBoardPoint
                .BoardAirportName = Airport.CityAirportName(.BoardPoint)
                .BoardCityName = Airport.CityName(.BoardPoint)
                .BoardAirportShortName = Airport.AirportShortname(.BoardPoint)
                .BoardCountryName = Airport.CountryName(.BoardPoint)
                .OffPoint = pOffPoint
                .OffPointAirportName = Airport.CityAirportName(.OffPoint)
                .OffPointCityName = Airport.CityName(.OffPoint)
                .OffPointAirportShortName = Airport.AirportShortname(.OffPoint)
                .offPointCountryName = Airport.CountryName(.OffPoint)
                .Status = pStatus
                .DepartTime = pDepartTime
                .ArriveTime = pArriveTime
                If pText.Length > 63 Then ' Len(pText) >= 60 And Mid(pText, 53, 1) = " " Then
                    .AirlineLocator = pText.Substring(53, 10).Trim '  Trim(Mid(pText, 54, 9))
                ElseIf pText.Length > 53 Then
                    .AirlineLocator = pText.Substring(53).Trim
                Else
                    .AirlineLocator = ""
                End If
                .Text = pText
                Try
                    mobjAirlineDate.IgnoreAmadeusRange = True
                    mobjAirlineDate.VBDate = .DepartureDate
                Catch ex As Exception
                    mobjAirlineDate.VBDate = DateAdd(DateInterval.Year, -1, .DepartureDate)
                End Try
                .DepartureDateIATA = mobjAirlineDate.IATA

                Try
                    mobjAirlineDate.IgnoreAmadeusRange = True
                    mobjAirlineDate.VBDate = .ArrivalDate
                Catch ex As Exception
                    mobjAirlineDate.VBDate = DateAdd(DateInterval.Year, -1, .ArrivalDate)
                End Try
                .ArrivalDateIATA = mobjAirlineDate.IATA
                mudtProps.Stopovers = ""
                mudtProps.ArriveTerminal = ""
                mudtProps.DepartTerminal = ""
                mudtProps.EstimatedFlyingTime = ""
                AnalyseSegDo(SegDo)
            End With

        End Sub
        Private Sub AnalyseSegDo(ByVal SegDo As String)

            Dim pSegDo() As String = SegDo.Split(vbCrLf)

            Dim pItinStarts As Integer = -1
            For i As Integer = 0 To pSegDo.GetUpperBound(0) - 1
                If pSegDo(i).IndexOf("*1A PLANNED FLIGHT INFO*") = 1 And pSegDo(i + 1).IndexOf("APT") = 1 Then
                    pItinStarts = i + 2
                    Exit For
                End If
            Next
            Dim pBoardStarts As Integer = -1
            If pItinStarts >= 0 Then
                For i As Integer = pItinStarts To pSegDo.GetUpperBound(0)
                    If pSegDo(i).Length > 3 AndAlso pSegDo(i).Substring(1, 3) = mudtProps.BoardPoint Then
                        pBoardStarts = i
                        Exit For
                    End If
                Next
            End If
            Dim pOffStarts As Integer = -1
            If pBoardStarts >= 0 Then
                For i As Integer = pBoardStarts + 1 To pSegDo.GetUpperBound(0)
                    If pSegDo(i).Length > 3 AndAlso pSegDo(i).Substring(1, 3) = mudtProps.OffPoint Then
                        pOffStarts = i
                        Exit For
                    End If
                    If pSegDo(i).Length > 3 AndAlso pSegDo(i).Substring(1, 3) <> "   " Then
                        If mudtProps.Stopovers <> "" Then
                            mudtProps.Stopovers &= vbCrLf
                        End If
                        mudtProps.Stopovers &= pSegDo(i).Substring(1, 3) & "-" & Airport.CityAirportName(pSegDo(i).Substring(1, 3))
                    End If
                Next
            End If
            If pOffStarts >= 0 Then
                If pSegDo(pOffStarts).Length > 63 Then
                    mudtProps.EstimatedFlyingTime = pSegDo(pOffStarts).Substring(60, 5)
                Else
                    mudtProps.EstimatedFlyingTime = ""
                End If
                For i As Integer = pOffStarts To pSegDo.GetUpperBound(0)
                    If pSegDo(i).IndexOf("- DEPARTS") > 0 Then
                        mudtProps.DepartTerminal = pSegDo(i).Substring(pSegDo(i).IndexOf("- DEPARTS") + 2)
                    ElseIf pSegDo(i).IndexOf("- ARRIVES") > 0 Then
                        mudtProps.ArriveTerminal = pSegDo(i).Substring(pSegDo(i).IndexOf("- ARRIVES") + 2)
                    End If
                Next
            End If
        End Sub
        Private Sub AnalyzeSVC1G(ByVal pSVC() As String)
            '
            ' *SVC for specific FF entry
            ' 
            mudtProps.Stopovers = ""
            mudtProps.ArriveTerminal = ""
            mudtProps.DepartTerminal = ""
            mudtProps.EstimatedFlyingTime = ""

            Dim pSeg As Integer = 0
            Dim pOperatedBy As String = ""
            Dim pRouting As String = ""
            Dim pBoardPoint As String = ""
            Dim pOffPoint As String = ""
            Dim pFlyingTime As Date = TimeSerial(0, 0, 0)
            For iSVC As Integer = 0 To pSVC.GetUpperBound(0)
                If IsNumeric(pSVC(iSVC).Trim.Substring(0, 1)) Then
                    If pSeg > 0 Then
                        ' add new entry to tickets
                        Dim x As String = ""
                        pSeg = 0
                        mudtProps.EstimatedFlyingTime = ""
                        pOperatedBy = ""
                        mudtProps.DepartTerminal = ""
                        mudtProps.ArriveTerminal = ""
                    End If
                    pSeg = pSVC(iSVC).Trim.Substring(0, pSVC(iSVC).Trim.IndexOf(" "))
                    pRouting = pSVC(iSVC).Substring(14, 3) & " " & pSVC(iSVC).Substring(3, 2) & " " & pSVC(iSVC).Substring(17, 3)
                    mudtProps.EstimatedFlyingTime = pSVC(iSVC).Trim.Substring(pSVC(iSVC).Trim.LastIndexOf(" ") + 1).PadLeft(5, "0")
                    pBoardPoint = pSVC(iSVC).Substring(14, 3)
                    pOffPoint = pSVC(iSVC).Substring(17, 3)
                    pFlyingTime = TimeSerial(mudtProps.EstimatedFlyingTime.Substring(0, mudtProps.EstimatedFlyingTime.IndexOf(":")), mudtProps.EstimatedFlyingTime.Substring(mudtProps.EstimatedFlyingTime.IndexOf(":") + 1, 2), 0)
                ElseIf pSeg > 0 Then
                    If pSVC(iSVC).IndexOf("OPERATED BY") >= 0 Then
                        pOperatedBy = pSVC(iSVC).Trim
                    ElseIf pSVC(iSVC).StartsWith(Space(14)) And pSVC(iSVC).Substring(14, 6).Replace(" ", "").Length = 6 And pSVC(iSVC).Substring(20, 1) = Space(1) Then
                        If mudtProps.Stopovers <> "" Then
                            mudtProps.Stopovers &= vbCrLf
                        End If
                        mudtProps.Stopovers &= pOffPoint & "-" & Airport.CityAirportName(pOffPoint)
                        pBoardPoint = pSVC(iSVC).Substring(14, 3)
                        pOffPoint = pSVC(iSVC).Substring(17, 3)
                        Dim pTime As String = pSVC(iSVC).Trim.Substring(pSVC(iSVC).Trim.LastIndexOf(" ") + 1)
                        pFlyingTime = DateAdd(DateInterval.Hour, CDbl(pTime.Substring(0, pTime.IndexOf(":"))), pFlyingTime)
                        pFlyingTime = DateAdd(DateInterval.Minute, CDbl(pTime.Substring(pTime.IndexOf(":") + 1, 2)), pFlyingTime)
                        mudtProps.EstimatedFlyingTime = Format(pFlyingTime, "HH:mm")
                    End If
                    If pSVC(iSVC).IndexOf("DEPARTS") >= 0 Then
                        If pSVC(iSVC).IndexOf("-") > 0 Then
                            mudtProps.DepartTerminal = pSVC(iSVC).Substring(0, pSVC(iSVC).IndexOf("-")).Trim
                        Else
                            mudtProps.DepartTerminal = pSVC(iSVC).Trim
                        End If
                    End If
                    If pSVC(iSVC).IndexOf("ARRIVES") >= 0 Then
                        If pSVC(iSVC).IndexOf("-") > 0 Then
                            mudtProps.ArriveTerminal = pSVC(iSVC).Substring(pSVC(iSVC).IndexOf("-") + 1).Trim
                        Else
                            mudtProps.ArriveTerminal = pSVC(iSVC).Trim
                        End If
                    End If
                ElseIf pSVC(iSVC).Substring(0, 1) <> " " Then
                    Exit For
                End If
            Next
        End Sub
    End Class
    Friend Class GDSSegColl
        Inherits Collections.Generic.Dictionary(Of String, GDSSeg.GDSSegItem)

        Private mMaxAirportNameLength As Integer = 11
        Private mMaxCityNameLength As Integer = 11
        Private mMaxAirportShortNameLength As Integer = 11
        Friend Function AddItem(ByVal pAirline As String, ByVal pBoardPoint As String, ByVal pClass As String, ByVal pDepartureDate As Date, ByVal pArrivalDate As Date, ByVal pElementNo As Short, ByVal pFlightNo As String, ByVal pOffPoint As String, ByVal pStatus As String, ByVal pDepartTime As Date, ByVal pArriveTime As Date, ByVal pText As String, ByVal SegDo As String) As GDSSeg.GDSSegItem

            Dim pobjClass As GDSSeg.GDSSegItem

            pobjClass = New GDSSeg.GDSSegItem

            pobjClass.SetValues(pAirline, pBoardPoint, pClass, pDepartureDate, pArrivalDate, pElementNo, pFlightNo, pOffPoint, pStatus, pDepartTime, pArriveTime, pText, SegDo)
            MyBase.Add(Format(pElementNo), pobjClass)

            SetNameLengths(pobjClass)

            Return pobjClass

        End Function
        Friend Function AddItem(ByVal pAirline As String, ByVal pBoardPoint As String, ByVal pClass As String, ByVal pDepartureDate As Date, ByVal pArrivalDate As Date, ByVal pElementNo As Short, ByVal pFlightNo As String, ByVal pOffPoint As String, ByVal pStatus As String, ByVal pDepartTime As Date, ByVal pArriveTime As Date, ByVal pVL() As String, ByVal pText As String, ByVal SVC() As String) As GDSSeg.GDSSegItem

            Dim pobjClass As GDSSeg.GDSSegItem

            pobjClass = New GDSSeg.GDSSegItem

            pobjClass.SetValues(pAirline, pBoardPoint, pClass, pDepartureDate, pArrivalDate, pElementNo, pFlightNo, pOffPoint, pStatus, pDepartTime, pArriveTime, pVL, pText, SVC)
            MyBase.Add(Format(pElementNo), pobjClass)

            SetNameLengths(pobjClass)

            Return pobjClass

        End Function
        Private Sub SetNameLengths(ByVal pobjClass As GDSSeg.GDSSegItem)

            mMaxAirportNameLength = Math.Max(pobjClass.BoardAirportName.Length, mMaxAirportNameLength)
            mMaxAirportNameLength = Math.Max(pobjClass.OffPointAirportName.Length, mMaxAirportNameLength)
            mMaxCityNameLength = Math.Max(pobjClass.BoardCityName.Length, mMaxCityNameLength)
            mMaxCityNameLength = Math.Max(pobjClass.OffPointCityName.Length, mMaxCityNameLength)
            mMaxAirportShortNameLength = Math.Max(pobjClass.BoardAirportShortName.Length, mMaxAirportShortNameLength)
            mMaxAirportShortNameLength = Math.Max(pobjClass.OffPointAirportShortName.Length, mMaxAirportShortNameLength)

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
                For Each Seg As GDSSeg.GDSSegItem In MyBase.Values
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
End Namespace
