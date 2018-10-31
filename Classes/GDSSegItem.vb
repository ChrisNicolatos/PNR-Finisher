Public Class GDSSegItem

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
        Dim BoardCountryCode As String
        Dim OffPoint As String
        Dim OffPointAirportName As String
        Dim OffPointCityName As String
        Dim OffPointAirportShortName As String
        Dim OffPointCountryName As String
        Dim OffPointCountryCode As String
        Dim DepartTime As Date
        Dim ArriveTime As Date
        Dim EstimatedFlyingTime As String
        Dim AirlineLocator As String
        Dim Text As String
        Dim OperatedBy As String
        Dim Stopovers As String
        Dim DepartTerminal As String
        Dim ArriveTerminal As String
        Dim Status As String
        Dim Equipment As String
    End Structure

    Private mudtProps As ClassProps
    Private mobjAirlineDate As New s1aAirlineDate.clsAirlineDate
    Public ReadOnly Property ElementNo As Short
        Get
            Return mudtProps.ElementNo
        End Get
    End Property
    Public ReadOnly Property Airline() As String
        Get
            Return mudtProps.Airline.Trim
        End Get
    End Property
    Public ReadOnly Property AirlineName As String
        Get
            Return mudtProps.AirlineName.Trim
        End Get
    End Property
    Public ReadOnly Property BoardPoint() As String
        Get
            Return mudtProps.BoardPoint.Trim
        End Get
    End Property
    Public ReadOnly Property BoardAirportName() As String
        Get
            Return mudtProps.BoardAirportName.Trim
        End Get
    End Property
    Public ReadOnly Property BoardCityName As String
        Get
            Return mudtProps.BoardCityName.Trim
        End Get
    End Property
    Public ReadOnly Property BoardAirportShortName As String
        Get
            If mudtProps.BoardAirportShortName = "" Then
                Return mudtProps.BoardCityName.Trim
            Else
                Return mudtProps.BoardAirportShortName.Trim
            End If
        End Get
    End Property
    Public ReadOnly Property BoardCountryName As String
        Get
            Return mudtProps.BoardCountryName
        End Get
    End Property
    Public ReadOnly Property OffPointAirportName() As String
        Get
            Return mudtProps.OffPointAirportName.Trim
        End Get
    End Property
    Public ReadOnly Property OffPointCityName As String
        Get
            Return mudtProps.OffPointCityName.Trim
        End Get
    End Property
    Public ReadOnly Property OffPointAirportShortName As String
        Get
            If mudtProps.OffPointAirportShortName = "" Then
                Return mudtProps.OffPointCityName.Trim
            Else
                Return mudtProps.OffPointAirportShortName.Trim
            End If
        End Get
    End Property
    Public ReadOnly Property OffPointCountryName As String
        Get
            Return mudtProps.OffPointCountryName
        End Get
    End Property
    Public ReadOnly Property Status As String
        Get
            Return mudtProps.Status.Trim
        End Get
    End Property
    Public ReadOnly Property ClassOfService() As String
        Get
            Return mudtProps.ClassOfService.Trim
        End Get
    End Property
    Public ReadOnly Property DepartureDate() As Date
        Get
            Return mudtProps.DepartureDate
        End Get
    End Property
    Public ReadOnly Property DepartureDateIATA As String
        Get
            Return mudtProps.DepartureDateIATA.Trim
        End Get
    End Property
    Public ReadOnly Property ArrivalDate As Date
        Get
            Return mudtProps.ArrivalDate
        End Get
    End Property
    Public ReadOnly Property ArrivalDateIATA As String
        Get
            Return mudtProps.ArrivalDateIATA.Trim
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
    Public ReadOnly Property Equipment As String
        Get
            Return mudtProps.Equipment.Trim
        End Get
    End Property
    Public ReadOnly Property FlightNo() As String
        Get

            Return mudtProps.FlightNo.Trim

        End Get
    End Property
    Public ReadOnly Property OffPoint() As String
        Get

            Return mudtProps.OffPoint.Trim

        End Get
    End Property
    Public ReadOnly Property DepartTime() As Date
        Get
            Return mudtProps.DepartTime
        End Get
    End Property
    Public ReadOnly Property ArriveTime() As Date
        Get
            Return mudtProps.ArriveTime
        End Get
    End Property
    Public ReadOnly Property EstimatedFlyingTime As String
        Get
            Return mudtProps.EstimatedFlyingTime
        End Get
    End Property
    Public ReadOnly Property AirlineLocator() As String
        Get
            Return mudtProps.AirlineLocator
        End Get
    End Property

    Public ReadOnly Property Text() As String
        Get

            Text = mudtProps.Text.Trim

        End Get
    End Property

    Public ReadOnly Property Stopovers As String
        Get
            Return mudtProps.Stopovers
        End Get
    End Property
    Public ReadOnly Property DepartTerminal As String
        Get
            Return mudtProps.DepartTerminal
        End Get
    End Property
    Public ReadOnly Property OperatedBy As String
        Get
            If mudtProps.OperatedBy <> "" Then
                Return mudtProps.OperatedBy
            Else
                OperatedBy = ""
                For i = 81 To mudtProps.Text.Length Step 80
                    If (mudtProps.Text & StrDup(80, " ")).Substring(i, 80).IndexOf("OPERATED BY") >= 0 Then
                        If OperatedBy <> "" Then
                            OperatedBy &= vbCrLf
                        End If
                        OperatedBy &= (mudtProps.Text.Trim & StrDup(80, " ")).Substring(i, 80)
                    End If
                Next
            End If
        End Get
    End Property
    Friend Sub SetValues(ByVal pAirline As String _
                       , ByVal pBoardPoint As String _
                       , ByVal pClass As String _
                       , ByVal pDepartureDate As Date _
                       , ByVal pArrivalDate As Date _
                       , ByVal pElementNo As Short _
                       , ByVal pFlightNo As String _
                       , ByVal pOffPoint As String _
                       , ByVal pStatus As String _
                       , ByVal pDepartTime As Date _
                       , ByVal pArriveTime As Date _
                       , ByVal pEquipment As String _
                       , ByVal pVL() As String _
                       , ByVal pText As String _
                       , ByVal pOperatedBy As String _
                       , ByVal SVC As String())
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
            .BoardCountryCode = Airport.CountryCode(.BoardPoint)
            .OffPoint = pOffPoint
            .OffPointAirportName = Airport.CityAirportName(.OffPoint)
            .OffPointCityName = Airport.CityName(.OffPoint)
            .OffPointAirportShortName = Airport.AirportShortname(.OffPoint)
            .OffPointCountryName = Airport.CountryName(.OffPoint)
            .OffPointCountryCode = Airport.CountryCode(.OffPoint)
            .Status = pStatus
            .DepartTime = pDepartTime
            .ArriveTime = pArriveTime
            .AirlineLocator = ""
            .Equipment = pEquipment
            If pVL.GetUpperBound(0) = 1 Then
                If pVL(1).IndexOf("/") > 0 Then
                    .AirlineLocator = pVL(1).Substring(5, pVL(1).IndexOf("/") - 5)
                Else
                    .AirlineLocator = pVL(1).Substring(5)
                End If
            Else
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
            End If
            .Text = pText
            .OperatedBy = pOperatedBy
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

    Friend Sub SetValues(ByVal pAirline As String, ByVal pBoardPoint As String, ByVal pClass As String, ByVal pDepartureDate As Date, ByVal pArrivalDate As Date, ByVal pElementNo As Short, ByVal pFlightNo As String, ByVal pOffPoint As String, ByVal pStatus As String, ByVal pDepartTime As Date, ByVal pArriveTime As Date, ByVal pEquipment As String, ByVal pText As String, ByVal SegDo As String)
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
            .BoardCountryCode = Airport.CountryCode(.BoardPoint)

            .OffPoint = pOffPoint
            .OffPointAirportName = Airport.CityAirportName(.OffPoint)
            .OffPointCityName = Airport.CityName(.OffPoint)
            .OffPointAirportShortName = Airport.AirportShortname(.OffPoint)
            .OffPointCountryName = Airport.CountryName(.OffPoint)
            .OffPointCountryCode = Airport.CountryCode(.OffPoint)

            .Status = pStatus
            .DepartTime = pDepartTime
            .ArriveTime = pArriveTime
            .Equipment = pEquipment
            If pText.Length > 63 Then
                .AirlineLocator = pText.Substring(53, 10).Trim
            ElseIf pText.Length > 53 Then
                .AirlineLocator = pText.Substring(53).Trim
            Else
                .AirlineLocator = ""
            End If
            .Text = pText
            .OperatedBy = ""
            'If .Text.Substring(35).StartsWith("FLWN") Then
            '    .DepartureDate = Date.MinValue
            '    .ArrivalDate = Date.MinValue
            'End If
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
            AnalyseSegDo1A(SegDo)
        End With

    End Sub
    Private Sub AnalyseSegDo1A(ByVal SegDo As String)

        Dim pSegDo() As String = SegDo.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

        Dim pItinStarts As Integer = -1
        For i As Integer = 0 To pSegDo.GetUpperBound(0) - 1
            If pSegDo(i).IndexOf("*1A PLANNED FLIGHT INFO*") >= 0 And pSegDo(i + 1).IndexOf("APT") >= 0 Then
                pItinStarts = i + 2
                Exit For
            End If
        Next
        Dim pBoardStarts As Integer = -1
        If pItinStarts >= 0 Then
            For i As Integer = pItinStarts To pSegDo.GetUpperBound(0)
                If pSegDo(i).Length > 3 AndAlso pSegDo(i).Substring(0, 3) = mudtProps.BoardPoint Then
                    pBoardStarts = i
                    Exit For
                End If
            Next
        End If
        Dim pOffStarts As Integer = -1
        If pBoardStarts >= 0 Then
            For i As Integer = pBoardStarts + 1 To pSegDo.GetUpperBound(0)
                If pSegDo(i).Length > 3 AndAlso pSegDo(i).Substring(0, 3) = mudtProps.OffPoint Then
                    pOffStarts = i
                    Exit For
                End If
                If pSegDo(i).Length > 3 AndAlso pSegDo(i).Substring(0, 3) <> "   " Then
                    If mudtProps.Stopovers <> "" Then
                        mudtProps.Stopovers &= vbCrLf
                    End If
                    mudtProps.Stopovers &= pSegDo(i).Substring(0, 3) & "-" & Airport.CityAirportName(pSegDo(i).Substring(0, 3))
                End If
            Next
        End If
        If pOffStarts >= 0 Then
            If pSegDo(pOffStarts).Length > 63 Then
                mudtProps.EstimatedFlyingTime = pSegDo(pOffStarts).Substring(pSegDo(pOffStarts).LastIndexOf(" ")).Trim.PadLeft(5)
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
        mudtProps.Equipment = ""

        Dim pSeg As Integer = 0
        Dim pOperatedBy As String = ""
        Dim pRouting As String = ""
        Dim pBoardPoint As String = ""
        Dim pOffPoint As String = ""
        Dim pFlyingTime As Date = TimeSerial(0, 0, 0)
        For iSVC As Integer = 0 To pSVC.GetUpperBound(0)
            If pSVC(iSVC).Substring(2, 1) = " " AndAlso (IsNumeric(pSVC(iSVC).Substring(0, 2)) Or IsNumeric(pSVC(iSVC).Substring(1, 1))) Then
                If pSeg > 0 Then
                    ' add new entry to tickets
                    Dim x As String = ""
                    pSeg = 0
                    mudtProps.EstimatedFlyingTime = ""
                    pOperatedBy = ""
                    mudtProps.DepartTerminal = ""
                    mudtProps.ArriveTerminal = ""
                End If

                pSeg = CInt(pSVC(iSVC).Trim.Substring(0, pSVC(iSVC).Trim.IndexOf(" ")))
                pRouting = pSVC(iSVC).Substring(14, 3) & " " & pSVC(iSVC).Substring(3, 2) & " " & pSVC(iSVC).Substring(17, 3)
                mudtProps.EstimatedFlyingTime = pSVC(iSVC).Trim.Substring(pSVC(iSVC).Trim.LastIndexOf(" ") + 1).PadLeft(5, "0"c)
                pBoardPoint = pSVC(iSVC).Substring(14, 3)
                pOffPoint = pSVC(iSVC).Substring(17, 3)
                pFlyingTime = TimeSerial(CInt(mudtProps.EstimatedFlyingTime.Substring(0, mudtProps.EstimatedFlyingTime.IndexOf(":"))), CInt(mudtProps.EstimatedFlyingTime.Substring(mudtProps.EstimatedFlyingTime.IndexOf(":") + 1, 2)), 0)
                mudtProps.Equipment = pSVC(iSVC).Substring(21, 5).Trim
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
                    Dim pTime As String = pSVC(iSVC).Trim.Substring(pSVC(iSVC).Trim.LastIndexOf(" ") + 1).PadLeft(5, "0"c)
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