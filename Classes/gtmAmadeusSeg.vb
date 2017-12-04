Option Strict Off
Option Explicit On
Friend Class gtmAmadeusSeg

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
		Dim OffPoint As String
        Dim OffPointAirportName As String
        Dim OffPointCityName As String
        Dim OffPointAirportShortName As String
        Dim FareBase As String
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
    Private mobjAirlines As New gtmAirlines
    Private mobjCityName As New gtmAirports

	Public ReadOnly Property Airline() As String
		Get
			
			Airline = Trim(mudtProps.Airline)
			
		End Get
	End Property
    'Public ReadOnly Property AirlineName() As String
    '	Get

    '		AirlineName = mudtProps.AirlineName

    '	End Get
    'End Property
    Public ReadOnly Property BoardPoint() As String
		Get
			
			BoardPoint = Trim(mudtProps.BoardPoint)
			
		End Get
	End Property
    Public ReadOnly Property BoardAirportName() As String
        Get

            BoardAirportName = mudtProps.BoardAirportName

        End Get
    End Property
    Public ReadOnly Property BoardCityName As String
        Get
            BoardCityName = mudtProps.BoardCityName
        End Get
    End Property
    Public ReadOnly Property BoardAirportShortName As String
        Get
            If mudtProps.BoardAirportShortName = "" Then
                BoardAirportShortName = mudtProps.BoardCityName
            Else
                BoardAirportShortName = mudtProps.BoardAirportShortName
            End If
        End Get
    End Property
    Public ReadOnly Property OffPointAirportName() As String
        Get
            OffPointAirportName = mudtProps.OffPointAirportName
        End Get
    End Property
    Public ReadOnly Property OffPointCityName As String
        Get
            OffPointCityName = mudtProps.OffPointCityName
        End Get
    End Property
    Public ReadOnly Property OffPointAirportShortName As String
        Get
            If mudtProps.OffPointAirportShortName = "" Then
                OffPointAirportShortName = mudtProps.OffPointCityName
            Else
                OffPointAirportShortName = mudtProps.OffPointAirportShortName
            End If
        End Get
    End Property
    Public ReadOnly Property Status As String
        Get
            Status = mudtProps.Status
        End Get
    End Property
    Public ReadOnly Property ClassOfService() As String
        Get

            ClassOfService = Trim(mudtProps.ClassOfService)

        End Get
    End Property
	Public ReadOnly Property DepartureDate() As Date
		Get
			
			DepartureDate = mudtProps.DepartureDate
			
		End Get
    End Property
    Public ReadOnly Property DepartureDateIATA As String
        Get
            DepartureDateIATA = mudtProps.DepartureDateIATA
        End Get
    End Property
    Public ReadOnly Property ArrivalDate As Date
        Get
            ArrivalDate = mudtProps.ArrivalDate
        End Get
    End Property
    Public ReadOnly Property ArrivalDateIATA As String
        Get
            ArrivalDateIATA = mudtProps.ArrivalDateIATA
        End Get
    End Property
    'Public ReadOnly Property DepartureDay() As String

    '    Get
    '        DepartureDay = WeekDaySeg(mudtProps.DepartureDate)
    '    End Get

    'End Property
    'Public ReadOnly Property ArrivalDay As String
    '    Get
    '        ArrivalDay = WeekDaySeg(mudtProps.ArrivalDate)
    '    End Get
    'End Property
    'Private Function WeekDaySeg(ByVal InputDate As Date) As String

    '    Try
    '        Select Case Weekday(InputDate)
    '            Case 1
    '                WeekDaySeg = "Sunday"
    '            Case 2
    '                WeekDaySeg = "Monday"
    '            Case 3
    '                WeekDaySeg = "Tuesday"
    '            Case 4
    '                WeekDaySeg = "Wednesday"
    '            Case 5
    '                WeekDaySeg = "Thursday"
    '            Case 6
    '                WeekDaySeg = "Friday"
    '            Case 7
    '                WeekDaySeg = "Saturday"
    '            Case Else
    '                WeekDaySeg = ""
    '        End Select
    '    Catch ex As Exception
    '        WeekDaySeg = ""
    '    End Try

    'End Function
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
    Friend Sub SetValues(ByRef pAirline As String, ByRef pBoardPoint As String, ByRef pClass As String, ByRef pDepartureDate As Date, ByRef pArrivalDate As Date, ByRef pElementNo As Short, ByRef pFlightNo As String, ByRef pOffPoint As String, ByVal pStatus As String, ByRef pFareBase As String, ByRef pDepartTime As Date, ByRef pArriveTime As Date, ByRef pText As String, ByVal SegDo As String)

        With mudtProps
            .ElementNo = pElementNo
            .Airline = pAirline
            .AirlineName = mobjAirlines.AirlineName(.Airline)
            .FlightNo = pFlightNo
            .ClassOfService = pClass
            .DepartureDate = pDepartureDate
            .ArrivalDate = pArrivalDate
            .BoardPoint = pBoardPoint
            .BoardAirportName = mobjCityName.CityAirportName(.BoardPoint)
            .BoardCityName = mobjCityName.CityName(.BoardPoint)
            .BoardAirportShortName = mobjCityName.AirportShortname(.BoardPoint)
            .OffPoint = pOffPoint
            .OffPointAirportName = mobjCityName.CityAirportName(.OffPoint)
            .OffPointCityName = mobjCityName.CityName(.OffPoint)
            .OffPointAirportShortName = mobjCityName.AirportShortname(.OffPoint)
            .Status = pStatus
            .FareBase = pFareBase
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
            AnalyseSegDo(SegDo)
        End With

    End Sub

    Private Sub AnalyseSegDo(ByVal SegDo As String)

        Dim pSegDo() As String = SegDo.Split(vbCrLf)

        mudtProps.Stopovers = ""
        mudtProps.ArriveTerminal = ""
        mudtProps.DepartTerminal = ""
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
                If pSegDo(i).Substring(1, 3) = mudtProps.BoardPoint Then
                    pBoardStarts = i
                    Exit For
                End If
            Next
        End If
        Dim pOffStarts As Integer = -1
        If pBoardStarts >= 0 Then
            For i As Integer = pBoardStarts + 1 To pSegDo.GetUpperBound(0)
                If pSegDo(i).Substring(1, 3) = mudtProps.OffPoint Then
                    pOffStarts = i
                    Exit For
                End If
                If pSegDo(i).Substring(1, 3) <> "   " Then
                    If mudtProps.Stopovers <> "" Then
                        mudtProps.Stopovers &= vbCrLf
                    End If
                    mudtProps.Stopovers &= pSegDo(i).Substring(1, 3) & "-" & mobjCityName.CityAirportName(pSegDo(i).Substring(1, 3))
                End If
            Next
        End If
        If pOffStarts >= 0 Then
            mudtProps.EstimatedFlyingTime = pSegDo(pOffStarts).Substring(60, 5)
            For i As Integer = pOffStarts To pSegDo.GetUpperBound(0)
                If pSegDo(i).IndexOf("- DEPARTS") > 0 Then
                    mudtProps.DepartTerminal = pSegDo(i).Substring(pSegDo(i).IndexOf("- DEPARTS") + 2)
                ElseIf pSegDo(i).IndexOf("- ARRIVES") > 0 Then
                    mudtProps.ArriveTerminal = pSegDo(i).Substring(pSegDo(i).IndexOf("- ARRIVES") + 2)
                End If
            Next
        End If
    End Sub
End Class