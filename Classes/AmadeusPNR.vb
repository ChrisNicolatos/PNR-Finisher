Option Strict Off
Option Explicit On
Friend Class AmadeusPNR
    Public Event ReadStatus(ByRef Status As Short, ByRef StatusDescription As String)
    Private Structure ClassProps
        Dim RequestedPNR As String
        Dim UserSignIn As String
        Dim PNRCreationdate As Date
        Dim Seats As String

        Dim isDirty As Boolean
        Dim isValid As Boolean
        Dim isNew As Boolean
    End Structure

    Private Structure TQTProps
        Dim TQTElement As Integer
        Dim Segment As Integer
        Dim Itin As String
        Dim Allowance As String
        Dim Pax As String
        Dim Status As String
    End Structure

    Private mudtProps As ClassProps
    Private mudtTQT() As TQTProps
    Private mudtAllowance() As TQTProps

    Private WithEvents mobjHostSession As k1aHostToolKit.HostSession
    Private mobjPNR As s1aPNR.PNR
    Private mobjPax As AmadeusPax.AmadeusPaxColl
    Private mobjSegs As AmadeusSeg.AmadeusSegColl
    Private mSegsLastElement As Integer
    Private mobjNumberParser As AmadeusNumberParser
    Private mstrVesselName As String
    Private mobjTickets As Ticket.TicketColl
    Private mobjFrequentFlyer As FrequentFlyer.FrequentFlyerColl
    Private mstrCC As String
    Private mstrCLN As String
    Private mstrCLA As string

    Private mflgCancelError As Boolean

    Private mintStatus As Short
    Private mstrStatus As String
    Private Sub Class_Initialize_Renamed()
        mobjPax = New AmadeusPax.AmadeusPaxColl
        mobjSegs = New AmadeusSeg.AmadeusSegColl
        mobjNumberParser = New AmadeusNumberParser
        mobjFrequentFlyer = New FrequentFlyer.FrequentFlyerColl
    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
    Public ReadOnly Property AllowanceForSegment(ByVal Origin As String, ByVal Destination As String, ByVal Airline As String) As String ', ByVal Pax As String) As String
        Get
            AllowanceForSegment = ""
            For i As Integer = 1 To mudtAllowance.GetUpperBound(0)
                If mudtAllowance(i).Itin = Origin & " " & Airline & " " & Destination Then
                    AllowanceForSegment = mudtAllowance(i).Allowance
                End If
            Next
        End Get
    End Property
    Public ReadOnly Property FrequentFlyerNumber(ByVal Airline As String, ByVal PaxName As String) As String
        Get
            FrequentFlyerNumber = ""
            For Each pItem As FrequentFlyer.FrequentFlyerItem In mobjFrequentFlyer.Values
                If pItem.Airline = Airline And pItem.PaxName = PaxName Then
                    FrequentFlyerNumber = pItem.Airline & " " & pItem.FrequentTravelerNo
                    Exit For
                End If
            Next
        End Get
    End Property
    Public ReadOnly Property VesselName() As String
        Get
            VesselName = mstrVesselName
        End Get
    End Property
    Public ReadOnly Property ClientName As String
        Get
            ClientName = mstrCLA
        End Get
    End Property
    Public ReadOnly Property ClientCode As String
        Get
            ClientCode = mstrCLN
        End Get
    End Property
    Public ReadOnly Property CostCentre As String
        Get
            CostCentre = mstrCC
        End Get
    End Property
    Public ReadOnly Property RequestedPNR() As String
        Get
            RequestedPNR = mudtProps.RequestedPNR
        End Get
    End Property
    Public ReadOnly Property Seats As String
        Get
            Seats = mudtProps.Seats
        End Get
    End Property
    Public Property CancelError() As Boolean
        Get
            CancelError = mflgCancelError
        End Get
        Set(ByVal Value As Boolean)
            mflgCancelError = Value
        End Set
    End Property

    Public ReadOnly Property Tickets() As Ticket.TicketColl
        Get
            Tickets = mobjTickets
        End Get
    End Property

    Public ReadOnly Property Segments() As AmadeusSeg.AmadeusSegColl
        Get
            Segments = mobjSegs
        End Get
    End Property
    Public ReadOnly Property HasSegments As Boolean
        Get
            HasSegments = (mSegsLastElement > -1)
        End Get
    End Property
    Public ReadOnly Property LastSegment As AmadeusSeg.AmadeusSegItem
        Get
            If mSegsLastElement = -1 Then
                LastSegment = New AmadeusSeg.AmadeusSegItem
            Else
                LastSegment = mobjSegs(Format(mSegsLastElement))
            End If
        End Get
    End Property
    Public ReadOnly Property MaxAirportNameLength As Integer
        Get
            MaxAirportNameLength = mobjSegs.MaxAirportNameLength
        End Get
    End Property
    Public ReadOnly Property MaxCityNameLength As Integer
        Get
            MaxCityNameLength = mobjSegs.MaxCityNameLength
        End Get
    End Property
    Public ReadOnly Property MaxAirportShortNameLength As Integer
        Get
            MaxAirportShortNameLength = mobjSegs.MaxAirportShortNameLength
        End Get
    End Property
    Public ReadOnly Property Passengers() As AmadeusPax.AmadeusPaxColl
        Get
            Passengers = mobjPax
        End Get
    End Property
    Public Function RetrievePNRsFromQueue(ByVal Queue As String) As String

        Dim pobjHostSessions As k1aHostToolKit.HostSessions
        Dim pQV As String = ""

        RetrievePNRsFromQueue = ""

        Try
            mstrStatus = ""
            pobjHostSessions = New k1aHostToolKit.HostSessions

            If pobjHostSessions.Count > 0 Then
                mobjHostSession = pobjHostSessions.UIActiveSession

                If Queue <> "" Then
                    mobjHostSession.Send("QI")
                    mobjHostSession.Send("IG")
                End If
                pQV &= mobjHostSession.Send("QV/" & Queue).Text
                Do While pQV.IndexOf(")>") = pQV.Length - 4
                    pQV &= mobjHostSession.Send("MDR").Text
                Loop
                Dim pLines() As String = pQV.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                Dim pPNRs As String = ""
                For i As Integer = 4 To pLines.GetUpperBound(0)
                    If pLines(i).Length >= 19 Then
                        pPNRs &= pLines(i).Substring(14, 6) & vbCrLf
                    End If
                Next
                RetrievePNRsFromQueue = pPNRs
            Else
                Throw New Exception("Amadeus not signed in")
            End If


        Catch ex As Exception
            mintStatus = 999
            mstrStatus = Err.Description
            RaiseEvent ReadStatus(mintStatus, mstrStatus)

            If CancelError Then
                Throw New Exception("RetrivePNRsFromQueue()" & vbCrLf & mstrStatus)
            End If
        End Try

    End Function
    Public Function ReadPNR(ByVal PNR As String, ByVal ForReportOnly As Boolean) As Boolean

        Dim pobjHostSessions As k1aHostToolKit.HostSessions

        Try
            mstrStatus = ""
            pobjHostSessions = New k1aHostToolKit.HostSessions

            If pobjHostSessions.Count > 0 Then
                mobjHostSession = pobjHostSessions.UIActiveSession

                If PNR <> "" Then
                    mobjHostSession.Send("QI")
                    mobjHostSession.Send("IG")
                End If
                mudtProps.RequestedPNR = PNR
                ReadPNR = RetrievePNR(ForReportOnly)
            Else
                Throw New Exception("Amadeus not signed in")
            End If

            If ReadPNR Then
                mintStatus = 0
                mstrStatus = "Amadeus read " & PNR & " OK"
            Else
                mintStatus = 1
                mstrStatus = "Amadeus " & PNR & " not found"
            End If
            mobjHostSession.SendSpecialKey(512 + 282) '(k1aHostConstantsLib.AmaKeyValues.keySHIFT + k1aHostConstantsLib.AmaKeyValues.keyPause)
            mobjHostSession.Send("RT")
            RaiseEvent ReadStatus(mintStatus, mstrStatus)
        Catch ex As Exception
            mintStatus = 999
            mstrStatus = Err.Description
            RaiseEvent ReadStatus(mintStatus, mstrStatus)

            If CancelError Then
                Throw New Exception("ReadPNR()" & vbCrLf & mstrStatus)
            End If
        End Try

    End Function

    Private Function RetrievePNR(ByVal ForReportOnly As Boolean) As Boolean

        Dim pintPNRStatus As Integer

        mobjPNR = New s1aPNR.PNR
        mobjTickets = New Ticket.TicketColl
        mstrVesselName = ""
        mstrCC = ""
        mstrCLA = ""
        mstrCLN = ""

        With mudtProps

            If .RequestedPNR = "" Then
                pintPNRStatus = mobjPNR.RetrieveCurrent(mobjHostSession)
            Else
                pintPNRStatus = mobjPNR.RetrievePNR(mobjHostSession, "RT" & .RequestedPNR)
            End If
            .PNRCreationdate = Today

            If pintPNRStatus = 0 Or pintPNRStatus = 1005 Then
                .RequestedPNR = setRecordLocator()
                If ForReportOnly Then
                    getPax()
                    getSegs(ForReportOnly)
                    getOtherServiceElements()
                    GetRMElements()
                Else
                    getTQT()
                    getPax()
                    getSegs(ForReportOnly)
                    getAutoTickets()
                    getOtherServiceElements()
                    getSSRElements()
                    GetRMElements()
                End If
                RetrievePNR = True
            Else
                RetrievePNR = False
            End If
        End With

    End Function
    Private Function setRecordLocator() As String
        Try
            setRecordLocator = mobjPNR.Header.RecordLocator
        Catch ex As Exception
            setRecordLocator = UCase(mudtProps.RequestedPNR)
        End Try
    End Function
    Private Sub getAutoTickets()

        Dim pobjFareAutoTktElement As s1aPNR.FareAutoTktElement
        Dim pobjFareOriginalIssueElement As s1aPNR.FareOriginalIssueElement

        For Each pobjFareOriginalIssueElement In mobjPNR.FareOriginalIssueElements
            parseFareOriginal(pobjFareOriginalIssueElement)
        Next pobjFareOriginalIssueElement

        For Each pobjFareAutoTktElement In mobjPNR.FareAutoTktElements
            parseFareAutoTktElement(pobjFareAutoTktElement)
        Next pobjFareAutoTktElement

    End Sub

    Private Sub parseFareOriginal(ByVal Element As s1aPNR.FareOriginalIssueElement)

        Dim i As Short
        Dim pflgIATAFound As Boolean
        Dim pstrText As String
        Dim pstrSplit1() As String
        Dim pstrSplit2() As String

        Try

            Dim SegAssociations As String = ""
            Dim PaxAssociations As String = ""

            Dim objSeg As Object
            Dim objPax As Object

            If Element.Associations.Segments.Count > 0 Then
                For Each objSeg In Element.Associations.segments
                    SegAssociations &= mobjSegs(objSeg.ElementNo).BoardPoint & " " & mobjSegs(objSeg.ElementNo).Airline & " " & mobjSegs(objSeg.ElementNo).OffPoint & vbCrLf
                Next
            Else
                For Each pSeg As AmadeusSeg.AmadeusSegItem In mobjSegs.Values
                    SegAssociations &= pSeg.BoardPoint & " " & pSeg.Airline & " " & pSeg.OffPoint & vbCrLf
                Next
            End If

            If Element.Associations.Passengers.Count > 0 Then
                For Each objPax In Element.Associations.Passengers
                    PaxAssociations &= mobjPax(objPax.ElementNo).PaxName & vbCrLf
                Next
            Else
                For Each pPax As AmadeusPax.AmadeusPaxitem In mobjPax.Values
                    PaxAssociations &= pPax.PaxName & vbCrLf
                Next
            End If

            pstrText = ConcatenateText(Element.Text)

            pstrSplit1 = Split(pstrText, " ")
            pstrSplit2 = Split(pstrText, "/")

            pflgIATAFound = False
            If IsArray(pstrSplit2) Then
                For i = LBound(pstrSplit2) To UBound(pstrSplit2)
                    If InStr(pstrSplit2(i), MySettings.IATANumber) > 0 Then
                        pflgIATAFound = True
                        Exit For
                    End If
                    If pflgIATAFound Then
                        Exit For
                    End If
                Next i
            End If

            If pflgIATAFound Then
                If IsArray(pstrSplit1) Then
                    For i = LBound(pstrSplit1) To UBound(pstrSplit1)
                        If Len(pstrSplit1(i)) >= 13 Then
                            With mobjNumberParser
                                If .TicketNumberText(pstrSplit1(i)) Then
                                    mobjTickets.addTicket("FO", .StockType, .DocumentNumber, .Books, .AirlineNumber, .AirlineNumber, False, SegAssociations, PaxAssociations, pstrSplit2(2))
                                End If
                            End With
                        End If
                    Next i
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub
    ''' <summary>
    ''' This is the new version which does not check the issuing office of the ticket
    ''' </summary>
    Private Sub parseFareAutoTktElement(ByVal Element As s1aPNR.FareAutoTktElement)

        Dim pstrSplit2() As String
        Dim pflgETicket As Boolean

        Try

            Dim SegAssociations As String = ""
            Dim PaxAssociations As String = ""

            Dim objSeg As Object
            Dim objPax As Object

            If Element.Associations.Segments.Count > 0 Then
                For Each objSeg In Element.Associations.segments
                    SegAssociations &= mobjSegs(objSeg.ElementNo).BoardPoint & " " & mobjSegs(objSeg.ElementNo).Airline & " " & mobjSegs(objSeg.ElementNo).OffPoint & vbCrLf
                Next
            Else
                For Each pSeg As AmadeusSeg.AmadeusSegItem In mobjSegs.Values
                    SegAssociations &= pSeg.BoardPoint & " " & pSeg.Airline & " " & pSeg.OffPoint & vbCrLf
                Next
            End If

            If Element.Associations.Passengers.Count > 0 Then
                For Each objPax In Element.Associations.Passengers
                    PaxAssociations &= mobjPax(objPax.ElementNo).PaxName & vbCrLf
                Next
            Else
                For Each pPax As AmadeusPax.AmadeusPaxitem In mobjPax.Values
                    PaxAssociations &= pPax.PaxName & vbCrLf
                Next
            End If


            Dim pstrText As String = Element.Text

            Dim pstrSplit1() As String = Split(pstrText, "/")
            pflgETicket = False

            If IsArray(pstrSplit1) Then

                pstrSplit2 = Split(pstrSplit1(0), " ")
                If UBound(pstrSplit1) > 0 Then
                    If Len(pstrSplit1(1)) = 4 And Left(pstrSplit1(1), 2) = "ET" Then
                        pflgETicket = True
                    End If
                End If
                If IsArray(pstrSplit2) Then
                    If UBound(pstrSplit2) >= 3 Then
                        With mobjNumberParser
                            If .TicketNumberText(pstrSplit2(3)) Then
                                mobjTickets.addTicket("FA", .StockType, .DocumentNumber, .Books, .AirlineNumber, pstrSplit1(1).Substring(2, 2), pflgETicket, SegAssociations, PaxAssociations, pstrSplit2(2))
                            End If
                        End With
                    End If
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub
    Private Sub getPax()

        Dim i As Short
        Dim j As Short
        Dim pstrID As String
        Dim pobjPax As s1aPNR.NameElement

        mobjPax.Clear()

        For Each pobjPax In mobjPNR.NameElements
            With pobjPax
                '                i = InStr(.Text, "(ID")
                i = InStr(.Text, "(")
                If i > 0 Then
                    j = InStrRev(.Text, ")")
                    If j = 0 Then
                        j = Len(.Text) + 1
                    End If
                    pstrID = .Text.Substring(i - 1, j - i + 1) ' Mid(.Text, i + 3, j - (i + 3))
                Else
                    pstrID = ""
                End If
                mobjPax.AddItem(.ElementNo, .Initial, .LastName, pstrID)
            End With
        Next pobjPax

    End Sub

    Private Sub getSegs(ByVal ForReportOnly As Boolean)

        Dim pobjSeg As Object

        mobjSegs.Clear()
        mSegsLastElement = -1

        For Each pobjSeg In mobjPNR.AllAirSegments
            Dim pElementNo As Short = airElementNo(pobjSeg)
            If ForReportOnly Then
                mobjSegs.AddItem(airAirline(pobjSeg), airBoardPoint(pobjSeg), airClass(pobjSeg), airDepartureDate(pobjSeg), airArrivalDate(pobjSeg), pElementNo, airFlightNo(pobjSeg), airOffPoint(pobjSeg), airStatus(pobjSeg), airDepartTime(pobjSeg), airArriveTime(pobjSeg), airText(pobjSeg), "")
            Else
                Dim pSegDo As k1aHostToolKit.CHostResponse = mobjHostSession.Send("DO" & pobjSeg.ElementNo)
                mobjSegs.AddItem(airAirline(pobjSeg), airBoardPoint(pobjSeg), airClass(pobjSeg), airDepartureDate(pobjSeg), airArrivalDate(pobjSeg), pElementNo, airFlightNo(pobjSeg), airOffPoint(pobjSeg), airStatus(pobjSeg), airDepartTime(pobjSeg), airArriveTime(pobjSeg), airText(pobjSeg), pSegDo.Text)
            End If
            If pElementNo > mSegsLastElement Then
                mSegsLastElement = pElementNo
            End If
        Next pobjSeg

    End Sub
    Private Function airStatus(ByRef pSegment As Object) As String

        Try
            airStatus = pSegment.text.substring(27, 2)
        Catch ex As Exception
            airStatus = ""
        End Try

    End Function

    Private Function airAirline(ByRef pSegment As Object) As String

        Try
            airAirline = pSegment.Airline
        Catch ex As Exception
            airAirline = ""
        End Try

    End Function

    Private Function airBoardPoint(ByRef pSegment As Object) As String

        Try
            airBoardPoint = pSegment.BoardPoint
        Catch ex As Exception
            airBoardPoint = ""
        End Try

    End Function

    Private Function airClass(ByRef pSegment As Object) As String

        Try
            airClass = pSegment.Class
        Catch ex As Exception
            airClass = ""
        End Try

    End Function

    Private Function airDepartureDate(ByRef pSegment As Object) As Date

        Dim pdteDate As Date

        Try
            pdteDate = pSegment.DepartureDate
            Do While pdteDate > DateAdd(Microsoft.VisualBasic.DateInterval.Year, 1, mudtProps.PNRCreationdate) Or pdteDate < System.DateTime.FromOADate(2)
                pdteDate = DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, pdteDate)
            Loop
            If pdteDate < System.DateTime.FromOADate(2) Then
                pdteDate = System.DateTime.FromOADate(0)
            End If

            airDepartureDate = pdteDate
        Catch ex As Exception
            airDepartureDate = System.DateTime.FromOADate(0)
        End Try

    End Function

    Private Function airArrivalDate(ByRef pSegment As Object) As Date

        Dim pdteDate As Date

        Try
            pdteDate = pSegment.ArrivalDate
            Do While pdteDate > DateAdd(Microsoft.VisualBasic.DateInterval.Year, 1, mudtProps.PNRCreationdate) Or pdteDate < System.DateTime.FromOADate(2)
                pdteDate = DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, pdteDate)
            Loop
            If pdteDate < System.DateTime.FromOADate(2) Then
                pdteDate = System.DateTime.FromOADate(0)
            End If

            airArrivalDate = pdteDate
        Catch ex As Exception
            airArrivalDate = System.DateTime.FromOADate(0)
        End Try

    End Function
    Private Function airElementNo(ByRef pSegment As Object) As Short

        Try
            airElementNo = pSegment.ElementNo
        Catch ex As Exception
            airElementNo = CShort("")
        End Try

    End Function

    Private Function airFlightNo(ByRef pSegment As Object) As String

        Try
            airFlightNo = pSegment.FlightNo
        Catch ex As Exception
            airFlightNo = ""
        End Try

    End Function

    Private Function airOffPoint(ByRef pSegment As Object) As String

        Try
            airOffPoint = pSegment.OffPoint
        Catch ex As Exception
            airOffPoint = ""
        End Try

    End Function
    Private Function airDepartTime(ByRef pSegment As Object) As Date

        Try
            airDepartTime = pSegment.DepartureTime
        Catch ex As Exception
            airDepartTime = System.DateTime.FromOADate(0)
        End Try

    End Function

    Private Function airArriveTime(ByRef pSegment As Object) As Date

        Try
            airArriveTime = pSegment.ArrivalTime
        Catch ex As Exception
            airArriveTime = System.DateTime.FromOADate(0)
        End Try

    End Function

    Private Function airText(ByRef pSegment As Object) As String

        Try
            airText = pSegment.Text
        Catch ex As Exception
            airText = ""
        End Try

    End Function

    Private Sub getOtherServiceElements()

        Dim pobjOtherServiceElement As s1aPNR.OtherServiceElement

        For Each pobjOtherServiceElement In mobjPNR.OtherServiceElements
            parseOtherServiceElements(pobjOtherServiceElement)
        Next pobjOtherServiceElement

    End Sub

    Private Sub parseOtherServiceElements(ByVal Element As s1aPNR.OtherServiceElement)

        Dim i As Short
        Dim j As Short

        Dim pintLen As Short
        Dim pstrText As String
        Dim pstrSplit() As String

        pstrText = ConcatenateText(Element.Text)
        pintLen = Len(pstrText)

        i = InStr(pstrText, "/SG")
        j = InStr(pstrText, "/P")
        If i > 0 And i - 1 < pintLen Then
            pintLen = i - 1
        End If
        If j > 0 And j - 1 < pintLen Then
            pintLen = j - 1
        End If

        pstrSplit = Split(Left(pstrText, pintLen), " ")

        If IsArray(pstrSplit) Then
            For i = LBound(pstrSplit) To UBound(pstrSplit)
                If pstrSplit(i) = "MV" Then
                    mstrVesselName = ""
                    For j = i + 1 To UBound(pstrSplit)
                        mstrVesselName = mstrVesselName & " " & Trim(pstrSplit(j))
                    Next j
                    Exit For
                ElseIf Left(pstrSplit(i), 11) = "SEMN/VESSEL" Then
                    mstrVesselName = Mid(pstrSplit(i), 13)
                    For j = i + 1 To UBound(pstrSplit)
                        If pstrSplit(j) <> "-" Then
                            mstrVesselName = mstrVesselName & " " & Trim(pstrSplit(j))
                        End If
                    Next j
                    Exit For
                End If
            Next i
        End If

    End Sub

    Private Sub getSSRElements()

        Dim pobjSSR As s1aPNR.SSRfqtvElement

        For Each pobjSSR In mobjPNR.SSRfqtvElements

            If pobjSSR.Associations.Passengers.Count > 0 Then
                For Each objPax In pobjSSR.Associations.Passengers
                    mobjFrequentFlyer.AddItem(mobjPax(objPax.ElementNo).PaxName, pobjSSR.Airline, pobjSSR.FrequentTravelerNo)
                Next
            Else
                For Each pPax As AmadeusPax.AmadeusPaxitem In mobjPax.Values
                    mobjFrequentFlyer.AddItem(pPax.PaxName, pobjSSR.Airline, pobjSSR.FrequentTravelerNo)
                Next
            End If

        Next

        Dim pTQTtext As k1aHostToolKit.CHostResponse = mobjHostSession.Send("RTSTR")
        If pTQTtext.Text.IndexOf("NO SEATS") = 0 Then
            mudtProps.Seats = ""
        Else
            Dim pTemp() As String = pTQTtext.Text.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
            Dim pTemp2 As String = ""
            If IsArray(pTemp) Then

                pTemp2 = pTemp(0)
                For i As Integer = 1 To pTemp.GetUpperBound(0)
                    If pTemp(i).Length > 10 Then
                        If pTemp2.Length > 0 Then
                            pTemp2 &= vbCrLf
                        End If
                        If pTemp(i).StartsWith(" ") Then
                            pTemp2 &= pTemp(i).Substring(0, 10) & " " & pTemp(i).Substring(13)
                        Else
                            pTemp2 &= pTemp(i)
                        End If

                    End If
                Next
            End If
            mudtProps.Seats = pTemp2
        End If

    End Sub

    Private Sub GetRMElements()

        Dim pobjRMElement As s1aPNR.RemarkElement

        For Each pobjRMElement In mobjPNR.RemarkElements
            parseRMElements(pobjRMElement)
        Next pobjRMElement

    End Sub
    Private Sub parseRMElements(ByVal Element As s1aPNR.RemarkElement)

        Dim pintLen As Short
        Dim pstrText As String
        Dim pstrSplit() As String

        pstrText = ConcatenateText(Element.Text)
        pintLen = Len(pstrText)
        pstrSplit = Split(Left(pstrText, pintLen), "/")

        If IsArray(pstrSplit) AndAlso pstrSplit.Length >= 2 Then
            If pstrSplit(1) = "CC" Then
                mstrCC = pstrSplit(2)
            ElseIf pstrSplit(1) = "CLN" Then
                mstrCLN = pstrSplit(2)
            ElseIf pstrSplit(1) = "CLA" Then
                mstrCLA = pstrSplit(2)
            End If
        End If

    End Sub
    Private Sub getTQT()

        Dim pTQTtext As k1aHostToolKit.CHostResponse = mobjHostSession.Send("TQT")
        Dim pTQT() As String = pTQTtext.Text.Split(vbCrLf)

        ReDim mudtAllowance(0)
        ReDim mudtTQT(0)

        If pTQT(0).StartsWith("T     P/S  NAME") Then
            For i As Integer = 1 To pTQT.GetUpperBound(0)
                If pTQT(i).Length > 62 AndAlso pTQT(i).Substring(1) <> " " Then
                    ReDim Preserve mudtTQT(mudtTQT.GetUpperBound(0) + 1)
                    If pTQT(i).Substring(1, pTQT(i).IndexOf(" ")) <> pTQT(i - 1).Substring(1, pTQT(i).IndexOf(" ")) AndAlso IsNumeric(pTQT(i).Substring(1, pTQT(i).IndexOf(" "))) Then
                        mudtTQT(mudtTQT.GetUpperBound(0)).TQTElement = pTQT(i).Substring(1, pTQT(i).IndexOf(" "))
                        Dim pSeg() As String
                        If i < pTQT.GetUpperBound(0) AndAlso pTQT(i + 1).Length > 2 AndAlso pTQT(i + 1).Substring(1, 1) = " " Then
                            pTQT(i) &= pTQT(i + 1).Trim
                            pSeg = pTQT(i).Substring(pTQT(i).LastIndexOf(" ")).Trim.Split(",")
                        Else
                            pSeg = pTQT(i).Substring(pTQT(i).LastIndexOf(" ")).Trim.Split(",")
                        End If
                        For i1 As Integer = 0 To pSeg.GetUpperBound(0)
                            Dim pSeg1() As String = pSeg(i1).Split("-")
                            If IsNumeric(pSeg1(0)) Then
                                mudtTQT(mudtTQT.GetUpperBound(0)).Segment = CInt(pSeg1(0))
                                If pSeg1.GetUpperBound(0) = 1 Then
                                    For i2 As Integer = CInt(pSeg1(0)) + 1 To CInt(pSeg1(1))
                                        ReDim Preserve mudtTQT(mudtTQT.GetUpperBound(0) + 1)
                                        mudtTQT(mudtTQT.GetUpperBound(0)).TQTElement = mudtTQT(mudtTQT.GetUpperBound(0) - 1).TQTElement
                                        mudtTQT(mudtTQT.GetUpperBound(0)).Segment = i2
                                    Next
                                Else

                                End If
                            End If

                        Next
                    End If

                    Dim pTSTText As k1aHostToolKit.CHostResponse = mobjHostSession.Send("TQT/T" & pTQT(i).Substring(1, pTQT(i).IndexOf(" ")))
                    Dim pTST() As String = pTSTText.Text.Split(vbCrLf)

                    SplitTQT(pTST)

                End If
            Next
        ElseIf pTQT(0).StartsWith("TST") Then
            SplitTQT(pTQT)
        End If

    End Sub

    Private Sub SplitTQT(ByVal pTQT() As String)

        Dim iSeg As Integer = 0
        For i As Integer = 0 To pTQT.GetUpperBound(0)
            If pTQT(i).Length > 4 AndAlso pTQT(i).Substring(5, 1) = "." Then
                iSeg = i + 1
            ElseIf iSeg > 0 Then
                Exit For
            End If
        Next
        If iSeg > 0 Then
            For i As Integer = iSeg To pTQT.GetUpperBound(0)
                If IsNumeric(pTQT(i).Substring(2, 1)) Then
                    If pTQT(i).Length > 60 Then
                        ReDim Preserve mudtAllowance(mudtAllowance.GetUpperBound(0) + 1)
                        mudtAllowance(mudtAllowance.GetUpperBound(0)).Itin = pTQT(i).Substring(6, 6) & " " & pTQT(i + 1).Substring(6, 3)
                        mudtAllowance(mudtAllowance.GetUpperBound(0)).Allowance = pTQT(i).Substring(61)
                        mudtAllowance(mudtAllowance.GetUpperBound(0)).Status = pTQT(i).Substring(31, 3)
                    End If
                Else
                    Exit For
                End If
            Next
        End If

    End Sub
    Private Function ConcatenateText(ByVal Text As String) As String

        Dim i As Short
        Dim j As Short
        Dim pintLen As Short
        Dim pstrTemp As String

        Try
            j = -1
            pintLen = Len(Text)
            For i = 1 To Len(Text)
                pstrTemp = Mid(Text, i, 1)
                If pstrTemp <> " " And (pstrTemp < "0" Or pstrTemp > "9") Then
                    j = i
                    Exit For
                End If
            Next i

            If j = -1 Then
                ConcatenateText = Text
            Else
                pstrTemp = Mid(Text, j, 60)
                j = j + 60
                Do While j <= pintLen
                    If Mid(Text, j, Math.Min(23, pintLen - j + 1)) & " " = " " & Mid(Text, j, Math.Min(23, pintLen - j + 1)) Then
                        j = j + 23
                        If j <= pintLen Then
                            pstrTemp = pstrTemp & Mid(Text, j, 57)
                            j = j + 57
                        End If
                    End If
                Loop
                ConcatenateText = pstrTemp
            End If
        Catch ex As Exception
            ConcatenateText = Text
        End Try

    End Function
End Class