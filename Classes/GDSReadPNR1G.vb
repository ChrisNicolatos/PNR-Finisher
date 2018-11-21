Public Class GDSReadPNR1G
    Event TerminalCommandSent(ByVal TerminalCommand As String)
    Private Structure PaxFFProps
        Dim PaxNumber As Short
        Dim Paxname As String
        Private _TicketNumber As String
        Private _DocumentNumber As String
        Private _Airline As String
        Dim Books As Integer
        Property TicketNumber As String
            Get
                TicketNumber = _TicketNumber
            End Get
            Set(value As String)
                value = value.Trim
                If value.Length = 10 Then
                    _TicketNumber = value
                    _DocumentNumber = value
                    _Airline = ""
                    Books = 1
                ElseIf value.Length = 13 AndAlso IsNumeric(value) Then
                    _TicketNumber = value
                    _DocumentNumber = value.Substring(3)
                    _Airline = value.Substring(0, 3)
                    Books = 1
                ElseIf value.Length = 17 AndAlso value.Substring(13, 1) = "-" Then
                    _TicketNumber = value
                    _DocumentNumber = value.Substring(3, 10)
                    _Airline = value.Substring(0, 3)
                    Dim pTemp As String = _DocumentNumber.Substring(0, 7) & value.Substring(14, 3)
                    Books = CDbl(pTemp) - CDbl(_DocumentNumber) + 1
                Else
                    _TicketNumber = ""
                    _DocumentNumber = ""
                    _Airline = ""
                    Books = 0
                End If
            End Set
        End Property
        ReadOnly Property DocumentNumber As String
            Get
                DocumentNumber = _DocumentNumber
            End Get
        End Property
        ReadOnly Property Airline As String
            Get
                Airline = _Airline
            End Get
        End Property
    End Structure
    Private Structure SegFFProps
        Dim SegNo As Short
        Dim BaggageAllowance As String
    End Structure
    Private WithEvents mobjSession1G As New Travelport.TravelData.Factory.GalileoDesktopFactory("SPG720", "MYCONNECTION", False, True, "SMRT")

    Private mobjPassengers As New GDSPaxColl
    Private mobjSegments As New GDSSegCollection
    Private mobjTickets As New GDSTicketCollection
    Private mobjPhones As New PhoneNumbersCollection
    Private mobjEmails As New EmailCollection
    Private mobjOpenSegments As New OpenSegmentCollection
    Private mobjDI As New DICollection
    Private mobjTicketElement As New TicketElementItem
    Private mobjOptionQueue As New OptionQueueCollection
    Private mobjSSR As New SSRCollection
    Private mobjFreqFlyer As New FrequentFlyerCollection
    Private mobjItinRemarks As New GDSItineraryRemarksCollection
    Private mstrOfficeOfResponsibility As String
    Private mstrPNRNumber As String
    Private mdteDepartureDate As Date
    Private mstrItinerary As String
    Private mSegsFirstElement As Integer
    Private mSegsLastElement As Integer
    Private mudtAllowance() As TQTItem
    Private mstrSeats As String

    Public Sub ReadRaw(ByVal RequestedPNR As String)

        Dim pPNRStatus() As String
        If RequestedPNR = "" Then
            pPNRStatus = SendTerminalCommand("*R")
        Else
            pPNRStatus = SendTerminalCommand("*" & RequestedPNR)
        End If

        If pPNRStatus(0).StartsWith("NO B.F.") Then
            Throw New Exception(pPNRStatus(0))
        End If

        ReadPNRElements()

    End Sub
    Private Sub ReadPNRElements()

        mobjPassengers.Clear()
        mobjSegments.Clear()
        mobjPhones.Clear()
        mobjEmails.Clear()
        mobjOpenSegments.Clear()
        mobjDI.Clear()
        mobjTicketElement.Clear()
        mobjOptionQueue.Clear()
        mobjSSR.Clear()
        mobjFreqFlyer.Clear()
        mobjItinRemarks.Clear()
        mstrSeats = ""

        GetOfficeOfResponsibility1G()
        GetPassengers1G()
        GetSegments1G()
        GetPhoneElement1G()
        GetEmailElement1G()
        GetTicketElement1G()
        GetOpenSegment1G()
        GetOptionQueueElement1G()
        GetTickets()
        GetEMDs()
        GetSSR1G()
        GetDI1G()
        GetFreqFlyers()
        GetRI()

    End Sub
    Private Function SendTerminalCommand(ByVal TerminalEntry As String) As String()
        Dim mstrPNR As ObjectModel.ReadOnlyCollection(Of String) = mobjSession1G.SendTerminalCommand(TerminalEntry)
        RaiseEvent TerminalCommandSent(TerminalEntry)
        Dim pRawIndex As Integer = -1
        Dim pSendTerminalCommand(0) As String
        Dim pRead As Boolean = True
        Do While pRead
            For i As Integer = 0 To mstrPNR.Count - 1
                If mstrPNR(i).Trim <> "" And mstrPNR(i).Trim <> ")>" And mstrPNR(i).Trim <> ">" Then
                    pRawIndex += 1
                    ReDim Preserve pSendTerminalCommand(pRawIndex)
                    pSendTerminalCommand(pRawIndex) = mstrPNR(i).TrimEnd
                End If
            Next
            If mstrPNR(mstrPNR.Count - 1) = ")>" Then
                mstrPNR = mobjSession1G.SendTerminalCommand("MR")
            Else
                pRead = False
            End If
        Loop
        Return pSendTerminalCommand
    End Function
    Private Sub GetOfficeOfResponsibility1G()

        Dim pPCC() As String = SendTerminalCommand("*HI")
        mstrOfficeOfResponsibility = MySettings.GDSPcc
        If pPCC.GetUpperBound(0) >= 1 Then
            If pPCC(pPCC.GetUpperBound(0)).StartsWith("CRDT-") Then
                Dim pItems() As String = pPCC(pPCC.GetUpperBound(0)).Substring(5).Split("/")
                If pItems.GetUpperBound(0) >= 2 Then
                    mstrOfficeOfResponsibility = pItems(1).Trim
                End If
            End If
        End If
    End Sub
    Public ReadOnly Property Seats As String
        Get
            Seats = mstrSeats
        End Get
    End Property
    Public ReadOnly Property Tickets As GDSTicketCollection
        Get
            Tickets = mobjTickets
        End Get
    End Property
    Public ReadOnly Property Allowance As TQTItem()
        Get
            Allowance = mudtAllowance
        End Get
    End Property
    Public ReadOnly Property OfficeOfResponsibility As String
        Get
            OfficeOfResponsibility = mstrOfficeOfResponsibility
        End Get
    End Property
    Public ReadOnly Property RequestedPNR As String
        Get
            RequestedPNR = mstrPNRNumber
        End Get
    End Property
    Public ReadOnly Property ItineraryRemarks As GDSItineraryRemarksCollection
        Get
            Return mobjItinRemarks
        End Get
    End Property
    Private Sub GetPassengers1G()

        Dim pPax() As String = SendTerminalCommand("*N")
        Dim pAllPax As String = ""
        If pPax(0).IndexOf(".") >= 1 And pPax(0).IndexOf(".") <= 2 Then
            mstrPNRNumber = "New PNR"
        Else
            mstrPNRNumber = pPax(0).Substring(0, 6)
            For i As Integer = 0 To pPax.GetUpperBound(0)
                If pPax(i).IndexOf("/") = 6 Then
                    mstrPNRNumber = pPax(i).Substring(0, 6)
                    Exit For
                End If
            Next
        End If
        For i As Integer = 0 To pPax.GetUpperBound(0)
            If pPax(i).IndexOf(".") >= 1 And pPax(i).IndexOf(".") <= 3 Then
                ' remove any / character after the one that separates last name and first name. Specifically, this could appear in the name field entry
                'If pPax(i).IndexOf("/") > -1 And pPax(i).IndexOf("/") < pPax(i).Length - 1 Then
                '    If pPax(i).IndexOf("/", pPax(i).IndexOf("/") + 1) > -1 Then
                '        pPax(i) = pPax(i).Substring(0, pPax(i).IndexOf("/") + 1) & pPax(i).Substring(pPax(i).IndexOf("/") + 1).Replace("/", "")
                '    End If
                'End If
                pAllPax &= pPax(i) & " "
            End If
        Next

        pPax = pAllPax.Split("/")
        For i As Integer = 0 To pPax.GetUpperBound(0) - 1
            Dim iStart As Integer = 0
            If pPax(i).LastIndexOf(" ") >= 0 Then
                iStart = pPax(i).LastIndexOf(" ") + 1
            End If
            Dim iPaxCount As Integer = pPax(i).Substring(pPax(i).IndexOf(".", iStart) + 1, 1)
            Dim iSurname As String = pPax(i).Substring(pPax(i).IndexOf(".", iStart) + 2)
            If IsNumeric(pPax(i).Substring(pPax(i).IndexOf(".", iStart) + 1, 2)) Then
                iPaxCount = pPax(i).Substring(pPax(i).IndexOf(".", iStart) + 1, 2)
                iSurname = pPax(i).Substring(pPax(i).IndexOf(".", iStart) + 3)
            End If
            For j As Integer = i + 1 To i + iPaxCount
                pPax(j) = iSurname & "/" & pPax(j)
            Next
            If iStart = 0 Then
                pPax(i) = ""
            Else
                pPax(i) = pPax(i).Substring(0, iStart).Trim
            End If
            i = i + iPaxCount - 1
        Next
        Dim pNameRemark As String = ""
        mobjPassengers.Clear()
        For i As Integer = 1 To pPax.GetUpperBound(0)
            pNameRemark = ""
            If pPax(i).IndexOf("*") > 0 Then
                pNameRemark = pPax(i).Substring(pPax(i).IndexOf("*") + 1)
                pPax(i) = pPax(i).Substring(0, pPax(i).IndexOf("*"))
            End If
            Dim pNames() As String = pPax(i).Split("/")
            mobjPassengers.AddItem(i, pNames(1), pNames(0), pNameRemark)
        Next
    End Sub
    Public ReadOnly Property Passengers As GDSPaxColl
        Get
            Passengers = mobjPassengers
        End Get
    End Property
    Private Sub GetSegments1G()

        Dim pVL() As String = SendTerminalCommand("*VL")
        Dim pOff As String = ""
        Dim pSegs() As String = SendTerminalCommand("*IA")
        mobjSegments.Clear()
        mdteDepartureDate = Date.MinValue
        mstrItinerary = ""
        mSegsLastElement = -1
        mSegsFirstElement = -1

        For i As Integer = 0 To pSegs.GetUpperBound(0)
            Dim pOrigin As String
            Dim pDestination As String
            Dim pDepartureDate As Date
            Dim pArrivalDate As Date
            Dim pDepartureTime As Date
            Dim pArrivalTime As Date
            Dim pCarrier As String
            Dim pFlightNumber As String
            Dim pClassOfService As String
            Dim pStatus As String
            Dim pOperatedBy As String
            Dim pEquipment As String = ""
            Dim pArrivalDays As Integer = 0
            Dim pobjSeg As GDSSegItem
            Dim pStart As Integer = pSegs(i).IndexOf(".")
            If pStart >= 1 And pSegs(i).Length >= 57 Then
                With pSegs(i)
                    Dim pElementNo As Integer = .Substring(0, pStart).Trim
                    pCarrier = .Substring(pStart + 2, 2).Trim
                    pFlightNumber = .Substring(pStart + 5, 4).Trim
                    pClassOfService = .Substring(pStart + 10, 1).Trim
                    pDepartureDate = DateFromIATA(.Substring(pStart + 13, 5))
                    pOrigin = .Substring(pStart + 19, 3).Trim
                    pDestination = .Substring(pStart + 22, 3).Trim
                    pStatus = .Substring(pStart + 26, 2).Trim
                    pDepartureTime = TimeSerial(.Substring(pStart + 31, 2), .Substring(pStart + 33, 2), 0)
                    pArrivalTime = TimeSerial(.Substring(pStart + 38, 2), .Substring(pStart + 40, 2), 0)
                    pEquipment = ""
                    pArrivalDays = 0
                    If .Length > pStart + 37 Then
                        If .Substring(pStart + 37, 1) = "#" Then
                            pArrivalDays = +1
                        ElseIf .Substring(pStart + 37, 1) = "*" Then
                            pArrivalDays = +2
                        ElseIf .Substring(pStart + 37, 1) = "-" Then
                            pArrivalDays = -1
                        End If
                    End If
                    pArrivalDate = DateAdd(DateInterval.Day, pArrivalDays, pDepartureDate)
                    pOperatedBy = ""
                    If i < pSegs.GetUpperBound(0) AndAlso pSegs(i + 1).StartsWith(Space(4)) Then
                        pOperatedBy = pSegs(i + 1).Trim
                    End If
                    pobjSeg = mobjSegments.AddItem(pCarrier, pOrigin, pClassOfService, pDepartureDate, pArrivalDate, pElementNo, pFlightNumber, pDestination, pStatus, pDepartureTime, pArrivalTime, pEquipment, pVL, pSegs(i), pOperatedBy, ReadSVC(pElementNo))
                    If mSegsFirstElement = -1 Then
                        mSegsFirstElement = pElementNo
                    End If
                    If pElementNo > mSegsLastElement Then
                        mSegsLastElement = pElementNo
                    End If

                End With
                With pobjSeg
                    If mstrItinerary = "" Then
                        mstrItinerary = .BoardPoint & "-" & .OffPoint
                    Else
                        If .BoardPoint = pOff Then
                            mstrItinerary &= "-" & .OffPoint
                        Else
                            mstrItinerary &= "-***-" & .BoardPoint & "-" & .OffPoint
                        End If
                    End If
                    pOff = .OffPoint
                    If mdteDepartureDate = Date.MinValue Then
                        mdteDepartureDate = .DepartureDate
                    End If
                End With
            End If
        Next
        If mdteDepartureDate > Date.MinValue Then
            Dim pDate As New s1aAirlineDate.clsAirlineDate
            pDate.SetFromString(mdteDepartureDate)
            mstrItinerary &= " (" & pDate.IATA & ")"
        End If

    End Sub
    Private Function ReadSVC(ByVal ElementNo As String) As String()
        ReadSVC = SendTerminalCommand("*SVC" & ElementNo)
    End Function
    Public ReadOnly Property Segments As GDSSegCollection
        Get
            Segments = mobjSegments
        End Get
    End Property
    Public ReadOnly Property SegsFirstElement As Integer
        Get
            SegsFirstElement = mSegsFirstElement
        End Get
    End Property
    Public ReadOnly Property SegsLastElement As Integer
        Get
            SegsLastElement = mSegsLastElement
        End Get
    End Property
    Public ReadOnly Property DepartureDate As Date
        Get
            DepartureDate = mdteDepartureDate
        End Get
    End Property
    Private Sub GetPhoneElement1G()

        Dim pPhones() As String = SendTerminalCommand("*P")

        For i As Integer = 0 To pPhones.GetUpperBound(0)
            Dim pElement As Short = 0
            Dim pStart As Integer = 0
            Dim pStar As Integer = 0
            Dim pCityCode As String = ""
            Dim pPhoneNumber As String = ""
            Dim pPhoneType As String = ""
            If i < pPhones.GetUpperBound(0) AndAlso pPhones(i + 1).Length > 5 AndAlso pPhones(i + 1).StartsWith("     ") Then
                pPhones(i) &= pPhones(i + 1).Substring(5)
                pPhones(i + 1) = ""
            End If
            If pPhones(i).StartsWith("FONE-") Then
                pElement = 1
                pStart = 5
            ElseIf pPhones(i).IndexOf(". ") >= 1 And pPhones(i).IndexOf(".") <= 3 Then
                pElement = pPhones(i).Substring(0, pPhones(i).IndexOf(".")).Trim
                pStart = pPhones(i).IndexOf(".") + 2
            End If
            pStar = pPhones(i).IndexOf("*")
            If pStart > 0 And pStar > pStart Then
                pCityCode = pPhones(i).Substring(pStart, 3)
                pPhoneType = pPhones(i).Substring(pStart + 3, 1)
                pPhoneNumber = pPhones(i).Substring(pStar + 1)
                mobjPhones.AddItem(pElement, pCityCode, pPhoneType, pPhoneNumber)
            End If
        Next

    End Sub
    Public ReadOnly Property PhoneNumbers As PhoneNumbersCollection
        Get
            PhoneNumbers = mobjPhones
        End Get
    End Property
    Private Sub GetEmailElement1G()

        Dim pEmails() As String = SendTerminalCommand("*EM")

        Dim pElementAddress As Short = 0
        Dim pElementComment As Short = 0
        Dim pFromAddress As String = ""
        Dim pToAddress As String = ""
        Dim pComment As String = ""
        For i As Integer = 0 To pEmails.GetUpperBound(0)
            If pEmails(i).StartsWith("FROM-") Then
                pFromAddress = pEmails(i).Substring(5).Trim
                mobjEmails.SetFromAddress(pFromAddress)
            ElseIf pEmails(i).StartsWith("TO-") Then
                If pElementAddress <> 0 Then
                    mobjEmails.AddItem(pElementAddress, pToAddress, pComment)
                    pElementAddress = 0
                    pToAddress = ""
                    pComment = ""
                End If
                pElementAddress = pEmails(i).Substring(3, 5).Trim
                pToAddress = pEmails(i).Substring(10).Trim
            ElseIf pEmails(i).StartsWith("COM-") Then
                pElementComment = pEmails(i).Substring(4, 5).Trim
                If pElementAddress <> pElementComment Then
                    mobjEmails.AddItem(pElementAddress, pToAddress, pComment)
                    pElementAddress = pElementComment
                    pToAddress = ""
                    pComment = ""
                End If
                pComment = pEmails(i).Substring(10).Trim
            End If
        Next
        If pElementAddress <> 0 Then
            mobjEmails.AddItem(pElementAddress, pToAddress, pComment)
        End If

    End Sub
    Public ReadOnly Property Emails As EmailCollection
        Get
            Emails = mobjEmails
        End Get
    End Property
    Private Sub GetTicketElement1G()
        Dim pTicket() As String = SendTerminalCommand("*TD")
        For i As Integer = 0 To pTicket.GetUpperBound(0)
            Dim pElement As Short = i + 1
            Dim pPCC As String = ""
            Dim pActionDate As Date = Today
            Dim pRemark As String = ""
            Dim pItems() As String = pTicket(i).Split("/")
            If pItems(0) = "TKTG-TAU" Then
                ' TKTG-TAU/750B/WE15AUG
                Dim pRem() As String = pItems(pItems.GetUpperBound(0)).Split("*")
                If pRem.GetUpperBound(0) = 0 Then
                    pRemark = ""
                Else
                    pRemark = pRem(1)
                End If
                pActionDate = DateFromIATA(pRem(0).Substring(pRem(0).Length - 5))
                If pItems.GetUpperBound(0) = 2 Then
                    pPCC = pItems(1)
                End If
                mobjTicketElement.SetValues(pElement, pPCC, pActionDate, pRemark)
            End If
        Next
    End Sub
    Public ReadOnly Property TicketElement As TicketElementItem
        Get
            TicketElement = mobjTicketElement
        End Get
    End Property
    Private Sub GetOptionQueueElement1G()
        Dim pOP() As String = SendTerminalCommand("*RB")
        For i As Integer = 1 To pOP.GetUpperBound(0)
            Dim pElement As Short
            Dim pPCC As String = ""
            Dim pActionDateTime As Date
            Dim pQueue As String
            Dim pRemark As String
            If pOP(i).StartsWith("RBKG-") Then
                pElement = 1
            Else
                pElement = pOP(i).Substring(0, 3).Trim
            End If
            Dim pItem() As String = pOP(i).Substring(5).Split("/")
            Dim pRem() As String = pItem(pItem.GetUpperBound(0)).Split("*")
            pQueue = pRem(0)
            If pRem.GetUpperBound(0) = 1 Then
                pRemark = pRem(1)
            Else
                pRemark = ""
            End If
            pPCC = pItem(0)
            pActionDateTime = DateFromIATA(pItem(1).Substring(pItem(1).Length - 5)) + TimeSerial(pItem(2).Substring(0, 2), pItem(2).Substring(2), 0).TimeOfDay
            mobjOptionQueue.AddItem(pElement, pPCC, pActionDateTime, pQueue, pRemark)
        Next
    End Sub
    Public ReadOnly Property OptionQueue As OptionQueueCollection
        Get
            OptionQueue = mobjOptionQueue
        End Get
    End Property
    Private Sub GetFreqFlyers()
        Dim pAirline As String = ""
        Dim pPaxName As String = ""
        Dim pMembershipNo As String = ""
        Dim pCrossAccrual As String = ""
        Dim pFF() As String = SendTerminalCommand("*MM")

        For i As Integer = 0 To pFF.Count - 1
            If pFF(i).Length > 4 AndAlso pFF(i).StartsWith("P") AndAlso pFF(i).Substring(4, 1) = "." Then
                pAirline = pFF(i).Substring(24, 2).Trim
                pPaxName = pFF(i).Substring(6, 18).Trim
                pMembershipNo = pFF(i).Substring(28).Trim
                If pMembershipNo.IndexOf("*") > 0 Then
                    pMembershipNo = pMembershipNo.Substring(0, pMembershipNo.IndexOf("*"))
                End If
                Dim pFound As Boolean = False
                Dim pFN As String = ""
                Dim pSN As String = ""
                If pPaxName.IndexOf("/") > 0 Then
                    pSN = pPaxName.Substring(0, pPaxName.IndexOf("/")).Replace("+", "")
                    pFN = pPaxName.Substring(pPaxName.IndexOf("/") + 1).Replace("+", "")
                End If
                For Each pPass As GDSPaxItem In Passengers.Values
                    If pPass.LastName.StartsWith(pSN) And pPass.Initial.StartsWith(pFN) Then
                        pPaxName = pPass.PaxName
                        pFound = True
                        Exit For
                    End If
                Next

                'For j As Integer = 0 To Passengers.Count - 1
                '    If Passengers(j).PaxName = pPaxName Then
                '        pFound = True
                '        Exit For
                '    End If
                'Next j
                If Not pFound Then
                    Dim pFirstName As String = ""
                    Dim pSurname As String = ""
                    If pPaxName.IndexOf("/") > 0 Then
                        pSurname = pPaxName.Substring(0, pPaxName.IndexOf("/")).Replace("+", "")
                        pFirstName = pPaxName.Substring(pPaxName.IndexOf("/") + 1).Replace("+", "")
                    End If
                    For Each pPass As GDSPaxItem In Passengers.Values
                        If pPass.LastName.StartsWith(pSurname) And pPass.Initial.StartsWith(pFirstName) Then
                            pPaxName = pPass.PaxName
                            pFound = True
                            Exit For
                        End If

                    Next
                End If
                If i < pFF.Count - 1 AndAlso pFF(i + 1).StartsWith(Space(28)) Then
                    pCrossAccrual = pFF(i + 1).Substring(28).Trim
                End If
                mobjFreqFlyer.AddItem(pPaxName, pAirline, pMembershipNo, pCrossAccrual)
            End If
        Next i
    End Sub
    Public ReadOnly Property FrequentFlyers() As FrequentFlyerCollection
        Get
            FrequentFlyers = mobjFreqFlyer
        End Get
    End Property
    Private Sub GetRI()
        Dim pRI() As String = SendTerminalCommand("*RI")
        mobjItinRemarks.Load1G(pRI)
    End Sub
    Private Sub GetTickets()
        Dim pFF() As String = SendTerminalCommand("*FF")
        ReDim mudtAllowance(0)
        mobjTickets.Clear()

        For i = 0 To pFF.GetUpperBound(0)
            If pFF(i).StartsWith("FQ") Or pFF(i).StartsWith("FB") Then
                Dim pFFid As Integer = pFF(i).Substring(2, pFF(i).IndexOf(" ") - 2)

                '
                ' *FFx for each FF element
                '
                Dim pFFx() As String = SendTerminalCommand("*FF" & pFFid)
                Dim pPax(0) As PaxFFProps
                pPax(0).PaxNumber = 0
                Dim pSeg(0) As SegFFProps
                pSeg(0).SegNo = 0
                Dim pSegNo As Integer = 0
                Dim pBaggageAllowance = ""
                For iPFF As Integer = 0 To pFFx.GetUpperBound(0)
                    If pFFx(iPFF).Length > 13 _
                        AndAlso pFFx(iPFF).StartsWith(" P") _
                        AndAlso IsNumeric(pFFx(iPFF).Substring(2, 1)) Then

                        ' If line starts with P and a number then it is a passenger name line
                        ' Next passenger
                        pPax(0).PaxNumber += 1
                        ReDim Preserve pPax(pPax(0).PaxNumber)
                        pPax(pPax(0).PaxNumber).PaxNumber = pFFx(iPFF).Substring(2, pFFx(iPFF).IndexOf(" ", 2))
                        pPax(pPax(0).PaxNumber).TicketNumber = ""

                        If pFFx(iPFF).IndexOf(" ", 5) > 5 Then
                            pPax(pPax(0).PaxNumber).Paxname = pFFx(iPFF).Substring(5, pFFx(iPFF).IndexOf(" ", 5) - 4).Trim
                        Else
                            pPax(pPax(0).PaxNumber).Paxname = pFFx(iPFF).Substring(5).Trim
                        End If

                        Dim pTemp1 As String = ""
                        Dim pTemp2 As String = ""
                        If iPFF < pFFx.GetUpperBound(0) Then
                            pTemp1 = pFFx(iPFF + 1).Substring(pFFx(iPFF + 1).LastIndexOf(" "))
                            If pTemp1.Length > 13 Then
                                pTemp1 = pTemp1.Substring(0, 13)
                            End If
                        End If
                        If pFFx(iPFF).Trim.Length > 13 Then
                            pTemp2 = pFFx(iPFF).Trim.Substring(pFFx(iPFF).Trim.Length - 13)
                        End If
                        If iPFF < pFFx.GetUpperBound(0) AndAlso pFFx(iPFF + 1).StartsWith(Space(13)) AndAlso pFFx(iPFF + 1).LastIndexOf(" ") <= pFFx(iPFF + 1).Length _
                            AndAlso IsNumeric(pTemp1) AndAlso Not IsNumeric(pTemp2) Then

                            ' if the next line starts with spaces, then the ticket number might be on the next line
                            pPax(pPax(0).PaxNumber).TicketNumber = pFFx(iPFF + 1).Trim.Substring(pFFx(iPFF + 1).Trim.LastIndexOf(" ")).Trim
                            pFFx(iPFF + 1) = ""
                        ElseIf IsNumeric(pFFx(iPFF).Trim.Substring(pFFx(iPFF).Trim.Length - 2)) Then
                            pPax(pPax(0).PaxNumber).TicketNumber = pFFx(iPFF).Trim.Substring(pFFx(iPFF).LastIndexOf(" "))
                        Else
                            pPax(pPax(0).PaxNumber).TicketNumber = ""
                        End If
                    ElseIf pFFx(iPFF).Length > 7 AndAlso pFFx(iPFF).StartsWith(" S") AndAlso pFFx(iPFF).Substring(3, 3) = Space(3) AndAlso IsNumeric(pFFx(iPFF).Substring(2, 1)) Then
                        pSegNo = pFFx(iPFF).Substring(2, pFFx(iPFF).IndexOf(" ", 2))
                        pBaggageAllowance = ""
                        For ipff1 = iPFF To pFFx.GetUpperBound(0)
                            If (ipff1 = iPFF Or pFFx(ipff1).StartsWith(Space(6))) AndAlso pFFx(ipff1).IndexOf("BG-") > 0 Then
                                pBaggageAllowance = pFFx(ipff1).Substring(pFFx(ipff1).IndexOf("BG-") + 3).Trim & " "
                                pBaggageAllowance = pBaggageAllowance.Substring(0, pBaggageAllowance.IndexOf(" "))
                                Exit For
                            ElseIf (ipff1 = iPFF Or pFFx(ipff1).StartsWith(Space(6))) AndAlso pFFx(ipff1).IndexOf(" B-") > 0 Then
                                pBaggageAllowance = pFFx(ipff1).Substring(pFFx(ipff1).IndexOf(" B-") + 3).Trim & " "
                                pBaggageAllowance = pBaggageAllowance.Substring(0, pBaggageAllowance.IndexOf(" "))
                                Exit For
                            ElseIf ipff1 > iPFF And Not pFFx(ipff1).StartsWith(Space(6)) Then
                                Exit For
                            End If
                        Next
                        If pBaggageAllowance <> "" Then
                            pSeg(0).SegNo += 1
                            ReDim Preserve pSeg(pSeg(0).SegNo)
                            pSeg(pSeg(0).SegNo).SegNo = pSegNo
                            pSeg(pSeg(0).SegNo).BaggageAllowance = pBaggageAllowance
                        End If
                    End If
                Next

                For i1 As Integer = 1 To pPax(0).PaxNumber
                    Dim pTktSeg As String = ""
                    For j1 As Integer = 1 To pSeg(0).SegNo
                        ReDim Preserve mudtAllowance(mudtAllowance.GetUpperBound(0) + 1)
                        mudtAllowance(mudtAllowance.GetUpperBound(0)) = New TQTItem
                        With mudtAllowance(mudtAllowance.GetUpperBound(0))
                            .TQTElement = pFFid
                            .Pax = pPax(i1).PaxNumber
                            .TicketNumber = pPax(i1).TicketNumber
                            .Segment = pSeg(j1).SegNo
                            .Allowance = pSeg(j1).BaggageAllowance
                            Try
                                .Itin = mobjSegments(.Segment).BoardPoint & " " & mobjSegments(.Segment).Airline & " " & mobjSegments(.Segment).OffPoint
                                .Status = mobjSegments(.Segment).Status
                                If pTktSeg <> "" Then
                                    pTktSeg &= vbCrLf
                                End If
                                pTktSeg &= .Itin
                            Catch ex As Exception
                                .Itin = ""
                                .Status = ""
                            End Try
                        End With
                    Next
                    mobjTickets.addTicket("FA", 1, CDbl("0" & pPax(i1).DocumentNumber), pPax(i1).Books, pPax(i1).Airline, Airlines.AirlineCode(pPax(i1).Airline), True, pTktSeg, pPax(i1).Paxname, "PAX", "")


                Next
            End If
        Next
        If mobjTickets.Count = 0 Then
            GetTicketsFromHTE()
        End If
        mstrSeats &= GetSeats()

    End Sub
    Private Sub GetTicketsFromHTE()
        Dim pHTE() As String = SendTerminalCommand("*HTE")
        'ReDim mudtAllowance(0)
        mobjTickets.Clear()
        If pHTE(0).StartsWith("TKT:") Then
            GetTicketsHTEParser(pHTE)
        Else
            For i As Integer = 0 To pHTE.GetUpperBound(0)
                If pHTE(i).StartsWith(">*TE") Then
                    Dim pTE() As String = SendTerminalCommand(pHTE(i).Substring(1, 6))
                    GetTicketsHTEParser(pTE)
                End If
            Next
        End If
    End Sub
    Private Sub GetTicketsHTEParser(ByVal pTE As String())
        Dim pPax(0) As PaxFFProps
        pPax(0).PaxNumber = 0
        Dim pSeg(0) As SegFFProps
        pSeg(0).SegNo = 0
        Dim pTktSeg As String = ""
        Dim pItin As String = ""
        For i1 As Integer = 0 To pTE.GetUpperBound(0)
            If pTE(i1).StartsWith("TKT") And pTE(i1).Length > 30 Then
                pPax(0).PaxNumber += 1
                ReDim Preserve pPax(pPax(0).PaxNumber)
                pPax(pPax(0).PaxNumber).PaxNumber = pPax(0).PaxNumber
                pPax(pPax(0).PaxNumber).Paxname = pTE(i1).Substring(31).Trim
                pPax(pPax(0).PaxNumber).TicketNumber = pTE(i1).Substring(5, 20).Replace(" ", "")
            ElseIf pTE(i1).Length > 32 And pTE(i1).StartsWith(Space(3)) And Not pTE(i1).StartsWith("   USE  CR FLT") And Not pTE(i1).StartsWith(Space(10)) Then
                pItin = pTE(i1).Substring(26, 3) & " " & pTE(i1).Substring(8, 2) & " " & pTE(i1).Substring(29, 3)
                If pTktSeg <> "" Then
                    pTktSeg &= vbCrLf
                End If
                pTktSeg &= pItin
            ElseIf pTktSeg <> "" And Not pTE(i1).StartsWith(Space(1)) Then
                Exit For
            End If
        Next
        mobjTickets.addTicket("FA", 1, CDbl("0" & pPax(pPax(0).PaxNumber).DocumentNumber), pPax(pPax(0).PaxNumber).Books, pPax(pPax(0).PaxNumber).Airline, Airlines.AirlineCode(pPax(pPax(0).PaxNumber).Airline), True, pTktSeg, pPax(pPax(0).PaxNumber).Paxname, "PAX", "")
    End Sub
    Private Sub GetEMDs()
        Dim pPax(0) As PaxFFProps
        pPax(0).PaxNumber = 0
        Dim pEMDDescription As String = ""
        Dim pEMD() As String = SendTerminalCommand("EMDL")
        If pEMD(0).StartsWith("EMDL - ") Then
            For i As Integer = 1 To pEMD.GetUpperBound(0)
                If pEMD(i).Length > 5 AndAlso pEMD(i).Substring(3, 1) = "." AndAlso IsNumeric(pEMD(i).Substring(0, 3).Trim) Then
                    Dim pEMDD() As String = SendTerminalCommand("EMDD" & pEMD(i).Substring(0, 3).Trim)
                    Dim pTemp() As String = pEMDD(0).Split({" "}, StringSplitOptions.RemoveEmptyEntries)
                    If pTemp.GetUpperBound(0) > 1 Then
                        pPax(0).PaxNumber += 1
                        ReDim Preserve pPax(pPax(0).PaxNumber)
                        pPax(pPax(0).PaxNumber).PaxNumber = pPax(0).PaxNumber
                        If i < pEMD.GetUpperBound(0) AndAlso pEMD(i + 1).IndexOf("/") > 0 Then
                            pPax(pPax(0).PaxNumber).Paxname = pEMD(i + 1)
                        Else
                            pPax(pPax(0).PaxNumber).Paxname = pTemp(1)
                        End If
                        pPax(pPax(0).PaxNumber).TicketNumber = pTemp(0)
                        pEMDDescription = pEMDD(2).Substring(4, 29).Trim
                        mobjTickets.addTicket("FA", 1, CDbl("0" & pPax(pPax(0).PaxNumber).DocumentNumber), pPax(pPax(0).PaxNumber).Books, pPax(pPax(0).PaxNumber).Airline, Airlines.AirlineCode(pPax(pPax(0).PaxNumber).Airline), False, "", pPax(pPax(0).PaxNumber).Paxname, "EMD", pEMDDescription)
                    End If
                End If
            Next
        End If
    End Sub
    Private Function GetSeats() As String
        Dim pSeats() As String = SendTerminalCommand("*SD")
        GetSeats = ""
        If pSeats(0).IndexOf("DATA NOT FOUND") = -1 Then
            For i As Integer = 1 To pSeats.Count - 1
                If pSeats(i).Length > 2 AndAlso pSeats(i).Substring(2, 1) = "." AndAlso IsNumeric(pSeats(i).Substring(1, 1)) Then
                    pSeats(i) = pSeats(i).Substring(0, 12) & " " & pSeats(i).Substring(15)
                End If
                GetSeats &= pSeats(i).Replace("NO CHARACTERISTICS EXIST", "") & vbCrLf
            Next
        End If
    End Function
    Private Sub GetSSR1G()
        Dim pSSR() As String = SendTerminalCommand("*SO")
        Dim pElementNo As Short = 0
        Dim pSpaces As Integer = 0
        Dim pSSRType As String = ""
        Dim pSSRCode As String = ""
        Dim pCarrierCode As String = ""
        Dim pStatusCode As String = ""
        Dim pText As String = ""
        Dim pLastName As String = ""
        Dim pFirstName As String = ""
        Dim pDateOfBirth As Date = Today
        Dim pPassportNumber As String = ""
        ' ** SPECIAL SERVICE REQUIREMENT **  
        ' SEGMENT/PASSENGER RELATED   
        '
        ' ** OTHER SUPPLEMENTARY INFORMATION **    
        ' CARRIER RELATED  
        '

        For i = 2 To pSSR.GetUpperBound(0)
            If pSSR(i) <> "" Then
                For j = i + 1 To pSSR.GetUpperBound(0)
                    If pSSR(j).StartsWith(Space(20)) And pSSR(i).Length > 1 Then
                        pSSR(i) = pSSR(i).Substring(0, pSSR(i).Length - 1) & pSSR(j).Substring(21)
                        'pSSR(i) &= pSSR(j).Substring(20)
                        pSSR(j) = ""
                    Else
                        Exit For
                    End If
                Next
                pElementNo = pSSR(i).Substring(0, 3).Trim
                If pSSR(i).Length > 8 AndAlso pSSR(i).Substring(5, 3) = "SSR" Then
                    pSSRType = "MANUAL SSR"
                    pSSRCode = pSSR(i).Substring(8, 4)
                    pCarrierCode = pSSR(i).Substring(12, 2)
                    pStatusCode = pSSR(i).Substring(15, 2)
                    pSpaces = pSSR(i).IndexOf("/")
                Else
                    pSSRType = "CARRIER RELATED"
                    pSSRCode = ""
                    pCarrierCode = pSSR(i).Substring(5, 2)
                    pSpaces = 9
                End If
                If pSpaces > 0 Then
                    For i1 As Integer = i + 1 To pSSR.GetUpperBound(0)
                        If pSSR(i1).StartsWith(Space(pSpaces)) Then
                            If pSSR(i).EndsWith("-") Then
                                pSSR(i) = pSSR(i).Substring(0, pSSR(i).Length - 1)
                            End If
                            pSSR(i) &= pSSR(i1).Substring(pSpaces)
                            pSSR(i1) = ""
                        Else
                            Exit For
                        End If
                    Next
                    pText = pSSR(i).Substring(pSpaces).TrimEnd
                    If pSSRCode = "DOCS" Then
                        Dim pTextItems() As String = pText.Split("/")
                        pPassportNumber = pTextItems(3)
                        pDateOfBirth = DateFromIATA(pTextItems(5))
                        pLastName = pTextItems(8)
                        pFirstName = pTextItems(9).Split("-")(0)
                    End If
                    mobjSSR.AddItem(pElementNo, pSSRType, pSSRCode, pCarrierCode, pStatusCode, pText, pLastName, pFirstName, pDateOfBirth, pPassportNumber)
                End If
            End If
        Next
    End Sub
    Public ReadOnly Property SSR As SSRCollection
        Get
            SSR = mobjSSR
        End Get
    End Property
    Private Sub GetOpenSegment1G()

        Dim pOpenSegs() As String = SendTerminalCommand("*IN")

        For i As Integer = 0 To pOpenSegs.GetUpperBound(0)
            If pOpenSegs(i).Length > 3 Then
                For j As Integer = i + 1 To pOpenSegs.GetUpperBound(0)
                    If pOpenSegs(j).StartsWith(Space(4)) Then
                        pOpenSegs(i) &= pOpenSegs(j).Substring(4)
                        pOpenSegs(j) = ""
                    Else
                        Exit For
                    End If
                Next
                Dim pElement As Short = 0
                Dim pStart As Short = pOpenSegs(i).IndexOf(".")

                If pStart > 0 And pOpenSegs(i).Substring(pStart + 2, 3) <> "HTL" _
                            And pOpenSegs(i).Substring(pStart + 2, 3) <> "CAR" _
                            And pOpenSegs(i).Substring(pStart + 2, 3) <> "SUR" Then '1G/PM0MMO   1GSW19CS
                    pElement = pOpenSegs(i).Substring(0, pStart).Trim
                    Dim pSegType As String = pOpenSegs(i).Substring(pStart + 2, 1)
                    Dim pRemType As String = pOpenSegs(i).Substring(pStart + 5, 13)
                    Dim pRemDate As Date = DateFromIATA(pOpenSegs(i).Substring(pStart + 18, 5))
                    Dim pRemark As String = pOpenSegs(i).Substring(pStart + 24).Trim
                    mobjOpenSegments.AddItem(pElement, pSegType, pRemType, pRemDate, pRemark)
                End If
            End If
        Next

    End Sub
    Public ReadOnly Property OpenSegments As OpenSegmentCollection
        Get
            OpenSegments = mobjOpenSegments
        End Get
    End Property
    Private Sub GetDI1G()

        Dim pDI() As String = SendTerminalCommand("*DI")

        If Not pDI(0).StartsWith("NO DOC") Then
            For i As Integer = 0 To pDI.GetUpperBound(0)
                Dim pElement As Short = 0
                Dim pCategory As String = ""
                Dim pRemark As String = ""
                If pDI(i).Length > 5 AndAlso Not pDI(i).StartsWith("     ") Then
                    If i < pDI.GetUpperBound(0) AndAlso pDI(i + 1).Length > 5 AndAlso pDI(i + 1).StartsWith("     ") Then
                        pDI(i) &= pDI(i + 1).Substring(5)
                        pDI(i + 1) = ""
                    End If
                    If pDI(i).StartsWith("DOCI-") Then
                        pElement = 1
                    Else
                        pElement = pDI(i).Substring(0, 3)
                    End If
                    If pDI(i).IndexOf("-", 5) > 5 Then
                        pCategory = pDI(i).Substring(5, pDI(i).IndexOf("-", 5) - 5)
                        pRemark = pDI(i).Substring(pDI(i).IndexOf("-", 5) + 1)
                        mobjDI.AddItem(pElement, pCategory, pRemark)
                    ElseIf pDI(i).IndexOf("*") > 4 Then
                        pCategory = pDI(i).Substring(5, pDI(i).IndexOf("*") - 4)
                        pRemark = pDI(i).Substring(pDI(i).IndexOf("*") + 1)
                        mobjDI.AddItem(pElement, pCategory, pRemark)
                    End If
                End If
            Next
        End If
    End Sub
    Public ReadOnly Property DIElements As DICollection
        Get
            DIElements = mobjDI
        End Get
    End Property
End Class
