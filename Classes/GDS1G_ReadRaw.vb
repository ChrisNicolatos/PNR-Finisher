Namespace GDS1G_ReadRaw
    Friend Class ReadRaw
        Event TerminalCommandSent(ByVal TerminalCommand As String)
        Private Structure PaxFFProps
            Dim PaxNumber As Short
            Dim Paxname As String
            Dim TicketNumber As String
            ReadOnly Property DocumentNumber As String
                Get
                    If TicketNumber.Length = 13 Then
                        DocumentNumber = TicketNumber.Substring(3)
                    Else
                        DocumentNumber = ""
                    End If
                End Get
            End Property
            ReadOnly Property Airline As String
                Get
                    If TicketNumber.Length = 13 Then
                        Airline = TicketNumber.Substring(0, 3)
                    Else
                        Airline = ""
                    End If
                End Get
            End Property
        End Structure
        Private Structure SegFFProps
            Dim SegNo As Short
            Dim BaggageAllowance As String
        End Structure
        Private mintRawIndex As Integer
        Private WithEvents mobjSession1G As New Travelport.TravelData.Factory.GalileoDesktopFactory("SPG720", "MYCONNECTION", False, True, "SMRT")

        Private mobjPassengers As New GDSPax.GDSPaxColl
        Private mobjSegments As New GDSSeg.GDSSegColl
        Private mobjTickets As New GDSTickets.GDSTicketCollection
        Private mobjPhones As New PhoneNumberCollection
        Private mobjEmails As New EmailCollection
        Private mobjOpenSegments As New OpenSegmentClass
        Private mobjDI As New DICollection
        Private mobjTicketElement As New TicketElementItem
        Private mobjOptionQueue As New OptionQueueCollection
        Private mobjSSR As New SSRCollection
        Private mobjFreqFlyer As New FrequentFlyer.FrequentFlyerColl

        Private mstrOfficeOfResponsibility As String
        Private mstrPNRNumber As String
        Private mflgNewPNR As Boolean
        Private mdteDepartureDate As Date
        Private mstrItinerary As String
        Private mflgExistsSegments As Boolean
        Private mSegsFirstElement As Integer
        Private mSegsLastElement As Integer
        Private mudtAllowance() As TQT
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
            GetSSR1G()
            GetDI1G()
            GetFreqFlyers()

        End Sub
        Private Function SendTerminalCommand(ByVal TerminalEntry As String) As String()
            Dim mstrPNR = mobjSession1G.SendTerminalCommand(TerminalEntry)
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
        Public ReadOnly Property Tickets As GDSTickets.GDSTicketCollection
            Get
                Tickets = mobjTickets
            End Get
        End Property
        Public ReadOnly Property Allowance As TQT()
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
        Private Sub GetPassengers1G()

            Dim pPax() As String = SendTerminalCommand("*N")
            Dim pAllPax As String = ""
            If pPax(0).IndexOf(".") >= 1 And pPax(0).IndexOf(".") <= 2 Then
                mstrPNRNumber = "New PNR"
                mflgNewPNR = True
            Else
                mstrPNRNumber = pPax(0).Substring(0, 6)
                mflgNewPNR = False
            End If
            For i As Integer = 0 To pPax.GetUpperBound(0)
                If pPax(i).IndexOf(".") >= 1 And pPax(i).IndexOf(".") <= 3 Then
                    pAllPax &= pPax(i) & " "
                End If
            Next
            pPax = pAllPax.Split("/")
            For i As Integer = 0 To pPax.GetUpperBound(0) - 1
                Dim iStart As Integer = 0
                If pPax(i).LastIndexOf(" ") >= 0 Then
                    iStart = pPax(i).LastIndexOf(" ") + 1
                End If
                Dim iPaxNo As Integer = pPax(i).Substring(iStart, pPax(i).IndexOf(".", iStart) - iStart)
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
            Dim pPassengerNumber As Short = 0
            Dim pFirstName As String = ""
            Dim pLastName As String = ""
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
        Public ReadOnly Property Passengers As GDSPax.GDSPaxColl
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
                Dim pArrivalDays As Integer = 0
                Dim pobjSeg As GDSSeg.GDSSegItem
                Dim pStart As Integer = pSegs(i).IndexOf(".")
                If pStart >= 1 And pSegs(i).Length >= 57 Then
                    With pSegs(i)
                        Dim pElementNo As Integer = .Substring(0, pStart).Trim
                        pCarrier = .Substring(pStart + 2, 2).Trim
                        pFlightNumber = .Substring(pStart + 5, 4).Trim
                        pClassOfService = .Substring(pStart + 10, 1).Trim
                        pDepartureDate = Utilities.DateFromIATA(.Substring(pStart + 13, 5))
                        pOrigin = .Substring(pStart + 19, 3).Trim
                        pDestination = .Substring(pStart + 22, 3).Trim
                        pStatus = .Substring(pStart + 26, 2).Trim
                        pDepartureTime = TimeSerial(.Substring(pStart + 31, 2), .Substring(pStart + 33, 2), 0)
                        pArrivalTime = TimeSerial(.Substring(pStart + 38, 2), .Substring(pStart + 40, 2), 0)
                        If .Substring(pStart + 37, 1) = "#" Then
                            pArrivalDays = +1
                        ElseIf .Substring(pStart + 37, 1) = "*" Then
                            pArrivalDays = +2
                        ElseIf .Substring(pStart + 37, 1) = "-" Then
                            pArrivalDays = -1
                        Else
                            pArrivalDays = 0
                        End If
                        pArrivalDate = DateAdd(DateInterval.Day, pArrivalDays, pDepartureDate)
                        pOperatedBy = ""
                        If i < pSegs.GetUpperBound(0) AndAlso .IndexOf(".") < 3 Then
                            pOperatedBy = pSegs(i + 1).Trim
                        End If
                        pobjSeg = mobjSegments.AddItem(pCarrier, pOrigin, pClassOfService, pDepartureDate, pArrivalDate, pElementNo, pFlightNumber, pDestination, pStatus, pDepartureTime, pArrivalTime, pVL, pSegs(i), ReadSVC(pElementNo))
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
            mflgExistsSegments = ((mobjSegments.Count) > 0)

            If mdteDepartureDate > Date.MinValue Then
                Dim pDate As New s1aAirlineDate.clsAirlineDate
                pDate.SetFromString(mdteDepartureDate)
                mstrItinerary &= " (" & pDate.IATA & ")"
            End If

        End Sub
        Private Function ReadSVC(ByVal ElementNo As String) As String()
            ReadSVC = SendTerminalCommand("*SVC" & ElementNo)
        End Function
        Public ReadOnly Property Segments As GDSSeg.GDSSegColl
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
                Dim pobjClass As New PhoneNumbersItem
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
        Public ReadOnly Property PhoneNumbers As PhoneNumberCollection
            Get
                PhoneNumbers = mobjPhones
            End Get
        End Property
        Private Sub GetEmailElement1G()

            Dim pEmails() As String = SendTerminalCommand("*EM")

            Dim pobjClass As New EmailItem
            Dim pElementAddress As Short = 0
            Dim pElementComment As Short = 0
            Dim pPrevElement As Short = 0
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
                    pActionDate = Utilities.DateFromIATA(pRem(0).Substring(pRem(0).Length - 5))
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
                pActionDateTime = Utilities.DateFromIATA(pItem(1).Substring(pItem(1).Length - 5)) + TimeSerial(pItem(2).Substring(0, 2), pItem(2).Substring(2), 0).TimeOfDay
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
                If pFF(i).StartsWith("P") And pFF(i).Substring(4, 1) = "." Then
                    pAirline = pFF(i).Substring(24, 2).Trim
                    pPaxName = pFF(i).Substring(6, 18).Trim
                    pMembershipNo = pFF(i).Substring(28).Trim
                    If i < pFF.Count - 1 AndAlso pFF(i + 1).StartsWith(Space(28)) Then
                        pCrossAccrual = pFF(i + 1).Substring(28).Trim
                    End If
                    mobjFreqFlyer.AddItem(pPaxName, pAirline, pMembershipNo, pCrossAccrual)
                End If
            Next i
        End Sub
        Public ReadOnly Property FrequentFlyers() As FrequentFlyer.FrequentFlyerColl
            Get
                FrequentFlyers = mobjFreqFlyer
            End Get
        End Property
        Private Sub GetTickets()
            Dim pFF() As String = SendTerminalCommand("*FF")
            ReDim mudtAllowance(0)
            mobjTickets.Clear()

            For i = 0 To pFF.GetUpperBound(0)
                If pFF(i).StartsWith("FQ") Or pFF(i).StartsWith("FB") Then
                    Dim pFFid As Integer = pFF(i).Substring(2, pFF(i).IndexOf(" ") - 2)
                    Dim pFFSegs As String = pFF(i).Substring(pFF(i).IndexOf("- S") + 3, pFF(i).IndexOf(" ", pFF(i).IndexOf("- S") + 4) - pFF(i).IndexOf("- S") - 2)

                    '
                    ' *FFx for each FF element
                    '
                    Dim pFFx() As String = SendTerminalCommand("*FF" & pFFid)
                    Dim pPax(0) As PaxFFProps
                    pPax(0).PaxNumber = 0
                    Dim pSeg(0) As SegFFProps
                    pSeg(0).SegNo = 0
                    Dim pPaxNo As Integer = 0
                    Dim pTicketNumber As String = ""
                    Dim pSegNo As Integer = 0
                    Dim pBaggageAllowance = ""
                    For iPFF As Integer = 0 To pFFx.GetUpperBound(0)
                        If pFFx(iPFF).Length > 13 AndAlso pFFx(iPFF).StartsWith(" P") AndAlso IsNumeric(pFFx(iPFF).Substring(2, 1)) Then
                            pPax(0).PaxNumber += 1
                            ReDim Preserve pPax(pPax(0).PaxNumber)
                            pPax(pPax(0).PaxNumber).PaxNumber = pFFx(iPFF).Substring(2, pFFx(iPFF).IndexOf(" ", 2))
                            If pFFx(iPFF).IndexOf(" ", 5) > 5 Then
                                pPax(pPax(0).PaxNumber).Paxname = pFFx(iPFF).Substring(5, pFFx(iPFF).IndexOf(" ", 5) - 4).Trim
                            Else
                                pPax(pPax(0).PaxNumber).Paxname = pFFx(iPFF).Substring(5).Trim
                            End If
                            If iPFF < pFFx.GetUpperBound(0) AndAlso pFFx(iPFF + 1).StartsWith(Space(13)) AndAlso IsNumeric(pFFx(iPFF + 1).Trim.Substring(pFFx(iPFF + 1).Trim.Length - 13)) AndAlso Not IsNumeric(pFFx(iPFF).Trim.Substring(pFFx(iPFF).Trim.Length - 13)) Then
                                pPax(pPax(0).PaxNumber).TicketNumber = pFFx(iPFF + 1).Trim.Substring(pFFx(iPFF + 1).Trim.Length - 13)
                                pFFx(iPFF + 1) = ""
                            ElseIf IsNumeric(pFFx(iPFF).Trim.Substring(pFFx(iPFF).Trim.Length - 13)) Then
                                pPax(pPax(0).PaxNumber).TicketNumber = pFFx(iPFF).Trim.Substring(pFFx(iPFF).Trim.Length - 13)
                            Else
                                pPax(pPax(0).PaxNumber).TicketNumber = ""
                            End If
                        ElseIf pFFx(iPFF).Length > 13 AndAlso pFFx(iPFF).StartsWith(" S") AndAlso IsNumeric(pFFx(iPFF).Substring(2, 1)) Then
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
                            mudtAllowance(mudtAllowance.GetUpperBound(0)) = New TQT
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
                        mobjTickets.addTicket("FA", 1, CDbl("0" & pPax(i1).DocumentNumber), 1, pPax(i1).Airline, Airlines.AirlineCode(pPax(i1).Airline), True, pTktSeg, pPax(i1).Paxname, "PAX")
                        mstrSeats &= GetSeats(pPax(i1).PaxNumber)

                    Next
                End If
            Next
        End Sub
        Private Function GetSeats(ByVal PaxNo As Short) As String
            Dim pSeats() As String = SendTerminalCommand("SC*P" & PaxNo)
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
                    pElementNo = pSSR(i).Substring(0, 3).Trim
                    If pSSR(i).Substring(5, 3) = "SSR" Then
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
                        pDateOfBirth = Utilities.DateFromIATA(pTextItems(5))
                        pLastName = pTextItems(8)
                        pFirstName = pTextItems(9).Split("-")(0)
                    End If
                    mobjSSR.AddItem(pElementNo, pSSRType, pSSRCode, pCarrierCode, pStatusCode, pText, pLastName, pFirstName, pDateOfBirth, pPassportNumber)
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
                        Dim pRemDate As Date = Utilities.DateFromIATA(pOpenSegs(i).Substring(pStart + 18, 5))
                        Dim pRemark As String = pOpenSegs(i).Substring(pStart + 24).Trim
                        mobjOpenSegments.AddItem(pElement, pSegType, pRemType, pRemDate, pRemark)
                    End If
                End If
            Next

        End Sub
        Public ReadOnly Property OpenSegments As OpenSegmentClass
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
                        pCategory = pDI(i).Substring(5, pDI(i).IndexOf("-", 5) - 5)
                        pRemark = pDI(i).Substring(pDI(i).IndexOf("-", 5) + 1)
                        mobjDI.AddItem(pElement, pCategory, pRemark)
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
    Friend Class PhoneNumbersItem
        Private Structure ClassProps
            Dim ElementNo As Short
            Dim CityCode As String
            Dim PhoneType As String
            Dim PhoneNumber As String
        End Structure
        Dim mudtProps As ClassProps
        Public ReadOnly Property ElementNo As Short
            Get
                ElementNo = mudtProps.ElementNo
            End Get
        End Property
        Public ReadOnly Property CityCode As String
            Get
                CityCode = mudtProps.CityCode
            End Get
        End Property
        Public ReadOnly Property PhoneType As String
            Get
                PhoneType = mudtProps.PhoneType
            End Get
        End Property
        Public ReadOnly Property PhoneNumber As String
            Get
                PhoneNumber = mudtProps.PhoneNumber
            End Get
        End Property
        Friend Sub SetValues(ByVal pElementNo As Short, ByVal pCityCode As String, ByVal pPhoneType As String, ByVal pPhoneNumber As String)
            With mudtProps
                .ElementNo = pElementNo
                .CityCode = pCityCode
                .PhoneType = pPhoneType
                .PhoneNumber = pPhoneNumber
            End With
        End Sub
    End Class
    Friend Class PhoneNumberCollection
        Inherits Collections.Generic.Dictionary(Of Short, PhoneNumbersItem)
        Public Sub AddItem(ByVal pElementNo As Short, ByVal pCityCode As String, ByVal pPhoneType As String, ByVal pPhoneNumber As String)
            Dim pobjClass As New PhoneNumbersItem
            pobjClass.SetValues(pElementNo, pCityCode, pPhoneType, pPhoneNumber)
            MyBase.Add(pobjClass.ElementNo, pobjClass)
        End Sub
    End Class
    Friend Class EmailItem
        Private Structure ClassProps
            Dim ElementNo As Short
            Dim EmailAddress As String
            Dim EmailComment As String
        End Structure
        Private mudtProps As ClassProps
        Public ReadOnly Property ElementNo As Short
            Get
                ElementNo = mudtProps.ElementNo
            End Get
        End Property
        Public ReadOnly Property EmailAddress As String
            Get
                EmailAddress = mudtProps.EmailAddress
            End Get
        End Property
        Public ReadOnly Property EmailComment As String
            Get
                EmailComment = mudtProps.EmailComment
            End Get
        End Property
        Friend Sub SetValues(ByVal pElementNo As Short, ByVal pEmailAddress As String, ByVal pEmailComment As String)
            With mudtProps
                .ElementNo = pElementNo
                .EmailAddress = pEmailAddress
                .EmailComment = pEmailComment
            End With
        End Sub
    End Class
    Friend Class EmailCollection
        Inherits Collections.Generic.Dictionary(Of Short, EmailItem)
        Private mFromAddress As String
        Public Sub SetFromAddress(ByVal pFromAddress As String)
            mFromAddress = pFromAddress
        End Sub
        Public ReadOnly Property FromAddress As String
            Get
                FromAddress = mFromAddress
            End Get
        End Property
        Public Sub AddItem(ByVal pElementNo As Short, ByVal pEmailAddress As String, ByVal pEmailComment As String)
            Dim pobjClass As New EmailItem
            pobjClass.SetValues(pElementNo, pEmailAddress, pEmailComment)
            MyBase.Add(pobjClass.ElementNo, pobjClass)
        End Sub
    End Class
    Friend Class OpenSegmentItem
        Private Structure ClassProps
            Dim ElementNo As Short
            Dim SegmentType As String
            Dim RemarkType As String
            Dim RemarkDate As Date
            Dim Remark As String
        End Structure
        Private mudtProps As ClassProps
        Public ReadOnly Property ElementNo As Short
            Get
                ElementNo = mudtProps.ElementNo
            End Get
        End Property
        Public ReadOnly Property SegmentType As String
            Get
                SegmentType = mudtProps.SegmentType
            End Get
        End Property
        Public ReadOnly Property RemarkType As String
            Get
                RemarkType = mudtProps.RemarkType
            End Get
        End Property
        Public ReadOnly Property RemarkDate As Date
            Get
                RemarkDate = mudtProps.RemarkDate
            End Get
        End Property
        Public ReadOnly Property Remark As String
            Get
                Remark = mudtProps.Remark
            End Get
        End Property
        Friend Sub SetValues(ByVal pElementNo As Short, ByVal pSegmentType As String, ByVal pRemarkType As String, ByVal pRemarkDate As Date, ByVal pRemark As String)
            With mudtProps
                .ElementNo = pElementNo
                .SegmentType = pSegmentType
                .RemarkType = pRemarkType
                .RemarkDate = pRemarkDate
                .Remark = pRemark
            End With
        End Sub
    End Class
    Friend Class OpenSegmentClass
        Inherits Collections.Generic.Dictionary(Of Short, OpenSegmentItem)
        Public Sub AddItem(ByVal pElementNo As Short, ByVal pSegmentType As String, ByVal pRemarkType As String, ByVal pRemarkDate As Date, ByVal pRemark As String)
            Dim pobjClass As New OpenSegmentItem
            pobjClass.SetValues(pElementNo, pSegmentType, pRemarkType, pRemarkDate, pRemark)
            MyBase.Add(pobjClass.ElementNo, pobjClass)
        End Sub
    End Class
    Friend Class DIItem
        Private Structure ClassProps
            Dim ElementNo As Short
            Dim Category As String
            Dim CategoryDescription As String
            Dim Remark As String
        End Structure
        Private mudtProps As ClassProps
        Public ReadOnly Property ElementNo As Short
            Get
                ElementNo = mudtProps.ElementNo
            End Get
        End Property
        Public ReadOnly Property Category As String
            Get
                Category = mudtProps.Category
            End Get
        End Property
        Public ReadOnly Property CategoryDescription As String
            Get
                CategoryDescription = mudtProps.CategoryDescription
            End Get
        End Property

        Public ReadOnly Property Remark As String
            Get
                Remark = mudtProps.Remark
            End Get
        End Property
        Friend Sub SetValues(ByVal pElementNo As Short, ByVal pCategory As String, ByVal pRemark As String)
            With mudtProps
                .ElementNo = pElementNo
                .CategoryDescription = pCategory
                Select Case pCategory
                    Case "FREE TEXT"
                        .Category = "FT"
                    Case Else
                        .Category = pCategory
                End Select
                .Remark = pRemark
            End With
        End Sub
    End Class
    Friend Class DICollection
        Inherits Collections.Generic.Dictionary(Of Short, DIItem)
        Public Sub AddItem(ByVal pElementNo As Short, ByVal pCategory As String, ByVal pRemark As String)
            Dim pobjClass As New DIItem
            pobjClass.SetValues(pElementNo, pCategory, pRemark)
            MyBase.Add(pobjClass.ElementNo, pobjClass)
        End Sub
    End Class
    Friend Class TicketElementItem
        Private Structure ClassProps
            Dim ElementNo As Short
            Dim PCC As String
            Dim ActionDateTime As Date
            Dim Remark As String
        End Structure
        Private mudtProps As ClassProps
        Public ReadOnly Property ElementNo As Short
            Get
                ElementNo = mudtProps.ElementNo
            End Get
        End Property
        Public ReadOnly Property PCC As String
            Get
                PCC = mudtProps.PCC
            End Get
        End Property
        Public ReadOnly Property ActionDateTime As Date
            Get
                ActionDateTime = mudtProps.ActionDateTime
            End Get
        End Property
        Public ReadOnly Property Remark As String
            Get
                Remark = mudtProps.Remark
            End Get
        End Property
        Friend Sub SetValues(ByVal pElementNo As Short, ByVal pPCC As String, ByVal pActionDateTime As Date, ByVal pRemark As String)
            With mudtProps
                .ElementNo = pElementNo
                .PCC = pPCC
                .ActionDateTime = pActionDateTime
                .Remark = pRemark
            End With
        End Sub
        Friend Sub Clear()
            With mudtProps
                .ElementNo = 0
                .PCC = ""
                .ActionDateTime = Now
                .Remark = ""
            End With
        End Sub
    End Class
    Friend Class OptionQueueItem
        Private Structure ClassProps
            Dim ElementNo As Short
            Dim PCC As String
            Dim ActionDateTime As Date
            Dim QueueNumber As String
            Dim Remark As String
        End Structure
        Private mudtProps As ClassProps
        Public ReadOnly Property ElementNo As Short
            Get
                ElementNo = mudtProps.ElementNo
            End Get
        End Property
        Public ReadOnly Property PCC As String
            Get
                PCC = mudtProps.PCC
            End Get
        End Property
        Public ReadOnly Property ActionDateTime As Date
            Get
                ActionDateTime = mudtProps.ActionDateTime
            End Get
        End Property
        Public ReadOnly Property QueueNumber As String
            Get
                QueueNumber = mudtProps.QueueNumber
            End Get
        End Property
        Public ReadOnly Property Remark As String
            Get
                Remark = mudtProps.Remark
            End Get
        End Property
        Friend Sub SetValues(ByVal pElementNo As Short, ByVal pPCC As String, ByVal pActionDateTime As Date, ByVal pQueueNumber As String, ByVal pRemark As String)
            With mudtProps
                .ElementNo = pElementNo
                .PCC = pPCC
                .ActionDateTime = pActionDateTime
                .QueueNumber = pQueueNumber
                .Remark = pRemark
            End With
        End Sub
    End Class
    Friend Class OptionQueueCollection
        Inherits Collections.Generic.Dictionary(Of Short, OptionQueueItem)
        Public Sub AddItem(ByVal pElementNo As Short, ByVal pPCC As String, ByVal pActionDateTime As Date, ByVal pQueueNumber As String, ByVal pRemark As String)
            Dim pobjClass As New OptionQueueItem
            pobjClass.SetValues(pElementNo, pPCC, pActionDateTime, pQueueNumber, pRemark)
            MyBase.Add(pobjClass.ElementNo, pobjClass)
        End Sub
    End Class
    Friend Class SSRitem
        Private Structure ClassProps
            Dim ElementNo As Short
            Dim SSRType As String
            Dim SSRCode As String
            Dim CarrierCode As String
            Dim StatusCode As String
            Dim Text As String
            Dim LastName As String
            Dim FirstName As String
            Dim DateOfBirth As Date
            Dim PassportNumber As String
        End Structure
        Private mudtProps As ClassProps
        Public ReadOnly Property ElementNo As Short
            Get
                ElementNo = mudtProps.ElementNo
            End Get
        End Property
        Public ReadOnly Property SSRType As String
            Get
                SSRType = mudtProps.SSRType
            End Get
        End Property
        Public ReadOnly Property SSRCode As String
            Get
                SSRCode = mudtProps.SSRCode
            End Get
        End Property
        Public ReadOnly Property CarrierCode As String
            Get
                CarrierCode = mudtProps.CarrierCode
            End Get
        End Property
        Public ReadOnly Property StatusCode As String
            Get
                StatusCode = mudtProps.StatusCode
            End Get
        End Property
        Public ReadOnly Property Text As String
            Get
                Text = mudtProps.Text
            End Get
        End Property
        Public ReadOnly Property LastName As String
            Get
                LastName = mudtProps.LastName
            End Get
        End Property
        Public ReadOnly Property FirstName As String
            Get
                FirstName = mudtProps.FirstName
            End Get
        End Property
        Public ReadOnly Property DateOfBirth As Date
            Get
                DateOfBirth = mudtProps.DateOfBirth
            End Get
        End Property
        Public ReadOnly Property PassportNumber As String
            Get
                PassportNumber = mudtProps.PassportNumber
            End Get
        End Property

        Friend Sub SetValues(ByVal pElementNo As Short, ByVal pSSRType As String, ByVal pSSRCode As String, ByVal pCarrierCode As String _
                             , ByVal pStatusCode As String, ByVal pText As String, ByVal pLastName As String, ByVal pFirstname As String _
                             , ByVal pDateOfBirth As Date, ByVal pPassportNumber As String)
            With mudtProps
                .ElementNo = pElementNo
                .SSRType = pSSRType
                .SSRCode = pSSRCode
                .CarrierCode = pCarrierCode
                .StatusCode = pStatusCode
                .Text = pText
                .LastName = pLastName
                .FirstName = pFirstname
                .DateOfBirth = pDateOfBirth
                .PassportNumber = pPassportNumber
            End With
        End Sub
    End Class
    Friend Class SSRCollection
        Inherits Collections.Generic.Dictionary(Of Short, SSRitem)
        Public Sub AddItem(ByVal pElementNo As Short, ByVal pSSRType As String, ByVal pSSRCode As String, ByVal pCarrierCode As String _
                             , ByVal pStatusCode As String, ByVal pText As String, ByVal pLastName As String, ByVal pFirstname As String _
                             , ByVal pDateOfBirth As Date, ByVal pPassportNumber As String)
            Dim pobjClass As New SSRitem
            pobjClass.SetValues(pElementNo, pSSRType, pSSRCode, pCarrierCode _
                             , pStatusCode, pText, pLastName, pFirstname _
                             , pDateOfBirth, pPassportNumber)
            MyBase.Add(pobjClass.ElementNo, pobjClass)
        End Sub
    End Class
End Namespace