Option Strict Off
Option Explicit On
Imports k1aHostToolKit

Friend Class GDSReadPNR
    Public Event ReadStatus(ByRef Status As Short, ByRef StatusDescription As String)
    Public Event TerminalCommandSent(ByVal TerminalCommand As String)
    Private Structure LineNumbers
        Dim Category As String
        Dim LineNumber As Integer
    End Structure
    Private Structure ClassProps
        Dim RequestedPNR As String
        Dim UserSignIn As String
        Dim PNRCreationdate As Date
        Dim Seats As String

        Dim isDirty As Boolean
        Dim isValid As Boolean
        Dim isNew As Boolean
        Friend Sub Clear()
            RequestedPNR = ""
            UserSignIn = ""
            PNRCreationdate = Date.MinValue
            Seats = ""
            isDirty = False
            isValid = False
            isNew = True
        End Sub
    End Structure
    'Private Structure TQTProps
    '    Dim TQTElement As Integer
    '    Dim Segment As Integer
    '    Dim Itin As String
    '    Dim Allowance As String
    '    Dim Pax As String
    '    Dim Status As String
    'End Structure
    Private WithEvents mobjSession1A As k1aHostToolKit.HostSession
    Private mobjPNR1A As s1aPNR.PNR
    Private mobjSession1G As New Travelport.TravelData.Factory.GalileoDesktopFactory("SPG720", "MYCONNECTION", False, True, "SMRT")
    Private WithEvents mobjPNR1GRaw As New GDS1G_ReadRaw.ReadRaw

    Private mobjPassengers As New GDSPax.GDSPaxColl
    Private mobjSegments As New GDSSeg.GDSSegColl
    Private mobjTickets As New GDSTickets.GDSTicketCollection

    Private mobjFrequentFlyer As New FrequentFlyer.FrequentFlyerColl
    Private mobjNumberParser As New GDSNumberParser

    Private mobjExistingGDSElements As New GDSExisting.Collection
    Private mobjNewGDSElements As New GDSNew.Collection
    Private mGDSCode As Utilities.EnumGDSCode

    Private mudtProps As ClassProps
    Private mudtTQT() As TQT
    Private mudtAllowance() As TQT

    Private mstrPNRResponse As String
    Private mstrPNRNumber As String
    Private mflgNewPNR As Boolean
    Private mstrGroupName As String
    Private mintGroupNamesCount As Integer
    Private mstrItinerary As String

    Private mstrOfficeOfResponsibility As String
    Private mflgQRSegment As Boolean = False
    Private mdteDepartureDate As Date
    Private mflgExistsSegments As Boolean
    Private mflgExistsSSRDocs As Boolean
    Private mstrSSRDocs As String

    Private mSegsFirstElement As Integer
    Private mSegsLastElement As Integer
    Private mstrVesselName As String
    Private mstrBookedBy As String
    Private mstrCC As String
    Private mstrCLN As String
    Private mstrCLA As String
    Private mflgCancelError As Boolean

    Private mintStatus As Short
    Private mstrStatus As String

    Public Sub New()
        mobjPNR1GRaw = New GDS1G_ReadRaw.ReadRaw

        mobjPassengers.Clear()
        mobjSegments.Clear()
        mobjTickets.Clear()

        mobjFrequentFlyer.Clear()


        mobjExistingGDSElements.Clear()
        mobjNewGDSElements.Clear()

        mudtProps.Clear()
        ReDim mudtTQT(0)
        mudtTQT(0) = New TQT
        ReDim mudtAllowance(0)
        mudtAllowance(0) = New TQT

        mstrPNRResponse = ""
        mstrPNRNumber = ""
        mflgNewPNR = False
        mstrGroupName = ""
        mintGroupNamesCount = 0
        mstrItinerary = ""

        mstrOfficeOfResponsibility = ""
        mflgQRSegment = False = False
        mdteDepartureDate = Date.MinValue
        mflgExistsSegments = False
        mflgExistsSSRDocs = False
        mstrSSRDocs = ""

        mSegsFirstElement = 0
        mSegsLastElement = 0
        mstrVesselName = ""
        mstrBookedBy = ""
        mstrCC = ""
        mstrCLN = ""
        mstrCLA = ""
        mflgCancelError = False

        mintStatus = 0
        mstrStatus = ""
    End Sub

    Private Sub mobjSession_ReceivedResponse(ByRef newResponse As CHostResponse) Handles mobjSession1A.ReceivedResponse
        mstrPNRResponse = newResponse.Text
    End Sub

    Public ReadOnly Property Segments As GDSSeg.GDSSegColl
        Get
            Segments = mobjSegments
        End Get
    End Property
    Public ReadOnly Property Passengers As GDSPax.GDSPaxColl
        Get
            Passengers = mobjPassengers
        End Get
    End Property
    Public ReadOnly Property AllowanceForSegment(ByVal Origin As String, ByVal Destination As String, ByVal Airline As String) As String
        Get
            AllowanceForSegment = ""
            If Not IsNothing(mudtAllowance) Then
                For i As Integer = 1 To mudtAllowance.GetUpperBound(0)
                    If mudtAllowance(i).Itin = Origin & " " & Airline & " " & Destination Then
                        AllowanceForSegment = mudtAllowance(i).Allowance
                    End If
                Next
            End If
        End Get
    End Property
    Public ReadOnly Property AllowanceForSegment(ByVal PaxNo As Short, ByVal SegNo As Short) As String
        Get
            AllowanceForSegment = ""
            If Not IsNothing(mudtAllowance) Then
                For i As Integer = 1 To mudtAllowance.GetUpperBound(0)
                    If mudtAllowance(i).Pax = PaxNo And mudtAllowance(i).Segment = SegNo Then
                        AllowanceForSegment = mudtAllowance(i).Allowance
                    End If
                Next
            End If
        End Get
    End Property
    Public ReadOnly Property AllowanceForSegment(ByVal SegNo As Short) As String
        Get
            AllowanceForSegment = ""
            If Not IsNothing(mudtAllowance) Then
                For i As Integer = 1 To mudtAllowance.GetUpperBound(0)
                    If mudtAllowance(i).Segment = SegNo Then
                        If AllowanceForSegment.IndexOf(mudtAllowance(i).Allowance) = -1 Then
                            If AllowanceForSegment.Length > 0 Then
                                AllowanceForSegment &= "/"
                            End If
                            AllowanceForSegment &= mudtAllowance(i).Allowance
                        End If
                    End If
                Next
            End If
        End Get
    End Property

    Public ReadOnly Property GroupName As String
        Get
            GroupName = mstrGroupName
        End Get
    End Property
    Public ReadOnly Property GroupNamesCount As Integer
        Get
            GroupNamesCount = mintGroupNamesCount
        End Get
    End Property
    Public ReadOnly Property NumberOfPax As Integer
        Get
            NumberOfPax = mobjPassengers.Count
        End Get
    End Property
    Public ReadOnly Property PaxLeadName As String
        Get
            PaxLeadName = mobjPassengers.LeadName
        End Get
    End Property
    Public ReadOnly Property IsGroup As Boolean
        Get
            IsGroup = (mstrGroupName <> "")
        End Get
    End Property
    Public ReadOnly Property HasSegments As Boolean
        Get
            HasSegments = (mSegsLastElement > -1)
        End Get
    End Property
    Public ReadOnly Property FirstSegment As GDSSeg.GDSSegItem
        Get
            If mSegsFirstElement = -1 Then
                FirstSegment = New GDSSeg.GDSSegItem
            Else
                FirstSegment = mobjSegments(Format(mSegsFirstElement))
            End If
        End Get
    End Property
    Public ReadOnly Property LastSegment As GDSSeg.GDSSegItem
        Get
            If mSegsLastElement = -1 Then
                LastSegment = New GDSSeg.GDSSegItem
            Else
                LastSegment = mobjSegments(Format(mSegsLastElement))
            End If
        End Get
    End Property
    Public ReadOnly Property Itinerary As String
        Get
            Itinerary = mstrItinerary
        End Get
    End Property
    Public ReadOnly Property Tickets() As GDSTickets.GDSTicketCollection
        Get
            Tickets = mobjTickets
        End Get
    End Property
    Public ReadOnly Property FrequentFlyerNumber(ByVal Airline As String, ByVal PaxName As String) As String
        Get
            FrequentFlyerNumber = ""
            For Each pItem As FrequentFlyer.FrequentFlyerItem In mobjFrequentFlyer.Values
                If pItem.PaxName.StartsWith(PaxName) Or PaxName.StartsWith(pItem.PaxName) Then
                    Dim pAirlineCode = Airlines.AirlineCode(Airline)
                    If pItem.Airline = Airline Or pItem.Airline = pAirlineCode Then
                        FrequentFlyerNumber = pItem.Airline & " " & pItem.FrequentTravelerNo
                        Exit For
                    ElseIf pItem.CrossAccrual = Airline Or pItem.CrossAccrual = pAirlineCode Then
                        FrequentFlyerNumber = pItem.Airline & " " & pItem.FrequentTravelerNo & " (Cross Accrual: " & pItem.CrossAccrual & ")"
                    End If
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
    Public ReadOnly Property BookedBy As String
        Get
            BookedBy = mstrBookedBy
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
    Public ReadOnly Property PnrNumber As String
        Get
            PnrNumber = mstrPNRNumber
        End Get
    End Property
    Public ReadOnly Property OfficeOfResponsibility As String
        Get
            OfficeOfResponsibility = mstrOfficeOfResponsibility
        End Get
    End Property
    Public ReadOnly Property DepartureDate As Date
        Get
            DepartureDate = mdteDepartureDate
        End Get
    End Property
    Public ReadOnly Property ExistingElements As GDSExisting.Collection
        Get
            ExistingElements = mobjExistingGDSElements
        End Get
    End Property
    Public ReadOnly Property NewElements As GDSNew.Collection
        Get
            NewElements = mobjNewGDSElements
        End Get
    End Property
    Public ReadOnly Property HasQRSegment As Boolean
        Get
            HasQRSegment = mflgQRSegment
        End Get
    End Property
    Public ReadOnly Property SegmentsExist As Boolean
        Get
            SegmentsExist = mflgExistsSegments
        End Get
    End Property
    Public ReadOnly Property SSRDocsExists As Boolean
        Get
            SSRDocsExists = mflgExistsSSRDocs
        End Get
    End Property
    Public ReadOnly Property SSRDocs As String
        Get
            SSRDocs = mstrSSRDocs
        End Get
    End Property
    Public ReadOnly Property GDSAbbreviation As String
        Get
            If GDSCode = Utilities.EnumGDSCode.Amadeus Then
                GDSAbbreviation = "1A"
            ElseIf GDSCode = Utilities.EnumGDSCode.Galileo Then
                GDSAbbreviation = "1G"
            Else
                GDSAbbreviation = ""
            End If
        End Get
    End Property
    Public ReadOnly Property GDSCode As Utilities.EnumGDSCode
        Get
            GDSCode = mGDSCode
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
    Public ReadOnly Property MaxAirportNameLength As Integer
        Get
            MaxAirportNameLength = mobjSegments.MaxAirportNameLength
        End Get
    End Property
    Public ReadOnly Property MaxCityNameLength As Integer
        Get
            MaxCityNameLength = mobjSegments.MaxCityNameLength
        End Get
    End Property
    Public ReadOnly Property MaxAirportShortNameLength As Integer
        Get
            MaxAirportShortNameLength = mobjSegments.MaxAirportShortNameLength
        End Get
    End Property

    Public ReadOnly Property NewPNR As Boolean
        Get
            NewPNR = mflgNewPNR
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
                mobjSession1A = pobjHostSessions.UIActiveSession

                If Queue <> "" Then
                    mobjSession1A.Send("QI")
                    mobjSession1A.Send("IG")
                End If
                pQV &= mobjSession1A.Send("QV/" & Queue).Text
                Do While pQV.IndexOf(")>") = pQV.Length - 4
                    pQV &= mobjSession1A.Send("MDR").Text
                Loop
                Dim pLines() As String = pQV.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
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
    Public Function Read(ByVal pGDSCode As Utilities.EnumGDSCode, ByVal PNR As String, ByVal ForReportOnly As Boolean) As Boolean
        ' from GDSPnr
        mGDSCode = pGDSCode

        If mGDSCode = Utilities.EnumGDSCode.Amadeus Then
            Read1A(PNR, ForReportOnly)
        ElseIf mGDSCode = Utilities.EnumGDSCode.Galileo Then
            ReadPNR1G(PNR, ForReportOnly)
        Else
            Throw New Exception("Incorrect GDS")
        End If

    End Function
    Public Function Read(ByVal GDSCode As Utilities.EnumGDSCode) As String
        ' from ReadPNR
        mGDSCode = GDSCode

        If mGDSCode = Utilities.EnumGDSCode.Amadeus Then
            Read = Read1A()
        ElseIf mGDSCode = Utilities.EnumGDSCode.Galileo Then
            Read = Read1G()
        Else
            Throw New Exception("ReadPNR.Read()" & vbCrLf & "NO GDS Specified")
        End If
    End Function
    Private Function Read1A(ByVal PNR As String, ByVal ForReportOnly As Boolean) As Boolean
        ' from GDSPnr
        Dim pobjHostSessions As k1aHostToolKit.HostSessions

        Try
            mstrStatus = ""
            pobjHostSessions = New k1aHostToolKit.HostSessions

            If pobjHostSessions.Count > 0 Then
                mobjSession1A = pobjHostSessions.UIActiveSession

                If PNR <> "" Then
                    mobjSession1A.Send("QI")
                    mobjSession1A.Send("IG")
                End If
                mudtProps.RequestedPNR = PNR
                Read1A = RetrievePNR1A(ForReportOnly)
            Else
                Throw New Exception("Amadeus not signed in")
            End If

            If Read1A Then
                mintStatus = 0
                mstrStatus = "Amadeus read " & PNR & " OK"
            Else
                mintStatus = 1
                mstrStatus = "Amadeus " & PNR & " not found"
            End If
            mobjSession1A.SendSpecialKey(512 + 282) '(k1aHostConstantsLib.AmaKeyValues.keySHIFT + k1aHostConstantsLib.AmaKeyValues.keyPause)
            mobjSession1A.Send("RT")
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

    Private Function Read1A() As String

        If mGDSCode <> Utilities.EnumGDSCode.Amadeus Then
            Throw New Exception("ReadPNR.Read1A()" & vbCrLf & "Selected GDS is not Amadeus")
        End If

        Dim Sessions As k1aHostToolKit.HostSessions

        Read1A = ""

        Try
            ' To be able to retrieve the PNR that have been created we need to link our '
            ' application to the current session of the FOS                             '
            Sessions = New k1aHostToolKit.HostSessions

            If Sessions.Count > 0 Then
                ' There is at least one session opened.                    '
                ' We link our application to the active session of the FOS '
                mobjSession1A = Sessions.UIActiveSession

                ' Initialize the PNR
                mobjPNR1A = New s1aPNR.PNR

                ' Retrieve the name elements, Air segments and Hotel Segments of the current PNR
                Dim pStatus As Integer = mobjPNR1A.RetrievePNR(mobjSession1A, "RT")
                mflgNewPNR = False

                If pStatus = 0 Or pStatus = 1005 Then
                    GetOfficeOfResponsibility1A()
                    GetPnrNumber1A()

                    GetGroup1A()
                    GetPassengers1A()
                    GetSegments1A()
                    GetPhoneElement1A()
                    GetEmailElement1A()
                    GetAOH1A()
                    GetOpenSegment1A()
                    GetTicketElement1A()
                    GetOptionQueueElement1A()
                    GetVesselOSI1A()
                    GetSSR1A()
                    GetRM1A()
                    GetTickets1A()
                    If mobjPNR1A.RawResponse.IndexOf("***  NHP  ***") >= 0 Then
                        Read1A = "               ***  NHP  ***"
                    Else
                        Read1A = CheckDMI1A()
                    End If
                Else
                    Throw New Exception("There is no active PNR" & vbCrLf & mstrPNRResponse)
                End If
            Else
                Throw New Exception("Please start Amadeus and retry")
            End If
        Catch ex As Exception
            Throw New Exception("Read1A()" & vbCrLf & ex.Message)
        End Try
    End Function
    Private Sub ReadPNR1G(ByVal PNR As String, ByVal ForReportOnly As Boolean)
        Try

            If PNR <> "" Then
                mobjSession1G.SendTerminalCommand("QXI+I")
            End If
            mudtProps.RequestedPNR = PNR
            Read1G()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Function Read1G() As String

        mobjPNR1GRaw = New GDS1G_ReadRaw.ReadRaw
        Dim pResponse = mobjSession1G.SendTerminalCommand("*R")
        If pResponse(0).Substring(6, 1) = "/" Then
            mudtProps.RequestedPNR = pResponse(0).Substring(0, 6)
        ElseIf Not pResponse(0).StartsWith(" ") Then
            Throw New Exception(pResponse(0))
        Else
            mudtProps.RequestedPNR = ""
        End If
        Read1G = ""
        If mudtProps.RequestedPNR.Trim <> "" Then
            mobjSession1G.SendTerminalCommand("*" & mudtProps.RequestedPNR)
        End If
        Try
            mobjTickets = New GDSTickets.GDSTicketCollection
            mstrVesselName = ""
            mstrBookedBy = ""
            mstrCC = ""
            mstrCLA = ""
            mstrCLN = ""

            mobjPNR1GRaw.ReadRaw(mudtProps.RequestedPNR)
            mudtProps.RequestedPNR = mobjPNR1GRaw.RequestedPNR
            mstrOfficeOfResponsibility = mobjPNR1GRaw.OfficeOfResponsibility
            mobjPassengers = mobjPNR1GRaw.Passengers
            mobjSegments = mobjPNR1GRaw.Segments
            mobjFrequentFlyer = mobjPNR1GRaw.FrequentFlyers
            mstrItinerary = mobjSegments.Itinerary
            mdteDepartureDate = mobjPNR1GRaw.DepartureDate
            mflgExistsSegments = (mobjSegments.Count > 0)
            mSegsFirstElement = mobjPNR1GRaw.SegsFirstElement
            mSegsLastElement = mobjPNR1GRaw.SegsLastElement
            mudtAllowance = mobjPNR1GRaw.Allowance
            mobjTickets = mobjPNR1GRaw.Tickets
            mudtProps.Seats = mobjPNR1GRaw.Seats

            GetPhoneElement1G()
            GetEmailElement1G()
            GetTicketElement1G()
            GetOpenSegment1G()
            GetOptionQueueElement1G()
            GetSSR1G()
            GetRM1G()

        Catch ex As Exception
            Throw New Exception("Read1G()" & vbCrLf & ex.Message)
        End Try

    End Function
    Public Sub PrepareNewGDSElements()
        mobjNewGDSElements = New GDSNew.Collection(OfficeOfResponsibility, DepartureDate, NumberOfPax, mGDSCode)
    End Sub
    Private Function CheckDMI1A() As String
        If mGDSCode <> Utilities.EnumGDSCode.Amadeus Then
            Throw New Exception("ReadPNR.CheckDMI1A()" & vbCrLf & "Selected GDS is not Amadeus")
        End If

        Try
            If mobjPNR1A.AirSegments.Count <= 1 Then
                Return ""
            End If

            Dim pDMI As String = mobjSession1A.Send("DMI").Text
            If pDMI.Contains("ITINERARY OK") Then
                Return ""
            Else
                Return pDMI
            End If
        Catch ex As Exception
            Return ""
        End Try

    End Function
    Private Sub RemoveOldGDSEntries1A()

        If mGDSCode <> Utilities.EnumGDSCode.Amadeus Then
            Throw New Exception("ReadPNR.RemoveOldGDSEntries1A()" & vbCrLf & "Selected GDS is not Amadeus")
        End If

        Dim pLineNumbers(0) As Integer

        ' the following elements remain as they are if they already exist in the PNR
        ClearExistingItems(mobjExistingGDSElements.PhoneElement, mobjNewGDSElements.PhoneElement)
        ClearExistingItems(mobjExistingGDSElements.EmailElement, mobjNewGDSElements.EmailElement)
        ClearExistingItems(mobjExistingGDSElements.AOH, mobjNewGDSElements.AOH)
        ClearExistingItems(mobjExistingGDSElements.OpenSegment, mobjNewGDSElements.OpenSegment)
        ClearExistingItems(mobjExistingGDSElements.OptionQueueElement, mobjNewGDSElements.OptionQueueElement)
        ClearExistingItems(mobjExistingGDSElements.TicketElement, mobjNewGDSElements.TicketElement)
        ClearExistingItems(mobjExistingGDSElements.AgentID, mobjNewGDSElements.AgentID)

        ' the following elements are removed and replaced if they exist in the PNR
        Utilities1A.PrepareLineNumbers1A(mobjExistingGDSElements.CustomerCode, pLineNumbers)
        Utilities1A.PrepareLineNumbers1A(mobjExistingGDSElements.CustomerName, pLineNumbers)
        Utilities1A.PrepareLineNumbers1A(mobjExistingGDSElements.SubDepartmentCode, pLineNumbers)
        Utilities1A.PrepareLineNumbers1A(mobjExistingGDSElements.SubDepartmentName, pLineNumbers)
        Utilities1A.PrepareLineNumbers1A(mobjExistingGDSElements.CRMCode, pLineNumbers)
        Utilities1A.PrepareLineNumbers1A(mobjExistingGDSElements.CRMName, pLineNumbers)
        Utilities1A.PrepareLineNumbers1A(mobjExistingGDSElements.VesselFlag, pLineNumbers)
        Utilities1A.PrepareLineNumbers1A(mobjExistingGDSElements.VesselName, pLineNumbers)
        Utilities1A.PrepareLineNumbers1A(mobjExistingGDSElements.VesselOSI, pLineNumbers)
        Utilities1A.PrepareLineNumbers1A(mobjExistingGDSElements.Reference, pLineNumbers)
        Utilities1A.PrepareLineNumbers1A(mobjExistingGDSElements.BookedBy, pLineNumbers)
        Utilities1A.PrepareLineNumbers1A(mobjExistingGDSElements.Department, pLineNumbers)
        Utilities1A.PrepareLineNumbers1A(mobjExistingGDSElements.ReasonForTravel, pLineNumbers)
        Utilities1A.PrepareLineNumbers1A(mobjExistingGDSElements.CostCentre, pLineNumbers)

        Dim pMax As Integer = 0
        Dim pMaxIndex As Integer = -1
        Dim pFound As Boolean = True
        Do While pFound
            pFound = False
            For i As Integer = 0 To pLineNumbers.GetUpperBound(0)
                If pLineNumbers(i) > pMax Then
                    pMax = pLineNumbers(i)
                    pMaxIndex = i
                    pFound = True
                End If
            Next
            If pMaxIndex > -1 Then
                mobjSession1A.Send("XE" & pMax)
                pLineNumbers(pMaxIndex) = 0
            End If
            pMax = 0
            pMaxIndex = -1
        Loop

    End Sub
    Private Sub RemoveOldGDSEntries1G()

        If mGDSCode <> Utilities.EnumGDSCode.Galileo Then
            Throw New Exception("ReadPNR.RemoveOldGDSEntries1G()" & vbCrLf & "Selected GDS is not Galileo")
        End If

        Dim pLineNumbers(0) As LineNumbers

        ' the following elements remain as they are if they already exist in the PNR
        ClearExistingItems(mobjExistingGDSElements.PhoneElement, mobjNewGDSElements.PhoneElement)
        ClearExistingItems(mobjExistingGDSElements.EmailElement, mobjNewGDSElements.EmailElement)
        ClearExistingItems(mobjExistingGDSElements.AOH, mobjNewGDSElements.AOH)
        'ClearExistingItems(mobjExistingGDSElements.OpenSegment, mobjNewGDSElements.OpenSegment)
        'ClearExistingItems(mobjExistingGDSElements.OptionQueueElement, mobjNewGDSElements.OptionQueueElement)
        'ClearExistingItems(mobjExistingGDSElements.TicketElement, mobjNewGDSElements.TicketElement)
        'ClearExistingItems(mobjExistingGDSElements.AgentID, mobjNewGDSElements.AgentID)

        ' the following elements are removed and replaced if they exist in the PNR
        PrepareLineNumbers1G(mobjExistingGDSElements.OpenSegment, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.AgentID, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.OptionQueueElement, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.TicketElement, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.CustomerCode, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.CustomerName, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.SubDepartmentCode, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.SubDepartmentName, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.CRMCode, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.CRMName, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.VesselFlag, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.VesselName, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.VesselOSI, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.Reference, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.BookedBy, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.Department, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.ReasonForTravel, pLineNumbers)
        PrepareLineNumbers1G(mobjExistingGDSElements.CostCentre, pLineNumbers)

        Dim pMax As Integer = 0
        Dim pMaxIndex As Integer = -1
        Dim pCategory As String = ""
        Dim pFound As Boolean = True

        Do While pFound
            If pCategory = "" Then
                For i As Integer = 0 To pLineNumbers.GetUpperBound(0)
                    If pLineNumbers(i).Category <> "" Then
                        pCategory = pLineNumbers(i).Category
                        pMax = pLineNumbers(i).LineNumber
                        pMaxIndex = i
                        Exit For
                    End If
                Next
            End If
            If pCategory <> "" Then
                For i As Integer = 0 To pLineNumbers.GetUpperBound(0)
                    If pLineNumbers(i).Category = pCategory And pLineNumbers(i).LineNumber > pMax Then
                        pMax = pLineNumbers(i).LineNumber
                        pMaxIndex = i
                        pFound = True
                    End If
                Next
                Dim pResponse
                If pMaxIndex > -1 Then
                    If pCategory = "Segment." Then
                        pResponse = mobjSession1G.SendTerminalCommand("X" & pMax)
                    Else
                        pResponse = mobjSession1G.SendTerminalCommand(pCategory & pMax & "@")
                    End If
                    If pResponse(0) = "INVALID ENTRY" Then
                        pResponse = mobjSession1G.SendTerminalCommand(pCategory & "@")
                    End If
                    pLineNumbers(pMaxIndex).Category = ""
                    pLineNumbers(pMaxIndex).LineNumber = 0
                Else
                    pCategory = ""
                End If
                pMax = 0
                pMaxIndex = -1
            Else
                pFound = False
            End If
        Loop

    End Sub

    Private Sub ClearExistingItems(ByRef ExistingItem As GDSExisting.Item, ByRef NewItem As GDSNew.Item)
        If ExistingItem.Exists Then
            NewItem.Clear()
        End If
    End Sub

    'Private Sub Utilities1A.PrepareLineNumbers1A(ByVal ExistingItem As GDSExisting.Item, ByRef pLineNumbers() As Integer)
    '    If ExistingItem.Exists Then
    '        ReDim Preserve pLineNumbers(pLineNumbers.GetUpperBound(0) + 1)
    '        pLineNumbers(pLineNumbers.GetUpperBound(0)) = ExistingItem.LineNumber
    '    End If
    'End Sub
    Private Sub PrepareLineNumbers1G(ByVal ExistingItem As GDSExisting.Item, ByRef pLineNumbers() As LineNumbers)
        If ExistingItem.Exists Then
            Dim pItems() As String = ExistingItem.Category.Split(".")
            If IsArray(pItems) AndAlso pItems(0) <> "" Then
                ReDim Preserve pLineNumbers(pLineNumbers.GetUpperBound(0) + 1)
                pLineNumbers(pLineNumbers.GetUpperBound(0)).Category = pItems(0) & "."
                pLineNumbers(pLineNumbers.GetUpperBound(0)).LineNumber = ExistingItem.LineNumber
            End If
        End If
    End Sub
    Public Sub SendGDSEntry1A(ByVal GDSEntry As String)

        If mGDSCode <> Utilities.EnumGDSCode.Amadeus Then
            Throw New Exception("ReadPNR.SendNewGDSEntries1A()" & vbCrLf & "Selected GDS is not Amadeus")
        End If

        If GDSEntry <> "" Then
            mobjSession1A.Send(GDSEntry)
        End If

    End Sub
    Public Sub SendGDSEntry1G(ByVal GDSEntry As String)

        If mGDSCode <> Utilities.EnumGDSCode.Galileo Then
            Throw New Exception("ReadPNR.SendNewGDSEntries1G()" & vbCrLf & "Selected GDS is not Galileo")
        End If

        If GDSEntry <> "" Then
            mobjSession1G.SendTerminalCommand(GDSEntry)
        End If

    End Sub
    Public Function SendAllGDSEntries(ByVal WritePNR As Boolean, ByVal WriteDocs As Boolean, ByVal mflgExpiryDateOK As Boolean, dgvApis As DataGridView, AirlineEntries As TextBox) As String

        SendAllGDSEntries = ""
        If mGDSCode = Utilities.EnumGDSCode.Amadeus Then
            SendAllGDSEntries1A(WritePNR, WriteDocs, mflgExpiryDateOK, dgvApis, AirlineEntries)
        ElseIf mGDSCode = Utilities.EnumGDSCode.Galileo Then
            SendAllGDSEntries = SendAllGDSEntries1G(WritePNR, WriteDocs, mflgExpiryDateOK, dgvApis, AirlineEntries)
        Else
            Throw New Exception("ReadPNR.SendAllGDSEntries()" & vbCrLf & "No GDS Selected")
        End If

    End Function
    Private Sub SendAllGDSEntries1A(ByVal WritePNR As Boolean, ByVal WriteDocs As Boolean, ByVal mflgExpiryDateOK As Boolean, dgvApis As DataGridView, AirlineEntries As TextBox)
        Try
            If WritePNR Then
                RemoveOldGDSEntries1A()

                SendGDSElement1A(mobjNewGDSElements.PhoneElement)
                SendGDSElement1A(mobjNewGDSElements.EmailElement)
                SendGDSElement1A(mobjNewGDSElements.AgentID)
                SendGDSElement1A(mobjNewGDSElements.AOH)
                SendGDSElement1A(mobjNewGDSElements.OpenSegment)
                SendGDSElement1A(mobjNewGDSElements.TicketElement)
                SendGDSElement1A(mobjNewGDSElements.OptionQueueElement)

                If mflgNewPNR Then
                    SendGDSElement1A(mobjNewGDSElements.SavingsElement)
                    SendGDSElement1A(mobjNewGDSElements.LossElement)
                End If

                SendGDSElement1A(mobjNewGDSElements.CustomerCode)
                SendGDSElement1A(mobjNewGDSElements.CustomerName)
                SendGDSElement1A(mobjNewGDSElements.SubDepartmentCode)
                SendGDSElement1A(mobjNewGDSElements.SubDepartmentName)
                SendGDSElement1A(mobjNewGDSElements.CRMCode)
                SendGDSElement1A(mobjNewGDSElements.CRMName)
                SendGDSElement1A(mobjNewGDSElements.VesselName)
                SendGDSElement1A(mobjNewGDSElements.VesselFlag)
                SendGDSElement1A(mobjNewGDSElements.VesselOSI)
                SendGDSElement1A(mobjNewGDSElements.Reference)
                SendGDSElement1A(mobjNewGDSElements.BookedBy)
                SendGDSElement1A(mobjNewGDSElements.Department)
                SendGDSElement1A(mobjNewGDSElements.ReasonForTravel)
                SendGDSElement1A(mobjNewGDSElements.CostCentre)

                Dim pAirlineEntries() As String = AirlineEntries.Text.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

                For i As Integer = 0 To pAirlineEntries.GetUpperBound(0)
                    pAirlineEntries(i) = pAirlineEntries(i).Replace(">", "").Trim
                    If pAirlineEntries(i).Trim <> "" Then
                        SendGDSAirlineItems1A(pAirlineEntries(i).Replace("> ", ""))
                    End If
                Next
            End If

            If WriteDocs Then
                APISUpdate1A(mflgExpiryDateOK, dgvApis)
            End If

            If WritePNR Or WriteDocs Then
                CloseOffPNR1A()
            End If
        Catch ex As Exception
            Throw New Exception("SendNewGDSEntries()" & vbCrLf & ex.Message)
        End Try
    End Sub
    Private Function SendAllGDSEntries1G(ByVal WritePNR As Boolean, ByVal WriteDocs As Boolean, ByVal mflgExpiryDateOK As Boolean, dgvApis As DataGridView, AirlineEntries As TextBox) As String
        Try
            SendAllGDSEntries1G = ""
            If WritePNR Then
                RemoveOldGDSEntries1G()

                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.PhoneElement, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.EmailElement, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.AgentID, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.AOH, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.OpenSegment, False)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.TicketElement, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.OptionQueueElement, True)

                If mflgNewPNR Then
                    SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.SavingsElement, True)
                    SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.LossElement, True)
                End If

                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.CustomerCode, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.CustomerName, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.SubDepartmentCode, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.SubDepartmentName, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.CRMCode, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.CRMName, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.VesselName, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.VesselFlag, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.VesselOSI, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.Reference, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.BookedBy, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.Department, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.ReasonForTravel, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.CostCentre, True)
                SendAllGDSEntries1G &= SendGDSElement1G(mobjNewGDSElements.GalileoTrackingCode, True)

                Dim pAirlineEntries() As String = AirlineEntries.Text.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

                For i As Integer = 0 To pAirlineEntries.GetUpperBound(0)
                    pAirlineEntries(i) = pAirlineEntries(i).Replace(">", "").Trim
                    If pAirlineEntries(i).Trim <> "" Then
                        SendAllGDSEntries1G &= SendGDSAirlineItems1G(pAirlineEntries(i).Replace("> ", ""))
                    End If
                Next
            End If

            If WriteDocs Then
                APISUpdate1G(mflgExpiryDateOK, dgvApis)
            End If

            If WritePNR Or WriteDocs Then
                SendAllGDSEntries1G &= CloseOffPNR1G()
            End If
        Catch ex As Exception
            Throw New Exception("SendNewGDSEntries()" & vbCrLf & ex.Message)
        End Try
    End Function
    Private Sub APISUpdate1A(ByVal mflgExpiryDateOK As Boolean, dgvApis As DataGridView)

        Dim pstrCommand As String
        Try
            For i = 0 To dgvApis.RowCount - 1
                With dgvApis.Rows(i)
                    If .ErrorText.IndexOf("Birth") = -1 Then
                        Dim pobjItem As New PaxApisDB.Item(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value,
                                                       Utilities.DateFromIATA(.Cells(6).Value), .Cells(7).Value, .Cells(3).Value,
                                                     .Cells(4).Value, Utilities.DateFromIATA(.Cells(8).Value), .Cells(5).Value)

                        pobjItem.Update(mflgExpiryDateOK)
                        pstrCommand = "SR DOCS YY HK1-P-" & pobjItem.IssuingCountry & "-" & pobjItem.PassportNumber & "-" & pobjItem.Nationality &
                    "-" & Utilities.DateToIATA(pobjItem.BirthDate) & "-" & pobjItem.Gender & "-"
                        If mflgExpiryDateOK Then
                            pstrCommand &= Utilities.DateToIATA(pobjItem.ExpiryDate)
                        Else
                            pstrCommand &= ""
                        End If
                        pstrCommand &= "-" & pobjItem.Surname & "-" & pobjItem.FirstName & "/P" & pobjItem.Id
                        SendGDSEntry1A(pstrCommand)
                    End If

                End With

            Next
        Catch ex As Exception
            Throw New Exception("APISUpdate()" & vbCrLf & ex.Message)
        End Try


    End Sub
    Private Sub APISUpdate1G(ByVal mflgExpiryDateOK As Boolean, dgvApis As DataGridView)

        If mGDSCode <> Utilities.EnumGDSCode.Galileo Then
            Throw New Exception("ReadPNR.SendGDSElement1G()" & vbCrLf & "Selected GDS is not Galileo")
        End If

        Dim pstrCommand As String
        Try
            For i = 0 To dgvApis.RowCount - 1
                With dgvApis.Rows(i)
                    If .ErrorText.IndexOf("Birth") = -1 Then
                        Dim pobjItem As New PaxApisDB.Item(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value,
                                                       Utilities.DateFromIATA(.Cells(6).Value), .Cells(7).Value, .Cells(3).Value,
                                                     .Cells(4).Value, Utilities.DateFromIATA(.Cells(8).Value), .Cells(5).Value)

                        pobjItem.Update(mflgExpiryDateOK)
                        'SI.P1/SSRDOCSBAHK1/P/GB/S12345678/GB/12JUL76/M/23OCT16/SMITH/JOHN/RICHARD
                        pstrCommand = "SI.P" & pobjItem.Id & "/SSRDOCSYYHK1/P/" & pobjItem.IssuingCountry & "/" & pobjItem.PassportNumber & "/" & pobjItem.Nationality &
                    "/" & Utilities.DateToIATA(pobjItem.BirthDate) & "/" & pobjItem.Gender & "/"
                        If mflgExpiryDateOK Then
                            pstrCommand &= Utilities.DateToIATA(pobjItem.ExpiryDate)
                        Else
                            pstrCommand &= ""
                        End If
                        pstrCommand &= "/" & pobjItem.Surname & "/" & pobjItem.FirstName
                        For Each pElement As GDS1G_ReadRaw.SSRitem In mobjPNR1GRaw.SSR.Values
                            If pElement.SSRCode = "DOCS" Then
                                If pElement.LastName = pobjItem.Surname And pElement.PassportNumber = pobjItem.PassportNumber And pElement.DateOfBirth = Utilities.DateToIATA(pobjItem.BirthDate) Then
                                    pstrCommand = ""
                                    Exit For
                                End If
                            End If
                        Next
                        If pstrCommand <> "" Then
                            SendGDSEntry1G(pstrCommand)
                        End If
                    End If
                End With
            Next
        Catch ex As Exception
            Throw New Exception("APISUpdate()" & vbCrLf & ex.Message)
        End Try
    End Sub
    Private Sub CloseOffPNR1A()
        If mGDSCode <> Utilities.EnumGDSCode.Amadeus Then
            Throw New Exception("ReadPNR.CloseOffPNR1A()" & vbCrLf & "Selected GDS is not Amadeus")
        End If

        Dim pCloseOffEntries As New CloseOffEntries.Collection

        pCloseOffEntries.Load(MySettings.GDSPcc, mstrOfficeOfResponsibility = MySettings.GDSPcc)

        For Each pCommand As CloseOffEntries.Item In pCloseOffEntries.Values
            mobjSession1A.Send(pCommand.CloseOffEntry)
        Next
        If mstrPNRResponse.Contains("WARNING: SECURE FLT PASSENGER DATA REQUIRED") Then
            MessageBox.Show(mstrPNRResponse)
        End If

    End Sub
    Private Function CloseOffPNR1G() As String
        If mGDSCode <> Utilities.EnumGDSCode.Galileo Then
            Throw New Exception("ReadPNR.CloseOffPNR1G()" & vbCrLf & "Selected GDS is not Amadeus")
        End If
        Dim pCloseOffEntries As New CloseOffEntries.Collection
        CloseOffPNR1G = ""
        pCloseOffEntries.Load(MySettings.GDSPcc, mstrOfficeOfResponsibility = MySettings.GDSPcc)

        Dim pResponse
        Dim pPNR As String
        pResponse = mobjSession1G.SendTerminalCommand("R.CN")
        pResponse = mobjSession1G.SendTerminalCommand("ER")
        If pResponse(0).ToString.Length > 9 AndAlso pResponse(0).ToString.Substring(6, 1) = "/" Then


            pPNR = pResponse(0).ToString.Substring(0, 6)
            pResponse = mobjSession1G.SendTerminalCommand("I")
            For Each pCommand As CloseOffEntries.Item In pCloseOffEntries.Values
                pResponse = mobjSession1G.SendTerminalCommand("*" & pPNR)
                pResponse = mobjSession1G.SendTerminalCommand(pCommand.CloseOffEntry)
                If pResponse(0) & pResponse(1) <> " *>" And pResponse(0).ToString.IndexOf("ON QUEUE") = -1 Then
                    MessageBox.Show(pCommand.CloseOffEntry & vbCrLf & pResponse(0) & pResponse(1))
                End If
                pResponse = mobjSession1G.SendTerminalCommand("I")
            Next
            pResponse = mobjSession1G.SendTerminalCommand("*" & pPNR)
            pResponse = mobjSession1G.SendTerminalCommand("IR")
            CloseOffPNR1G = pPNR
        Else
            MessageBox.Show(pResponse(0) & vbCrLf & pResponse(1), "ERROR IN PNR UPDATE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            mobjSession1G.SendTerminalCommand("IR")
            Throw New Exception("Error in PNR Update")
        End If
    End Function
    Private Sub SendGDSElement1A(ByVal pElement As GDSNew.Item)
        If mGDSCode <> Utilities.EnumGDSCode.Amadeus Then
            Throw New Exception("ReadPNR.SendGDSElement1A()" & vbCrLf & "Selected GDS is not Amadeus")
        End If

        If pElement.GDSCommand <> "" Then
            mobjSession1A.Send(pElement.GDSCommand)
        End If

    End Sub
    Private Function SendGDSElement1G(ByVal pElement As GDSNew.Item, ByVal ShowResponse As Boolean) As String
        If mGDSCode <> Utilities.EnumGDSCode.Galileo Then
            Throw New Exception("ReadPNR.SendGDSElement1G()" & vbCrLf & "Selected GDS is not Galileo")
        End If
        SendGDSElement1G = ""
        Dim pResponse
        If pElement.GDSCommand <> "" Then
            pResponse = mobjSession1G.SendTerminalCommand(pElement.GDSCommand)
            If pResponse(0) & pResponse(1) <> " *>" Then
                SendGDSElement1G = vbCrLf & pElement.GDSCommand
                For i As Integer = 0 To pResponse.count - 1
                    SendGDSElement1G &= vbCrLf & pResponse(i)
                Next
                If ShowResponse Then
                    MessageBox.Show(pElement.GDSCommand & vbCrLf & pResponse(0) & pResponse(1))
                End If
            End If
        End If

    End Function
    Private Sub SendGDSAirlineItems1A(ByVal pItemToSend As String)
        If mGDSCode <> Utilities.EnumGDSCode.Amadeus Then
            Throw New Exception("ReadPNR.SendGDSAirlineItems1A()" & vbCrLf & "Selected GDS is not Amadeus")
        End If

        If pItemToSend.StartsWith("OS ") Then
            If mobjPNR1A.RawResponse.Replace(vbCrLf, "").Replace(" ", "").IndexOf(("OSI " & pItemToSend.Substring(3)).Replace(" ", "")) = -1 Then
                mobjSession1A.Send(pItemToSend)
            End If
        ElseIf pItemToSend.StartsWith("R") Then
            If mobjPNR1A.RawResponse.Replace(vbCrLf, "").Replace(" ", "").IndexOf(pItemToSend.Replace(" ", "")) = -1 Then
                mobjSession1A.Send(pItemToSend)
            End If
        ElseIf pItemToSend.StartsWith("S") Then
            Dim pString As String
            pString = pItemToSend.Replace(" ", "").Replace("SRCKIN-", "")
            If mobjPNR1A.RawResponse.Replace(vbCrLf, "").Replace(" ", "").IndexOf(pString) = -1 Then
                mobjSession1A.Send(pItemToSend)
            End If
        Else
            mobjSession1A.Send(pItemToSend)
        End If

    End Sub
    Private Function SendGDSAirlineItems1G(ByVal pItemToSend As String) As String
        If mGDSCode <> Utilities.EnumGDSCode.Galileo Then
            Throw New Exception("ReadPNR.SendGDSAirlineItems1G()" & vbCrLf & "Selected GDS is not Galileo")
        End If
        SendGDSAirlineItems1G = ""
        Dim pResponse
        If pItemToSend <> "" Then
            If pItemToSend.StartsWith("DI.") Then
                For Each pElement As GDS1G_ReadRaw.DIItem In mobjPNR1GRaw.DIElements.Values
                    If pElement.Category & pElement.Remark = pItemToSend.Replace(" ", "") Then
                        pItemToSend = ""
                        Exit For
                    End If
                Next
            ElseIf pItemToSend.StartsWith("SI.") Then
                For Each pElement As GDS1G_ReadRaw.SSRitem In mobjPNR1GRaw.SSR.Values
                    If ("SI." & pElement.CarrierCode & "*" & pElement.Text).Replace(" ", "") = pItemToSend.Replace(" ", "") Or "SI.SSR" & pElement.SSRCode & pElement.CarrierCode & pElement.StatusCode & "1" & pElement.Text = pItemToSend Then
                        pItemToSend = ""
                        Exit For
                    End If
                Next
            End If
            If pItemToSend <> "" Then
                pResponse = mobjSession1G.SendTerminalCommand(pItemToSend)
                SendGDSAirlineItems1G = vbCrLf & pItemToSend
                For i As Integer = 0 To pResponse.count - 1
                    SendGDSAirlineItems1G &= vbCrLf & pResponse(i)
                Next
                If pResponse(0) & pResponse(1) <> " *>" Then
                    MessageBox.Show(pItemToSend & vbCrLf & pResponse(0) & pResponse(1))
                End If
            End If
        End If

    End Function
    Private Function RetrievePNR1A(ByVal ForReportOnly As Boolean) As Boolean

        Dim pintPNRStatus As Integer

        mobjPNR1A = New s1aPNR.PNR
        mobjTickets = New GDSTickets.GDSTicketCollection
        mstrVesselName = ""
        mstrBookedBy = ""
        mstrCC = ""
        mstrCLA = ""
        mstrCLN = ""

        With mudtProps

            If .RequestedPNR = "" Then
                pintPNRStatus = mobjPNR1A.RetrieveCurrent(mobjSession1A)
            Else
                pintPNRStatus = mobjPNR1A.RetrievePNR(mobjSession1A, "RT" & .RequestedPNR)
            End If
            .PNRCreationdate = Today

            If pintPNRStatus = 0 Or pintPNRStatus = 1005 Then
                .RequestedPNR = setRecordLocator1A()
                If ForReportOnly Then
                    GetGroup1A()
                    GetPax1A()
                    GetSegs1A(ForReportOnly)
                    GetOtherServiceElements1A()
                    GetRMElements1A()
                Else
                    GetTQT1A()
                    GetGroup1A()
                    GetPax1A()
                    GetSegs1A(ForReportOnly)
                    GetAutoTickets1A()
                    GetOtherServiceElements1A()
                    GetSSRElements1A()
                    GetRMElements1A()
                End If
                RetrievePNR1A = True
            Else
                RetrievePNR1A = False
            End If
        End With

    End Function

    Private Sub GetPnrNumber1A()

        Try
            mstrPNRNumber = mobjPNR1A.Header.RecordLocator
        Catch ex As Exception
            mstrPNRNumber = ""
        End Try

        If mstrPNRNumber = "" Then
            mstrPNRNumber = "New PNR"
            mflgNewPNR = True
        End If
    End Sub
    Private Function setRecordLocator1A() As String
        Try
            setRecordLocator1A = mobjPNR1A.Header.RecordLocator
        Catch ex As Exception
            setRecordLocator1A = UCase(mudtProps.RequestedPNR)
        End Try
    End Function
    'Private Sub GetPnrNumber1G()

    '    Try
    '        mstrPNRNumber = mobjPNR1G.RecordLocator
    '    Catch ex As Exception
    '        mstrPNRNumber = ""
    '    End Try

    '    If mstrPNRNumber = "" Then
    '        mstrPNRNumber = "New PNR"
    '        mflgNewPNR = True
    '    End If
    'End Sub
    Private Sub GetOfficeOfResponsibility1A()

        Try
            mstrOfficeOfResponsibility = mobjPNR1A.Header.OfficeOfResponsability
        Catch ex As Exception
            mstrOfficeOfResponsibility = MySettings.GDSPcc
        End Try

    End Sub
    'Private Sub GetOfficeOfResponsibility1G()
    '    Try
    '        mstrOfficeOfResponsibility = mobjPNR1G.CurrentAgencyPcc
    '    Catch ex As Exception
    '        mstrOfficeOfResponsibility = MySettings.GDSPcc
    '    End Try
    'End Sub
    Private Sub GetGroup1AGDS()

        mstrGroupName = ""
        mintGroupNamesCount = 0

        For Each pGroup As s1aPNR.GroupNameElement In mobjPNR1A.GroupNameElements
            mstrGroupName = pGroup.GroupName
            mintGroupNamesCount = pGroup.NbrOfAssignedNames + pGroup.NbrNamesMissing
            Exit For
        Next
        If mobjPNR1A.GroupNameElements.Count > 1 Then
            mstrGroupName &= "x" & mobjPNR1A.GroupNameElements.Count
        End If

    End Sub

    Private Sub GetGroup1A()

        mstrGroupName = ""
        mintGroupNamesCount = 0

        For Each pGroup As s1aPNR.GroupNameElement In mobjPNR1A.GroupNameElements
            mstrGroupName = pGroup.GroupName
            mintGroupNamesCount = pGroup.NbrOfAssignedNames + pGroup.NbrNamesMissing
            Exit For
        Next
        If mobjPNR1A.GroupNameElements.Count > 1 Then
            mstrGroupName &= "x" & mobjPNR1A.GroupNameElements.Count
        End If

    End Sub
    Private Sub GetPassengers1A()
        mobjPassengers.Clear()
        For Each Pax As s1aPNR.NameElement In mobjPNR1A.NameElements
            With Pax
                mobjPassengers.AddItem(.ElementNo, .Initial, .LastName, If(IsNothing(.ID), "", .ID))
                'Exit For
            End With
        Next
    End Sub
    'Private Sub GetPassengers1G()

    '    mobjPassengers.Clear()
    '    For Each pobjPax As Travelport.TravelData.Person In mobjPNR1G.Passengers
    '        With pobjPax
    '            mobjPassengers.AddItem(.PassengerNumber, .FirstName, .LastName, If(IsNothing(.NameRemark), "", .NameRemark))
    '        End With
    '    Next
    'End Sub
    Private Sub GetSegments1A()

        mobjSegments.Clear()
        mdteDepartureDate = Date.MinValue
        mstrItinerary = ""
        Dim pOff As String = ""

        For Each pSeg As s1aPNR.AirFlownSegment In mobjPNR1A.AirFlownSegments
            With pSeg
                If mstrItinerary = "" Then
                    mstrItinerary = .BoardPoint & "-" & .OffPoint
                Else
                    If .BoardPoint = pOff Then
                        mstrItinerary &= "-" & .OffPoint
                    Else
                        mstrItinerary &= "-***-" & .BoardPoint & "-" & .OffPoint
                    End If
                End If
                If .Airline = "QR" Then
                    mflgQRSegment = True
                End If
                pOff = .OffPoint
                Dim pDate As New s1aAirlineDate.clsAirlineDate
                pDate.SetFromString(.DepartureDate)
                If mdteDepartureDate = Date.MinValue Then
                    mdteDepartureDate = pDate.VBDate
                End If
                mobjSegments.AddItem(Utilities1A.airAirline1A(pSeg), Utilities1A.airBoardPoint1A(pSeg), Utilities1A.airClass1A(pSeg), Utilities1A.airDepartureDate1A(pSeg), Utilities1A.airArrivalDate1A(pSeg), .ElementNo, Utilities1A.airFlightNo1A(pSeg), Utilities1A.airOffPoint1A(pSeg), Utilities1A.airStatus1A(pSeg), Utilities1A.airDepartTime1A(pSeg), Utilities1A.airArriveTime1A(pSeg), Utilities1A.airText1A(pSeg), "")
            End With
        Next

        For Each pSeg As s1aPNR.AirSegment In mobjPNR1A.AirSegments
            With pSeg
                If mstrItinerary = "" Then
                    mstrItinerary = pSeg.BoardPoint & "-" & pSeg.OffPoint
                Else
                    If pSeg.BoardPoint = pOff Then
                        mstrItinerary &= "-" & pSeg.OffPoint
                    Else
                        mstrItinerary &= "-***-" & pSeg.BoardPoint & "-" & pSeg.OffPoint
                    End If
                End If
                pOff = pSeg.OffPoint
                Dim pDate As New s1aAirlineDate.clsAirlineDate
                pDate.SetFromString(pSeg.DepartureDate)
                If mdteDepartureDate = Date.MinValue Then
                    mdteDepartureDate = pDate.VBDate
                End If
                mobjSegments.AddItem(Utilities1A.airAirline1A(pSeg), Utilities1A.airBoardPoint1A(pSeg), Utilities1A.airClass1A(pSeg), Utilities1A.airDepartureDate1A(pSeg), Utilities1A.airArrivalDate1A(pSeg), .ElementNo, Utilities1A.airFlightNo1A(pSeg), Utilities1A.airOffPoint1A(pSeg), Utilities1A.airStatus1A(pSeg), Utilities1A.airDepartTime1A(pSeg), Utilities1A.airArriveTime1A(pSeg), Utilities1A.airText1A(pSeg), "")
            End With
        Next
        mflgExistsSegments = ((mobjPNR1A.AirFlownSegments.Count + mobjPNR1A.AirSegments.Count) > 0)

        If mdteDepartureDate > Date.MinValue Then
            Dim pDate As New s1aAirlineDate.clsAirlineDate
            pDate.SetFromString(mdteDepartureDate)
            mstrItinerary &= " (" & pDate.IATA & ")"
        End If

    End Sub

    'Private Sub GetSegments1G()

    '    Dim pOff As String = ""

    '    mobjSegments.Clear()
    '    mdteDepartureDate = Date.MinValue
    '    mstrItinerary = ""

    '    For Each pSeg As Travelport.TravelData.AirSegment In mobjPNR1G.AirSegments
    '        With pSeg
    '            If mstrItinerary = "" Then
    '                mstrItinerary = .Origin.Code & "-" & .Destination.Code
    '            Else
    '                If .Origin.Code = pOff Then
    '                    mstrItinerary &= "-" & .Destination.Code
    '                Else
    '                    mstrItinerary &= "-***-" & .Origin.Code & "-" & .Destination.Code
    '                End If
    '            End If
    '            pOff = .Destination.Code
    '            If mdteDepartureDate = Date.MinValue Then
    '                mdteDepartureDate = .EndDateTime.Date
    '            End If
    '            mobjSegments.AddItem(.Carrier.Code, .Origin.Code, If(IsNothing(.ClassOfService), "", .ClassOfService), .StartDateTime, .EndDateTime, .SegmentNumber, .FlightNumber, .Destination.Code, .RequestStatus, .StartDateTime, .EndDateTime, .ToString, "")
    '        End With
    '    Next pSeg
    '    mflgExistsSegments = ((mobjPNR1G.AirSegments.Count) > 0)

    '    If mdteDepartureDate > Date.MinValue Then
    '        Dim pDate As New s1aAirlineDate.clsAirlineDate
    '        pDate.SetFromString(mdteDepartureDate)
    '        mstrItinerary &= " (" & pDate.IATA & ")"
    '    End If

    'End Sub
    Private Sub GetOpenSegment1A()

        For Each pSeg As s1aPNR.MemoSegment In mobjPNR1A.MemoSegments
            If pSeg.Text.Contains(MySettings.GDSValue("TextMISSegmentLookup") & mobjPNR1A.NameElements.Count & " " & MySettings.OfficeCityCode) Then
                mobjExistingGDSElements.OpenSegment.SetValues(True, pSeg.ElementNo, MySettings.GDSElement("TextMISSegmentLookup"), "", "")
                Exit For
            End If
        Next

    End Sub
    Private Sub GetOpenSegment1G()

        For Each pOpenSeg As GDS1G_ReadRaw.OpenSegmentItem In mobjPNR1GRaw.OpenSegments.Values
            If pOpenSeg.SegmentType = "T" Then
                mobjExistingGDSElements.OpenSegment.SetValues(True, pOpenSeg.ElementNo, "Segment", pOpenSeg.Remark.ToString, "")
            End If
        Next
    End Sub
    Private Sub GetPhoneElement1A()

        For Each pField As s1aPNR.PhoneElement In mobjPNR1A.PhoneElements
            If pField.Text.Replace(" ", "").Contains(MySettings.GDSValue("TextAP").Replace(" ", "")) Then
                mobjExistingGDSElements.PhoneElement.SetValues(True, pField.Text.Substring(0, pField.Text.IndexOf(pField.ElementID) - 1), MySettings.GDSElement("TextAP"), "", "")
                Exit For
            End If
        Next

    End Sub
    Private Sub GetPhoneElement1G()

        For Each pPhone As GDS1G_ReadRaw.PhoneNumbersItem In mobjPNR1GRaw.PhoneNumbers.Values
            If "P." & pPhone.CityCode & "T*" & pPhone.PhoneNumber = MySettings.GDSValue("TextAP") Then
                mobjExistingGDSElements.PhoneElement.SetValues(True, pPhone.ElementNo, MySettings.GDSElement("TextAP"), pPhone.PhoneNumber, pPhone.PhoneNumber)
            ElseIf "P." & pPhone.CityCode & "T*" & pPhone.PhoneNumber = MySettings.GDSValue("TextAOH") Then
                mobjExistingGDSElements.AOH.SetValues(True, pPhone.ElementNo, MySettings.GDSElement("TextAOH"), pPhone.PhoneNumber, pPhone.PhoneNumber)
            End If
        Next
    End Sub
    Private Sub GetEmailElement1A()

        For Each pField As s1aPNR.PhoneElement In mobjPNR1A.PhoneElements
            If pField.Text.Contains(MySettings.GDSValue("TextAPE_ToFind")) Then
                mobjExistingGDSElements.EmailElement.SetValues(True, pField.Text.Substring(0, pField.Text.IndexOf(pField.ElementID) - 1), MySettings.GDSElement("TextAPE_ToFind"), "", "")
            End If
        Next
    End Sub
    Private Sub GetEmailElement1G()
        For Each pEmail As GDS1G_ReadRaw.EmailItem In mobjPNR1GRaw.Emails.Values
            If "MT." & pEmail.EmailAddress = MySettings.GDSValue("TextAPE") Then
                mobjExistingGDSElements.EmailElement.SetValues(True, pEmail.ElementNo, MySettings.GDSElement("TextAPE"), pEmail.EmailAddress, pEmail.EmailAddress)
            End If
        Next
    End Sub
    Private Sub GetAOH1A()
        For Each pElement As s1aPNR.SSRElement In mobjPNR1A.SSRElements
            If pElement.Text.Contains(MySettings.GDSValue("TextAOH_ToFind")) Then
                mobjExistingGDSElements.AOH.SetValues(True, pElement.Text.Substring(0, pElement.Text.IndexOf(pElement.ElementID) - 1), MySettings.GDSElement("TextAOH_ToFind"), "", "")
            End If
        Next
    End Sub

    Private Sub GetTicketElement1A()
        For Each pElement As s1aPNR.TicketElement In mobjPNR1A.TicketElements
            mobjExistingGDSElements.TicketElement.SetValues(True, pElement.Text.Substring(0, pElement.Text.IndexOf(pElement.ElementID) - 1), "TKT", "", "")
        Next
    End Sub
    Private Sub GetTicketElement1G()

        If mobjPNR1GRaw.TicketElement.ElementNo = 1 Then
            mobjExistingGDSElements.TicketElement.SetValues(True, 1, "T.", mobjPNR1GRaw.TicketElement.ActionDateTime, "")
        End If
    End Sub

    Private Sub GetOptionQueueElement1A()
        For Each pElement As s1aPNR.OptionQueueElement In mobjPNR1A.OptionQueueElements
            If pElement.Text.Contains(MySettings.GDSValue("TextOP")) Then
                mobjExistingGDSElements.OptionQueueElement.SetValues(True, pElement.Text.Substring(0, pElement.Text.IndexOf(pElement.ElementID) - 1), MySettings.GDSElement("TextOP"), "", "")
                Exit For
            End If
        Next
    End Sub
    Private Sub GetOptionQueueElement1G()
        For Each pField As GDS1G_ReadRaw.OptionQueueItem In mobjPNR1GRaw.OptionQueue.Values
            Dim pFullText As String = MySettings.GDSValue("TextOP") & "/DDMMM/0001/Q" & MySettings.AgentOPQueue
            If pFullText.StartsWith(MySettings.GDSValue("TextOP")) And pFullText.EndsWith("/0001/Q" & MySettings.AgentOPQueue) Then
                mobjExistingGDSElements.OptionQueueElement.SetValues(True, pField.ElementNo, MySettings.GDSElement("TextOP"), pField.QueueNumber, pField.QueueNumber)
            End If
        Next
    End Sub
    Private Sub GetVesselOSI1A()
        For Each pOSI As s1aPNR.OtherServiceElement In mobjPNR1A.OtherServiceElements
            If pOSI.Text.Contains(MySettings.GDSValue("TextVOSI")) Then
                If mobjExistingGDSElements.VesselOSI.Exists Then
                    Throw New Exception("Please check PNR. Duplicate OSI Vessel defined" & vbCrLf & mobjExistingGDSElements.VesselOSI.RawText & vbCrLf & pOSI.Text)
                Else
                    Dim pVesselNameOSI As String = pOSI.Text.Substring(pOSI.Text.IndexOf(MySettings.GDSValue("TextVSL")) + MySettings.GDSValue("TextVSL").Length)
                    mobjExistingGDSElements.VesselOSI.SetValues(True, pOSI.Text.Substring(0, pOSI.Text.IndexOf(pOSI.ElementID) - 1), MySettings.GDSElement("TextVSL"), pOSI.Text, pVesselNameOSI)
                End If
            End If
        Next
    End Sub
    Private Sub GetSSRElements1A()

        Dim pobjSSR As s1aPNR.SSRfqtvElement

        For Each pobjSSR In mobjPNR1A.SSRfqtvElements

            If pobjSSR.Associations.Passengers.Count > 0 Then
                For Each objPax In pobjSSR.Associations.Passengers
                    mobjFrequentFlyer.AddItem(mobjPassengers(objPax.ElementNo).PaxName, pobjSSR.Airline, pobjSSR.FrequentTravelerNo, "")
                Next
            Else
                For Each pPax As GDSPax.GDSPaxItem In mobjPassengers.Values
                    mobjFrequentFlyer.AddItem(pPax.PaxName, pobjSSR.Airline, pobjSSR.FrequentTravelerNo, "")
                Next
            End If

        Next

        Dim pTQTtext As k1aHostToolKit.CHostResponse = mobjSession1A.Send("RTSTR")
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

    Private Sub GetSSR1A()
        mflgExistsSSRDocs = False
        mstrSSRDocs = ""
        For Each pSSR As s1aPNR.SSRElement In mobjPNR1A.SSRElements
            If pSSR.Text.IndexOf("SSR DOCS") > 0 And pSSR.Text.IndexOf("SSR DOCS") < 10 Then
                mstrSSRDocs &= pSSR.Text & vbCrLf
                mflgExistsSSRDocs = True
            End If
        Next
    End Sub
    Private Sub GetSSR1G()
        mflgExistsSSRDocs = False
        mstrSSRDocs = ""
        For Each pobjSSR As GDS1G_ReadRaw.SSRitem In mobjPNR1GRaw.SSR.Values
            With pobjSSR
                '"SEMN/VESSEL-CHRISTOS"
                If (("SI." & .CarrierCode & "*" & .Text).StartsWith(MySettings.GDSValue("TextVOSI"))) Then
                    Dim pVesselNameOSI As String = ("SI." & .CarrierCode & "*" & .Text).Substring(MySettings.GDSValue("TextVOSI").Length).Trim
                    mobjExistingGDSElements.VesselOSI.SetValues(True, pobjSSR.ElementNo, MySettings.GDSElement("TextVOSI"), pobjSSR.Text, pVesselNameOSI)
                    mstrVesselName = .Text.Substring(12).Trim
                ElseIf .SSRCode = "DOCS" Then
                    mstrSSRDocs &= "SI.SSR" & .SSRCode & .CarrierCode & .StatusCode & "1" & .Text.Split("-")(0) & vbCrLf
                    mflgExistsSSRDocs = True
                End If
            End With
        Next pobjSSR
    End Sub

    Private Sub GetRMElements1A()

        Dim pobjRMElement As s1aPNR.RemarkElement

        For Each pobjRMElement In mobjPNR1A.RemarkElements
            parseRMElements1A(pobjRMElement)
        Next pobjRMElement

    End Sub
    Private Sub parseRMElements1A(ByVal Element As s1aPNR.RemarkElement)

        Dim pintLen As Short
        Dim pstrText As String
        Dim pstrSplit() As String

        pstrText = ConcatenateText(Element.Text)
        pintLen = Len(pstrText)
        pstrSplit = Split(Left(pstrText, pintLen), "/")
        ' TODO - make necessary changes for Cyprus Discovery remarks
        If IsArray(pstrSplit) AndAlso pstrSplit.Length >= 2 Then
            If pstrSplit(1) = "CC" Then
                mstrCC = pstrSplit(2)
            ElseIf pstrSplit(1) = "CLN" Then
                mstrCLN = pstrSplit(2)
            ElseIf pstrSplit(1) = "CLA" Then
                mstrCLA = pstrSplit(2)
            ElseIf pstrSplit(1) = "BBY" Then
                mstrBookedBy = pstrSplit(2)
            End If
        End If
        pstrSplit = Split(Left(pstrText, pintLen), "-")
        If IsArray(pstrSplit) AndAlso pstrSplit.Length >= 2 Then
            If pstrSplit(0).IndexOf("D,BOOKED") > 0 Then
                mstrBookedBy = pstrSplit(1)
            ElseIf pstrSplit(0).IndexOf("D,AC") > 0 Then
                mstrCLN = pstrSplit(1)
            End If
        End If


    End Sub
    Private Sub GetRM1A()
        For Each pRemark As s1aPNR.RemarkElement In mobjPNR1A.RemarkElements
            If pRemark.Text.Contains(MySettings.GDSValue("TextAGT")) Then
                mobjExistingGDSElements.AgentID.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextAGT"), pRemark.Text, "")
            ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextCLN")) Then
                If mobjExistingGDSElements.CustomerCode.Exists Then
                    Throw New Exception("Please check PNR. Duplicate customer defined" & vbCrLf & mobjExistingGDSElements.CustomerCode.RawText & vbCrLf & pRemark.Text)
                Else
                    Dim pCustomerCode As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextCLN")) + MySettings.GDSValue("TextCLN").Length)
                    mobjExistingGDSElements.CustomerCode.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextCLN"), pRemark.Text, pCustomerCode)
                End If
            ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextSBN")) Then
                If mobjExistingGDSElements.SubDepartmentCode.Exists Then
                    Throw New Exception("Please check PNR. Duplicate subdepartment defined" & vbCrLf & mobjExistingGDSElements.SubDepartmentCode.LineNumber & vbCrLf & pRemark.Text)
                Else
                    Dim pSubDepartment As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextSBN")) + MySettings.GDSValue("TextSBN").Length)
                    mobjExistingGDSElements.SubDepartmentCode.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextSBN"), "", pSubDepartment)
                End If
            ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextCRN")) Then
                If mobjExistingGDSElements.CRMCode.Exists Then
                    Throw New Exception("Please check PNR. Duplicate CRM defined" & vbCrLf & mobjExistingGDSElements.CRMCode.LineNumber & vbCrLf & pRemark.Text)
                Else
                    Dim pCRM As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextCRN")) + MySettings.GDSValue("TextCRN").Length)
                    mobjExistingGDSElements.CRMCode.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextCRN"), "", pCRM)
                End If
            ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextVSL")) Then
                If mobjExistingGDSElements.VesselName.Exists Then
                    Throw New Exception("Please check PNR. Duplicate vessel defined" & vbCrLf & mobjExistingGDSElements.VesselName.LineNumber & vbCrLf & pRemark.Text)
                Else
                    Dim pVesselName As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextVSL")) + MySettings.GDSValue("TextVSL").Length)
                    mobjExistingGDSElements.VesselName.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextVSL"), "", pVesselName)
                End If
            ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextVSR")) Then
                If mobjExistingGDSElements.VesselFlag.Exists Then
                    Throw New Exception("Please check PNR. Duplicate vessel registration defined" & vbCrLf & mobjExistingGDSElements.VesselFlag.LineNumber & vbCrLf & pRemark.Text)
                Else
                    Dim pVesselRegistration As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextVSR")) + MySettings.GDSValue("TextVSR").Length)
                    mobjExistingGDSElements.VesselFlag.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextVSR"), "", pVesselRegistration)
                End If
            ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextREF")) Then
                Dim pReference As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextREF")) + MySettings.GDSValue("TextREF").Length)
                mobjExistingGDSElements.Reference.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextREF"), "", pReference)
            ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextBBY")) Then
                Dim pBookedBy As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextBBY")) + MySettings.GDSValue("TextBBY").Length)
                mobjExistingGDSElements.BookedBy.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextBBY"), "", pBookedBy)
            ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextDPT")) Then
                Dim pDepartment As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextDPT")) + MySettings.GDSValue("TextDPT").Length)
                mobjExistingGDSElements.Department.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextDPT"), True, pDepartment)
            ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextRFT")) Then
                Dim pReasonForTravel As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextRFT")) + MySettings.GDSValue("TextRFT").Length)
                mobjExistingGDSElements.ReasonForTravel.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextRFT"), "", pReasonForTravel)
            ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextCC")) Then
                Dim pCostCentre As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextCC")) + MySettings.GDSValue("TextCC").Length)
                mobjExistingGDSElements.CostCentre.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextCC"), "", pCostCentre)
            ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextCLA")) Then
                If mobjExistingGDSElements.CustomerName.Exists Then
                    Throw New Exception("Please check PNR. Duplicate customer name defined" & vbCrLf & mobjExistingGDSElements.CustomerName.LineNumber & vbCrLf & pRemark.Text)
                Else
                    mobjExistingGDSElements.CustomerName.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextCLA"), "", "")
                End If
            ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextSBA")) Then
                mobjExistingGDSElements.SubDepartmentName.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextSBA"), "", "")
            ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextCRA")) Then
                mobjExistingGDSElements.CRMName.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextCRA"), "", "")
            End If
        Next
    End Sub
    Private Sub GetRM1G()
        For Each pRemark As GDS1G_ReadRaw.DIItem In mobjPNR1GRaw.DIElements.Values
            With pRemark
                Dim pFullText As String = "DI." & .Category & "-" & .Remark
                If pFullText.StartsWith(MySettings.GDSValue("TextAGT")) Then
                    mobjExistingGDSElements.AgentID.SetValues(True, .ElementNo, MySettings.GDSElement("TextAGT"), .Remark, pFullText.Substring(MySettings.GDSValue("TextAGT").Length))
                ElseIf pFullText.StartsWith(MySettings.GDSValue("TextBBY")) Then
                    Dim pBookedBy As String = pFullText.Substring(MySettings.GDSValue("TextBBY").Length)
                    mstrBookedBy = .Remark.Substring(10)
                    mobjExistingGDSElements.BookedBy.SetValues(True, .ElementNo, MySettings.GDSElement("TextBBY"), .Remark, pBookedBy)
                ElseIf pFullText.StartsWith(MySettings.GDSValue("TextCC")) Then
                    mstrCC = .Remark.Substring(9)
                    mobjExistingGDSElements.CostCentre.SetValues(True, .ElementNo, MySettings.GDSElement("TextCC"), .Remark, pFullText.Substring(MySettings.GDSValue("TextCC").Length))
                ElseIf pFullText.StartsWith(MySettings.GDSValue("TextCLA")) Then
                    If mobjExistingGDSElements.CustomerName.Exists Then
                        Throw New Exception("Please check PNR. Duplicate customer name defined" & vbCrLf & mobjExistingGDSElements.CustomerName.LineNumber & vbCrLf & .Remark)
                    Else
                        mstrCLA = .Remark.Substring(10)
                        mobjExistingGDSElements.CustomerName.SetValues(True, .ElementNo, MySettings.GDSElement("TextCLA"), .Remark, pFullText.Substring(MySettings.GDSValue("TextCLA").Length))
                    End If
                ElseIf pFullText.StartsWith(MySettings.GDSValue("TextCLN")) Then
                    Dim pCustomerCode As String = pFullText.Substring(MySettings.GDSValue("TextCLN").Length)
                    If mobjExistingGDSElements.CustomerCode.Exists Then
                        Throw New Exception("Please check PNR. Duplicate customer defined" & vbCrLf & mobjExistingGDSElements.CustomerCode.RawText & vbCrLf & .Remark)
                    Else
                        mstrCLN = .Remark.Substring(10)
                        mobjExistingGDSElements.CustomerCode.SetValues(True, .ElementNo, MySettings.GDSElement("TextCLN"), .Remark, pCustomerCode)
                    End If
                ElseIf pFullText.StartsWith(MySettings.GDSValue("TextCRA")) Then
                    mobjExistingGDSElements.CRMName.SetValues(True, .ElementNo, MySettings.GDSElement("TextCRA"), .Remark, pFullText.Substring(MySettings.GDSValue("TextCRA").Length))
                ElseIf pFullText.StartsWith(MySettings.GDSValue("TextCRN")) Then
                    mobjExistingGDSElements.CRMCode.SetValues(True, .ElementNo, MySettings.GDSElement("TextCRN"), .Remark, pFullText.Substring(MySettings.GDSValue("TextCRN").Length))
                ElseIf pFullText.StartsWith(MySettings.GDSValue("TextDPT")) Then
                    mobjExistingGDSElements.Department.SetValues(True, .ElementNo, MySettings.GDSElement("TextDPT"), .Remark, pFullText.Substring(MySettings.GDSValue("TextDPT").Length))
                ElseIf pFullText.StartsWith(MySettings.GDSValue("TextREF")) Then
                    mobjExistingGDSElements.Reference.SetValues(True, .ElementNo, MySettings.GDSElement("TextREF"), .Remark, pFullText.Substring(MySettings.GDSValue("TextREF").Length))
                ElseIf pFullText.StartsWith(MySettings.GDSValue("TextRFT")) Then
                    mobjExistingGDSElements.ReasonForTravel.SetValues(True, .ElementNo, MySettings.GDSElement("TextRFT"), .Remark, pFullText.Substring(MySettings.GDSValue("TextRFT").Length))
                ElseIf pFullText.StartsWith(MySettings.GDSValue("TextSBA")) Then
                    mobjExistingGDSElements.SubDepartmentName.SetValues(True, .ElementNo, MySettings.GDSElement("TextSBA"), .Remark, pFullText.Substring(MySettings.GDSValue("TextSBA").Length))
                ElseIf pFullText.StartsWith(MySettings.GDSValue("TextSBN")) Then
                    mobjExistingGDSElements.SubDepartmentCode.SetValues(True, .ElementNo, MySettings.GDSElement("TextSBN"), .Remark, pFullText.Substring(MySettings.GDSValue("TextSBN").Length))
                ElseIf pFullText.StartsWith(MySettings.GDSValue("TextVSL")) Then
                    mobjExistingGDSElements.VesselName.SetValues(True, .ElementNo, MySettings.GDSElement("TextVSL"), .Remark, pFullText.Substring(MySettings.GDSValue("TextVSL").Length))
                ElseIf pFullText.StartsWith(MySettings.GDSValue("TextVSR")) Then
                    mobjExistingGDSElements.VesselFlag.SetValues(True, .ElementNo, MySettings.GDSElement("TextVSR"), .Remark, pFullText.Substring(MySettings.GDSValue("TextVSR").Length))
                ElseIf pFullText.StartsWith("D,BOOKED") > 0 Then
                    mobjExistingGDSElements.BookedBy.SetValues(True, .ElementNo, "D,BOOKED", .Remark, "DI.")
                ElseIf pFullText.StartsWith("D,AC") > 0 Then
                    Dim pCustomerCode As String = .Remark.Substring(10)
                    If mobjExistingGDSElements.CustomerCode.Exists Then
                        Throw New Exception("Please check PNR. Duplicate customer defined" & vbCrLf & mobjExistingGDSElements.CustomerCode.RawText & vbCrLf & .Remark)
                    Else
                        mobjExistingGDSElements.CustomerCode.SetValues(True, .ElementNo, "D,AC", .Remark, "DI.")
                    End If
                End If
                If .Remark.StartsWith("D,BOOKED") > 0 Then
                    mstrBookedBy = .Remark.Substring(8)
                ElseIf .Remark.StartsWith("D,AC") > 0 Then
                    mstrCLN = .Remark.Substring(4)
                End If

            End With
        Next
    End Sub
    Private Sub GetTickets1A()
        mobjTickets = New GDSTickets.GDSTicketCollection(mobjPNR1A)
    End Sub
    Private Sub GetPax1A()

        Dim i As Short
        Dim j As Short
        Dim pstrID As String
        Dim pobjPax As s1aPNR.NameElement

        mobjPassengers.Clear()

        For Each pobjPax In mobjPNR1A.NameElements
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
                mobjPassengers.AddItem(.ElementNo, .Initial, .LastName, pstrID)
            End With
        Next pobjPax

    End Sub

    Private Sub GetSegs1A(ByVal ForReportOnly As Boolean)

        Dim pobjSeg As Object

        mobjSegments.Clear()
        mSegsLastElement = -1
        mSegsFirstElement = -1

        For Each pobjSeg In mobjPNR1A.AllAirSegments
            Dim pElementNo As Short = Utilities1A.airElementNo1A(pobjSeg)
            If ForReportOnly Then
                mobjSegments.AddItem(Utilities1A.airAirline1A(pobjSeg), Utilities1A.airBoardPoint1A(pobjSeg), Utilities1A.airClass1A(pobjSeg), Utilities1A.airDepartureDate1A(pobjSeg), Utilities1A.airArrivalDate1A(pobjSeg), pElementNo, Utilities1A.airFlightNo1A(pobjSeg), Utilities1A.airOffPoint1A(pobjSeg), Utilities1A.airStatus1A(pobjSeg), Utilities1A.airDepartTime1A(pobjSeg), Utilities1A.airArriveTime1A(pobjSeg), Utilities1A.airText1A(pobjSeg), "")
            Else
                Dim pSegDoTemp As k1aHostToolKit.CHostResponse = mobjSession1A.Send("DO" & pobjSeg.ElementNo)
                Dim pSegDo As String = pSegDoTemp.Text
                Do While pSegDo.IndexOf(")>") > 0
                    pSegDoTemp = mobjSession1A.Send("MDR")
                    pSegDo = pSegDo.Replace(")>" & vbCrLf, "") & pSegDoTemp.Text
                Loop
                mobjSegments.AddItem(Utilities1A.airAirline1A(pobjSeg), Utilities1A.airBoardPoint1A(pobjSeg), Utilities1A.airClass1A(pobjSeg), Utilities1A.airDepartureDate1A(pobjSeg), Utilities1A.airArrivalDate1A(pobjSeg), pElementNo, Utilities1A.airFlightNo1A(pobjSeg), Utilities1A.airOffPoint1A(pobjSeg), Utilities1A.airStatus1A(pobjSeg), Utilities1A.airDepartTime1A(pobjSeg), Utilities1A.airArriveTime1A(pobjSeg), Utilities1A.airText1A(pobjSeg), pSegDo)
            End If
                If mSegsFirstElement = -1 Then
                mSegsFirstElement = pElementNo
            End If
            If pElementNo > mSegsLastElement Then
                mSegsLastElement = pElementNo
            End If
        Next pobjSeg

    End Sub
    'Private Sub GetSegs1G(ByVal ForReportOnly As Boolean)

    '    Dim pobjSeg As Travelport.TravelData.AirSegment

    '    mobjSegments.Clear()
    '    mSegsLastElement = -1
    '    mSegsFirstElement = -1

    '    For Each pobjSeg In mobjPNR1G.AirSegments
    '        With pobjSeg
    '            Dim pElementNo As Short = .SegmentNumber
    '            If ForReportOnly Then
    '                mobjSegments.AddItem(.Carrier.Code, .Origin.Code, .ClassOfService, .StartDateTime, .EndDateTime, .SegmentNumber, .FlightNumber, .Destination.Code, .FlightStatus, .StartDateTime, .EndDateTime, .ToString, "")
    '            Else
    '                '                    Dim pSegDo As k1aHostToolKit.CHostResponse = mobjHostSession.Send("DO" & pobjSeg.ElementNo)
    '                mobjSegments.AddItem(.Carrier.Code, .Origin.Code, If(IsNothing(.ClassOfService), "", .ClassOfService), .StartDateTime, .EndDateTime, .SegmentNumber, .FlightNumber, .Destination.Code, .RequestStatus, .StartDateTime, .EndDateTime, .ToString, "")
    '            End If
    '            If mSegsFirstElement = -1 Then
    '                mSegsFirstElement = pElementNo
    '            End If
    '            If pElementNo > mSegsLastElement Then
    '                mSegsLastElement = pElementNo
    '            End If
    '        End With
    '    Next pobjSeg

    'End Sub

    Private Sub GetAutoTickets1A()

        Dim pobjFareAutoTktElement As s1aPNR.FareAutoTktElement
        Dim pobjFareOriginalIssueElement As s1aPNR.FareOriginalIssueElement

        For Each pobjFareOriginalIssueElement In mobjPNR1A.FareOriginalIssueElements
            parseFareOriginal(pobjFareOriginalIssueElement)
        Next pobjFareOriginalIssueElement

        For Each pobjFareAutoTktElement In mobjPNR1A.FareAutoTktElements
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
                    SegAssociations &= mobjSegments(objSeg.ElementNo).BoardPoint & " " & mobjSegments(objSeg.ElementNo).Airline & " " & mobjSegments(objSeg.ElementNo).OffPoint & vbCrLf
                Next
            Else
                For Each pSeg As GDSSeg.GDSSegItem In mobjSegments.Values
                    SegAssociations &= pSeg.BoardPoint & " " & pSeg.Airline & " " & pSeg.OffPoint & vbCrLf
                Next
            End If

            If Element.Associations.Passengers.Count > 0 Then
                For Each objPax In Element.Associations.Passengers
                    PaxAssociations &= mobjPassengers(objPax.ElementNo).PaxName & vbCrLf
                Next
            Else
                For Each pPax As GDSPax.GDSPaxItem In mobjPassengers.Values
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
                    SegAssociations &= mobjSegments(objSeg.ElementNo).BoardPoint & " " & mobjSegments(objSeg.ElementNo).Airline & " " & mobjSegments(objSeg.ElementNo).OffPoint & vbCrLf
                Next
            Else
                For Each pSeg As GDSSeg.GDSSegItem In mobjSegments.Values
                    SegAssociations &= pSeg.BoardPoint & " " & pSeg.Airline & " " & pSeg.OffPoint & vbCrLf
                Next
            End If

            If Element.Associations.Passengers.Count > 0 Then
                For Each objPax In Element.Associations.Passengers
                    PaxAssociations &= mobjPassengers(objPax.ElementNo).PaxName & vbCrLf
                Next
            Else
                For Each pPax As GDSPax.GDSPaxItem In mobjPassengers.Values
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
    Private Sub GetOtherServiceElements1A()

        Dim pobjOtherServiceElement As s1aPNR.OtherServiceElement

        For Each pobjOtherServiceElement In mobjPNR1A.OtherServiceElements
            parseOtherServiceElements1A(pobjOtherServiceElement)
        Next pobjOtherServiceElement

    End Sub

    Private Sub parseOtherServiceElements1A(ByVal Element As s1aPNR.OtherServiceElement)

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
    Private Sub GetTQT1A()

        Dim pTQTtext As k1aHostToolKit.CHostResponse = mobjSession1A.Send("TQT")
        Dim pTQT() As String = pTQTtext.Text.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

        ReDim mudtAllowance(0)
        mudtAllowance(0) = New TQT
        ReDim mudtTQT(0)
        mudtTQT(0) = New TQT
        If pTQT(0).StartsWith("T     P/S  NAME") Then
            For i As Integer = 1 To pTQT.GetUpperBound(0)
                If pTQT(i).Length > 62 AndAlso pTQT(i).Substring(0) <> " " Then
                    ReDim Preserve mudtTQT(mudtTQT.GetUpperBound(0) + 1)
                    mudtTQT(mudtTQT.GetUpperBound(0)) = New TQT
                    If pTQT(i).Substring(0, pTQT(i).IndexOf(" ")) <> pTQT(i - 1).Substring(0, pTQT(i - 1).IndexOf(" ")) AndAlso IsNumeric(pTQT(i).Substring(0, pTQT(i).IndexOf(" "))) Then
                        mudtTQT(mudtTQT.GetUpperBound(0)).TQTElement = pTQT(i).Substring(0, pTQT(i).IndexOf(" "))
                        Dim pSeg() As String
                        If i < pTQT.GetUpperBound(0) AndAlso pTQT(i + 1).Length > 2 AndAlso pTQT(i + 1).Substring(0, 1) = " " Then
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
                                        mudtTQT(mudtTQT.GetUpperBound(0)) = New TQT With {
                                            .TQTElement = mudtTQT(mudtTQT.GetUpperBound(0) - 1).TQTElement,
                                            .Segment = i2
                                        }
                                    Next
                                Else

                                End If
                            End If

                        Next
                    End If

                    Dim pTSTText As k1aHostToolKit.CHostResponse = mobjSession1A.Send("TQT/T" & pTQT(i).Substring(0, pTQT(i).IndexOf(" ")))
                    Dim pTST() As String = pTSTText.Text.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

                    SplitTQT1A(pTST)

                End If
            Next
        ElseIf pTQT(0).StartsWith("TST") Then
            SplitTQT1A(pTQT)
        End If

    End Sub
    Private Sub SplitTQT1A(ByVal pTQT() As String)

        Dim iSeg As Integer = 0
        For i As Integer = 0 To pTQT.GetUpperBound(0)
            If pTQT(i).Length > 3 AndAlso pTQT(i).Substring(4, 1) = "." Then
                iSeg = i + 1
            ElseIf iSeg > 0 Then
                Exit For
            End If
        Next
        If iSeg > 0 Then
            For i As Integer = iSeg To pTQT.GetUpperBound(0)
                If IsNumeric(pTQT(i).Substring(1, 1)) Then
                    If pTQT(i).Length > 60 Then
                        ReDim Preserve mudtAllowance(mudtAllowance.GetUpperBound(0) + 1)
                        mudtAllowance(mudtAllowance.GetUpperBound(0)) = New TQT With {
                            .Itin = pTQT(i).Substring(5, 6) & " " & pTQT(i + 1).Substring(5, 3),
                            .Allowance = pTQT(i).Substring(60),
                            .Status = pTQT(i).Substring(30, 3)
                        }
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

    Private Sub mobjPNR1GRaw_TerminalCommandSent(TerminalCommand As String) Handles mobjPNR1GRaw.TerminalCommandSent
        RaiseEvent TerminalCommandSent(TerminalCommand)
    End Sub
    'Private Sub GetPax1G()

    '    Dim pobjPax As Travelport.TravelData.Person

    '    mobjPassengers.Clear()

    '    For Each pobjPax In mobjPNR1G.Passengers
    '        With pobjPax
    '            mobjPassengers.AddItem(.PassengerNumber, .FirstName, .LastName, If(IsNothing(.NameRemark), "", .NameRemark))
    '        End With
    '    Next pobjPax

    'End Sub


    'Private Function setRecordLocator1G() As String
    '    Try
    '        setRecordLocator1G = mobjPNR1G.RecordLocator
    '    Catch ex As Exception
    '        setRecordLocator1G = UCase(mudtProps.RequestedPNR)
    '    End Try
    'End Function

End Class
