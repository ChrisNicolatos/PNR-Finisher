Option Strict Off
Option Explicit On
Imports s1aPNR
Friend Class ReadPNR
    'Event Timer1G(ByVal Loop1G As Integer)
    'Private Structure LineNumbers
    '    Dim Category As String
    '    Dim LineNumber As Integer
    'End Structure

    'Private WithEvents mobjSession As k1aHostToolKit.HostSession
    'Private mobjPNR1A As s1aPNR.PNR
    'Private mobjPNR1G As Travelport.TravelData.BookingFile
    'Private mobjPNR1GRaw As New GDS1GReadRaw

    'Private mobjPassengers As New GDSPax.GDSPaxColl
    'Private mobjSegments As New GDSSeg.GDSSegColl
    'Private mobjTicketElements As GDSTickets.Collection

    'Private mobjExistingGDSElements As New GDSExisting.Collection
    'Private mobjNewGDSElements As GDSNew.Collection

    'Private mGDSCode As Config.GDSCode

    'Private mstrPNRResponse As String
    'Private mstrPNRNumber As String
    'Private mflgNewPNR As Boolean
    'Private mstrGroupName As String
    'Private mintGroupNamesCount As Integer
    'Private mstrItinerary As String

    'Private mstrOfficeOfResponsibility As String
    'Private mflgQRSegment As Boolean = False
    'Private mdteDepartureDate As Date
    'Private mflgExistsSegments As Boolean
    'Private mflgExistsSSRDocs As Boolean
    'Private mstrSSRDocs As String
    'Public ReadOnly Property Segments As GDSSeg.GDSSegColl
    '    Get
    '        Segments = mobjSegments
    '    End Get
    'End Property
    'Public ReadOnly Property Passengers As GDSPax.GDSPaxColl
    '    Get
    '        Passengers = mobjPassengers
    '    End Get
    'End Property
    'Public ReadOnly Property GroupName As String
    '    Get
    '        GroupName = mstrGroupName
    '    End Get
    'End Property
    'Public ReadOnly Property GroupNamesCount As Integer
    '    Get
    '        GroupNamesCount = mintGroupNamesCount
    '    End Get
    'End Property
    'Public ReadOnly Property NumberOfPax As Integer
    '    Get
    '        NumberOfPax = mobjPassengers.Count
    '    End Get
    'End Property
    'Public ReadOnly Property PaxLeadName As String
    '    Get
    '        PaxLeadName = mobjPassengers.LeadName
    '    End Get
    'End Property
    'Public ReadOnly Property IsGroup As Boolean
    '    Get
    '        IsGroup = (mstrGroupName <> "")
    '    End Get
    'End Property
    'Public ReadOnly Property Itinerary As String
    '    Get
    '        Itinerary = mstrItinerary
    '    End Get
    'End Property
    'Public ReadOnly Property PnrNumber As String
    '    Get
    '        PnrNumber = mstrPNRNumber
    '    End Get
    'End Property
    'Public ReadOnly Property OfficeOfResponsibility As String
    '    Get
    '        OfficeOfResponsibility = mstrOfficeOfResponsibility
    '    End Get
    'End Property
    'Public ReadOnly Property DepartureDate As Date
    '    Get
    '        DepartureDate = mdteDepartureDate
    '    End Get
    'End Property
    'Public ReadOnly Property ExistingElements As GDSExisting.Collection
    '    Get
    '        ExistingElements = mobjExistingGDSElements
    '    End Get
    'End Property
    'Public ReadOnly Property NewElements As GDSNew.Collection
    '    Get
    '        NewElements = mobjNewGDSElements
    '    End Get
    'End Property
    'Public ReadOnly Property HasQRSegment As Boolean
    '    Get
    '        HasQRSegment = mflgQRSegment
    '    End Get
    'End Property
    'Public ReadOnly Property SegmentsExist As Boolean
    '    Get
    '        SegmentsExist = mflgExistsSegments
    '    End Get
    'End Property
    'Public ReadOnly Property SSRDocsExists As Boolean
    '    Get
    '        SSRDocsExists = mflgExistsSSRDocs
    '    End Get
    'End Property
    'Public ReadOnly Property SSRDocs As String
    '    Get
    '        SSRDocs = mstrSSRDocs
    '    End Get
    'End Property

    'Public ReadOnly Property GDSCode As Config.GDSCode
    '    Get
    '        GDSCode = mGDSCode
    '    End Get
    'End Property
    'Public ReadOnly Property NewPNR As Boolean
    '    Get
    '        NewPNR = mflgNewPNR
    '    End Get
    'End Property
    ''Public ReadOnly Property Tickets As GDSTickets.Collection
    ''    Get
    ''        Tickets = mobjTicketElements
    ''    End Get
    ''End Property
    'Private Sub mobjSession_ReceivedResponse(ByRef newResponse As k1aHostToolKit.CHostResponse) Handles mobjSession.ReceivedResponse
    '    mstrPNRResponse = newResponse.Text
    'End Sub
    'Public Function Read(ByVal GDSCode As Config.GDSCode) As String

    '    mGDSCode = GDSCode

    '    If mGDSCode = Config.GDSCode.GDSisAmadeus Then
    '        Read = Read1A()
    '    ElseIf mGDSCode = Config.GDSCode.GDSisGalileo Then
    '        Read = Read1G()
    '    Else
    '        Throw New Exception("ReadPNR.Read()" & vbCrLf & "NO GDS Specified")
    '    End If
    'End Function
    'Private Function Read1A() As String

    '    If mGDSCode <> Config.GDSCode.GDSisAmadeus Then
    '        Throw New Exception("ReadPNR.Read1A()" & vbCrLf & "Selected GDS is not Amadeus")
    '    End If

    '    Dim Sessions As k1aHostToolKit.HostSessions

    '    Read1A = ""

    '    Try
    '        ' To be able to retrieve the PNR that have been created we need to link our '
    '        ' application to the current session of the FOS                             '
    '        Sessions = New k1aHostToolKit.HostSessions

    '        If Sessions.Count > 0 Then
    '            ' There is at least one session opened.                    '
    '            ' We link our application to the active session of the FOS '
    '            mobjSession = Sessions.UIActiveSession

    '            ' Initialize the PNR
    '            mobjPNR1A = New s1aPNR.PNR

    '            ' Retrieve the name elements, Air segments and Hotel Segments of the current PNR
    '            Dim pStatus As Integer = mobjPNR1A.RetrievePNR(mobjSession, "RT")
    '            mflgNewPNR = False

    '            If pStatus = 0 Or pStatus = 1005 Then
    '                GetOfficeOfResponsibility1A()
    '                GetPnrNumber1A()

    '                GetGroup1A()
    '                GetPassengers1A()
    '                GetSegments1A()
    '                GetPhoneElement1A()
    '                GetEmailElement1A()
    '                GetAOH1A()
    '                GetOpenSegment1A()
    '                GetTicketElement1A()
    '                GetOptionQueueElement1A()
    '                GetVesselOSI1A()
    '                GetSSR1A()
    '                GetRM1A()
    '                GetTickets1A()
    '                If mobjPNR1A.RawResponse.IndexOf("***  NHP  ***") >= 0 Then
    '                    Read1A = "               ***  NHP  ***"
    '                Else
    '                    Read1A = CheckDMI1A()
    '                End If
    '            Else
    '                Throw New Exception("There is no active PNR" & vbCrLf & mstrPNRResponse)
    '            End If
    '        Else
    '            Throw New Exception("Please start Amadeus and retry")
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception("Read1A()" & vbCrLf & ex.Message)
    '    End Try
    'End Function
    'Private Function Read1G() As String

    '    Read1G = ""
    '    Try
    '        Dim Session As New Travelport.TravelData.Factory.GalileoDesktopFactory("SPG720", "MyConnection", False, True, "SMRT")
    '        Dim pErr As Integer = 1

    '        Do While pErr < 90
    '            Try
    '                mobjPNR1GRaw.ReadRaw()
    '                'Dim mstrPNR = Session.SendTerminalCommand("*ALL")
    '                'Read1G = ""
    '                'For i = 0 To mstrPNR.Count - 1
    '                '    Read1G &= mstrPNR(i) & vbCrLf
    '                'Next
    '                'Return Read1G

    '                'mobjPNR1G = Session.RetrieveCurrentBookingFile
    '                pErr = 99
    '            Catch ex As Exception
    '                System.Threading.Thread.Sleep(2000)
    '                pErr += 1
    '                RaiseEvent Timer1G(pErr)
    '            End Try
    '        Loop
    '        If pErr < 99 Then
    '            Throw New Exception("Galileo communication problem. Please try again or contact your system administrator")
    '        End If

    '        If mobjPNR1G.IsEmpty Then
    '            Throw New Exception("No B.F. to display")
    '        Else
    '            GetOfficeOfResponsibility1G()
    '            GetPnrNumber1G()

    '            'GetGroup1G()
    '            GetPassengers1G()
    '            GetSegments1G()
    '            GetPhoneElement1G()
    '            GetEmailElement1G()
    '            GetOpenSegment1G()
    '            GetTicketElement1G()
    '            GetOptionQueueElement1G()
    '            GetVesselOSI1G()
    '            GetSSR1G()
    '            GetRM1G()
    '            'GetTickets1G()
    '            If mobjPNR1G.ToString.IndexOf("***  NHP  ***") >= 0 Then
    '                Read1G = "               ***  NHP  ***"
    '            Else
    '                'Read1G = CheckDMI1G()
    '            End If
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception("Read1G()" & vbCrLf & ex.Message)
    '    End Try

    'End Function
    'Public Sub PrepareNewGDSElements()
    '    mobjNewGDSElements = New GDSNew.Collection(OfficeOfResponsibility, DepartureDate, NumberOfPax, mGDSCode)
    'End Sub
    'Private Function CheckDMI1A() As String
    '    If mGDSCode <> Config.GDSCode.GDSisAmadeus Then
    '        Throw New Exception("ReadPNR.CheckDMI1A()" & vbCrLf & "Selected GDS is not Amadeus")
    '    End If

    '    Try
    '        If mobjPNR1A.AirSegments.Count <= 1 Then
    '            Return ""
    '        End If

    '        Dim pDMI As String = mobjSession.Send("DMI").Text
    '        If pDMI.Contains("ITINERARY OK") Then
    '            Return ""
    '        Else
    '            Return pDMI
    '        End If
    '    Catch ex As Exception
    '        Return ""
    '    End Try

    'End Function
    'Private Sub RemoveOldGDSEntries1A()

    '    If mGDSCode <> Config.GDSCode.GDSisAmadeus Then
    '        Throw New Exception("ReadPNR.RemoveOldGDSEntries1A()" & vbCrLf & "Selected GDS is not Amadeus")
    '    End If

    '    Dim pLineNumbers(0) As Integer

    '    ' the following elements remain as they are if they already exist in the PNR
    '    ClearExistingItems(mobjExistingGDSElements.PhoneElement, mobjNewGDSElements.PhoneElement)
    '    ClearExistingItems(mobjExistingGDSElements.EmailElement, mobjNewGDSElements.EmailElement)
    '    ClearExistingItems(mobjExistingGDSElements.AOH, mobjNewGDSElements.AOH)
    '    ClearExistingItems(mobjExistingGDSElements.OpenSegment, mobjNewGDSElements.OpenSegment)
    '    ClearExistingItems(mobjExistingGDSElements.OptionQueueElement, mobjNewGDSElements.OptionQueueElement)
    '    ClearExistingItems(mobjExistingGDSElements.TicketElement, mobjNewGDSElements.TicketElement)
    '    ClearExistingItems(mobjExistingGDSElements.AgentID, mobjNewGDSElements.AgentID)

    '    ' the following elements are removed and replaced if they exist in the PNR
    '    PrepareLineNumbers1A(mobjExistingGDSElements.CustomerCode, pLineNumbers)
    '    PrepareLineNumbers1A(mobjExistingGDSElements.CustomerName, pLineNumbers)
    '    PrepareLineNumbers1A(mobjExistingGDSElements.SubDepartmentCode, pLineNumbers)
    '    PrepareLineNumbers1A(mobjExistingGDSElements.SubDepartmentName, pLineNumbers)
    '    PrepareLineNumbers1A(mobjExistingGDSElements.CRMCode, pLineNumbers)
    '    PrepareLineNumbers1A(mobjExistingGDSElements.CRMName, pLineNumbers)
    '    PrepareLineNumbers1A(mobjExistingGDSElements.VesselFlag, pLineNumbers)
    '    PrepareLineNumbers1A(mobjExistingGDSElements.VesselName, pLineNumbers)
    '    PrepareLineNumbers1A(mobjExistingGDSElements.VesselOSI, pLineNumbers)
    '    PrepareLineNumbers1A(mobjExistingGDSElements.Reference, pLineNumbers)
    '    PrepareLineNumbers1A(mobjExistingGDSElements.BookedBy, pLineNumbers)
    '    PrepareLineNumbers1A(mobjExistingGDSElements.Department, pLineNumbers)
    '    PrepareLineNumbers1A(mobjExistingGDSElements.ReasonForTravel, pLineNumbers)
    '    PrepareLineNumbers1A(mobjExistingGDSElements.CostCentre, pLineNumbers)

    '    Dim pMax As Integer = 0
    '    Dim pMaxIndex As Integer = -1
    '    Dim pFound As Boolean = True
    '    Do While pFound
    '        pFound = False
    '        For i As Integer = 0 To pLineNumbers.GetUpperBound(0)
    '            If pLineNumbers(i) > pMax Then
    '                pMax = pLineNumbers(i)
    '                pMaxIndex = i
    '                pFound = True
    '            End If
    '        Next
    '        If pMaxIndex > -1 Then
    '            mobjSession.Send("XE" & pMax)
    '            pLineNumbers(pMaxIndex) = 0
    '        End If
    '        pMax = 0
    '        pMaxIndex = -1
    '    Loop

    'End Sub
    'Private Sub RemoveOldGDSEntries1G()

    '    If mGDSCode <> Config.GDSCode.GDSisGalileo Then
    '        Throw New Exception("ReadPNR.RemoveOldGDSEntries1G()" & vbCrLf & "Selected GDS is not Galileo")
    '    End If

    '    Dim pLineNumbers(0) As LineNumbers

    '    ' the following elements remain as they are if they already exist in the PNR
    '    ClearExistingItems(mobjExistingGDSElements.PhoneElement, mobjNewGDSElements.PhoneElement)
    '    ClearExistingItems(mobjExistingGDSElements.EmailElement, mobjNewGDSElements.EmailElement)
    '    ClearExistingItems(mobjExistingGDSElements.AOH, mobjNewGDSElements.AOH)
    '    'ClearExistingItems(mobjExistingGDSElements.OpenSegment, mobjNewGDSElements.OpenSegment)
    '    'ClearExistingItems(mobjExistingGDSElements.OptionQueueElement, mobjNewGDSElements.OptionQueueElement)
    '    'ClearExistingItems(mobjExistingGDSElements.TicketElement, mobjNewGDSElements.TicketElement)
    '    'ClearExistingItems(mobjExistingGDSElements.AgentID, mobjNewGDSElements.AgentID)

    '    ' the following elements are removed and replaced if they exist in the PNR
    '    PrepareLineNumbers1G(mobjExistingGDSElements.OpenSegment, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.AgentID, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.OptionQueueElement, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.TicketElement, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.CustomerCode, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.CustomerName, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.SubDepartmentCode, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.SubDepartmentName, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.CRMCode, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.CRMName, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.VesselFlag, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.VesselName, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.VesselOSI, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.Reference, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.BookedBy, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.Department, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.ReasonForTravel, pLineNumbers)
    '    PrepareLineNumbers1G(mobjExistingGDSElements.CostCentre, pLineNumbers)

    '    Dim pMax As Integer = 0
    '    Dim pMaxIndex As Integer = -1
    '    Dim pCategory As String = ""
    '    Dim pFound As Boolean = True

    '    Do While pFound
    '        If pCategory = "" Then
    '            For i As Integer = 0 To pLineNumbers.GetUpperBound(0)
    '                If pLineNumbers(i).Category <> "" Then
    '                    pCategory = pLineNumbers(i).Category
    '                    pMax = pLineNumbers(i).LineNumber
    '                    pMaxIndex = i
    '                    Exit For
    '                End If
    '            Next
    '        End If
    '        If pCategory <> "" Then
    '            For i As Integer = 0 To pLineNumbers.GetUpperBound(0)
    '                If pLineNumbers(i).Category = pCategory And pLineNumbers(i).LineNumber > pMax Then
    '                    pMax = pLineNumbers(i).LineNumber
    '                    pMaxIndex = i
    '                    pFound = True
    '                End If
    '            Next
    '            Dim Session As New Travelport.TravelData.Factory.GalileoDesktopFactory("SPG720", "MYCONNECTION", False, True, "SMRT")
    '            Dim pResponse
    '            If pMaxIndex > -1 Then
    '                If pCategory = "Segment." Then
    '                    pResponse = Session.SendTerminalCommand("X" & pMax)
    '                Else
    '                    pResponse = Session.SendTerminalCommand(pCategory & pMax & "@")
    '                End If
    '                If pResponse(0) = "INVALID ENTRY" Then
    '                    pResponse = Session.SendTerminalCommand(pCategory & "@")
    '                End If
    '                pLineNumbers(pMaxIndex).Category = ""
    '                pLineNumbers(pMaxIndex).LineNumber = 0
    '            Else
    '                pCategory = ""
    '            End If
    '            pMax = 0
    '            pMaxIndex = -1
    '        Else
    '            pFound = False
    '        End If
    '    Loop

    'End Sub

    'Private Sub ClearExistingItems(ByRef ExistingItem As GDSExisting.Item, ByRef NewItem As GDSNew.Item)
    '    If ExistingItem.Exists Then
    '        NewItem.Clear()
    '    End If
    'End Sub

    ''Private Sub PrepareLineNumbers1A(ByVal ExistingItem As GDSExisting.Item, ByRef pLineNumbers() As Integer)
    ''    If ExistingItem.Exists Then
    ''        ReDim Preserve pLineNumbers(pLineNumbers.GetUpperBound(0) + 1)
    ''        pLineNumbers(pLineNumbers.GetUpperBound(0)) = ExistingItem.LineNumber
    ''    End If
    ''End Sub
    'Private Sub PrepareLineNumbers1G(ByVal ExistingItem As GDSExisting.Item, ByRef pLineNumbers() As LineNumbers)
    '    If ExistingItem.Exists Then
    '        Dim pItems() As String = ExistingItem.Category.Split(".")
    '        If IsArray(pItems) AndAlso pItems(0) <> "" Then
    '            ReDim Preserve pLineNumbers(pLineNumbers.GetUpperBound(0) + 1)
    '            pLineNumbers(pLineNumbers.GetUpperBound(0)).Category = pItems(0) & "."
    '            pLineNumbers(pLineNumbers.GetUpperBound(0)).LineNumber = ExistingItem.LineNumber
    '        End If
    '    End If
    'End Sub
    'Public Sub SendGDSEntry1A(ByVal GDSEntry As String)

    '    If mGDSCode <> Config.GDSCode.GDSisAmadeus Then
    '        Throw New Exception("ReadPNR.SendNewGDSEntries1A()" & vbCrLf & "Selected GDS is not Amadeus")
    '    End If

    '    If GDSEntry <> "" Then
    '        mobjSession.Send(GDSEntry)
    '    End If

    'End Sub
    'Public Sub SendGDSEntry1G(ByVal GDSEntry As String)

    '    If mGDSCode <> Config.GDSCode.GDSisGalileo Then
    '        Throw New Exception("ReadPNR.SendNewGDSEntries1G()" & vbCrLf & "Selected GDS is not Galileo")
    '    End If
    '    Dim Session As New Travelport.TravelData.Factory.GalileoDesktopFactory("SPG720", "MYCONNECTION", False, True, "SMRT")

    '    If GDSEntry <> "" Then
    '        Session.SendTerminalCommand(GDSEntry)
    '    End If

    'End Sub
    'Public Function SendAllGDSEntries(ByVal WritePNR As Boolean, ByVal WriteDocs As Boolean, ByVal mflgExpiryDateOK As Boolean, dgvApis As DataGridView, AirlineEntries As TextBox) As String

    '    SendAllGDSEntries = ""
    '    If mGDSCode = Config.GDSCode.GDSisAmadeus Then
    '        SendAllGDSEntries1A(WritePNR, WriteDocs, mflgExpiryDateOK, dgvApis, AirlineEntries)
    '    ElseIf mGDSCode = Config.GDSCode.GDSisGalileo Then
    '        SendAllGDSEntries = SendAllGDSEntries1G(WritePNR, WriteDocs, mflgExpiryDateOK, dgvApis, AirlineEntries)
    '    Else
    '        Throw New Exception("ReadPNR.SendAllGDSEntries()" & vbCrLf & "No GDS Selected")
    '    End If

    'End Function
    'Private Sub SendAllGDSEntries1A(ByVal WritePNR As Boolean, ByVal WriteDocs As Boolean, ByVal mflgExpiryDateOK As Boolean, dgvApis As DataGridView, AirlineEntries As TextBox)
    '    Try
    '        If WritePNR Then
    '            RemoveOldGDSEntries1A()

    '            SendGDSElement1A(mobjNewGDSElements.PhoneElement)
    '            SendGDSElement1A(mobjNewGDSElements.EmailElement)
    '            SendGDSElement1A(mobjNewGDSElements.AgentID)
    '            SendGDSElement1A(mobjNewGDSElements.AOH)
    '            SendGDSElement1A(mobjNewGDSElements.OpenSegment)
    '            SendGDSElement1A(mobjNewGDSElements.TicketElement)
    '            SendGDSElement1A(mobjNewGDSElements.OptionQueueElement)

    '            If mflgNewPNR Then
    '                SendGDSElement1A(mobjNewGDSElements.SavingsElement)
    '                SendGDSElement1A(mobjNewGDSElements.LossElement)
    '            End If

    '            SendGDSElement1A(mobjNewGDSElements.CustomerCode)
    '            SendGDSElement1A(mobjNewGDSElements.CustomerName)
    '            SendGDSElement1A(mobjNewGDSElements.SubDepartmentCode)
    '            SendGDSElement1A(mobjNewGDSElements.SubDepartmentName)
    '            SendGDSElement1A(mobjNewGDSElements.CRMCode)
    '            SendGDSElement1A(mobjNewGDSElements.CRMName)
    '            SendGDSElement1A(mobjNewGDSElements.VesselName)
    '            SendGDSElement1A(mobjNewGDSElements.VesselFlag)
    '            SendGDSElement1A(mobjNewGDSElements.VesselOSI)
    '            SendGDSElement1A(mobjNewGDSElements.Reference)
    '            SendGDSElement1A(mobjNewGDSElements.BookedBy)
    '            SendGDSElement1A(mobjNewGDSElements.Department)
    '            SendGDSElement1A(mobjNewGDSElements.ReasonForTravel)
    '            SendGDSElement1A(mobjNewGDSElements.CostCentre)

    '            Dim pAirlineEntries() As String = AirlineEntries.Text.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

    '            For i As Integer = 0 To pAirlineEntries.GetUpperBound(0)
    '                pAirlineEntries(i) = pAirlineEntries(i).Replace(">", "").Trim
    '                If pAirlineEntries(i).Trim <> "" Then
    '                    SendGDSAirlineItems1A(pAirlineEntries(i).Replace("> ", ""))
    '                End If
    '            Next
    '        End If

    '        If WriteDocs Then
    '            APISUpdate1A(mflgExpiryDateOK, dgvApis)
    '        End If

    '        If WritePNR Or WriteDocs Then
    '            CloseOffPNR1A()
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception("SendNewGDSEntries()" & vbCrLf & ex.Message)
    '    End Try
    'End Sub
    'Private Function SendAllGDSEntries1G(ByVal WritePNR As Boolean, ByVal WriteDocs As Boolean, ByVal mflgExpiryDateOK As Boolean, dgvApis As DataGridView, AirlineEntries As TextBox) As String
    '    Try
    '        SendAllGDSEntries1G = ""
    '        If WritePNR Then
    '            RemoveOldGDSEntries1G()

    '            SendGDSElement1G(mobjNewGDSElements.PhoneElement, True)
    '            SendGDSElement1G(mobjNewGDSElements.EmailElement, True)
    '            SendGDSElement1G(mobjNewGDSElements.AgentID, True)
    '            SendGDSElement1G(mobjNewGDSElements.AOH, True)
    '            SendGDSElement1G(mobjNewGDSElements.OpenSegment, False)
    '            SendGDSElement1G(mobjNewGDSElements.TicketElement, True)
    '            SendGDSElement1G(mobjNewGDSElements.OptionQueueElement, True)

    '            If mflgNewPNR Then
    '                SendGDSElement1G(mobjNewGDSElements.SavingsElement, True)
    '                SendGDSElement1G(mobjNewGDSElements.LossElement, True)
    '            End If

    '            SendGDSElement1G(mobjNewGDSElements.CustomerCode, True)
    '            SendGDSElement1G(mobjNewGDSElements.CustomerName, True)
    '            SendGDSElement1G(mobjNewGDSElements.SubDepartmentCode, True)
    '            SendGDSElement1G(mobjNewGDSElements.SubDepartmentName, True)
    '            SendGDSElement1G(mobjNewGDSElements.CRMCode, True)
    '            SendGDSElement1G(mobjNewGDSElements.CRMName, True)
    '            SendGDSElement1G(mobjNewGDSElements.VesselName, True)
    '            SendGDSElement1G(mobjNewGDSElements.VesselFlag, True)
    '            SendGDSElement1G(mobjNewGDSElements.VesselOSI, True)
    '            SendGDSElement1G(mobjNewGDSElements.Reference, True)
    '            SendGDSElement1G(mobjNewGDSElements.BookedBy, True)
    '            SendGDSElement1G(mobjNewGDSElements.Department, True)
    '            SendGDSElement1G(mobjNewGDSElements.ReasonForTravel, True)
    '            SendGDSElement1G(mobjNewGDSElements.CostCentre, True)
    '            SendGDSElement1G(mobjNewGDSElements.GalileoTrackingCode, True)

    '            Dim pAirlineEntries() As String = AirlineEntries.Text.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

    '            For i As Integer = 0 To pAirlineEntries.GetUpperBound(0)
    '                pAirlineEntries(i) = pAirlineEntries(i).Replace(">", "").Trim
    '                If pAirlineEntries(i).Trim <> "" Then
    '                    SendGDSAirlineItems1G(pAirlineEntries(i).Replace("> ", ""))
    '                End If
    '            Next
    '        End If

    '        If WriteDocs Then
    '            APISUpdate1G(mflgExpiryDateOK, dgvApis)
    '        End If

    '        If WritePNR Or WriteDocs Then
    '            SendAllGDSEntries1G = CloseOffPNR1G()
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception("SendNewGDSEntries()" & vbCrLf & ex.Message)
    '    End Try
    'End Function
    'Private Sub APISUpdate1A(ByVal mflgExpiryDateOK As Boolean, dgvApis As DataGridView)

    '    Dim pstrCommand As String
    '    Try
    '        For i = 0 To dgvApis.RowCount - 1
    '            With dgvApis.Rows(i)
    '                If .ErrorText.IndexOf("Birth") = -1 Then
    '                    Dim pobjItem As New PaxApisDB.Item(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value,
    '                                                   APISDateFromIATA(.Cells(6).Value), .Cells(7).Value, .Cells(3).Value,
    '                                                 .Cells(4).Value, APISDateFromIATA(.Cells(8).Value), .Cells(5).Value)

    '                    pobjItem.Update(mflgExpiryDateOK)
    '                    pstrCommand = "SR DOCS YY HK1-P-" & pobjItem.IssuingCountry & "-" & pobjItem.PassportNumber & "-" & pobjItem.Nationality &
    '                "-" & APISDateToIATA(pobjItem.BirthDate) & "-" & pobjItem.Gender & "-"
    '                    If mflgExpiryDateOK Then
    '                        pstrCommand &= APISDateToIATA(pobjItem.ExpiryDate)
    '                    Else
    '                        pstrCommand &= ""
    '                    End If
    '                    pstrCommand &= "-" & pobjItem.Surname & "-" & pobjItem.FirstName & "/P" & pobjItem.Id
    '                    SendGDSEntry1A(pstrCommand)
    '                End If

    '            End With

    '        Next
    '    Catch ex As Exception
    '        Throw New Exception("APISUpdate()" & vbCrLf & ex.Message)
    '    End Try


    'End Sub
    'Private Sub APISUpdate1G(ByVal mflgExpiryDateOK As Boolean, dgvApis As DataGridView)

    '    If mGDSCode <> Config.GDSCode.GDSisGalileo Then
    '        Throw New Exception("ReadPNR.SendGDSElement1G()" & vbCrLf & "Selected GDS is not Galileo")
    '    End If
    '    Dim Session As New Travelport.TravelData.Factory.GalileoDesktopFactory("SPG720", "MYCONNECTION", False, True, "SMRT")

    '    Dim pstrCommand As String
    '    Try
    '        For i = 0 To dgvApis.RowCount - 1
    '            With dgvApis.Rows(i)
    '                If .ErrorText.IndexOf("Birth") = -1 Then
    '                    Dim pobjItem As New PaxApisDB.Item(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value,
    '                                                   APISDateFromIATA(.Cells(6).Value), .Cells(7).Value, .Cells(3).Value,
    '                                                 .Cells(4).Value, APISDateFromIATA(.Cells(8).Value), .Cells(5).Value)

    '                    pobjItem.Update(mflgExpiryDateOK)
    '                    'SI.P1/SSRDOCSBAHK1/P/GB/S12345678/GB/12JUL76/M/23OCT16/SMITH/JOHN/RICHARD
    '                    pstrCommand = "SI.P" & pobjItem.Id & "/SSRDOCSYYHK1/P/" & pobjItem.IssuingCountry & "/" & pobjItem.PassportNumber & "/" & pobjItem.Nationality &
    '                "/" & APISDateToIATA(pobjItem.BirthDate) & "/" & pobjItem.Gender & "/"
    '                    If mflgExpiryDateOK Then
    '                        pstrCommand &= APISDateToIATA(pobjItem.ExpiryDate)
    '                    Else
    '                        pstrCommand &= ""
    '                    End If
    '                    pstrCommand &= "/" & pobjItem.Surname & "/" & pobjItem.FirstName
    '                    For Each pElement As Travelport.TravelData.BookingFileManualSsr In mobjPNR1G.ManualSsrs
    '                        If pElement.TextLastName = pobjItem.Surname And pElement.TextAddress = pobjItem.PassportNumber And pElement.TextDateOfBirth = APISDateToIATA(pobjItem.BirthDate) Then
    '                            pstrCommand = ""
    '                        End If
    '                    Next
    '                    If pstrCommand <> "" Then
    '                        SendGDSEntry1G(pstrCommand)
    '                    End If
    '                End If
    '            End With
    '        Next
    '    Catch ex As Exception
    '        Throw New Exception("APISUpdate()" & vbCrLf & ex.Message)
    '    End Try
    'End Sub
    'Private Sub CloseOffPNR1A()
    '    If mGDSCode <> Config.GDSCode.GDSisAmadeus Then
    '        Throw New Exception("ReadPNR.CloseOffPNR1A()" & vbCrLf & "Selected GDS is not Amadeus")
    '    End If

    '    Dim pCloseOffEntries As New CloseOffEntries.Collection

    '    pCloseOffEntries.Load(MySettings.GDSPcc, mstrOfficeOfResponsibility = MySettings.GDSPcc)

    '    For Each pCommand As CloseOffEntries.Item In pCloseOffEntries.Values
    '        mobjSession.Send(pCommand.CloseOffEntry)
    '    Next
    '    If mstrPNRResponse.Contains("WARNING: SECURE FLT PASSENGER DATA REQUIRED") Then
    '        MessageBox.Show(mstrPNRResponse)
    '    End If

    'End Sub
    'Private Function CloseOffPNR1G() As String
    '    If mGDSCode <> Config.GDSCode.GDSisGalileo Then
    '        Throw New Exception("ReadPNR.CloseOffPNR1G()" & vbCrLf & "Selected GDS is not Amadeus")
    '    End If
    '    Dim Session As New Travelport.TravelData.Factory.GalileoDesktopFactory("SPG720", "MYCONNECTION", False, True, "SMRT")
    '    Dim pCloseOffEntries As New CloseOffEntries.Collection
    '    CloseOffPNR1G = ""
    '    pCloseOffEntries.Load(MySettings.GDSPcc, mstrOfficeOfResponsibility = MySettings.GDSPcc)

    '    Dim pResponse
    '    Dim pPNR As String
    '    pResponse = Session.SendTerminalCommand("R.CN")
    '    pResponse = Session.SendTerminalCommand("ER")
    '    If pResponse(0).ToString.Length > 9 AndAlso pResponse(0).ToString.Substring(6, 1) = "/" Then


    '        pPNR = pResponse(0).ToString.Substring(0, 6)
    '        pResponse = Session.SendTerminalCommand("I")
    '        For Each pCommand As CloseOffEntries.Item In pCloseOffEntries.Values
    '            pResponse = Session.SendTerminalCommand("*" & pPNR)
    '            pResponse = Session.SendTerminalCommand(pCommand.CloseOffEntry)
    '            If pResponse(0) & pResponse(1) <> " *>" And pResponse(0).ToString.IndexOf("ON QUEUE") = -1 Then
    '                MessageBox.Show(pCommand.CloseOffEntry & vbCrLf & pResponse(0) & pResponse(1))
    '            End If
    '            pResponse = Session.SendTerminalCommand("I")
    '        Next
    '        pResponse = Session.SendTerminalCommand("*" & pPNR)
    '        pResponse = Session.SendTerminalCommand("IR")
    '        CloseOffPNR1G = pPNR
    '    Else
    '        MessageBox.Show(pResponse(0) & vbCrLf & pResponse(1), "ERROR IN PNR UPDATE", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Session.SendTerminalCommand("IR")
    '        Throw New Exception("Error in PNR Update")
    '    End If
    'End Function
    'Private Sub SendGDSElement1A(ByVal pElement As GDSNew.Item)
    '    If mGDSCode <> Config.GDSCode.GDSisAmadeus Then
    '        Throw New Exception("ReadPNR.SendGDSElement1A()" & vbCrLf & "Selected GDS is not Amadeus")
    '    End If

    '    If pElement.GDSCommand <> "" Then
    '        mobjSession.Send(pElement.GDSCommand)
    '    End If

    'End Sub
    'Private Sub SendGDSElement1G(ByVal pElement As GDSNew.Item, ByVal ShowResponse As Boolean)
    '    If mGDSCode <> Config.GDSCode.GDSisGalileo Then
    '        Throw New Exception("ReadPNR.SendGDSElement1G()" & vbCrLf & "Selected GDS is not Galileo")
    '    End If
    '    Dim Session As New Travelport.TravelData.Factory.GalileoDesktopFactory("SPG720", "MYCONNECTION", False, True, "SMRT")
    '    Dim pResponse
    '    If pElement.GDSCommand <> "" Then
    '        pResponse = Session.SendTerminalCommand(pElement.GDSCommand)
    '        If ShowResponse AndAlso pResponse(0) & pResponse(1) <> " *>" Then
    '            MessageBox.Show(pElement.GDSCommand & vbCrLf & pResponse(0) & pResponse(1))
    '        End If
    '    End If

    'End Sub
    'Private Sub SendGDSAirlineItems1A(ByVal pItemToSend As String)
    '    If mGDSCode <> Config.GDSCode.GDSisAmadeus Then
    '        Throw New Exception("ReadPNR.SendGDSAirlineItems1A()" & vbCrLf & "Selected GDS is not Amadeus")
    '    End If

    '    If pItemToSend.StartsWith("OS ") Then
    '        If mobjPNR1A.RawResponse.Replace(vbCrLf, "").Replace(" ", "").IndexOf(("OSI " & pItemToSend.Substring(3)).Replace(" ", "")) = -1 Then
    '            mobjSession.Send(pItemToSend)
    '        End If
    '    ElseIf pItemToSend.StartsWith("R") Then
    '        If mobjPNR1A.RawResponse.Replace(vbCrLf, "").Replace(" ", "").IndexOf(pItemToSend.Replace(" ", "")) = -1 Then
    '            mobjSession.Send(pItemToSend)
    '        End If
    '    ElseIf pItemToSend.StartsWith("S") Then
    '        Dim pString As String
    '        pString = pItemToSend.Replace(" ", "").Replace("SRCKIN-", "")
    '        If mobjPNR1A.RawResponse.Replace(vbCrLf, "").Replace(" ", "").IndexOf(pString) = -1 Then
    '            mobjSession.Send(pItemToSend)
    '        End If
    '    Else
    '        mobjSession.Send(pItemToSend)
    '    End If

    'End Sub
    'Private Sub SendGDSAirlineItems1G(ByVal pItemToSend As String)
    '    If mGDSCode <> Config.GDSCode.GDSisGalileo Then
    '        Throw New Exception("ReadPNR.SendGDSAirlineItems1G()" & vbCrLf & "Selected GDS is not Galileo")
    '    End If

    '    Dim pResponse
    '    If pItemToSend <> "" Then
    '        Dim Session As New Travelport.TravelData.Factory.GalileoDesktopFactory("SPG720", "MYCONNECTION", False, True, "SMRT")
    '        If pItemToSend.StartsWith("DI.") Then
    '            For Each pElement As Travelport.TravelData.BookingFileRemark In mobjPNR1G.InvoiceRemarks
    '                If pElement.RemarkType & pElement.Text = pItemToSend.Replace(" ", "") Then
    '                    pItemToSend = ""
    '                    Exit For
    '                End If
    '            Next
    '        ElseIf pItemToSend.StartsWith("SI.") Then
    '            For Each pElement In mobjPNR1G.OtherSupplementaryInformationRemarks
    '                If ("SI." & pElement.Carrier.Code & "*" & pElement.Message).Replace(" ", "") = pItemToSend.Replace(" ", "") Then
    '                    pItemToSend = ""
    '                    Exit For
    '                End If
    '            Next
    '            For Each pElement As Travelport.TravelData.BookingFileManualSsr In mobjPNR1G.ManualSsrs
    '                If "SI.SSR" & pElement.Code & pElement.VendorCode & pElement.StatusCode & "1" & pElement.Text = pItemToSend Then
    '                    pItemToSend = ""
    '                End If
    '            Next
    '        End If
    '        If pItemToSend <> "" Then
    '            pResponse = Session.SendTerminalCommand(pItemToSend)
    '            If pResponse(0) & pResponse(1) <> " *>" Then
    '                MessageBox.Show(pItemToSend & vbCrLf & pResponse(0) & pResponse(1))
    '            End If
    '        End If
    '    End If

    'End Sub
    'Private Sub GetPnrNumber1A()

    '    Try
    '        mstrPNRNumber = mobjPNR1A.Header.RecordLocator
    '    Catch ex As Exception
    '        mstrPNRNumber = ""
    '    End Try

    '    If mstrPNRNumber = "" Then
    '        mstrPNRNumber = "New PNR"
    '        mflgNewPNR = True
    '    End If
    'End Sub
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
    'Private Sub GetOfficeOfResponsibility1A()

    '    Try
    '        mstrOfficeOfResponsibility = mobjPNR1A.Header.OfficeOfResponsability
    '    Catch ex As Exception
    '        mstrOfficeOfResponsibility = MySettings.GDSPcc
    '    End Try

    'End Sub
    'Private Sub GetOfficeOfResponsibility1G()
    '    Try
    '        mstrOfficeOfResponsibility = mobjPNR1G.CurrentAgencyPcc
    '    Catch ex As Exception
    '        mstrOfficeOfResponsibility = MySettings.GDSPcc
    '    End Try
    'End Sub
    'Private Sub GetGroup1A()

    '    mstrGroupName = ""
    '    mintGroupNamesCount = 0

    '    For Each pGroup As s1aPNR.GroupNameElement In mobjPNR1A.GroupNameElements
    '        mstrGroupName = pGroup.GroupName
    '        mintGroupNamesCount = pGroup.NbrOfAssignedNames + pGroup.NbrNamesMissing
    '        Exit For
    '    Next
    '    If mobjPNR1A.GroupNameElements.Count > 1 Then
    '        mstrGroupName &= "x" & mobjPNR1A.GroupNameElements.Count
    '    End If

    'End Sub
    'Private Sub GetPassengers1A()
    '    mobjPassengers.Clear()
    '    For Each Pax As s1aPNR.NameElement In mobjPNR1A.NameElements
    '        With Pax
    '            mobjPassengers.AddItem(.ElementNo, .Initial, .LastName, If(IsNothing(.ID), "", .ID))
    '            'Exit For
    '        End With
    '    Next
    'End Sub
    'Private Sub GetPassengers1G()

    '    mobjPassengers.Clear()
    '    For Each pobjPax As Travelport.TravelData.Person In mobjPNR1G.Passengers
    '        With pobjPax
    '            mobjPassengers.AddItem(.PassengerNumber, .FirstName, .LastName, If(IsNothing(.NameRemark), "", .NameRemark))
    '        End With
    '    Next
    'End Sub
    'Private Sub GetSegments1A()

    '    mobjSegments.Clear()
    '    mdteDepartureDate = Date.MinValue
    '    mstrItinerary = ""
    '    Dim pOff As String = ""

    '    For Each pSeg As s1aPNR.AirFlownSegment In mobjPNR1A.AirFlownSegments
    '        With pSeg
    '            If mstrItinerary = "" Then
    '                mstrItinerary = .BoardPoint & "-" & .OffPoint
    '            Else
    '                If .BoardPoint = pOff Then
    '                    mstrItinerary &= "-" & .OffPoint
    '                Else
    '                    mstrItinerary &= "-***-" & .BoardPoint & "-" & .OffPoint
    '                End If
    '            End If
    '            If .Airline = "QR" Then
    '                mflgQRSegment = True
    '            End If
    '            pOff = .OffPoint
    '            Dim pDate As New s1aAirlineDate.clsAirlineDate
    '            pDate.SetFromString(.DepartureDate)
    '            If mdteDepartureDate = Date.MinValue Then
    '                mdteDepartureDate = pDate.VBDate
    '            End If
    '            mobjSegments.AddItem(airAirline1A(pSeg), airBoardPoint1A(pSeg), airClass1A(pSeg), airDepartureDate1A(pSeg), airArrivalDate1A(pSeg), .ElementNo, airFlightNo1A(pSeg), airOffPoint1A(pSeg), airStatus1A(pSeg), airDepartTime1A(pSeg), airArriveTime1A(pSeg), airText1A(pSeg), "")
    '        End With
    '    Next

    '    For Each pSeg As s1aPNR.AirSegment In mobjPNR1A.AirSegments
    '        With pSeg
    '            If mstrItinerary = "" Then
    '                mstrItinerary = pSeg.BoardPoint & "-" & pSeg.OffPoint
    '            Else
    '                If pSeg.BoardPoint = pOff Then
    '                    mstrItinerary &= "-" & pSeg.OffPoint
    '                Else
    '                    mstrItinerary &= "-***-" & pSeg.BoardPoint & "-" & pSeg.OffPoint
    '                End If
    '            End If
    '            pOff = pSeg.OffPoint
    '            Dim pDate As New s1aAirlineDate.clsAirlineDate
    '            pDate.SetFromString(pSeg.DepartureDate)
    '            If mdteDepartureDate = Date.MinValue Then
    '                mdteDepartureDate = pDate.VBDate
    '            End If
    '            mobjSegments.AddItem(airAirline1A(pSeg), airBoardPoint1A(pSeg), airClass1A(pSeg), airDepartureDate1A(pSeg), airArrivalDate1A(pSeg), .ElementNo, airFlightNo1A(pSeg), airOffPoint1A(pSeg), airStatus1A(pSeg), airDepartTime1A(pSeg), airArriveTime1A(pSeg), airText1A(pSeg), "")
    '        End With
    '    Next
    '    mflgExistsSegments = ((mobjPNR1A.AirFlownSegments.Count + mobjPNR1A.AirSegments.Count) > 0)

    '    If mdteDepartureDate > Date.MinValue Then
    '        Dim pDate As New s1aAirlineDate.clsAirlineDate
    '        pDate.SetFromString(mdteDepartureDate)
    '        mstrItinerary &= " (" & pDate.IATA & ")"
    '    End If

    'End Sub

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
    'Private Sub GetOpenSegment1A()

    '    For Each pSeg As s1aPNR.MemoSegment In mobjPNR1A.MemoSegments
    '        If pSeg.Text.Contains(MySettings.GDSValue("TextMISSegmentLookup") & mobjPNR1A.NameElements.Count & " " & MySettings.OfficeCityCode) Then
    '            mobjExistingGDSElements.OpenSegment.SetValues(True, pSeg.ElementNo, MySettings.GDSElement("TextMISSegmentLookup"), "", "")
    '            Exit For
    '        End If
    '    Next

    'End Sub
    'Private Sub GetOpenSegment1G()

    '    For i = 0 To mobjPNR1G.Segments.Count - 1
    '        If mobjPNR1G.Segments(i).SegmentType = Travelport.TravelData.SegmentType.DueOrPaid Then
    '            Dim pSeg As Travelport.TravelData.DueOrPaidSegment = mobjPNR1G.Segments(i)
    '            mobjExistingGDSElements.OpenSegment.SetValues(True, pSeg.SegmentNumber, "Segment", pSeg.Description.ToString, "")
    '        End If
    '    Next
    'End Sub
    'Private Sub GetPhoneElement1A()

    '    For Each pField As s1aPNR.PhoneElement In mobjPNR1A.PhoneElements
    '        If pField.Text.Replace(" ", "").Contains(MySettings.GDSValue("TextAP").Replace(" ", "")) Then
    '            mobjExistingGDSElements.PhoneElement.SetValues(True, pField.Text.Substring(0, pField.Text.IndexOf(pField.ElementID) - 1), MySettings.GDSElement("TextAP"), "", "")
    '            Exit For
    '        End If
    '    Next

    'End Sub
    'Private Sub GetPhoneElement1G()

    '    For Each pField As Travelport.TravelData.BookingFilePhone In mobjPNR1G.PhoneNumbers
    '        If "P." & pField.CityCode.Code & "T*" & pField.PhoneNumber = MySettings.GDSValue("TextAP") Then
    '            mobjExistingGDSElements.PhoneElement.SetValues(True, pField.Number, MySettings.GDSElement("TextAP"), pField.PhoneNumber, pField.PhoneNumber)
    '        ElseIf "P." & pField.CityCode.Code & "T*" & pField.PhoneNumber = MySettings.GDSValue("TextAOH") Then
    '            mobjExistingGDSElements.AOH.SetValues(True, pField.Number, MySettings.GDSElement("TextAOH"), pField.PhoneNumber, pField.PhoneNumber)
    '        End If
    '    Next
    'End Sub
    'Private Sub GetEmailElement1A()

    '    For Each pField As s1aPNR.PhoneElement In mobjPNR1A.PhoneElements
    '        If pField.Text.Contains(MySettings.GDSValue("TextAPE_ToFind")) Then
    '            mobjExistingGDSElements.EmailElement.SetValues(True, pField.Text.Substring(0, pField.Text.IndexOf(pField.ElementID) - 1), MySettings.GDSElement("TextAPE_ToFind"), "", "")
    '        End If
    '    Next
    'End Sub
    'Private Sub GetEmailElement1G()
    '    For Each pField As Travelport.TravelData.BookingFileEmailAddress In mobjPNR1G.EmailAddresses
    '        If "MT." & pField.Address = MySettings.GDSValue("TextAPE") Then
    '            mobjExistingGDSElements.EmailElement.SetValues(True, pField.Number, MySettings.GDSElement("TextAPE"), pField.Address, pField.Address)
    '        End If
    '    Next
    'End Sub
    'Private Sub GetAOH1A()
    '    For Each pElement As s1aPNR.SSRElement In mobjPNR1A.SSRElements
    '        If pElement.Text.Contains(MySettings.GDSValue("TextAOH_ToFind")) Then
    '            mobjExistingGDSElements.AOH.SetValues(True, pElement.Text.Substring(0, pElement.Text.IndexOf(pElement.ElementID) - 1), MySettings.GDSElement("TextAOH_ToFind"), "", "")
    '        End If
    '    Next
    'End Sub

    'Private Sub GetTicketElement1A()
    '    For Each pElement As s1aPNR.TicketElement In mobjPNR1A.TicketElements
    '        mobjExistingGDSElements.TicketElement.SetValues(True, pElement.Text.Substring(0, pElement.Text.IndexOf(pElement.ElementID) - 1), "TKT", "", "")
    '    Next
    'End Sub
    'Private Sub GetTicketElement1G()
    '    If Not mobjPNR1G.TicketIssue.TicketHaveBeenIssued And mobjPNR1G.TicketIssue.ActionDateTime > Now Then
    '        mobjExistingGDSElements.TicketElement.SetValues(True, 1, "T.", mobjPNR1G.TicketIssue.ActionDateTime, mobjPNR1G.TicketIssue.ActionDateTime)
    '    End If
    'End Sub

    'Private Sub GetOptionQueueElement1A()
    '    For Each pElement As s1aPNR.OptionQueueElement In mobjPNR1A.OptionQueueElements
    '        If pElement.Text.Contains(MySettings.GDSValue("TextOP")) Then
    '            mobjExistingGDSElements.OptionQueueElement.SetValues(True, pElement.Text.Substring(0, pElement.Text.IndexOf(pElement.ElementID) - 1), MySettings.GDSElement("TextOP"), "", "")
    '            Exit For
    '        End If
    '    Next
    'End Sub
    'Private Sub GetOptionQueueElement1G()
    '    For Each pField As Travelport.TravelData.BookingFileReminder In mobjPNR1G.Reminders
    '        Dim pFullText As String = MySettings.GDSValue("TextOP") & "/DDMMM/0001/Q" & MySettings.AgentOPQueue
    '        If pFullText.StartsWith(MySettings.GDSValue("TextOP")) And pFullText.EndsWith("/0001/Q" & MySettings.AgentOPQueue) Then
    '            'MySettings.GDSValue("TextOP") & "/" & pDateReminder.IATA & "/0001/Q" & MySettings.AgentOPQueue
    '            mobjExistingGDSElements.OptionQueueElement.SetValues(True, pField.Number, MySettings.GDSElement("TextOP"), pField.QueueNumber, pField.QueueNumber)
    '        End If
    '    Next
    'End Sub
    'Private Sub GetVesselOSI1A()
    '    For Each pOSI As s1aPNR.OtherServiceElement In mobjPNR1A.OtherServiceElements
    '        If pOSI.Text.Contains(MySettings.GDSValue("TextVOSI")) Then
    '            If mobjExistingGDSElements.VesselOSI.Exists Then
    '                Throw New Exception("Please check PNR. Duplicate OSI Vessel defined" & vbCrLf & mobjExistingGDSElements.VesselOSI.RawText & vbCrLf & pOSI.Text)
    '            Else
    '                Dim pVesselNameOSI As String = pOSI.Text.Substring(pOSI.Text.IndexOf(MySettings.GDSValue("TextVSL")) + MySettings.GDSValue("TextVSL").Length)
    '                mobjExistingGDSElements.VesselOSI.SetValues(True, pOSI.Text.Substring(0, pOSI.Text.IndexOf(pOSI.ElementID) - 1), MySettings.GDSElement("TextVSL"), pOSI.Text, pVesselNameOSI)
    '            End If
    '        End If
    '    Next
    'End Sub
    'Private Sub GetVesselOSI1G()
    '    Dim pobjOtherServiceElement As Travelport.TravelData.BookingFileOtherSupplementaryInformation
    '    For Each pobjOtherServiceElement In mobjPNR1G.OtherSupplementaryInformationRemarks
    '        With pobjOtherServiceElement
    '            '"SEMN/VESSEL-CHRISTOS"
    '            If (("SI." & .Carrier.Code & "*" & .Message).StartsWith(MySettings.GDSValue("TextVOSI"))) Then
    '                Dim pVesselNameOSI As String = ("SI." & .Carrier.Code & "*" & .Message).Substring(MySettings.GDSValue("TextVOSI").Length).Trim
    '                mobjExistingGDSElements.VesselOSI.SetValues(True, pobjOtherServiceElement.Number, MySettings.GDSElement("TextVOSI"), pobjOtherServiceElement.Message, pVesselNameOSI)
    '            End If
    '        End With
    '    Next pobjOtherServiceElement
    'End Sub
    'Private Sub GetSSR1A()
    '    mflgExistsSSRDocs = False
    '    mstrSSRDocs = ""
    '    For Each pSSR As s1aPNR.SSRElement In mobjPNR1A.SSRElements
    '        If pSSR.Text.IndexOf("SSR DOCS") > 0 And pSSR.Text.IndexOf("SSR DOCS") < 10 Then
    '            mstrSSRDocs &= pSSR.Text & vbCrLf
    '            mflgExistsSSRDocs = True
    '        End If
    '    Next
    'End Sub
    'Private Sub GetSSR1G()
    '    mflgExistsSSRDocs = False
    '    For Each pElement As Travelport.TravelData.BookingFileManualSsr In mobjPNR1G.ManualSsrs
    '        If pElement.Code = "DOCS" Then
    '            mstrSSRDocs &= "SI.SSR" & pElement.Code & pElement.VendorCode & pElement.StatusCode & "1" & pElement.Text & vbCrLf
    '            mflgExistsSSRDocs = True
    '        End If
    '    Next

    'End Sub
    'Private Sub GetRM1A()
    '    For Each pRemark As s1aPNR.RemarkElement In mobjPNR1A.RemarkElements
    '        If pRemark.Text.Contains(MySettings.GDSValue("TextAGT")) Then
    '            mobjExistingGDSElements.AgentID.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextAGT"), pRemark.Text, "")
    '        ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextCLN")) Then
    '            If mobjExistingGDSElements.CustomerCode.Exists Then
    '                Throw New Exception("Please check PNR. Duplicate customer defined" & vbCrLf & mobjExistingGDSElements.CustomerCode.RawText & vbCrLf & pRemark.Text)
    '            Else
    '                Dim pCustomerCode As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextCLN")) + MySettings.GDSValue("TextCLN").Length)
    '                mobjExistingGDSElements.CustomerCode.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextCLN"), pRemark.Text, pCustomerCode)
    '            End If
    '        ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextSBN")) Then
    '            If mobjExistingGDSElements.SubDepartmentCode.Exists Then
    '                Throw New Exception("Please check PNR. Duplicate subdepartment defined" & vbCrLf & mobjExistingGDSElements.SubDepartmentCode.LineNumber & vbCrLf & pRemark.Text)
    '            Else
    '                Dim pSubDepartment As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextSBN")) + MySettings.GDSValue("TextSBN").Length)
    '                mobjExistingGDSElements.SubDepartmentCode.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextSBN"), "", pSubDepartment)
    '            End If
    '        ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextCRN")) Then
    '            If mobjExistingGDSElements.CRMCode.Exists Then
    '                Throw New Exception("Please check PNR. Duplicate CRM defined" & vbCrLf & mobjExistingGDSElements.CRMCode.LineNumber & vbCrLf & pRemark.Text)
    '            Else
    '                Dim pCRM As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextCRN")) + MySettings.GDSValue("TextCRN").Length)
    '                mobjExistingGDSElements.CRMCode.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextCRN"), "", pCRM)
    '            End If
    '        ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextVSL")) Then
    '            If mobjExistingGDSElements.VesselName.Exists Then
    '                Throw New Exception("Please check PNR. Duplicate vessel defined" & vbCrLf & mobjExistingGDSElements.VesselName.LineNumber & vbCrLf & pRemark.Text)
    '            Else
    '                Dim pVesselName As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextVSL")) + MySettings.GDSValue("TextVSL").Length)
    '                mobjExistingGDSElements.VesselName.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextVSL"), "", pVesselName)
    '            End If
    '        ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextVSR")) Then
    '            If mobjExistingGDSElements.VesselFlag.Exists Then
    '                Throw New Exception("Please check PNR. Duplicate vessel registration defined" & vbCrLf & mobjExistingGDSElements.VesselFlag.LineNumber & vbCrLf & pRemark.Text)
    '            Else
    '                Dim pVesselRegistration As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextVSR")) + MySettings.GDSValue("TextVSR").Length)
    '                mobjExistingGDSElements.VesselFlag.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextVSR"), "", pVesselRegistration)
    '            End If
    '        ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextREF")) Then
    '            Dim pReference As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextREF")) + MySettings.GDSValue("TextREF").Length)
    '            mobjExistingGDSElements.Reference.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextREF"), "", pReference)
    '        ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextBBY")) Then
    '            Dim pBookedBy As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextBBY")) + MySettings.GDSValue("TextBBY").Length)
    '            mobjExistingGDSElements.BookedBy.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextBBY"), "", pBookedBy)
    '        ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextDPT")) Then
    '            Dim pDepartment As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextDPT")) + MySettings.GDSValue("TextDPT").Length)
    '            mobjExistingGDSElements.Department.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextDPT"), True, pDepartment)
    '        ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextRFT")) Then
    '            Dim pReasonForTravel As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextRFT")) + MySettings.GDSValue("TextRFT").Length)
    '            mobjExistingGDSElements.ReasonForTravel.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextRFT"), "", pReasonForTravel)
    '        ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextCC")) Then
    '            Dim pCostCentre As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.GDSValue("TextCC")) + MySettings.GDSValue("TextCC").Length)
    '            mobjExistingGDSElements.CostCentre.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextCC"), "", pCostCentre)
    '        ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextCLA")) Then
    '            If mobjExistingGDSElements.CustomerName.Exists Then
    '                Throw New Exception("Please check PNR. Duplicate customer name defined" & vbCrLf & mobjExistingGDSElements.CustomerName.LineNumber & vbCrLf & pRemark.Text)
    '            Else
    '                mobjExistingGDSElements.CustomerName.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextCLA"), "", "")
    '            End If
    '        ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextSBA")) Then
    '            mobjExistingGDSElements.SubDepartmentName.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextSBA"), "", "")
    '        ElseIf pRemark.Text.Contains(MySettings.GDSValue("TextCRA")) Then
    '            mobjExistingGDSElements.CRMName.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), MySettings.GDSElement("TextCRA"), "", "")
    '        End If
    '    Next
    'End Sub
    'Private Sub GetRM1G()
    '    For Each pRemark As Travelport.TravelData.BookingFileRemark In mobjPNR1G.InvoiceRemarks
    '        With pRemark
    '            Dim pFullText As String = "DI." & pRemark.Category & "-" & pRemark.Text
    '            If pFullText.StartsWith(MySettings.GDSValue("TextAGT")) Then
    '                mobjExistingGDSElements.AgentID.SetValues(True, .Number, MySettings.GDSElement("TextAGT"), .Text, pFullText.Substring(MySettings.GDSValue("TextAGT").Length))
    '            ElseIf pFullText.StartsWith(MySettings.GDSValue("TextBBY")) Then
    '                Dim pBookedBy As String = pFullText.Substring(MySettings.GDSValue("TextBBY").Length)
    '                mobjExistingGDSElements.BookedBy.SetValues(True, .Number, MySettings.GDSElement("TextBBY"), .Text, pBookedBy)
    '            ElseIf pFullText.StartsWith(MySettings.GDSValue("TextCC")) Then
    '                mobjExistingGDSElements.CostCentre.SetValues(True, .Number, MySettings.GDSElement("TextCC"), .Text, pFullText.Substring(MySettings.GDSValue("TextCC").Length))
    '            ElseIf pFullText.StartsWith(MySettings.GDSValue("TextCLA")) Then
    '                If mobjExistingGDSElements.CustomerName.Exists Then
    '                    Throw New Exception("Please check PNR. Duplicate customer name defined" & vbCrLf & mobjExistingGDSElements.CustomerName.LineNumber & vbCrLf & .Text)
    '                Else
    '                    mobjExistingGDSElements.CustomerName.SetValues(True, .Number, MySettings.GDSElement("TextCLA"), .Text, pFullText.Substring(MySettings.GDSValue("TextCLA").Length))
    '                End If
    '            ElseIf pFullText.StartsWith(MySettings.GDSValue("TextCLN")) Then
    '                Dim pCustomerCode As String = pFullText.Substring(MySettings.GDSValue("TextCLN").Length)
    '                If mobjExistingGDSElements.CustomerCode.Exists Then
    '                    Throw New Exception("Please check PNR. Duplicate customer defined" & vbCrLf & mobjExistingGDSElements.CustomerCode.RawText & vbCrLf & .Text)
    '                Else
    '                    mobjExistingGDSElements.CustomerCode.SetValues(True, .Number, MySettings.GDSElement("TextCLN"), .Text, pCustomerCode)
    '                End If
    '            ElseIf pFullText.StartsWith(MySettings.GDSValue("TextCRA")) Then
    '                mobjExistingGDSElements.CRMName.SetValues(True, .Number, MySettings.GDSElement("TextCRA"), .Text, pFullText.Substring(MySettings.GDSValue("TextCRA").Length))
    '            ElseIf pFullText.StartsWith(MySettings.GDSValue("TextCRN")) Then
    '                mobjExistingGDSElements.CRMCode.SetValues(True, .Number, MySettings.GDSElement("TextCRN"), .Text, pFullText.Substring(MySettings.GDSValue("TextCRN").Length))
    '            ElseIf pFullText.StartsWith(MySettings.GDSValue("TextDPT")) Then
    '                mobjExistingGDSElements.Department.SetValues(True, .Number, MySettings.GDSElement("TextDPT"), .Text, pFullText.Substring(MySettings.GDSValue("TextDPT").Length))
    '            ElseIf pFullText.StartsWith(MySettings.GDSValue("TextREF")) Then
    '                mobjExistingGDSElements.Reference.SetValues(True, .Number, MySettings.GDSElement("TextREF"), .Text, pFullText.Substring(MySettings.GDSValue("TextREF").Length))
    '            ElseIf pFullText.StartsWith(MySettings.GDSValue("TextRFT")) Then
    '                mobjExistingGDSElements.ReasonForTravel.SetValues(True, .Number, MySettings.GDSElement("TextRFT"), .Text, pFullText.Substring(MySettings.GDSValue("TextRFT").Length))
    '            ElseIf pFullText.StartsWith(MySettings.GDSValue("TextSBA")) Then
    '                mobjExistingGDSElements.SubDepartmentName.SetValues(True, .Number, MySettings.GDSElement("TextSBA"), .Text, pFullText.Substring(MySettings.GDSValue("TextSBA").Length))
    '            ElseIf pFullText.StartsWith(MySettings.GDSValue("TextSBN")) Then
    '                mobjExistingGDSElements.SubDepartmentCode.SetValues(True, .Number, MySettings.GDSElement("TextSBN"), .Text, pFullText.Substring(MySettings.GDSValue("TextSBN").Length))
    '            ElseIf pFullText.StartsWith(MySettings.GDSValue("TextVSL")) Then
    '                mobjExistingGDSElements.VesselName.SetValues(True, .Number, MySettings.GDSElement("TextVSL"), .Text, pFullText.Substring(MySettings.GDSValue("TextVSL").Length))
    '            ElseIf pFullText.StartsWith(MySettings.GDSValue("TextVSR")) Then
    '                mobjExistingGDSElements.VesselFlag.SetValues(True, .Number, MySettings.GDSElement("TextVSR"), .Text, pFullText.Substring(MySettings.GDSValue("TextVSR").Length))
    '            ElseIf pFullText.StartsWith("D,BOOKED") > 0 Then
    '                mobjExistingGDSElements.BookedBy.SetValues(True, .Number, "D,BOOKED", .Text, "DI.")
    '            ElseIf pFullText.StartsWith("D,AC") > 0 Then
    '                Dim pCustomerCode As String = .Text.Substring(10)
    '                If mobjExistingGDSElements.CustomerCode.Exists Then
    '                    Throw New Exception("Please check PNR. Duplicate customer defined" & vbCrLf & mobjExistingGDSElements.CustomerCode.RawText & vbCrLf & .Text)
    '                Else
    '                    mobjExistingGDSElements.CustomerCode.SetValues(True, .Number, "D,AC", .Text, "DI.")
    '                End If
    '            End If
    '        End With
    '    Next
    'End Sub
    'Private Sub GetTickets1A()
    '    mobjTicketElements = New GDSTickets.Collection(mobjPNR1A)
    'End Sub

End Class
