Option Strict Off
Option Explicit On
Public Class ReadPNR

    Private WithEvents mobjSession As k1aHostToolKit.HostSession

    Private mobjPNR As s1aPNR.PNR
    Private mstrPNRResponse As String
    Private mstrPNRNumber As String
    Private mflgNewPNR As Boolean
    Private mstrPaxNames As String
    Private mstrGroupName As String
    Private mintGroupNamesCount As Integer
    Private mstrItinerary As String

    Private mstrOfficeOfResponsibility As String
    Private mdteCreationDate As Date
    Private mflgQRSegment As Boolean = False
    Private mdteDepartureDate As Date
    Private mflgExistsSegments As Boolean
    Private mflgExistsSSRDocs As Boolean
    Private mobjExistingAmadeusElements As New AmadeusExisting.Collection
    Private mobjNewAmadeusElements As AmadeusNew.Collection
    Private mobjTicketElements As AmadeusTickets.Collection

    Public ReadOnly Property PNR As s1aPNR.PNR
        Get
            PNR = mobjPNR
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
    Public ReadOnly Property CreationDate As Date
        Get
            CreationDate = mdteCreationDate
        End Get
    End Property
    Public ReadOnly Property NumberOfPax As Integer
        Get
            NumberOfPax = mobjPNR.NameElements.Count
        End Get
    End Property
    Public ReadOnly Property PaxName As String
        Get
            PaxName = mstrPaxNames
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
    Public ReadOnly Property IsGroup As Boolean
        Get
            IsGroup = (mstrGroupName <> "")
        End Get
    End Property
    Public ReadOnly Property DepartureDate As Date
        Get
            DepartureDate = mdteDepartureDate
        End Get
    End Property
    Public ReadOnly Property Itinerary As String
        Get
            Itinerary = mstrItinerary
        End Get
    End Property
    Public ReadOnly Property ExistingElements As AmadeusExisting.Collection
        Get
            ExistingElements = mobjExistingAmadeusElements
        End Get
    End Property
    Public ReadOnly Property NewElements As AmadeusNew.Collection
        Get
            NewElements = mobjNewAmadeusElements
        End Get
    End Property
    Public ReadOnly Property AirSegments As Object
        Get
            AirSegments = mobjPNR.AirSegments
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
    Public Function Read() As String

        Dim Sessions As k1aHostToolKit.HostSessions

        Read = ""

        Try
            ' To be able to retrieve the PNR that have been created we need to link our '
            ' application to the current session of the FOS                             '
            Sessions = New k1aHostToolKit.HostSessions

            If Sessions.Count > 0 Then
                ' There is at least one session opened.                    '
                ' We link our application to the active session of the FOS '
                mobjSession = Sessions.UIActiveSession

                ' Initialize the PNR
                mobjPNR = New s1aPNR.PNR

                ' Retrieve the name elements, Air segments and Hotel Segments of the current PNR
                Dim pStatus As Integer = mobjPNR.RetrievePNR(mobjSession, "RT")
                mflgNewPNR = False

                If pStatus = 0 Or pStatus = 1005 Then
                    GetOfficeOfResponsibility()
                    GetPnrNumber()
                    GetCreationDate()

                    GetGroup()
                    GetPassengers()
                    GetSegments()
                    GetPhoneElement()
                    GetEmailElement()
                    GetAOH()
                    GetOpenSegment()
                    GetTicketElement()
                    GetOptionQueueElement()
                    GetVesselOSI()
                    GetSSR()
                    GetRM()

                    GetTickets()
                    If mobjPNR.RawResponse.IndexOf("***  NHP  ***") >= 0 Then
                        Read = "               ***  NHP  ***"
                    Else
                        Read = CheckDMI()
                    End If
                Else
                    Throw New Exception("There is no active PNR" & vbCrLf & mstrPNRResponse)
                End If
            Else
                Throw New Exception("Please start Amadeus and retry")
            End If
        Catch ex As Exception
            Throw New Exception("Read()" & vbCrLf & ex.Message)
        End Try

    End Function
    Public Sub PrepareNewAmadeusElements()
        mobjNewAmadeusElements = New AmadeusNew.Collection(OfficeOfResponsibility, CreationDate, DepartureDate, NumberOfPax)
    End Sub
    Private Function CheckDMI() As String

        Try
            If mobjPNR.AirSegments.Count <= 1 Then
                Return ""
            End If

            Dim pDMI As String = mobjSession.Send("DMI").Text

            If pDMI.Contains("ITINERARY OK") Then
                Return ""
            Else
                Return pDMI
            End If
        Catch ex As Exception
            Return ""
        End Try


    End Function
    Private Sub RemoveOldAmadeusEntries()

        Dim pLineNumbers(0) As Integer

        ' the following elements remain as they are if they already exist in the PNR
        ClearExistingItems(mobjExistingAmadeusElements.PhoneElement, mobjNewAmadeusElements.PhoneElement)
        ClearExistingItems(mobjExistingAmadeusElements.EmailElement, mobjNewAmadeusElements.EmailElement)
        ClearExistingItems(mobjExistingAmadeusElements.AOH, mobjNewAmadeusElements.AOH)
        ClearExistingItems(mobjExistingAmadeusElements.OpenSegment, mobjNewAmadeusElements.OpenSegment)
        ClearExistingItems(mobjExistingAmadeusElements.OptionQueueElement, mobjNewAmadeusElements.OptionQueueElement)
        ClearExistingItems(mobjExistingAmadeusElements.TicketElement, mobjNewAmadeusElements.TicketElement)
        ClearExistingItems(mobjExistingAmadeusElements.AgentID, mobjNewAmadeusElements.AgentID)

        ' the following elements are removed and replaced if they exist in the PNR
        PrepareLineNumbers(mobjExistingAmadeusElements.CustomerCode, pLineNumbers)
        PrepareLineNumbers(mobjExistingAmadeusElements.CustomerName, pLineNumbers)
        PrepareLineNumbers(mobjExistingAmadeusElements.SubDepartmentCode, pLineNumbers)
        PrepareLineNumbers(mobjExistingAmadeusElements.SubDepartmentName, pLineNumbers)
        PrepareLineNumbers(mobjExistingAmadeusElements.CRMCode, pLineNumbers)
        PrepareLineNumbers(mobjExistingAmadeusElements.CRMName, pLineNumbers)
        PrepareLineNumbers(mobjExistingAmadeusElements.VesselFlag, pLineNumbers)
        PrepareLineNumbers(mobjExistingAmadeusElements.VesselName, pLineNumbers)
        PrepareLineNumbers(mobjExistingAmadeusElements.VesselOSI, pLineNumbers)
        PrepareLineNumbers(mobjExistingAmadeusElements.Reference, pLineNumbers)
        PrepareLineNumbers(mobjExistingAmadeusElements.BookedBy, pLineNumbers)
        PrepareLineNumbers(mobjExistingAmadeusElements.Department, pLineNumbers)
        PrepareLineNumbers(mobjExistingAmadeusElements.ReasonForTravel, pLineNumbers)
        PrepareLineNumbers(mobjExistingAmadeusElements.CostCentre, pLineNumbers)

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
                mobjSession.Send("XE" & pMax)
                pLineNumbers(pMaxIndex) = 0
            End If
            pMax = 0
            pMaxIndex = -1
        Loop

    End Sub
    Private Sub ClearExistingItems(ByRef ExistingItem As AmadeusExisting.Item, ByRef NewItem As AmadeusNew.Item)
        If ExistingItem.Exists Then
            NewItem.Clear()
        End If
    End Sub

    Private Sub PrepareLineNumbers(ByVal ExistingItem As AmadeusExisting.Item, ByRef pLineNumbers() As Integer)
        If ExistingItem.Exists Then
            ReDim Preserve pLineNumbers(pLineNumbers.GetUpperBound(0) + 1)
            pLineNumbers(pLineNumbers.GetUpperBound(0)) = ExistingItem.LineNumber
        End If
    End Sub
    Public Sub SendNewAmadeusEntries(ByVal AmadeusEntry As String)

        If AmadeusEntry <> "" Then
            mobjSession.Send(AmadeusEntry)
        End If

    End Sub
    Public Sub SendNewAmadeusEntries(ByVal WritePNR As Boolean, ByVal WriteDocs As Boolean, ByVal mflgExpiryDateOK As Boolean, dgvApis As DataGridView, AirlineEntries As TextBox)

        Try
            If WritePNR Then
                RemoveOldAmadeusEntries()

                SendAmadeusElement(mobjNewAmadeusElements.PhoneElement)
                SendAmadeusElement(mobjNewAmadeusElements.EmailElement)
                SendAmadeusElement(mobjNewAmadeusElements.AgentID)
                SendAmadeusElement(mobjNewAmadeusElements.AOH)
                SendAmadeusElement(mobjNewAmadeusElements.OpenSegment)
                SendAmadeusElement(mobjNewAmadeusElements.TicketElement)
                SendAmadeusElement(mobjNewAmadeusElements.OptionQueueElement)

                If mflgNewPNR Then
                    SendAmadeusElement(mobjNewAmadeusElements.SavingsElement)
                    SendAmadeusElement(mobjNewAmadeusElements.LossElement)
                End If

                SendAmadeusElement(mobjNewAmadeusElements.CustomerCode)
                SendAmadeusElement(mobjNewAmadeusElements.CustomerName)
                SendAmadeusElement(mobjNewAmadeusElements.SubDepartmentCode)
                SendAmadeusElement(mobjNewAmadeusElements.SubDepartmentName)
                SendAmadeusElement(mobjNewAmadeusElements.CRMCode)
                SendAmadeusElement(mobjNewAmadeusElements.CRMName)
                SendAmadeusElement(mobjNewAmadeusElements.VesselName)
                SendAmadeusElement(mobjNewAmadeusElements.VesselFlag)
                SendAmadeusElement(mobjNewAmadeusElements.VesselOSI)
                SendAmadeusElement(mobjNewAmadeusElements.Reference)
                SendAmadeusElement(mobjNewAmadeusElements.BookedBy)
                SendAmadeusElement(mobjNewAmadeusElements.Department)
                SendAmadeusElement(mobjNewAmadeusElements.ReasonForTravel)
                SendAmadeusElement(mobjNewAmadeusElements.CostCentre)

                Dim pAirlineEntries() As String = AirlineEntries.Text.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

                For i As Integer = 0 To pAirlineEntries.GetUpperBound(0)
                    pAirlineEntries(i) = pAirlineEntries(i).Replace(">", "").Trim
                    If pAirlineEntries(i).Trim <> "" Then
                        SendAmadeusAirlineItems(pAirlineEntries(i).Replace("> ", ""))
                    End If
                Next
            End If

            If WriteDocs Then
                APISUpdate(mflgExpiryDateOK, dgvApis)
            End If

            If WritePNR Or WriteDocs Then
                CloseOffPNR()
            End If
        Catch ex As Exception
            Throw New Exception("SendNewAmadeusEntries()" & vbCrLf & ex.Message)
        End Try


    End Sub

    Private Sub APISUpdate(ByVal mflgExpiryDateOK As Boolean, dgvApis As DataGridView)

        Dim pstrCommand As String
        Try
            For i = 0 To dgvApis.RowCount - 1
                With dgvApis.Rows(i)
                    If .ErrorText = "" Then
                        Dim pobjItem As New PaxApisDB.Item(.Cells(0).Value, .Cells(1).Value, .Cells(2).Value,
                                                       APISDateFromIATA(.Cells(6).Value), .Cells(7).Value, .Cells(3).Value,
                                                     .Cells(4).Value, APISDateFromIATA(.Cells(8).Value), .Cells(5).Value)

                        pobjItem.Update(mflgExpiryDateOK)
                        pstrCommand = "SR DOCS YY HK1-P-" & pobjItem.IssuingCountry & "-" & pobjItem.PassportNumber & "-" & pobjItem.Nationality &
                    "-" & APISDateToIATA(pobjItem.BirthDate) & "-" & pobjItem.Gender & "-"
                        If mflgExpiryDateOK Then
                            pstrCommand &= APISDateToIATA(pobjItem.ExpiryDate)
                        Else
                            pstrCommand &= ""
                        End If
                        pstrCommand &= "-" & pobjItem.Surname & "-" & pobjItem.FirstName & "/P" & pobjItem.Id
                        SendNewAmadeusEntries(pstrCommand)
                    End If

                End With

            Next
        Catch ex As Exception
            Throw New Exception("APISUpdate()" & vbCrLf & ex.Message)
        End Try


    End Sub

    Private Sub CloseOffPNR()

        Dim pCloseOffEntries As New CloseOffEntries.Collection

        pCloseOffEntries.Load(MySettings.AmadeusPCC, mstrOfficeOfResponsibility = MySettings.AmadeusPCC)

        For Each pCommand As CloseOffEntries.Item In pCloseOffEntries.Values
            mobjSession.Send(pCommand.CloseOffEntry)
        Next
        If mstrPNRResponse.Contains("WARNING: SECURE FLT PASSENGER DATA REQUIRED") Then
            MessageBox.Show(mstrPNRResponse)
        End If

    End Sub

    Private Sub SendAmadeusElement(ByVal pElement As AmadeusNew.Item)

        If pElement.AmadeusCommand <> "" Then
            mobjSession.Send(pElement.AmadeusCommand)
        End If

    End Sub
    Private Sub SendAmadeusAirlineItems(ByVal pItemToSend As String)

        If pItemToSend.StartsWith("OS ") Then
            If mobjPNR.RawResponse.Replace(vbCrLf, "").Replace(" ", "").IndexOf(("OSI " & pItemToSend.Substring(3)).Replace(" ", "")) = -1 Then
                mobjSession.Send(pItemToSend)
            End If
        ElseIf pItemToSend.StartsWith("R") Then
            If mobjPNR.RawResponse.Replace(vbCrLf, "").Replace(" ", "").IndexOf(pItemToSend.Replace(" ", "")) = -1 Then
                mobjSession.Send(pItemToSend)
            End If
        ElseIf pItemToSend.StartsWith("S") Then
            Dim pString As String
            pString = pItemToSend.Replace(" ", "").Replace("SRCKIN-", "")
            If mobjPNR.RawResponse.Replace(vbCrLf, "").Replace(" ", "").IndexOf(pString) = -1 Then
                mobjSession.Send(pItemToSend)
            End If
        Else
            mobjSession.Send(pItemToSend)
        End If

    End Sub

    Public Function CopyAmadeusEntries(AirlineNotes As CheckedListBox, AirlinePoints As CheckedListBox) As String

        Dim pEntries As String = ""

        pEntries &= mobjNewAmadeusElements.PhoneElement.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.EmailElement.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.AgentID.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.AOH.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.TicketElement.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.OptionQueueElement.AmadeusCommand & vbCrLf

        If mflgNewPNR Then
            pEntries &= mobjNewAmadeusElements.SavingsElement.AmadeusCommand & vbCrLf
            pEntries &= mobjNewAmadeusElements.LossElement.AmadeusCommand & vbCrLf
        End If

        pEntries &= mobjNewAmadeusElements.CustomerCode.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.CustomerName.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.SubDepartmentCode.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.SubDepartmentName.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.CRMCode.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.CRMName.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.VesselName.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.VesselFlag.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.VesselOSI.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.Reference.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.BookedBy.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.Department.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.ReasonForTravel.AmadeusCommand & vbCrLf
        pEntries &= mobjNewAmadeusElements.CostCentre.AmadeusCommand & vbCrLf

        For i As Integer = 0 To AirlinePoints.CheckedItems.Count - 1
            pEntries &= AirlinePoints.Items(i).ToString & vbCrLf
        Next
        For i As Integer = 0 To AirlineNotes.CheckedItems.Count - 1
            pEntries &= AirlineNotes.Items(i).ToString & vbCrLf
        Next

        pEntries &= mobjNewAmadeusElements.OpenSegment.AmadeusCommand & vbCrLf

        While pEntries.IndexOf(vbCrLf & vbCrLf) >= 0
            pEntries = pEntries.Replace(vbCrLf & vbCrLf, vbCrLf)
        End While


        Return pEntries

    End Function

    Private Sub GetPnrNumber()

        Try
            mstrPNRNumber = mobjPNR.Header.RecordLocator
        Catch ex As Exception
            mstrPNRNumber = ""
        End Try

        If mstrPNRNumber = "" Then
            mstrPNRNumber = "New PNR"
            mflgNewPNR = True
        End If
    End Sub
    Public ReadOnly Property NewPNR As Boolean
        Get
            NewPNR = mflgNewPNR
        End Get
    End Property
    Private Sub GetOfficeOfResponsibility()

        Try
            mstrOfficeOfResponsibility = mobjPNR.Header.OfficeOfResponsability
        Catch ex As Exception
            mstrOfficeOfResponsibility = MySettings.AmadeusPCC
        End Try

    End Sub

    Private Sub GetCreationDate()

        Try
            mdteCreationDate = Today
            '            mdteCreationDate = mobjPNR.Header.CreationDate
        Catch ex As Exception
            mdteCreationDate = Today
        End Try

        If mdteCreationDate < DateSerial(2000, 1, 1) Then
            mdteCreationDate = Today
        End If

    End Sub
    Private Sub GetGroup()

        mstrGroupName = ""
        mintGroupNamesCount = 0

        For Each pGroup As s1aPNR.GroupNameElement In mobjPNR.GroupNameElements
            mstrGroupName = pGroup.GroupName
            mintGroupNamesCount = pGroup.NbrOfAssignedNames + pGroup.NbrNamesMissing
            Exit For
        Next
        If mobjPNR.GroupNameElements.Count > 1 Then
            mstrGroupName &= "x" & mobjPNR.GroupNameElements.Count
        End If

    End Sub
    Private Sub GetPassengers()

        mstrPaxNames = ""

        For Each Pax As s1aPNR.NameElement In mobjPNR.NameElements
            mstrPaxNames = Pax.LastName & "/" & Pax.Initial
            Exit For
        Next
        If mobjPNR.NameElements.Count > 1 Then
            mstrPaxNames &= " x " & mobjPNR.NameElements.Count
        End If
        'mflgExistsPassengers = (mobjPNR.NameElements.Count > 0)

    End Sub

    Private Sub GetSegments()

        mdteDepartureDate = Date.MinValue
        mstrItinerary = ""
        Dim pOff As String = ""

        For Each pSeg As s1aPNR.AirFlownSegment In mobjPNR.AirFlownSegments
            If mstrItinerary = "" Then
                mstrItinerary = pSeg.BoardPoint & "-" & pSeg.OffPoint
            Else
                If pSeg.BoardPoint = pOff Then
                    mstrItinerary &= "-" & pSeg.OffPoint
                Else
                    mstrItinerary &= "-***-" & pSeg.BoardPoint & "-" & pSeg.OffPoint
                End If
            End If
            If pSeg.Airline = "QR" Then
                mflgQRSegment = True
            End If
            pOff = pSeg.OffPoint
            Dim pDate As New s1aAirlineDate.clsAirlineDate
            pDate.SetFromString(pSeg.DepartureDate)
            If mdteDepartureDate = Date.MinValue Then
                mdteDepartureDate = pDate.VBDate
            End If
        Next

        For Each pSeg As s1aPNR.AirSegment In mobjPNR.AirSegments
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
        Next
        mflgExistsSegments = ((mobjPNR.AirFlownSegments.Count + mobjPNR.AirSegments.Count) > 0)

        If mdteDepartureDate > Date.MinValue Then
            Dim pDate As New s1aAirlineDate.clsAirlineDate
            pDate.SetFromString(mdteDepartureDate)
            mstrItinerary &= " (" & pDate.IATA & ")"
        End If

    End Sub

    Private Sub GetOpenSegment()

        For Each pSeg As s1aPNR.MemoSegment In mobjPNR.MemoSegments
            If pSeg.Text.Contains(MySettings.AmadeusValue("TextMISSegmentLookup") & mobjPNR.NameElements.Count & " " & MySettings.OfficeCityCode) Then
                mobjExistingAmadeusElements.OpenSegment.SetValues(True, pSeg.ElementNo, "", "")
                Exit For
            End If
        Next

    End Sub

    Private Sub GetPhoneElement()

        For Each pField As s1aPNR.PhoneElement In mobjPNR.PhoneElements
            If pField.Text.Replace(" ", "").Contains(MySettings.AmadeusValue("TextAP").Replace(" ", "")) Then
                mobjExistingAmadeusElements.PhoneElement.SetValues(True, pField.Text.Substring(0, pField.Text.IndexOf(pField.ElementID) - 1), "", "")
                Exit For
            End If
        Next

    End Sub
    Private Sub GetEmailElement()

        For Each pField As s1aPNR.PhoneElement In mobjPNR.PhoneElements
            If pField.Text.Contains(MySettings.AmadeusValue("TextAPE_ToFind")) Then
                mobjExistingAmadeusElements.EmailElement.SetValues(True, pField.Text.Substring(0, pField.Text.IndexOf(pField.ElementID) - 1), "", "")
            End If
        Next
    End Sub

    Private Sub GetAOH()

        For Each pElement As s1aPNR.SSRElement In mobjPNR.SSRElements
            If pElement.Text.Contains(MySettings.AmadeusValue("TextAOH_ToFind")) Then
                mobjExistingAmadeusElements.AOH.SetValues(True, pElement.Text.Substring(0, pElement.Text.IndexOf(pElement.ElementID) - 1), "", "")
            End If
        Next

    End Sub

    Private Sub GetTicketElement()

        For Each pElement As s1aPNR.TicketElement In mobjPNR.TicketElements
            mobjExistingAmadeusElements.TicketElement.SetValues(True, pElement.Text.Substring(0, pElement.Text.IndexOf(pElement.ElementID) - 1), "", "")
        Next

    End Sub

    Private Sub GetOptionQueueElement()

        For Each pElement As s1aPNR.OptionQueueElement In mobjPNR.OptionQueueElements
            If pElement.Text.Contains(MySettings.AmadeusValue("TextOP")) Then
                mobjExistingAmadeusElements.OptionQueueElement.SetValues(True, pElement.Text.Substring(0, pElement.Text.IndexOf(pElement.ElementID) - 1), "", "")
                Exit For
            End If
        Next

    End Sub

    Private Sub GetVesselOSI()

        For Each pOSI As s1aPNR.OtherServiceElement In mobjPNR.OtherServiceElements
            If pOSI.Text.Contains(MySettings.AmadeusValue("TextVOSI")) Then
                If mobjExistingAmadeusElements.VesselOSI.Exists Then
                    Throw New Exception("Please check PNR. Duplicate OSI Vessel defined" & vbCrLf & mobjExistingAmadeusElements.VesselOSI.RawText & vbCrLf & pOSI.Text)
                Else
                    Dim pVesselNameOSI As String = pOSI.Text.Substring(pOSI.Text.IndexOf(MySettings.AmadeusValue("TextVSL")) + MySettings.AmadeusValue("TextVSL").Length)
                    mobjExistingAmadeusElements.VesselOSI.SetValues(True, pOSI.Text.Substring(0, pOSI.Text.IndexOf(pOSI.ElementID) - 1), pOSI.Text, pVesselNameOSI)
                End If
            End If
        Next

    End Sub
    Private Sub GetSSR()

        mflgExistsSSRDocs = False
        For Each pSSR As s1aPNR.SSRElement In mobjPNR.SSRElements
            If pSSR.Text.Contains("SSR DOCS") Then
                mflgExistsSSRDocs = True
            End If
        Next

    End Sub
    Private Sub GetRM()

        For Each pRemark As s1aPNR.RemarkElement In mobjPNR.RemarkElements
            If pRemark.Text.Contains(MySettings.AmadeusValue("TextAGT")) Then
                mobjExistingAmadeusElements.AgentID.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), pRemark.Text, "")
            End If
            If pRemark.Text.Contains(MySettings.AmadeusValue("TextCLN")) Then
                If mobjExistingAmadeusElements.CustomerCode.Exists Then
                    Throw New Exception("Please check PNR. Duplicate customer defined" & vbCrLf & mobjExistingAmadeusElements.CustomerCode.RawText & vbCrLf & pRemark.Text)
                Else
                    Dim pCustomerCode As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.AmadeusValue("TextCLN")) + MySettings.AmadeusValue("TextCLN").Length)
                    mobjExistingAmadeusElements.CustomerCode.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), pRemark.Text, pCustomerCode)
                End If
            ElseIf pRemark.Text.Contains(MySettings.AmadeusValue("TextSBN")) Then
                If mobjExistingAmadeusElements.SubDepartmentCode.Exists Then
                    Throw New Exception("Please check PNR. Duplicate subdepartment defined" & vbCrLf & mobjExistingAmadeusElements.SubDepartmentCode.LineNumber & vbCrLf & pRemark.Text)
                Else
                    Dim pSubDepartment As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.AmadeusValue("TextSBN")) + MySettings.AmadeusValue("TextSBN").Length)
                    mobjExistingAmadeusElements.SubDepartmentCode.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), "", pSubDepartment)
                End If

            ElseIf pRemark.Text.Contains(MySettings.AmadeusValue("TextCRN")) Then
                If mobjExistingAmadeusElements.CRMCode.Exists Then
                    Throw New Exception("Please check PNR. Duplicate CRM defined" & vbCrLf & mobjExistingAmadeusElements.CRMCode.LineNumber & vbCrLf & pRemark.Text)
                Else
                    Dim pCRM As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.AmadeusValue("TextCRN")) + MySettings.AmadeusValue("TextCRN").Length)
                    mobjExistingAmadeusElements.CRMCode.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), "", pCRM)
                End If
            ElseIf pRemark.Text.Contains(MySettings.AmadeusValue("TextVSL")) Then
                If mobjExistingAmadeusElements.VesselName.Exists Then
                    Throw New Exception("Please check PNR. Duplicate vessel defined" & vbCrLf & mobjExistingAmadeusElements.VesselName.LineNumber & vbCrLf & pRemark.Text)
                Else
                    Dim pVesselName As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.AmadeusValue("TextVSL")) + MySettings.AmadeusValue("TextVSL").Length)
                    mobjExistingAmadeusElements.VesselName.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), "", pVesselName)
                End If
            ElseIf pRemark.Text.Contains(MySettings.AmadeusValue("TextVSR")) Then
                If mobjExistingAmadeusElements.VesselFlag.Exists Then
                    Throw New Exception("Please check PNR. Duplicate vessel registration defined" & vbCrLf & mobjExistingAmadeusElements.VesselFlag.LineNumber & vbCrLf & pRemark.Text)
                Else
                    Dim pVesselRegistration As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.AmadeusValue("TextVSR")) + MySettings.AmadeusValue("TextVSR").Length)
                    mobjExistingAmadeusElements.VesselFlag.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), "", pVesselRegistration)
                End If
            ElseIf pRemark.Text.Contains(MySettings.AmadeusValue("TextREF")) Then
                Dim pReference As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.AmadeusValue("TextREF")) + MySettings.AmadeusValue("TextREF").Length)
                mobjExistingAmadeusElements.Reference.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), "", pReference)
            ElseIf pRemark.Text.Contains(MySettings.AmadeusValue("TextBBY")) Then
                Dim pBookedBy As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.AmadeusValue("TextBBY")) + MySettings.AmadeusValue("TextBBY").Length)
                mobjExistingAmadeusElements.BookedBy.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), "", pBookedBy)
            ElseIf pRemark.Text.Contains(MySettings.AmadeusValue("TextDPT")) Then
                Dim pDepartment As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.AmadeusValue("TextDPT")) + MySettings.AmadeusValue("TextDPT").Length)
                mobjExistingAmadeusElements.Department.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), True, pDepartment)
            ElseIf pRemark.Text.Contains(MySettings.AmadeusValue("TextRFT")) Then
                Dim pReasonForTravel As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.AmadeusValue("TextRFT")) + MySettings.AmadeusValue("TextRFT").Length)
                mobjExistingAmadeusElements.ReasonForTravel.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), "", pReasonForTravel)
            ElseIf pRemark.Text.Contains(MySettings.AmadeusValue("TextCC")) Then
                Dim pCostCentre As String = pRemark.Text.Substring(pRemark.Text.IndexOf(MySettings.AmadeusValue("TextCC")) + MySettings.AmadeusValue("TextCC").Length)
                mobjExistingAmadeusElements.CostCentre.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), "", pCostCentre)
            ElseIf pRemark.Text.Contains(MySettings.AmadeusValue("TextCLA")) Then
                If mobjExistingAmadeusElements.CustomerName.Exists Then
                    Throw New Exception("Please check PNR. Duplicate customer name defined" & vbCrLf & mobjExistingAmadeusElements.CustomerName.LineNumber & vbCrLf & pRemark.Text)
                Else
                    mobjExistingAmadeusElements.CustomerName.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), "", "")
                End If
            ElseIf pRemark.Text.Contains(MySettings.AmadeusValue("TextSBA")) Then
                mobjExistingAmadeusElements.SubDepartmentName.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), "", "")
            ElseIf pRemark.Text.Contains(MySettings.AmadeusValue("TextCRA")) Then
                mobjExistingAmadeusElements.CRMName.SetValues(True, pRemark.Text.Substring(0, pRemark.Text.IndexOf(pRemark.ElementID) - 1), "", "")
            End If
        Next

    End Sub

    Private Sub GetTickets()

        mobjTicketElements = New AmadeusTickets.Collection(mobjPNR)

    End Sub

    Public ReadOnly Property Tickets As AmadeusTickets.Collection
        Get
            Tickets = mobjTicketElements
        End Get
    End Property
    Private Sub mobjSession_ReceivedResponse(ByRef newResponse As k1aHostToolKit.CHostResponse) Handles mobjSession.ReceivedResponse

        mstrPNRResponse = newResponse.Text

    End Sub

End Class
