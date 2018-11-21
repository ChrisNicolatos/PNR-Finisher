Option Strict On
Option Explicit On
Friend Class ItnRTBDoc
    Private mobjPNR As GDSReadPNR
    Private mintMaxString As Integer
    Private mstrRemarks As String
    Private mintHeaderLength As Integer = 0
    Public Sub New(ByRef pPNR As GDSReadPNR, ByVal pMaxString As Integer, ByRef pItnRemarks As CheckedListBox)
        mobjPNR = pPNR
        mintMaxString = pMaxString
        mstrRemarks = ""
        For iRem As Integer = 0 To pItnRemarks.CheckedItems.Count - 1
            Dim pItem As RemarksItem
            pItem = CType(pItnRemarks.CheckedItems(iRem), RemarksItem)
            mstrRemarks &= pItem.Remark & vbCrLf
        Next

    End Sub
    Public ReadOnly Property RTBDocPassengers As String
        Get
            Dim pString As New System.Text.StringBuilder
            With mobjPNR
                If .Passengers.Count > 0 Then
                    If MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefs Or MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefsWithCode Then
                        pString.AppendLine("FOR PASSENGER" & If(.Passengers.Count > 1, "(S)", ""))
                    End If
                    For Each pobjPax In .Passengers.Values
                        If MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefs Or MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefsWithCode Then
                            pString.AppendLine(pobjPax.PaxName)
                        Else
                            pString.AppendLine(pobjPax.ElementNo & " " & pobjPax.PaxName & " " & pobjPax.PaxID)
                        End If
                    Next pobjPax
                ElseIf .IsGroup Then
                    pString.AppendLine("GROUP: " & .GroupName & " " & .GroupNamesCount)
                Else
                    pString.AppendLine("PASSENGER INFORMATION NOT AVAILABLE")
                End If
                RTBDocPassengers = pString.ToString
            End With
        End Get
    End Property
    Public ReadOnly Property makeRTBDoc() As String
        Get
            Dim pString As New System.Text.StringBuilder

            pString.Clear()
            mintMaxString = 80

            Try
                'TODO - Fix length of output line total 78 characters including spaces
                pString.Append(MakeRTBDocPart1)
                pString.Append(MakeRTBDocTickets)
                If MySettings.ShowItinRemarks Then
                    pString.Append(MakeRTBDocItinRemarks)
                End If
                If Not (MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefs Or MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefsWithCode) And mintMaxString > 0 Then
                    pString.AppendLine(StrDup(mintHeaderLength, "-"))
                End If
                pString.AppendLine()
                pString.Append(mstrRemarks)
            Catch ex As Exception
                Throw New Exception("makeRTBDoc()" & vbCrLf & ex.Message)
            End Try
            makeRTBDoc = pString.ToString
        End Get
    End Property
    Public ReadOnly Property MakeRTBMSReport() As String
        Get
            Try
                Dim pString As New System.Text.StringBuilder
                With mobjPNR
                    If .HasSegments Then
                        Dim pDepTime As String = ""
                        Dim pArrTime As String = ""
                        For Each pobjPax In .Passengers.Values
                            If .LastSegment.Text.Substring(35, 4) = "FLWN" Then
                                pDepTime = "FLOWN"
                                pArrTime = "FLOWN"
                            Else
                                pDepTime = Format(.LastSegment.DepartTime, "HH:mm")
                                pArrTime = Format(.LastSegment.ArriveTime, "HH:mm")
                            End If
                            pString.AppendLine(pobjPax.LastName & vbTab & pobjPax.Initial & vbTab & pobjPax.IdNo & vbTab & pobjPax.Department & vbTab & .VesselName &
                                               vbTab & .LastSegment.DepartureDateIATA & vbTab & .LastSegment.Airline & vbTab & .LastSegment.FlightNo & vbTab & pDepTime &
                                               vbTab & .LastSegment.BoardPoint & vbTab & pArrTime & vbTab & .LastSegment.OffPoint & vbTab & .RequestedPNR & vbTab & pobjPax.ElementNo)
                        Next pobjPax
                        Return pString.ToString
                    Else
                        Return ""
                    End If
                End With
                MakeRTBMSReport = pString.ToString
            Catch ex As Exception
                Throw New Exception("MakeRTBMSReport()" & vbCrLf & ex.Message)
            End Try
        End Get
    End Property
    Private ReadOnly Property MakeRTBDocPart1 As String
        Get
            Try
                Dim pString As New System.Text.StringBuilder
                Dim pAirlineLocator As String = ""
                Dim pobjSeg As GDSSegItem
                pString.Clear()
                With mobjPNR

                    pString.Clear()
                    Dim pTemp As String = ""
                    If Not (MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefs Or MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefsWithCode) And MySettings.ShowVessel And .VesselName <> "" Then
                        pTemp &= "VESSEL     : " & .VesselName
                    End If
                    If Not (MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefs Or MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefsWithCode) And MySettings.ShowCostCentre And .CostCentre <> "" Then
                        If pTemp <> "" Then
                            pTemp &= vbCrLf
                        End If
                        pTemp &= "COST CENTRE: " & .CostCentre
                    End If
                    If pTemp <> "" Then
                        pString.AppendLine(" ")
                        pString.AppendLine(pTemp)
                        pString.AppendLine(" ")
                    End If
                    Dim pHeader As New System.Text.StringBuilder

                    If MySettings.FormatStyle = Utilities.EnumItnFormat.DefaultFormat Then
                        pHeader.Append("Flight ")
                        If MySettings.ShowClassOfService Then
                            pHeader.Append("C ")
                        End If
                        pHeader.Append("Date  ")
                        Select Case MySettings.AirportName
                            Case 0
                                pHeader.Append("Org Dest")
                            Case 1
                                pHeader.Append("Origin " & StrDup(.MaxAirportNameLength - 5, " ") & "Destination" & StrDup(.MaxAirportNameLength - 9, " "))
                            Case 2
                                pHeader.Append("Origin " & StrDup(.MaxAirportNameLength - 1, " ") & "Destination" & StrDup(.MaxAirportNameLength - 5, " "))
                            Case 3
                                pHeader.Append("Origin " & StrDup(.MaxAirportShortNameLength - 5, " ") & "Destination" & StrDup(.MaxAirportShortNameLength - 9, " "))
                            Case 4
                                pHeader.Append("Origin " & StrDup(.MaxAirportShortNameLength - 1, " ") & "Destination" & StrDup(.MaxAirportShortNameLength - 5, " "))
                        End Select
                        pHeader.Append("Dep   ")
                        pHeader.Append("Arr   ")
                        If MySettings.ShowFlyingTime Then
                            pHeader.Append(" EFT  ")
                        End If
                        pHeader.Append("ArrDte ")
                        pHeader.Append(If(MySettings.ShowAirlineLocator, "AL Locator", ""))
                        pHeader.Append(" - BagAl")
                        If MySettings.ShowCabinDescription Then
                            pHeader.Append(" Class")
                        End If

                        mintHeaderLength = pHeader.Length

                        pString.AppendLine(StrDup(mintHeaderLength, "-"))
                        pString.AppendLine(pHeader.ToString)
                        pString.AppendLine(StrDup(mintHeaderLength, "-"))
                    ElseIf MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefs Or MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefsWithCode Then
                        pHeader.Append("Flight ")
                        pHeader.Append("Date  ")
                        If MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefsWithCode Then
                            pHeader.Append("Org    " & StrDup(.MaxAirportShortNameLength - 1, " ") & "Dest       " & StrDup(.MaxAirportShortNameLength - 5, " "))
                        Else
                            pHeader.Append("Org    " & StrDup(.MaxAirportShortNameLength - 5, " ") & "Dest       " & StrDup(.MaxAirportShortNameLength - 9, " "))
                        End If
                        pHeader.Append("Dep   ")
                        pHeader.Append("Arr   ")
                        pHeader.Append("Term   ")
                        pHeader.Append("Status")
                        pHeader.Append("   BagAl")
                        mintHeaderLength = pHeader.Length

                        pString.AppendLine(StrDup(mintHeaderLength, "-"))
                        pString.AppendLine(pHeader.ToString)
                        pString.AppendLine(StrDup(mintHeaderLength, "-"))
                    End If

                    Dim iSegCount As Integer = 0
                    'Dim pPrevOff As String = ""
                    For Each pobjSeg In .Segments.Values
                        iSegCount = iSegCount + 1
                        Dim pSeg As New System.Text.StringBuilder

                        If MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefs Or MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefsWithCode Then
                            pSeg.Append(pobjSeg.Airline & pobjSeg.FlightNo.PadLeft(4) & " ")
                            pSeg.Append(pobjSeg.DepartureDateIATA & " ")
                            If MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefsWithCode Then
                                pSeg.Append(pobjSeg.BoardPoint & " " & pobjSeg.BoardAirportShortName.PadRight(.MaxAirportShortNameLength + 1, " "c).Substring(0, .MaxAirportShortNameLength + 1) & " ")
                                pSeg.Append(pobjSeg.OffPoint & " " & pobjSeg.OffPointAirportShortName.PadRight(.MaxAirportShortNameLength + 1, " "c).Substring(0, .MaxAirportShortNameLength + 1) & " ")
                            Else
                                pSeg.Append(pobjSeg.BoardAirportShortName.PadRight(.MaxAirportShortNameLength + 1, " "c).Substring(0, .MaxAirportShortNameLength + 1) & " ")
                                pSeg.Append(pobjSeg.OffPointAirportShortName.PadRight(.MaxAirportShortNameLength + 1, " "c).Substring(0, .MaxAirportShortNameLength + 1) & " ")
                            End If
                            If pobjSeg.Text.Length > 35 AndAlso pobjSeg.Text.Substring(35, 4) = "FLWN" Then
                                pSeg.Append("FLWN")
                            Else
                                pSeg.Append(Format(pobjSeg.DepartTime, "HHmm") & "  ")
                                pSeg.Append(Format(pobjSeg.ArriveTime, "HHmm"))
                                If DateDiff(DateInterval.Day, pobjSeg.DepartureDate, pobjSeg.ArrivalDate) > 0 Then
                                    pSeg.Append("+" & DateDiff(DateInterval.Day, pobjSeg.DepartureDate, pobjSeg.ArrivalDate) & " ")
                                ElseIf DateDiff(DateInterval.Day, pobjSeg.DepartureDate, pobjSeg.ArrivalDate) < 0 Then
                                    pSeg.Append(DateDiff(DateInterval.Day, pobjSeg.DepartureDate, pobjSeg.ArrivalDate) & " ")
                                Else
                                    pSeg.Append("   ")
                                End If
                                If pobjSeg.DepartTerminal <> "" Then
                                    If pobjSeg.DepartTerminal.LastIndexOf(" ") > -1 Then
                                        pSeg.Append(pobjSeg.DepartTerminal.Substring(pobjSeg.DepartTerminal.LastIndexOf(" ")).PadLeft(3))
                                    Else
                                        pSeg.Append("   ")
                                    End If
                                Else
                                    pSeg.Append("   ")
                                End If

                                If pobjSeg.Status = "HL" Then
                                    pSeg.Append("      HL")
                                Else
                                    pSeg.Append("      OK")
                                End If
                                pSeg.Append(mobjPNR.AllowanceForSegment(pobjSeg.BoardPoint, pobjSeg.OffPoint, pobjSeg.Airline).PadLeft(8))
                                'If MySettings.ShowCabinDescription Then
                                '    pSeg.Append(" " & GetClassOfService(pobjSeg.Airline, pobjSeg.BoardPoint, pobjSeg.OffPoint, pobjSeg.ClassOfService).CabinDescription)
                                'End If
                                If pAirlineLocator.IndexOf(pobjSeg.AirlineLocator.Trim) = -1 Then
                                    If pAirlineLocator = "" Then
                                        pAirlineLocator = "AIRLINE REF: " & pobjSeg.AirlineLocator.Trim '& "(" & pobjSeg.Airline & " " & pobjSeg.AirlineName & ")"
                                    Else
                                        pAirlineLocator &= vbCrLf & "             " & pobjSeg.AirlineLocator.Trim '& "(" & pobjSeg.Airline & " " & pobjSeg.AirlineName & ")"
                                    End If
                                    '                                    pAirlineLocator &= pobjSeg.AirlineLocator.Trim & "(" & pobjSeg.Airline & " " & pobjSeg.AirlineName & ")"
                                End If
                            End If
                        Else
                            pSeg.Append(pobjSeg.Airline & pobjSeg.FlightNo.PadLeft(4) & " ")
                            If MySettings.ShowClassOfService Then
                                pSeg.Append(pobjSeg.ClassOfService & " ")
                            End If
                            pSeg.Append(pobjSeg.DepartureDateIATA & " ")
                            Select Case MySettings.AirportName
                                Case 0 'code
                                    pSeg.Append(pobjSeg.BoardPoint & " " & pobjSeg.OffPoint & " ")
                                Case 1 'airport name
                                    pSeg.Append(pobjSeg.BoardAirportName.PadRight(.MaxAirportNameLength + 1, " "c).Substring(0, .MaxAirportNameLength + 1) & " " &
                                                pobjSeg.OffPointAirportName.PadRight(.MaxAirportNameLength + 1, " "c).Substring(0, .MaxAirportNameLength + 1) & " ")
                                Case 2 'code and airport
                                    pSeg.Append(pobjSeg.BoardPoint & " " & pobjSeg.BoardAirportName.PadRight(.MaxAirportNameLength + 1, " "c).Substring(0, .MaxAirportNameLength + 1) & " " &
                                                pobjSeg.OffPoint & " " & pobjSeg.OffPointAirportName.PadRight(.MaxAirportNameLength + 1, " "c).Substring(0, .MaxAirportNameLength + 1) & " ")
                                Case 3 'city name
                                    pSeg.Append(pobjSeg.BoardAirportShortName.PadRight(.MaxAirportShortNameLength + 1, " "c).Substring(0, .MaxAirportShortNameLength + 1) & " " &
                                                pobjSeg.OffPointAirportShortName.PadRight(.MaxAirportShortNameLength + 1, " "c).Substring(0, .MaxAirportShortNameLength + 1) & " ")
                                Case 4 'code and city
                                    pSeg.Append(pobjSeg.BoardPoint & " " & pobjSeg.BoardAirportShortName.PadRight(.MaxAirportShortNameLength + 1, " "c).Substring(0, .MaxAirportShortNameLength + 1) & " " &
                                                pobjSeg.OffPoint & " " & pobjSeg.OffPointAirportShortName.PadRight(.MaxAirportShortNameLength + 1, " "c).Substring(0, .MaxAirportShortNameLength + 1) & " ")
                            End Select
                            If pobjSeg.Text.Length > 35 AndAlso pobjSeg.Text.Substring(35, 4) = "FLWN" Then
                                pSeg.Append("FLWN")
                            Else
                                'pSeg.Append(pobjSeg.Status.PadRight(3))
                                pSeg.Append(Format(pobjSeg.DepartTime, "HHmm") & "  ")
                                pSeg.Append(Format(pobjSeg.ArriveTime, "HHmm") & "  ")
                                If MySettings.ShowFlyingTime Then
                                    pSeg.Append(pobjSeg.EstimatedFlyingTime & " ")
                                End If
                                pSeg.Append(pobjSeg.ArrivalDateIATA & "   ")
                                pSeg.Append(If(MySettings.ShowAirlineLocator, pobjSeg.AirlineLocator.PadRight(9, " "c), ""))
                                pSeg.Append(" - " & mobjPNR.AllowanceForSegment(pobjSeg.BoardPoint, pobjSeg.OffPoint, pobjSeg.Airline).PadLeft(5))
                                If pobjSeg.Status = "HL" Then
                                    pSeg.Append("   WAITLISTED")
                                End If
                                If MySettings.ShowCabinDescription Then
                                    pSeg.Append(" " & GetClassOfService(pobjSeg.Airline, pobjSeg.BoardPoint, pobjSeg.OffPoint, pobjSeg.ClassOfService).CabinDescription)
                                End If

                                If MySettings.ShowTerminal And pobjSeg.DepartTerminal <> "" Then
                                    pSeg.Append("   " & pobjSeg.DepartTerminal)
                                End If
                            End If
                        End If

                        pString.AppendLine(pSeg.ToString)
                        If pobjSeg.Equipment = "TRN" Then
                            pString.AppendLine("             ***     TRAIN     ****  ")
                        End If

                        If Not MySettings.FormatStyle = Utilities.EnumItnFormat.Plain Then
                            If pobjSeg.OperatedBy <> "" Then
                                pString.AppendLine(StrDup(13, " ") & pobjSeg.OperatedBy)
                            End If
                            If (MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefs Or MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefsWithCode Or MySettings.ShowStopovers) And pobjSeg.Stopovers <> "" Then
                                pString.AppendLine("             *INTERMEDIATE STOP*  " & pobjSeg.Stopovers)
                            End If
                        End If

                        If pSeg.ToString.Length > mintMaxString Then
                            mintMaxString = pSeg.ToString.Length
                        End If
                    Next pobjSeg

                    If iSegCount = 0 Then
                        pString.AppendLine("ROUTING INFORMATION NOT AVAILABLE")
                    End If

                    If .RequestedPNR <> "" Then
                        pString.AppendLine(" ")
                        If MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefs Or MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefsWithCode Then
                            pString.AppendLine("ATPI REF   : " & .GDSAbbreviation & "/" & .RequestedPNR)
                            If pAirlineLocator <> "" Then
                                pString.AppendLine(pAirlineLocator)
                            End If
                        Else

                            pString.AppendLine("ATPI Booking Reference: " & .GDSAbbreviation & "/" & .RequestedPNR)
                        End If
                    End If

                End With

                Return pString.ToString
            Catch ex As Exception
                Throw New Exception("MakeRTBDocPart1()" & vbCrLf & ex.Message)
            End Try
        End Get
    End Property
    Private ReadOnly Property MakeRTBDocTickets As String
        Get
            Try
                Dim pString As New System.Text.StringBuilder
                pString.Clear()
                Dim pHeader As String = ""
                Dim pAncServicesTitle As String = ""
                If MySettings.FormatStyle = Utilities.EnumItnFormat.DefaultFormat Then
                    pHeader = "Ticket Number   "
                    If MySettings.ShowPaxSegPerTkt Then
                        pHeader &= "Routing      Passenger"
                    End If
                    pAncServicesTitle = "Ancillary Services"
                End If

                With mobjPNR
                    If MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefs Or MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefsWithCode Then
                        For Each pobjPax In .Passengers.Values
                            pString.AppendLine()
                            pString.AppendLine(pobjPax.PaxName)
                            For Each tkt As GDSTicketItem In .Tickets.Values
                                If tkt.TicketType = "PAX" Then
                                    If tkt.Pax.Trim = pobjPax.PaxName.Trim Or tkt.Pax.Trim.StartsWith(pobjPax.PaxName.Trim) Or pobjPax.PaxName.Trim.StartsWith(tkt.Pax.Trim) Then
                                        Dim pFF As String = mobjPNR.FrequentFlyerNumber(tkt.AirlineCode, tkt.Pax.Substring(0, tkt.Pax.Length - 2).Trim)
                                        If pFF <> "" Then
                                            pFF = " - Frequent Flyer Number: " & pFF
                                        End If
                                        If tkt.Document > 0 Then
                                            pString.AppendLine("ETICKET NUMBER: " _
                                                               & tkt.IssuingAirline & "-" & tkt.Document & " " & tkt.AirlineCode & " " & Airlines.AirlineName(tkt.AirlineCode) & pFF)
                                        Else
                                            pString.AppendLine(pFF)
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    Else
                        For iTickType = 1 To 2 ' 1 for tickets, 2 for EMD
                            For Each tkt As GDSTicketItem In .Tickets.Values
                                If MySettings.ShowTickets And iTickType = 1 And tkt.TicketType = "PAX" Then
                                    ' Tickets
                                    If pHeader <> "" Then
                                        pString.AppendLine(StrDup(mintHeaderLength, "-"))
                                        pString.AppendLine(pHeader)
                                        pString.AppendLine(StrDup(mintHeaderLength, "-"))
                                        pHeader = ""
                                    End If
                                    If MySettings.ShowPaxSegPerTkt Then
                                        Dim pFF As String = mobjPNR.FrequentFlyerNumber(tkt.AirlineCode, tkt.Pax.PadRight(3).Substring(0, tkt.Pax.PadRight(3).Length - 2).Trim)
                                        If pFF <> "" Then
                                            pFF = " - Frequent Flyer Number: " & pFF
                                        End If

                                        Dim pTemp As String = tkt.IssuingAirline & "-" & tkt.Document & If(tkt.Books > 1, tkt.Conjunction, "") & "  "
                                        pString.AppendLine(pTemp & tkt.Segs.PadRight(10).Substring(0, 10) & "   " & tkt.Pax.PadRight(3).Substring(0, tkt.Pax.PadRight(3).Length - 2) & pFF)
                                        pTemp = Space(pTemp.Length)
                                        For i As Integer = 12 To tkt.Segs.Length - 10 Step 12
                                            pString.AppendLine(pTemp & tkt.Segs.Substring(i, 10))
                                        Next
                                    Else
                                        pString.AppendLine(tkt.IssuingAirline & "-" & tkt.Document & If(tkt.Books > 1, tkt.Conjunction, ""))
                                    End If
                                ElseIf MySettings.ShowEMD And iTickType = 2 And tkt.TicketType <> "PAX" Then
                                    'EMDs
                                    If pAncServicesTitle <> "" Then
                                        pString.AppendLine(StrDup(mintHeaderLength, "-"))
                                        pString.AppendLine(pAncServicesTitle)
                                        pString.AppendLine(StrDup(mintHeaderLength, "-"))
                                        pAncServicesTitle = ""
                                    End If
                                    If MySettings.ShowPaxSegPerTkt Then
                                        Dim pTemp As String = tkt.IssuingAirline & "-" & tkt.Document & If(tkt.Books > 1, tkt.Conjunction, "") & "  "
                                        pString.AppendLine(pTemp & tkt.Segs.PadRight(10).Substring(0, 10) & "   " & tkt.Pax.PadRight(3).Substring(0, tkt.Pax.PadRight(3).Length - 2))
                                        pTemp = Space(pTemp.Length)
                                        For i As Integer = 12 To tkt.Segs.Length - 10 Step 12
                                            pString.AppendLine(pTemp & tkt.Segs.Substring(i, 10))
                                        Next
                                        If tkt.ServicesDescription <> "" Then
                                            pString.AppendLine(tkt.ServicesDescription)
                                        End If
                                    Else
                                        pString.AppendLine(tkt.IssuingAirline & "-" & tkt.Document & If(tkt.Books > 1, tkt.Conjunction, ""))
                                        If tkt.ServicesDescription <> "" Then
                                            pString.AppendLine(tkt.ServicesDescription)
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If

                    If MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefs Or MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefsWithCode Or MySettings.ShowSeating Then
                        If .Seats <> "" Then
                            If Not MySettings.FormatStyle = Utilities.EnumItnFormat.Plain Then
                                pString.AppendLine(StrDup(mintHeaderLength, "-"))
                            End If
                            pString.AppendLine("Seat Assignment")
                            If Not MySettings.FormatStyle = Utilities.EnumItnFormat.Plain Then
                                pString.AppendLine(StrDup(mintHeaderLength, "-"))
                            End If
                            pString.AppendLine(.Seats & vbCrLf)
                        End If
                    End If

                End With

                Return pString.ToString
            Catch ex As Exception
                Throw New Exception("MakeRTBDocTickets()" & vbCrLf & ex.Message)
            End Try
        End Get
    End Property
    Private ReadOnly Property MakeRTBDocItinRemarks As String
        Get
            Dim pString As New System.Text.StringBuilder
            Dim pFound As Boolean = False
            pString.Clear()
            For Each pItem As GDSItineraryRemarksItem In mobjPNR.ItineraryRemarks
                If Not pFound Then
                    pString.AppendLine(vbCrLf)
                    pFound = True
                End If
                pString.AppendLine(pItem.FreeFlow)
            Next
            Return pString.ToString
        End Get
    End Property
    Public ReadOnly Property RTBMSReportHeader(ByVal FromDate As String, ByVal ToDate As String) As String
        Get
            RTBMSReportHeader = "FROM " & FromDate & " : To " & ToDate & vbCrLf
            RTBMSReportHeader &= "Last Name" & vbTab & "First Name" & vbTab & "ID No." & vbTab & "Department" & vbTab & "Vessel Name" & vbTab & "Date Of Travel" & vbTab & "Airline" & vbTab & "Flight No." & vbTab & "Dep.Time" & vbTab & "Dep.City" & vbTab & "Arr.Time" & vbTab & "Arr.City" & vbTab & "PNR" & vbTab & "PaxNo" & vbCrLf
        End Get
    End Property
    Public ReadOnly Property MakeRTBMSReportOutsiderange() As String
        Get
            Try
                Dim pString As New System.Text.StringBuilder

                With mobjPNR
                    If .HasSegments Then
                        Dim pDepTime As String = ""
                        Dim pArrTime As String = ""
                        For Each pobjPax As GDSPaxItem In .Passengers.Values
                            For Each pSeg As GDSSegItem In .Segments.Values
                                If pSeg.Text.Substring(35, 4) = "FLWN" Then
                                    pDepTime = "FLOWN"
                                    pArrTime = "FLOWN"
                                Else
                                    pDepTime = Format(pSeg.DepartTime, "HH:mm")
                                    pArrTime = Format(pSeg.ArriveTime, "HH:mm")
                                End If
                                pString.AppendLine(pobjPax.LastName & vbTab & pobjPax.Initial & vbTab & pobjPax.IdNo & vbTab & pobjPax.Department & vbTab & .VesselName &
                                                   vbTab & pSeg.DepartureDateIATA & vbTab & pSeg.Airline & vbTab & pSeg.FlightNo & vbTab & pDepTime &
                                                   vbTab & pSeg.BoardPoint & vbTab & pArrTime & vbTab & pSeg.OffPoint & vbTab & .RequestedPNR & vbTab & pobjPax.ElementNo)
                            Next pSeg
                        Next pobjPax
                        pString.AppendLine(" ")
                        Return pString.ToString
                    Else
                        Return ""
                    End If
                End With
            Catch ex As Exception
                Throw New Exception("MakeRTBMSReportOutsiderange()" & vbCrLf & ex.Message)
            End Try
        End Get
    End Property
End Class
