Option Strict On
Option Explicit On
Namespace GDSNew
    Friend Class Item
        Private Structure NewItemClass
            Dim GDSCommand As String
            Dim TextRequested As String
            Friend Sub Clear()
                GDSCommand = ""
                TextRequested = ""
            End Sub

        End Structure
        Private mudtProps As NewItemClass

        Public ReadOnly Property GDSCommand As String
            Get
                GDSCommand = mudtProps.GDSCommand
            End Get
        End Property
        Public ReadOnly Property TextRequested As String
            Get
                TextRequested = mudtProps.TextRequested
            End Get
        End Property

        Friend Sub SetText(ByVal Value As String)
            mudtProps.TextRequested = Value
        End Sub

        Friend Sub SetText(ByVal Value As String, ByVal pGDSCommand As String)
            mudtProps.TextRequested = Value
            mudtProps.GDSCommand = pGDSCommand
        End Sub

        Friend Sub Clear()
            mudtProps.Clear()
        End Sub

    End Class
    Friend Class Collection
        Private mobjOpenSegment As New Item
        Private mobjPhoneElement As New Item
        Private mobjEmailElement As New Item
        Private mobjTicketElement As New Item

        Private mobjOptionQueueElement As New Item
        Private mobjAOH As New Item
        Private mobjAgentID As New Item
        Private mobjSavingsElement As New Item
        Private mobjLossElement As New Item

        Private mobjCustomerCodeAI As New Item
        Private mobjCustomerCode As New Item
        Private mobjCustomerName As New Item
        Private mobjSubDepartmentCode As New Item
        Private mobjSubDepartmentName As New Item
        Private mobjCRMCode As New Item
        Private mobjCRMName As New Item
        Private mobjVesselName As New Item
        Private mobjVesselFlag As New Item
        Private mobjVesselOSI As New Item
        Private mobjReference As New Item
        Private mobjBookedBy As New Item
        Private mobjDepartment As New Item
        Private mobjReasonForTravel As New Item
        Private mobjCostCentre As New Item
        Private mobjTRId As New Item

        Private mobjVesselNameForPNR As New Item
        Private mobjVesselFlagForPNR As New Item

        Private mobjGalTracking As New Item
        Private mobjGreekToLatin As New GreekToLatin

        Private mstrOfficeOfResponsibility As String
        Private mdteDepartureDate As Date
        Private mintNumberOfPax As Integer
        Private mGDSCode As Utilities.EnumGDSCode

        Public Sub New(ByVal pOfficeOfResponsibility As String, ByVal pDepartureDate As Date, ByVal pNumberOfPax As Integer, ByVal pGDSCode As Utilities.EnumGDSCode)
            mstrOfficeOfResponsibility = pOfficeOfResponsibility
            mdteDepartureDate = pDepartureDate
            mintNumberOfPax = pNumberOfPax
            mGDSCode = pGDSCode
            PrepareCommands()
        End Sub
        Public Sub New()
            mstrOfficeOfResponsibility = ""
            mdteDepartureDate = Date.MinValue
            mintNumberOfPax = 0
            mGDSCode = Utilities.EnumGDSCode.Unknown
        End Sub
        Public ReadOnly Property OpenSegment As Item
            Get
                OpenSegment = mobjOpenSegment
            End Get
        End Property
        Public ReadOnly Property PhoneElement As Item
            Get
                PhoneElement = mobjPhoneElement
            End Get
        End Property
        Public ReadOnly Property EmailElement As Item
            Get
                EmailElement = mobjEmailElement
            End Get
        End Property
        Public ReadOnly Property TicketElement As Item
            Get
                TicketElement = mobjTicketElement
            End Get
        End Property
        Public ReadOnly Property OptionQueueElement As Item
            Get
                OptionQueueElement = mobjOptionQueueElement
            End Get
        End Property
        Public ReadOnly Property AOH As Item
            Get
                AOH = mobjAOH
            End Get
        End Property
        Public ReadOnly Property AgentID As Item
            Get
                AgentID = mobjAgentID
            End Get
        End Property
        Public ReadOnly Property CustomerCodeAI As Item
            Get
                Return mobjCustomerCodeAI
            End Get
        End Property
        Public ReadOnly Property CustomerCode As Item
            Get
                CustomerCode = mobjCustomerCode
            End Get
        End Property
        Public ReadOnly Property SavingsElement As Item
            Get
                SavingsElement = mobjSavingsElement
            End Get
        End Property
        Public ReadOnly Property LossElement As Item
            Get
                LossElement = mobjLossElement
            End Get
        End Property
        Public ReadOnly Property CustomerName As Item
            Get
                CustomerName = mobjCustomerName
            End Get
        End Property
        Public ReadOnly Property SubDepartmentCode As Item
            Get
                SubDepartmentCode = mobjSubDepartmentCode
            End Get
        End Property
        Public ReadOnly Property SubDepartmentName As Item
            Get
                SubDepartmentName = mobjSubDepartmentName
            End Get
        End Property
        Public ReadOnly Property CRMCode As Item
            Get
                CRMCode = mobjCRMCode
            End Get
        End Property
        Public ReadOnly Property CRMName As Item
            Get
                CRMName = mobjCRMName
            End Get
        End Property
        Public ReadOnly Property VesselName As Item
            Get
                VesselName = mobjVesselName
            End Get
        End Property
        Public ReadOnly Property VesselFlag As Item
            Get
                VesselFlag = mobjVesselFlag
            End Get
        End Property
        Public ReadOnly Property VesselOSI As Item
            Get
                VesselOSI = mobjVesselOSI
            End Get
        End Property
        Public ReadOnly Property VesselNameForPNR As Item
            Get
                VesselNameForPNR = mobjVesselNameForPNR
            End Get
        End Property
        Public ReadOnly Property VesselFlagForPNR As Item
            Get
                VesselFlagForPNR = mobjVesselFlagForPNR
            End Get
        End Property
        Public ReadOnly Property Reference As Item
            Get
                Reference = mobjReference
            End Get
        End Property
        Public ReadOnly Property BookedBy As Item
            Get
                BookedBy = mobjBookedBy
            End Get
        End Property
        Public ReadOnly Property GalileoTrackingCode As Item
            Get
                GalileoTrackingCode = mobjGalTracking
            End Get
        End Property
        Public ReadOnly Property Department As Item
            Get
                Department = mobjDepartment
            End Get
        End Property
        Public ReadOnly Property ReasonForTravel As Item
            Get
                ReasonForTravel = mobjReasonForTravel
            End Get
        End Property
        Public ReadOnly Property CostCentre As Item
            Get
                CostCentre = mobjCostCentre
            End Get
        End Property
        Public ReadOnly Property TRId As Item
            Get
                TRId = mobjTRId
            End Get
        End Property

        Public Sub SetItem(ByVal Item As Customers.CustomerItem)

            mobjCustomerCodeAI.Clear()
            mobjCustomerCode.Clear()
            mobjCustomerName.Clear()
            If Not Item Is Nothing Then
                If Item.Code <> "" Then
                    mobjCustomerCode.SetText(Item.Code, MySettings.GDSValue("TextCLN") & Item.Code)
                    mobjCustomerCodeAI.SetText(Item.Code, MySettings.GDSValue("TextAI") & Item.Code)
                End If
                If Item.Name <> "" Then
                    mobjCustomerName.SetText(Item.Name, MySettings.GDSValue("TextCLA") & mobjGreekToLatin.Convert(Item.Name))
                End If
                If Item.GalileoTrackingCode <> "" Then
                    mobjGalTracking.SetText(Item.GalileoTrackingCode, MySettings.GDSValue("TextGalTrackingCode") & Item.GalileoTrackingCode)
                End If
            End If

        End Sub
        Public Sub SetItem(ByVal Item As SubDepartments.Item)

            mobjSubDepartmentCode.Clear()
            mobjSubDepartmentName.Clear()
            If Not Item Is Nothing Then
                If Item.ID > 0 Then
                    mobjSubDepartmentCode.SetText(Item.Code, MySettings.GDSValue("TextSBN") & Item.ID)
                End If
                If Item.Name <> "" Then
                    mobjSubDepartmentName.SetText(Item.Name, MySettings.GDSValue("TextSBA") & mobjGreekToLatin.Convert(Item.Name))
                End If
            End If

        End Sub
        Public Sub SetItem(ByVal Item As CRM.Item)

            mobjCRMCode.Clear()
            mobjCRMName.Clear()
            If Not Item Is Nothing Then
                If Item.ID > 0 Then
                    mobjCRMCode.SetText(Item.Code, MySettings.GDSValue("TextCRN") & Item.Code)
                End If
                If Item.Name <> "" Then
                    mobjCRMName.SetText(Item.Name, MySettings.GDSValue("TextCRA") & mobjGreekToLatin.Convert(Item.Name))
                End If
            End If

        End Sub
        Public Sub SetItem(ByVal Item As Vessels.Item)

            mobjVesselName.Clear()
            mobjVesselFlag.Clear()
            If Not Item Is Nothing Then
                mobjVesselName.SetText(Item.Name, MySettings.GDSValue("TextVSL") & Item.Name)
                If Item.Flag <> "" Then
                    mobjVesselFlag.SetText(Item.Flag, MySettings.GDSValue("TextVSR") & Item.Flag)
                End If
                mobjVesselOSI.SetText("", MySettings.GDSValue("TextVOS") & Item.Name)
                mobjVesselNameForPNR.Clear()
                mobjVesselFlagForPNR.Clear()
            End If

        End Sub
        Public Sub SetVesselForPNR(ByVal pVesselName As String, ByVal pVesselFlag As String)

            mobjVesselName.Clear()
            mobjVesselFlag.Clear()
            If pVesselName <> "" Then
                mobjVesselName.SetText("", MySettings.GDSValue("TextVSL") & mobjVesselNameForPNR.TextRequested)
                mobjVesselOSI.SetText("", MySettings.GDSValue("TextVOS") & mobjVesselNameForPNR.TextRequested)
                If pVesselFlag <> "" Then
                    mobjVesselFlag.SetText("", MySettings.GDSValue("TextVSR") & mobjVesselFlagForPNR.TextRequested)
                    If mobjVesselOSI.GDSCommand <> "" Then
                        mobjVesselOSI.SetText("", MySettings.GDSValue("TextVOS") & mobjVesselFlagForPNR.TextRequested)
                    End If
                End If
            End If

            mobjVesselNameForPNR.SetText(pVesselName)
            mobjVesselFlagForPNR.SetText(pVesselFlag)

        End Sub
        Public Sub SetReference(ByVal Text As String)
            Text = Text.Trim
            If Text <> "" Then
                mobjReference.SetText(Text, MySettings.GDSValue("TextREF") & Text)
            Else
                mobjReference.Clear()
            End If
        End Sub
        Public Sub SetBookedBy(ByVal Text As String)
            Text = Text.Trim
            If Text <> "" Then
                mobjBookedBy.SetText(Text, MySettings.GDSValue("TextBBY") & Text)
            Else
                mobjBookedBy.Clear()
            End If
        End Sub
        Public Sub SetGalileoTracking(ByVal Text As String)
            Text = Text.Trim
            If Text <> "" Then
                mobjGalTracking.SetText(Text, MySettings.GDSValue("TextGalTrackingCode") & Text)
            Else
                mobjGalTracking.Clear()
            End If
        End Sub
        Public Sub SetDepartment(ByVal Text As String)
            Text = Text.Trim
            If Text <> "" Then
                mobjDepartment.SetText(Text, MySettings.GDSValue("TextDPT") & Text)
            Else
                mobjDepartment.Clear()
            End If
        End Sub
        Public Sub SetReasonForTravel(ByVal Text As String)
            Text = Text.Trim
            If Text <> "" Then
                mobjReasonForTravel.SetText(Text, MySettings.GDSValue("TextRFT") & Text)
            Else
                mobjReasonForTravel.Clear()
            End If
        End Sub
        Public Sub SetCostCentre(ByVal Text As String)
            Text = Text.Trim
            If Text <> "" Then
                mobjCostCentre.SetText(Text, MySettings.GDSValue("TextCC") & Text)
            Else
                mobjCostCentre.Clear()
            End If
        End Sub
        Public Sub SetTRId(ByVal Text As String)
            Text = Text.Trim
            If Text <> "" Then
                mobjTRId.SetText(Text, MySettings.GDSValue("TextTRID") & Text)
            Else
                mobjTRId.Clear()
            End If
        End Sub
        Private Sub PrepareCommands()

            Dim pDateTimeLimit As New s1aAirlineDate.clsAirlineDate
            Dim pDateReminder As New s1aAirlineDate.clsAirlineDate
            Dim pDateRetain As New s1aAirlineDate.clsAirlineDate

            If mdteDepartureDate > DateAdd(DateInterval.Day, 3, Today) Then
                Try
                    pDateTimeLimit.VBDate = DateAdd(DateInterval.Day, -3, mdteDepartureDate)
                Catch ex As Exception
                    pDateTimeLimit.VBDate = Today
                End Try
            Else
                pDateTimeLimit.VBDate = Today
            End If

            If mdteDepartureDate > Today Then
                Try
                    pDateReminder.VBDate = DateAdd(DateInterval.Day, 1, mdteDepartureDate)
                Catch ex As Exception
                    pDateReminder.VBDate = Today
                End Try
            Else
                pDateReminder.VBDate = Today
            End If

            Try
                pDateRetain.VBDate = DateAdd(DateInterval.Month, 11, Today)
            Catch ex As Exception
                pDateRetain.VBDate = Today
            End Try

            mobjPhoneElement.SetText("", (MySettings.GDSValue("TextAP").Replace("  ", " ")))
            mobjEmailElement.SetText("", MySettings.GDSValue("TextAPE"))
            mobjAgentID.SetText("", MySettings.GDSValue("TextAGT"))

            If mGDSCode = Utilities.EnumGDSCode.Amadeus Then
                Dim pTTLString As String
                If mstrOfficeOfResponsibility <> MySettings.GDSPcc Then
                    pTTLString = MySettings.GDSValue("TextTTL") & pDateTimeLimit.IATA & "/" & MySettings.GDSPcc
                Else
                    pTTLString = MySettings.GDSValue("TextTTL") & pDateTimeLimit.IATA
                End If
                mobjTicketElement.SetText("", pTTLString)
                mobjOpenSegment.SetText("",
                                    MySettings.GDSValue("TextMISSegmentCommand") &
                                    IIf(mintNumberOfPax = 0, 1, mintNumberOfPax).ToString & " " &
                                    MySettings.OfficeCityCode & " " &
                                    pDateRetain.IATA & "-" & MySettings.GDSValue("TextMISSegmentText"))
                mobjOptionQueueElement.SetText("", MySettings.GDSValue("TextOP") & MySettings.GDSPcc & "/" & pDateReminder.IATA & "/" & MySettings.AgentOPQueue)
            ElseIf mGDSCode = Utilities.EnumGDSCode.Galileo Then
                mobjTicketElement.SetText("", MySettings.GDSValue("TextTTL") & pDateTimeLimit.IATA)
                mobjOpenSegment.SetText("", MySettings.GDSValue("TextMISSegmentCommand") & pDateRetain.IATA & "*" & MySettings.GDSValue("TextMISSegmentText"))
                mobjOptionQueueElement.SetText("", MySettings.GDSValue("TextOP") & "/" & pDateReminder.IATA & "/0001/Q" & MySettings.AgentOPQueue)
            Else
                Throw New Exception("GDSNew.PrepareCommands()" & vbCrLf & "GDS Not selected")
            End If
            mobjAOH.SetText("", MySettings.GDSValue("TextAOH"))

        End Sub
        Public Sub Clear()

            mobjOpenSegment.Clear()
            mobjPhoneElement.Clear()
            mobjEmailElement.Clear()
            mobjTicketElement.Clear()

            mobjOptionQueueElement.Clear()
            mobjAOH.Clear()
            mobjAgentID.Clear()
            mobjSavingsElement.Clear()
            mobjLossElement.Clear()

            ClearCustomerElements()

        End Sub
        Public Sub ClearCustomerElements()
            mobjCustomerCodeAI.Clear()
            mobjCustomerCode.Clear()
            mobjCustomerName.Clear()
            mobjSubDepartmentCode.Clear()
            mobjSubDepartmentName.Clear()
            mobjCRMCode.Clear()
            mobjCRMName.Clear()
            mobjVesselName.Clear()
            mobjVesselFlag.Clear()
            mobjVesselOSI.Clear()
            mobjReference.Clear()
            mobjBookedBy.Clear()
            mobjGalTracking.Clear()
            mobjDepartment.Clear()
            mobjReasonForTravel.Clear()
            mobjCostCentre.Clear()
            mobjVesselNameForPNR.Clear()
            mobjVesselFlagForPNR.Clear()
            mobjTRId.Clear()
        End Sub
    End Class

End Namespace
