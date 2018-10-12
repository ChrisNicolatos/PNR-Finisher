Option Strict Off
Option Explicit On
Public Class frmPNR
    Private Const VersionText As String = "Athens PNR Finisher (12/10/2018 12:31) "
    Private Const OptionPriceOptimizer As Boolean = False ' This blocks users' access to Price Optimiser until it is ready
    Private Structure PaxNamesPos
        Dim StartPos As Integer
        Dim EndPos As Integer
    End Structure

    Private WithEvents mobjPNR As New GDSReadPNR
    Private mSelectedPNRGDSCode As Utilities.EnumGDSCode
    Private mSelectedItnGDSCode As Utilities.EnumGDSCode

    Private mflgReadPNR As Boolean
    Private mintMaxString As Integer = 80

    Private mobjAirlinePoints As New AirlinePoints.Collection
    Private mobjAirlineNotes As New AirlineNotes.Collection
    Private mstrAirlineAlert As String
    Private mobjConditionalEntry As New ConditionalGDSEntry.Collection

    Private mobjCustomerSelected As Customers.CustomerItem
    Private mobjCustomers As New Customers.CustomerCollection

    Private mobjSubDepartmentSelected As SubDepartments.Item
    Private mobjCRMSelected As CRM.Item
    Private mobjVesselSelected As Vessels.Item
    Private mobjAveragePrice As New AveragePrice.Collection
    Private mobjGender As New PaxApisDB.GenderCollection
    Private mudtPaxNames() As PaxNamesPos

    Private mOSMPax As New osmPax.PaxCollection
    Private mOSMAgents As New osmVessels.EmailCollection
    Private mOSMAgentIndex As Integer = -1

    Private mItnFromDate As Date
    Private mItnToDate As Date

    Private mflgExpiryDateOK As Boolean
    Private mflgAPISUpdate As Boolean

    Private mflgLoading As Boolean
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click, cmdItnExit.Click
        Me.Close()
    End Sub
    Private Sub cmdPNRRead1APNR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPNRRead1APNR.Click
        Try
            mSelectedPNRGDSCode = Utilities.EnumGDSCode.Amadeus
            ClearForm()
            ReadPNR(Utilities.EnumGDSCode.Amadeus)
            ShowPriceOptimiser()
            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub ShowPriceOptimiser()

        If OptionPriceOptimizer Then
            If Not MySettings Is Nothing Then
                If MySettings.GDSPcc <> "" And MySettings.GDSUser <> "" Then
                    Dim pFrm As New frmPriceOptimiser(MySettings.GDSPcc, MySettings.GDSUser)
                    pFrm.Show()
                End If
            End If
        End If
    End Sub
    Private Sub ClearForm()

        Try

            mobjCustomerSelected = New Customers.CustomerItem
            mobjSubDepartmentSelected = New SubDepartments.Item
            mobjCRMSelected = New CRM.Item
            mobjVesselSelected = New Vessels.Item

            lblPNR.Text = ""
            lblPax.Text = ""
            lblSegs.Text = ""

            txtCustomer.Clear()
            txtSubdepartment.Clear()
            txtCRM.Clear()
            txtVessel.Clear()
            lstAirlineEntries.Items.Clear()

            lstVessels.Items.Clear()

            lstSubDepartments.Items.Clear()
            txtSubdepartment.Enabled = (lstSubDepartments.Items.Count > 0)

            lstCRM.Items.Clear()
            txtCRM.Enabled = (lstCRM.Items.Count > 0)

            txtReference.Clear()
            cmbDepartment.Items.Clear()
            cmbDepartment.Text = ""
            cmbDepartment.Tag = Nothing
            cmbBookedby.Items.Clear()
            cmbBookedby.Text = ""
            cmbBookedby.Tag = Nothing
            cmbReasonForTravel.Items.Clear()
            cmbReasonForTravel.Text = ""
            cmbReasonForTravel.Tag = Nothing
            cmbCostCentre.Items.Clear()
            cmbCostCentre.Text = ""
            cmbCostCentre.Tag = Nothing
            txtTrId.Clear()
            txtTrId.Tag = Nothing

            cmdPNRWrite.Enabled = False
            cmdPNRWriteWithDocs.Enabled = False
            cmdPNROnlyDocs.Enabled = False
            cmdPriceOptimiser.Enabled = False
            cmdPriceOptimiser.Visible = OptionPriceOptimizer

            mobjPNR.ExistingElements.Clear()

            mflgAPISUpdate = False
            mflgExpiryDateOK = False

            UtilitiesAPIS.APISPrepareGrid(dgvApis)

        Catch ex As Exception
            Throw New Exception("ClearForm()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub SetEnabled()

        Dim pProps As CustomProperties.Item

        Try
            ' read PNR and Exit are always enabled
            cmdPNRRead1APNR.Enabled = True
            cmdExit.Enabled = True

            cmdAdmin.Enabled = MySettings.Administrator
            cmdAdmin.Visible = cmdAdmin.Enabled

            ' customer based entries are enabled if a PNR has been read and there is data available
            txtCustomer.Enabled = mflgReadPNR And (lstCustomers.Items.Count > 0)
            lstCustomers.Enabled = mflgReadPNR And (lstCustomers.Items.Count > 0)
            cmdCostCentre.Enabled = mflgReadPNR And (lstCustomers.Items.Count > 0)

            txtSubdepartment.Enabled = mflgReadPNR And (lstSubDepartments.Items.Count > 0)
            lstSubDepartments.Enabled = mflgReadPNR And (lstSubDepartments.Items.Count > 0)

            txtCRM.Enabled = mflgReadPNR And (lstCRM.Items.Count > 0)
            lstCRM.Enabled = mflgReadPNR And (lstCRM.Items.Count > 0)

            txtVessel.Enabled = mflgReadPNR And (lstVessels.Items.Count > 0)
            lstVessels.Enabled = mflgReadPNR And (lstVessels.Items.Count > 0)

            ' the exception is the one time vessel which is always enabled for any PNR
            cmdOneTimeVessel.Enabled = mflgReadPNR

            ' Update is enabled if a PNR has been read and if mandatory fields have been entered
            cmdPNRWrite.Enabled = mflgReadPNR
            cmdPriceOptimiser.Enabled = (OptionPriceOptimizer And mflgReadPNR)

            ' Customer is always needed

            txtCustomer.BackColor = lstCustomers.BackColor
            txtSubdepartment.BackColor = lstCustomers.BackColor
            txtCRM.BackColor = lstCustomers.BackColor
            If Not mobjPNR.NewElements Is Nothing Then
                If mobjPNR.NewElements.CustomerCode.GDSCommand = "" Then
                    cmdPNRWrite.Enabled = False
                    txtCustomer.BackColor = Color.Red
                End If

                ' if subdepartments exist they are by default madatory
                If mobjPNR.NewElements.CustomerCode.GDSCommand <> "" And lstSubDepartments.Items.Count > 0 And mobjPNR.NewElements.SubDepartmentCode.GDSCommand = "" Then
                    cmdPNRWrite.Enabled = False
                    txtSubdepartment.BackColor = Color.Red
                End If

                ' the code above is complete validation but allow entry without CRM in any case
                If mobjPNR.NewElements.CustomerCode.GDSCommand <> "" And lstCRM.Items.Count > 0 And mobjPNR.NewElements.CRMCode.GDSCommand = "" Then
                    txtCRM.BackColor = Color.Pink
                End If
                If mobjPNR.NewElements.BookedBy.GDSCommand = "" And cmbBookedby.Enabled Then
                    pProps = CType(cmbBookedby.Tag, CustomProperties.Item)
                    If Not pProps Is Nothing AndAlso pProps.RequiredType = Utilities.CustomPropertyRequiredType.PropertyReqToSave Then
                        cmdPNRWrite.Enabled = False
                    End If
                End If
                If mobjPNR.NewElements.CostCentre.GDSCommand = "" And cmbCostCentre.Enabled Then
                    pProps = CType(cmbCostCentre.Tag, CustomProperties.Item)
                    If Not pProps Is Nothing AndAlso pProps.RequiredType = Utilities.CustomPropertyRequiredType.PropertyReqToSave Then
                        cmdPNRWrite.Enabled = False
                    End If
                End If
                If mobjPNR.NewElements.ReasonForTravel.GDSCommand = "" And cmbReasonForTravel.Enabled Then
                    pProps = CType(cmbReasonForTravel.Tag, CustomProperties.Item)
                    If Not pProps Is Nothing AndAlso pProps.RequiredType = Utilities.CustomPropertyRequiredType.PropertyReqToSave Then
                        cmdPNRWrite.Enabled = False
                    End If
                End If
                If mobjPNR.NewElements.TRId.GDSCommand = "" And txtTrId.Enabled Then
                    pProps = CType(txtTrId.Tag, CustomProperties.Item)
                    If Not pProps Is Nothing AndAlso pProps.RequiredType = Utilities.CustomPropertyRequiredType.PropertyReqToSave Then
                        cmdPNRWrite.Enabled = False
                    End If
                End If
            End If

            cmdPNRWriteWithDocs.Enabled = cmdPNRWrite.Enabled And mflgAPISUpdate
            cmdPNROnlyDocs.Enabled = mflgAPISUpdate And Not mobjPNR.NewPNR
            dgvApis.Enabled = True

            txtReference.Enabled = True

            lblBookedByHighlight.Enabled = (cmbBookedby.Enabled)
            lblDepartmentHighlight.Enabled = (cmbDepartment.Enabled)
            lblReasonForTravelHighLight.Enabled = (cmbReasonForTravel.Enabled)
            lblCostCentreHighlight.Enabled = (cmbCostCentre.Enabled)
            lblTRIDHighLight.Enabled = (txtTrId.Enabled)

            SetLabelColor(lblBookedByHighlight, cmbBookedby.Tag)
            SetLabelColor(lblDepartmentHighlight, cmbDepartment.Tag)
            SetLabelColor(lblReasonForTravelHighLight, cmbReasonForTravel.Tag)
            SetLabelColor(lblCostCentreHighlight, cmbCostCentre.Tag)
            SetLabelColor(lblTRIDHighLight, txtTrId.Tag)

        Catch ex As Exception
            Throw New Exception("SetEnabled()" & vbCrLf & ex.Message)
        End Try

    End Sub
    Private Sub SetLabelColor(ByRef TextLabel As Label, ByVal CustomProps As CustomProperties.Item)
        Try
            If TextLabel.Enabled Then
                If Not CustomProps Is Nothing AndAlso CustomProps.RequiredType = Utilities.CustomPropertyRequiredType.PropertyReqToSave Then
                    TextLabel.BackColor = Color.FromArgb(255, 128, 128)
                Else
                    TextLabel.BackColor = Color.Cyan
                End If
            Else
                TextLabel.BackColor = Color.Silver
            End If
        Catch ex As Exception
            Throw New Exception("SetLabelColor()" & vbCrLf & ex.Message)
        End Try
    End Sub
    Private Sub ReadPNR(ByVal GDSCode As Utilities.EnumGDSCode)
        Dim pDMI As String
        Try
            With mobjPNR
                mflgReadPNR = False
                Dim mGDSUser As New GDSUser(GDSCode)
                InitSettings(mGDSUser)
                SetupPCCOptions()
                pDMI = .Read(GDSCode)
                If .NumberOfPax = 0 And Not .IsGroup Then
                    Throw New Exception("Need passenger names")
                End If
                If pDMI <> "" Then
                    If MessageBox.Show("There is a problem with your itinerary. Do you want to cancel the PNR Finisher?" & vbCrLf & vbCrLf & pDMI, "Itinerary Check", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                        Throw New Exception("PNR Finisher cancelled because of itinerary check")
                    End If
                End If

                mflgReadPNR = True
                .PrepareNewGDSElements()
                lblPNR.Text = .PnrNumber
                If .IsGroup Then
                    lblPax.Text = "Group:" & .GroupName & " " & .GroupNamesCount
                Else
                    lblPax.Text = .PaxLeadName
                End If

                lblSegs.Text = .Itinerary
                If .Segments.AirlineAlert <> "" Then
                    MessageBox.Show(.Segments.AirlineAlert, "AIRLINE ALERT", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If

                Dim pFromDate As Date = DateAdd(DateInterval.Month, -3, Today)

                pFromDate = DateSerial(Year(pFromDate), Month(pFromDate), 1)

                mobjAveragePrice.SetValues(pFromDate, .Itinerary)
                PrepareAirlinePoints()
            End With
            DisplayCustomer()
            APISDisplayPax()

        Catch ex As Exception
            Throw New Exception("ReadPNR()" & vbCrLf & ex.Message)
        End Try
    End Sub
    Private Sub DisplayCustomer()

        Dim pstrCustomerCode As String
        Dim pintSubDepartment As Integer
        Dim pstrCRM As String
        Dim pstrVesselName As String
        Dim pstrVesselRegistration As String

        Try
            With mobjPNR.ExistingElements
                pstrCustomerCode = .CustomerCode.Key
                pintSubDepartment = If(IsNumeric(.SubDepartmentCode.Key), CInt(.SubDepartmentCode.Key), 0)
                pstrCRM = .CRMCode.Key
                pstrVesselName = .VesselName.Key
                pstrVesselRegistration = .VesselFlag.Key

                mobjPNR.NewElements.ClearCustomerElements()

                txtCustomer.Clear()
                txtSubdepartment.Clear()
                txtCRM.Clear()
                txtVessel.Clear()

                txtReference.Text = .Reference.Key
                txtSubdepartment.Text = .SubDepartmentCode.Key
                txtCRM.Text = .CRMCode.Key
            End With

            If pstrCustomerCode <> "" Then
                Dim pCustomer As New Customers.CustomerItem
                pCustomer.Load(pstrCustomerCode)
                txtCustomer.Text = pCustomer.Code
                If pintSubDepartment <> 0 Then
                    Dim pSub As New SubDepartments.Item
                    pSub.Load(pintSubDepartment)
                    txtSubdepartment.Text = pSub.Code & " " & pSub.Name
                End If
                If pstrCRM.Length > 0 Then
                    Dim pSub As New CRM.Item
                    pSub.Load(pstrCRM)
                    txtCRM.Text = pSub.Code & " " & pSub.Name
                End If

                If pstrVesselName <> "" Then
                    Dim pVessel As New Vessels.Item
                    If pVessel.Load(pstrCustomerCode, pstrVesselName) Then
                        mobjPNR.NewElements.VesselNameForPNR.Clear()
                        mobjPNR.NewElements.VesselFlagForPNR.Clear()
                        txtVessel.Text = pVessel.Name
                    Else
                        mobjPNR.NewElements.SetVesselForPNR(pstrVesselName, pstrVesselRegistration)
                        txtVessel.Text = mobjPNR.NewElements.VesselNameForPNR.TextRequested & " REG " & mobjPNR.NewElements.VesselFlagForPNR.TextRequested
                    End If
                End If

                DisplayOldCustomProperty(cmbBookedby, mobjPNR.ExistingElements.BookedBy)
                DisplayOldCustomProperty(cmbDepartment, mobjPNR.ExistingElements.Department)
                DisplayOldCustomProperty(cmbReasonForTravel, mobjPNR.ExistingElements.ReasonForTravel)
                DisplayOldCustomProperty(cmbCostCentre, mobjPNR.ExistingElements.CostCentre)
                DisplayOldCustomProperty(txtTrId, mobjPNR.ExistingElements.TRId)

                txtReference.Text = mobjPNR.ExistingElements.Reference.Key
                PrepareAirlinePoints()
            End If
        Catch ex As Exception
            Throw New Exception("DisplayCustomer()" & vbCrLf & ex.Message)
        End Try

    End Sub
    Private Sub DisplayOldCustomProperty(ByRef cmbList As ComboBox, ByVal Item As GDSExisting.Item)
        Try
            If Item.Key <> "" Then
                If cmbList.DropDownStyle = ComboBoxStyle.DropDown Then
                    If Item.Key <> "" Then
                        cmbList.Text = Item.Key
                    End If
                Else
                    For i As Integer = 0 To cmbList.Items.Count - 1
                        If Item.Key.ToUpper = cmbList.Items(i).ToString.ToUpper Then
                            cmbList.SelectedIndex = i
                            Exit For
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Throw New Exception("DisplayOldCustomProperty(ByRef cmbList As ComboBox, ByVal Item As PNR_Finisher.GDSExisting.Item)" & vbCrLf & ex.Message)
        End Try

    End Sub
    Private Sub DisplayOldCustomProperty(ByRef txtText As TextBox, ByVal Item As GDSExisting.Item)
        Try
            txtText.Text = Item.Key
        Catch ex As Exception
            Throw New Exception("DisplayOldCustomProperty(ByRef txtText As TextBox, ByVal Item As GDSExisting.Item)" & vbCrLf & ex.Message)
        End Try

    End Sub
    Private Sub DisplayOldCustomProperty(ByRef cmbList As ComboBox, ByVal Item As String)
        Try
            If Item <> "" Then
                If cmbList.DropDownStyle = ComboBoxStyle.DropDown Then
                    cmbList.Text = Item
                Else
                    For i As Integer = 0 To cmbList.Items.Count - 1
                        If cmbList.Items(i).ToString.ToUpper.StartsWith(Item.ToUpper) Then
                            cmbList.SelectedIndex = i
                            Exit For
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Throw New Exception("DisplayOldCustomProperty(ByRef cmbList As ComboBox, ByVal Item As String)" & vbCrLf & ex.Message)
        End Try

    End Sub
    Private Sub PrepareAirlinePoints()
        Try
            Dim pFound As Boolean = False
            lstAirlineEntries.Items.Clear()

            For Each pSeg As GDSSeg.GDSSegItem In mobjPNR.Segments.Values
                mobjAirlinePoints.Load(mobjCustomerSelected.ID, pSeg.Airline, mobjPNR.GDSCode)
                For Each pItem As AirlinePoints.Item In mobjAirlinePoints.Values
                    pFound = False
                    For i As Integer = 0 To lstAirlineEntries.Items.Count - 1
                        If lstAirlineEntries.Items(i).ToString = pItem.ToString Then
                            pFound = True
                            Exit For
                        End If
                    Next
                    If Not pFound Then
                        lstAirlineEntries.Items.Add(pItem, True)
                    End If
                Next
            Next

            If mflgReadPNR Then
                For Each pSeg As GDSSeg.GDSSegItem In mobjPNR.Segments.Values
                    mobjAirlineNotes.Load(pSeg.Airline, mobjPNR.GDSCode)
                    For Each pItem As AirlineNotes.Item In mobjAirlineNotes.Values
                        With pItem
                            If Not .Seaman Or Not mobjVesselSelected Is Nothing Then
                                Dim pGDSText As String = .GDSText

                                If pGDSText.Contains("<?VESSEL NAME>") Then
                                    If Not mobjVesselSelected Is Nothing Then
                                        If mobjVesselSelected.Name Is Nothing Then
                                            pGDSText = pGDSText.Replace("<?VESSEL NAME>", mobjVesselSelected.Name)
                                        Else
                                            pGDSText = pGDSText.Replace("<?VESSEL NAME>", mobjVesselSelected.Name.Replace("(", "-").Replace(")", "-").Replace("&", "-"))
                                        End If
                                    End If
                                End If

                                If pGDSText.Contains("<?VESSEL REGISTRATION>") Then
                                    If Not mobjVesselSelected Is Nothing Then
                                        If mobjVesselSelected.Flag Is Nothing Then
                                            pGDSText = pGDSText.Replace("<?VESSEL REGISTRATION>", mobjVesselSelected.Flag)
                                        Else
                                            pGDSText = pGDSText.Replace("<?VESSEL REGISTRATION>", mobjVesselSelected.Flag.Replace("(", "-").Replace(")", "-").Replace("&", "-"))
                                        End If
                                    End If
                                End If

                                If pGDSText.Contains("<?NBR OF PSGRS>") Then
                                    pGDSText = pGDSText.Replace("<?NBR OF PSGRS>", CStr(mobjPNR.NumberOfPax))
                                End If

                                If pGDSText.Contains("<?Segment selection>") Then
                                    pGDSText = pGDSText.Replace("<?Segment selection>", CStr(pSeg.ElementNo))
                                End If

                                Dim pGDSCommand As String
                                If .GDSElement.StartsWith("R") Then
                                    pGDSCommand = .GDSElement & " " & .AirlineCode & " " & pGDSText
                                ElseIf .GDSElement.StartsWith("S") Then
                                    pGDSCommand = .GDSElement & "-" & pGDSText
                                Else
                                    pGDSCommand = .GDSElement & " " & pGDSText
                                End If
                                pFound = False
                                For i As Integer = 0 To lstAirlineEntries.Items.Count - 1
                                    If lstAirlineEntries.Items(i).ToString = pGDSCommand Then
                                        pFound = True
                                        Exit For
                                    End If
                                Next
                                If Not pFound Then
                                    lstAirlineEntries.Items.Add(pGDSCommand, True)
                                End If

                            End If
                        End With
                    Next
                Next

                If Not mobjCustomerSelected Is Nothing And Not mobjVesselSelected Is Nothing Then
                    mobjConditionalEntry.Load(MySettings.PCCBackOffice, mobjCustomerSelected.ID, mobjVesselSelected.Name)
                    For Each pItem As ConditionalGDSEntry.Item In mobjConditionalEntry.Values
                        Dim pGDSCommand As String = ""
                        If mSelectedPNRGDSCode = Utilities.EnumGDSCode.Amadeus Then
                            pGDSCommand = pItem.ConditionalEntry1A
                        ElseIf mSelectedPNRGDSCode = Utilities.EnumGDSCode.Galileo Then
                            pGDSCommand = pItem.ConditionalEntry1G
                        Else
                            pGDSCommand = ""
                        End If
                        If pGDSCommand <> "" Then
                            pFound = False
                            For i As Integer = 0 To lstAirlineEntries.Items.Count - 1
                                If lstAirlineEntries.Items(i).ToString = pGDSCommand Then
                                    pFound = True
                                    Exit For
                                End If
                            Next
                            If Not pFound Then
                                lstAirlineEntries.Items.Add(pGDSCommand, True)
                            End If

                        End If

                    Next
                End If
            End If

        Catch ex As Exception
            Throw New Exception("PrepareAirlinePoints()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub frmPNR_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            mflgLoading = True
            Text = VersionText
            dgvApis.VirtualMode = False
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            mflgLoading = False
        End Try

    End Sub
    Private Sub SetupPCCOptions()

        Try

            mflgLoading = True
            Dim pText As String = ""
            Text = VersionText
            If MySettings.GDSPcc <> "" And MySettings.GDSUser <> "" Then
                pText &= MySettings.GDSPcc & " " & MySettings.GDSUser
                If mSelectedPNRGDSCode = Utilities.EnumGDSCode.Amadeus Then
                    lblPNRAmadeus.Text = pText
                ElseIf mSelectedPNRGDSCode = Utilities.EnumGDSCode.Galileo Then
                    lblPNRGalileo.Text = pText
                End If
                If MySettings.GDSPcc <> RequestedPCC Or MySettings.GDSUser <> RequestedUser Then
                    pText &= " (Jump in to " & RequestedPCC & " as user " & RequestedUser & ")"
                End If
                If MySettings.GDSPcc <> MyHomeSettings.GDSPcc Or MySettings.GDSUser <> MyHomeSettings.GDSUser Then
                    pText &= " (Jump in from " & MyHomeSettings.GDSPcc & " user " & MyHomeSettings.GDSUser & ")"
                End If

            Else
                Throw New Exception("No GDS signed in")
            End If

            If CheckOptions() Then
                ' finisher tab
                mflgReadPNR = False
                ClearForm()
                SetEnabled()
                PrepareForm()
                UtilitiesAPIS.APISPrepareGrid(dgvApis)

                ' itinerary tab
                LoadRemarks()
                If MySettings.AirportName = 0 Then
                    optItnAirportCode.Checked = True
                ElseIf MySettings.AirportName = 1 Then
                    optItnAirportname.Checked = True
                ElseIf MySettings.AirportName = 2 Then
                    optItnAirportBoth.Checked = True
                ElseIf MySettings.AirportName = 3 Then
                    optItnAirportCityName.Checked = True
                ElseIf MySettings.AirportName = 4 Then
                    optItnAirportCityBoth.Checked = True
                End If

                Select Case MySettings.FormatStyle
                    Case Utilities.EnumItnFormat.DefaultFormat
                        optItnFormatDefault.Checked = True
                    Case Utilities.EnumItnFormat.Plain
                        optItnFormatPlain.Checked = True
                    Case Utilities.EnumItnFormat.SeaChefs
                        optItnFormatSeaChefs.Checked = True
                    Case Utilities.EnumItnFormat.SeaChefsWithCode
                        optItnFormatSeaChefsWith3LetterCode.Checked = True
                    Case Utilities.EnumItnFormat.Euronav
                        optItnFormatEuronav.Checked = True
                End Select
                SetITNEnabled(True)

                chkItnVessel.Checked = MySettings.ShowVessel
                chkItnClass.Checked = MySettings.ShowClassOfService
                chkItnAirlineLocator.Checked = MySettings.ShowAirlineLocator
                chkItnTickets.Checked = MySettings.ShowTickets
                chkItnPaxSegPerTicket.Checked = MySettings.ShowPaxSegPerTkt
                chkItnSeating.Checked = MySettings.ShowSeating
                chkItnStopovers.Checked = MySettings.ShowStopovers
                chkItnTerminal.Checked = MySettings.ShowTerminal
                chkItnFlyingTime.Checked = MySettings.ShowFlyingTime
                chkItnCostCentre.Checked = MySettings.ShowCostCentre

                chkItnElecItemsBan.Checked = MySettings.ShowBanElectricalEquipment
                chkItnBrazilText.Checked = MySettings.ShowBrazilText
                chkItnUSAText.Checked = MySettings.ShowUSAText

                cmdItn1AReadPNR.Enabled = False
                cmdItn1AReadQueue.Enabled = False
                cmdItn1GReadPNR.Enabled = False
                cmdItn1GReadQueue.Enabled = False
                optItnFormatMSReport.Enabled = cmdItn1AReadQueue.Enabled
            Else
                Throw New Exception("User not authorized for this PCC")
            End If
        Catch ex As Exception
        Finally
            mflgLoading = False
        End Try

    End Sub
    Private Sub LoadRemarks()

        Try
            With lstItnRemarks.Items()
                .Clear()
                .Add("SEAMAN FARE DOES NOT PERMIT UPGRADING")
                .Add("SEAMAN FARE DOES NOT PERMIT PRESEATING BUT WITH UPGRADING")
                .Add("SEAMAN FARE WITH UPGRADING")
                .Add("SEAMAN FARE WITH UPGRADING AND PRESEATING")
                .Add("PLEASE CHECK BELOW AND ADVISE IF OK TO ISSUE")
                .Add("ALL BOOKINGS ON TIME LIMIT")
                .Add("ALL FARES ON TODAY'S RATE/ADVANCE PURCHASE")
            End With

        Catch ex As Exception
            Throw New Exception("LoadRemarks()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Function CheckOptions() As Boolean
        Try
            With MySettings
                While Not .isValid
                    If MessageBox.Show("Please enter your details", "Options Missing", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.Cancel Then
                        Return False
                    End If
                    ShowOptionsForm()
                End While
                Return True
            End With
        Catch ex As Exception
            Throw New Exception("CheckOptions()" & vbCrLf & ex.Message)
        End Try

    End Function

    Private Sub PrepareForm()

        Try
            PrepareLists()
            PopulateCustomerList("")
        Catch ex As Exception
            Throw New Exception("PrepareForms()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub PrepareLists()

        Try
            lstCustomers.Items.Clear()

            lstSubDepartments.Items.Clear()
            mobjSubDepartmentSelected = Nothing

            lstCRM.Items.Clear()
            mobjCRMSelected = Nothing

            lstVessels.Items.Clear()
            mobjVesselSelected = Nothing

            cmdPNRWrite.Enabled = False
            cmdPNRWriteWithDocs.Enabled = False
            cmdPNROnlyDocs.Enabled = False

        Catch ex As Exception
            Throw New Exception("PrepareLists()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub PopulateCustomerList(ByVal SearchString As String)

        Try
            mobjCustomers.Load(SearchString)

            lstCustomers.Items.Clear()
            For Each pCustomer As Customers.CustomerItem In mobjCustomers.Values
                If SearchString = "" Or pCustomer.ToString.ToUpper.Contains(SearchString.ToUpper) Then
                    lstCustomers.Items.Add(pCustomer)
                End If
            Next

            If lstCustomers.Items.Count = 1 Then
                Try
                    mflgLoading = True
                    Dim pCust As Customers.CustomerItem = CType(lstCustomers.Items(0), Customers.CustomerItem)
                    SelectCustomer(pCust)
                    txtCustomer.Text = lstCustomers.Items(0).ToString
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                Finally
                    mflgLoading = False
                End Try
            End If
        Catch ex As Exception
            Throw New Exception("PopulateCustomerList()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub PopulateSubdepartmentsList(ByVal SearchString As String)

        Try
            Dim pobjSubDepartments As New SubDepartments.Collection

            If SearchString = "" Then
                mobjSubDepartmentSelected = Nothing
                mobjPNR.NewElements.SetItem(mobjSubDepartmentSelected)
            End If
            lstSubDepartments.Items.Clear()

            If Not mobjCustomerSelected Is Nothing Then
                pobjSubDepartments.Load(mobjCustomerSelected.ID)

                For Each pSubDepartment As SubDepartments.Item In pobjSubDepartments.Values
                    If SearchString = "" Or pSubDepartment.ToString.ToUpper.Contains(SearchString.ToUpper) Then
                        lstSubDepartments.Items.Add(pSubDepartment)
                    End If
                Next

                If lstSubDepartments.Items.Count = 1 Then
                    Try
                        mflgLoading = True
                        SelectSubDepartment(lstSubDepartments.Items(0))
                        txtSubdepartment.Text = lstSubDepartments.Items(0).ToString
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    Finally
                        mflgLoading = False
                    End Try
                End If
            End If
        Catch ex As Exception
            Throw New Exception("PopulateSubdepartmentsList()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub PopulateCRMList(ByVal SearchString As String)

        Try
            Dim pobjCRM As New CRM.Collection

            If SearchString = "" Then
                mobjCRMSelected = Nothing
                mobjPNR.NewElements.SetItem(mobjCRMSelected)
            End If
            lstCRM.Items.Clear()

            If Not mobjCustomerSelected Is Nothing Then
                pobjCRM.Load(mobjCustomerSelected.ID)

                For Each pCRM As CRM.Item In pobjCRM.Values
                    If SearchString = "" Or pCRM.ToString.ToUpper.Contains(SearchString.ToUpper) Then
                        lstCRM.Items.Add(pCRM)
                    End If
                Next
                If mobjPNR.NewElements.CRMCode.TextRequested <> "" And lstCRM.Items.Count = 1 Then
                    Try
                        mflgLoading = True
                        SelectCRM(lstCRM.Items(0))
                        txtCRM.Text = lstCRM.Items(0).ToString
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    Finally
                        mflgLoading = False
                    End Try
                End If
            End If
        Catch ex As Exception
            Throw New Exception("PopulateCRMList()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub PopulateVesselsList()

        Try
            Dim pobjVessels As New Vessels.Collection

            lstVessels.Items.Clear()

            If Not mobjCustomerSelected Is Nothing Then

                pobjVessels.Load(mobjCustomerSelected.ID)

                For Each pVessel As Vessels.Item In pobjVessels.Values
                    If mobjPNR.NewElements.VesselName.TextRequested = "" Or pVessel.ToString.ToUpper.Contains(mobjPNR.NewElements.VesselName.TextRequested.ToUpper) Then
                        lstVessels.Items.Add(pVessel)
                    End If
                Next
                If lstVessels.Items.Count = 1 Then
                    Try
                        mflgLoading = True
                        SelectVessel(lstVessels.Items(0))
                        txtVessel.Text = lstVessels.Items(0).ToString
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    Finally
                        mflgLoading = False
                    End Try
                End If
            End If
        Catch ex As Exception
            Throw New Exception("PopulateVesselsList()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub PopulateCustomProperties()

        Try
            cmbBookedby.Items.Clear()
            cmbDepartment.Items.Clear()
            cmbReasonForTravel.Items.Clear()
            cmbCostCentre.Items.Clear()
            cmbBookedby.Enabled = False
            cmbDepartment.Enabled = False
            cmbReasonForTravel.Enabled = False
            cmbCostCentre.Enabled = False
            txtTrId.Enabled = False

            If Not mobjCustomerSelected Is Nothing Then
                For Each pProp As CustomProperties.Item In mobjCustomerSelected.CustomerProperties.Values
                    If pProp.CustomPropertyID = Utilities.EnumCustomPropertyID.BookedBy Then
                        PrepareCustomProperty(cmbBookedby, pProp)
                    ElseIf pProp.CustomPropertyID = Utilities.EnumCustomPropertyID.Department Then
                        PrepareCustomProperty(cmbDepartment, pProp)
                    ElseIf pProp.CustomPropertyID = Utilities.EnumCustomPropertyID.ReasonFortravel Then
                        PrepareCustomProperty(cmbReasonForTravel, pProp)
                    ElseIf pProp.CustomPropertyID = Utilities.EnumCustomPropertyID.CostCentre Then
                        PrepareCustomProperty(cmbCostCentre, pProp)
                    ElseIf pProp.CustomPropertyID = Utilities.EnumCustomPropertyID.TRId Then
                        PrepareCustomProperty(txtTrId, pProp)
                    End If
                Next
            End If
        Catch ex As Exception
            Throw New Exception("PopulateCustomproperties()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub PrepareCustomProperty(ByRef cmbCombo As ComboBox, ByRef pProp As CustomProperties.Item)

        Try
            cmbCombo.Enabled = True
            cmbCombo.Tag = pProp
            If pProp.LimitToLookup Then
                cmbCombo.DropDownStyle = ComboBoxStyle.DropDownList
            Else
                cmbCombo.DropDownStyle = ComboBoxStyle.DropDown
            End If
            cmbCombo.AutoCompleteSource = AutoCompleteSource.ListItems
            cmbCombo.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            For i As Integer = 0 To pProp.ValuesCount - 1
                cmbCombo.Items.Add(pProp.Value(i))
            Next
        Catch ex As Exception
            Throw New Exception("PrepareCustomProperty()" & vbCrLf & ex.Message)
        End Try

    End Sub
    Private Sub PrepareCustomProperty(ByRef txtText As TextBox, ByRef pProp As CustomProperties.Item)

        Try
            txtText.Enabled = True
            txtText.Tag = pProp
        Catch ex As Exception
            Throw New Exception("PrepareCustomProperty()" & vbCrLf & ex.Message)
        End Try

    End Sub
    Private Sub txtCustomer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustomer.TextChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    PopulateCustomerList(txtCustomer.Text)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub SelectCustomer(ByVal pCustomer As Customers.CustomerItem)

        Try
            'TODO
            mobjPNR.NewElements.ClearCustomerElements()
            mobjAirlinePoints.Clear()
            mobjAirlineNotes.Clear()
            mobjConditionalEntry.Clear()
            mobjCustomerSelected = pCustomer
            txtCustomer.Text = pCustomer.ToString
            mobjPNR.NewElements.SetItem(mobjCustomerSelected)

            txtSubdepartment.Clear()
            lstSubDepartments.Items.Clear()
            mobjSubDepartmentSelected = Nothing

            txtCRM.Clear()
            lstCRM.Items.Clear()
            mobjCRMSelected = Nothing

            txtVessel.Clear()
            lstVessels.Items.Clear()
            mobjVesselSelected = Nothing

            txtReference.Clear()

            cmbBookedby.Text = ""
            cmbDepartment.Text = ""
            txtTrId.Clear()

            If mobjCustomerSelected.HasVessels Then
                PopulateVesselsList()
            End If

            If mobjCustomerSelected.HasDepartments Then
                PopulateSubdepartmentsList("")
            End If

            PopulateCRMList("")
            PopulateCustomProperties()
            PrepareAirlinePoints()

            SetEnabled()

            If pCustomer.Alert <> "" Then

                MessageBox.Show(pCustomer.Alert, pCustomer.Code & " " & pCustomer.Name, MessageBoxButtons.OK, MessageBoxIcon.Error)

            End If
        Catch ex As Exception
            Throw New Exception("SelectCustomer()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub txtSubdepartment_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSubdepartment.TextChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    PopulateSubdepartmentsList(txtSubdepartment.Text)
                End If
            End If

            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub txtCRM_TextChanged(sender As Object, e As EventArgs) Handles txtCRM.TextChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    PopulateCRMList(txtCRM.Text)
                End If
            End If

            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub SelectSubDepartment(ByVal pSubDepartment As SubDepartments.Item)

        Try
            mobjSubDepartmentSelected = pSubDepartment
            txtSubdepartment.Text = pSubDepartment.ToString
            mobjPNR.NewElements.SetItem(mobjSubDepartmentSelected)

            SetEnabled()
        Catch ex As Exception
            Throw New Exception("SelectSubDepartment()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub SelectCRM(ByVal pCRM As CRM.Item)

        Try
            mobjCRMSelected = pCRM
            txtCRM.Text = pCRM.ToString
            mobjPNR.NewElements.SetItem(mobjCRMSelected)

            SetEnabled()
            If pCRM.Alert <> "" Then
                MessageBox.Show(pCRM.Alert, pCRM.Code & " " & pCRM.Name, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            Throw New Exception("SelectCRM()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub SelectVessel(ByVal pVessel As Vessels.Item)

        Try
            mobjVesselSelected = pVessel
            txtVessel.Text = pVessel.ToString
            mobjPNR.NewElements.SetItem(mobjVesselSelected)
            PrepareAirlinePoints()
            SetEnabled()
        Catch ex As Exception
            Throw New Exception("SelectVessel()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub lstCustomers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstCustomers.SelectedIndexChanged

        Try
            If lstCustomers.SelectedIndex >= 0 Then
                mflgLoading = True
                Dim pCust As Customers.CustomerItem = CType(lstCustomers.SelectedItem, Customers.CustomerItem)
                SelectCustomer(pCust)
                txtCustomer.Text = lstCustomers.SelectedItem.ToString
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            mflgLoading = False
        End Try

    End Sub

    Private Sub lstSubDepartments_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstSubDepartments.SelectedIndexChanged

        Try
            If Not lstSubDepartments.SelectedItem Is Nothing Then
                mflgLoading = True
                SelectSubDepartment(lstSubDepartments.SelectedItem)
                txtSubdepartment.Text = lstSubDepartments.SelectedItem.ToString
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            mflgLoading = False
        End Try

    End Sub

    Private Sub lstCRM_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstCRM.SelectedIndexChanged

        Try
            If Not lstCRM.SelectedItem Is Nothing Then
                mflgLoading = True
                SelectCRM(lstCRM.SelectedItem)
                txtCRM.Text = lstCRM.SelectedItem.ToString
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            mflgLoading = False
        End Try

    End Sub

    Private Sub txtVessel_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVessel.TextChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    mobjPNR.NewElements.SetVesselForPNR("", "")
                    mobjPNR.NewElements.VesselName.SetText(txtVessel.Text)
                    PopulateVesselsList()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub lstVessels_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstVessels.SelectedIndexChanged

        Try
            If lstVessels.SelectedIndex >= 0 Then
                mflgLoading = True
                SelectVessel(lstVessels.SelectedItem)
                txtVessel.Text = lstVessels.SelectedItem.ToString
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            mflgLoading = False
        End Try

    End Sub

    Private Sub cmdPNRWrite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPNRWrite.Click

        Try
            PNRWrite(True, False)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmdPNRWriteWithDocs_Click(sender As Object, e As EventArgs) Handles cmdPNRWriteWithDocs.Click
        Try
            PNRWrite(True, True)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub cmdPNROnlyDocs_Click(sender As Object, e As EventArgs) Handles cmdPNROnlyDocs.Click
        Try
            PNRWrite(False, True)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Function PNRWrite(ByVal WritePNR As Boolean, ByVal WriteDocs As Boolean) As String

        Try
            PNRWrite = UpdatePNR(WritePNR, WriteDocs)
            If mSelectedPNRGDSCode = Utilities.EnumGDSCode.Galileo And PNRWrite.Length > 6 Then
                MessageBox.Show("Please enter *R or *ALL in Galileo to show the PNR" & If(PNRWrite <> "", vbCrLf & vbCrLf & "PNR: " & PNRWrite, ""), "Galileo Information for PNR")
            End If
            mflgReadPNR = False
            ClearForm()
            SetEnabled()
        Catch ex As Exception
            Throw New Exception("PNRWrite(" & WritePNR & ", " & WriteDocs & ")" & vbCrLf & ex.Message)
        End Try

    End Function

    Private Function UpdatePNR(ByVal WritePNR As Boolean, ByVal WriteDocs As Boolean) As String
        Try
            UpdatePNR = mobjPNR.SendAllGDSEntries(WritePNR, WriteDocs, mflgExpiryDateOK, dgvApis, lstAirlineEntries)
            Dim pPNR As String = mobjPNR.PnrNumber
            Dim pNewEntry = False
            If pPNR = "New PNR" Or pPNR = "" Then
                If UpdatePNR.LastIndexOf(" ") > -1 Then
                    pPNR = UpdatePNR.Substring(UpdatePNR.LastIndexOf(" ")).Trim
                ElseIf UpdatePNR.Length = 6 Then
                    pPNR = UpdatePNR
                End If
                pNewEntry = True
            End If
            Dim pClient As String = mobjPNR.ClientCode
            If pClient = "" Then
                pClient = mobjPNR.NewElements.CustomerCode.TextRequested
            End If
            If pPNR <> "" Then
                Dim pTrans As New PNRFinisherTransactions
                pTrans.UpdateTransactions(pPNR, MySettings.GDSAbbreviation, MySettings.GDSPcc, MySettings.GDSUser, Now, mobjPNR.Passengers.AllPassengers, mobjPNR.Segments.FullItinerary, "", pClient, pNewEntry)
            End If
        Catch ex As Exception
            Throw New Exception("UpdatePNR()" & vbCrLf & ex.Message)
        End Try

    End Function

    Private Sub ShowOptionsForm()
        Try
            Dim pFrm As New frmShowOptions
            pFrm.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub llbOptions_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbOptions.LinkClicked

        Try
            ShowOptionsForm()

            If Not CheckOptions() Then
                Me.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmdOneTimeVessel_Click(sender As Object, e As EventArgs) Handles cmdOneTimeVessel.Click

        Try
            Dim pFrm As New frmVesselForPNR

            If pFrm.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
                With mobjPNR.NewElements
                    .SetVesselForPNR(pFrm.VesselName, pFrm.Registration)
                    mflgLoading = True
                    txtVessel.Text = .VesselNameForPNR.TextRequested & If(.VesselFlagForPNR.TextRequested <> "", " REG " & .VesselFlagForPNR.TextRequested, "")
                End With
            End If
            pFrm.Dispose()
            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            mflgLoading = False
        End Try
    End Sub

    Private Sub cmbBookedby_TextChanged(sender As Object, e As EventArgs) Handles cmbBookedby.TextChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    mobjPNR.NewElements.SetBookedBy(cmbBookedby.Text)
                End If
            End If
            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmbReasonForTravel_TextChanged(sender As Object, e As EventArgs) Handles cmbReasonForTravel.TextChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    mobjPNR.NewElements.SetReasonForTravel(cmbReasonForTravel.Text)
                End If
            End If
            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmbCostCentre_TextChanged(sender As Object, e As EventArgs) Handles cmbCostCentre.TextChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    mobjPNR.NewElements.SetCostCentre(cmbCostCentre.Text)
                End If
            End If
            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub txtTrId_TextChanged(sender As Object, e As EventArgs) Handles txtTrId.TextChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    mobjPNR.NewElements.SetTRId(txtTrId.Text)
                End If
            End If
            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub txtReference_TextChanged(sender As Object, e As EventArgs) Handles txtReference.TextChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    mobjPNR.NewElements.SetReference(txtReference.Text)
                End If
            End If
            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmbDepartment_TextChanged(sender As Object, e As EventArgs) Handles cmbDepartment.TextChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    mobjPNR.NewElements.SetDepartment(cmbDepartment.Text)
                End If
            End If
            SetEnabled()
        Catch ex As Exception
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmdItn1AReadPNR_Click(sender As Object, e As EventArgs) Handles cmdItn1AReadPNR.Click

        mSelectedItnGDSCode = Utilities.EnumGDSCode.Amadeus
        Try
            Cursor = System.Windows.Forms.Cursors.WaitCursor
            Dim mGDSUser As New GDSUser(Utilities.EnumGDSCode.Amadeus)
            InitSettings(mGDSUser)
            SetupPCCOptions()
            lblItnPNRCounter.Text = ""
            ProcessRequestedPNRs(txtItnPNR)
            CopyItinToClipboard()
            cmdItnRefresh.Enabled = False
            cmdItnFormatOSMLoG.Enabled = True
            Cursor = Cursors.Default
            MessageBox.Show("Ready", "Read PNR", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmdItnReadQueue_Click(sender As Object, e As EventArgs) Handles cmdItn1AReadQueue.Click

        Try
            lblItnPNRCounter.Text = ""
            If optItnFormatMSReport.Checked Then
                If ItnReadFromToDates() = Windows.Forms.DialogResult.Cancel Then
                    Exit Sub
                End If
            End If
            Cursor = System.Windows.Forms.Cursors.WaitCursor
            txtItnPNR.Text = mobjPNR.RetrievePNRsFromQueue(txtItnPNR.Text)
            mSelectedItnGDSCode = Utilities.EnumGDSCode.Amadeus
            Dim mGDSUser As New GDSUser(mSelectedItnGDSCode)
            InitSettings(mGDSUser)
            SetupPCCOptions()
            ProcessRequestedPNRs(txtItnPNR)
            CopyItinToClipboard()
            cmdItnRefresh.Enabled = False
            cmdItnFormatOSMLoG.Enabled = False
            Cursor = Cursors.Default
            MessageBox.Show("Ready", "Read PNR", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Function ItnReadFromToDates() As DialogResult

        Try
            Dim pFrm As New frmItnMSReportDates

            ItnReadFromToDates = pFrm.ShowDialog()
            mItnFromDate = pFrm.FromDate
            mItnToDate = pFrm.ToDate
            pFrm.Close()
        Catch ex As Exception
            Throw New Exception(message:=$"ItnReadFromToDates(){vbCrLf}{ex.Message}")
        End Try

    End Function
    Private Sub ProcessRequestedPNRs(ByVal RefreshOnly As Boolean)

        Try

            If Not RefreshOnly Then
                ReDim mudtPaxNames(0)
                readGDS("")
            End If
            If optItnFormatEuronav.Checked Then
                webItnDoc.Width = rtbItnDoc.Width
                webItnDoc.Height = rtbItnDoc.Height
                webItnDoc.Left = rtbItnDoc.Left
                webItnDoc.Top = rtbItnDoc.Top
                webItnDoc.Visible = True
                webItnDoc.BringToFront()
                rtbItnDoc.Visible = False
                webItnDoc.DocumentText = makeWebHead() & makeWebDoc() & makeWebClose()
            Else
                webItnDoc.Visible = False
                rtbItnDoc.Visible = True
                rtbItnDoc.Clear()
                makeRTBDoc()
            End If
        Catch ex As Exception
            Throw New Exception("ProcessRequestedPNRs(RefreshOnly)" & vbCrLf & ex.Message)
        End Try

    End Sub
    Private Sub ProcessRequestedPNRs(ByVal txtPNR As TextBox)

        Try
            Dim pPNR() As String = txtPNR.Text.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Dim pPNRsOutsideRange As New System.Text.StringBuilder
            Dim pWebItn As String = ""

            pPNRsOutsideRange.Clear()

            ReDim mudtPaxNames(0)
            If optItnFormatEuronav.Checked Then
                webItnDoc.Width = rtbItnDoc.Width
                webItnDoc.Height = rtbItnDoc.Height
                webItnDoc.Left = rtbItnDoc.Left
                webItnDoc.Top = rtbItnDoc.Top
                webItnDoc.Visible = True
                rtbItnDoc.Visible = False
                pWebItn = makeWebHead()
            Else
                webItnDoc.Visible = False
                rtbItnDoc.Visible = True
                rtbItnDoc.Clear()
            End If
            For i As Integer = pPNR.GetLowerBound(0) To pPNR.GetUpperBound(0)
                lblItnPNRCounter.Text = i + 1 & " of " & pPNR.GetUpperBound(0) + 1
                If pPNR(i).Trim <> "" Then
                    readGDS(pPNR(i).Trim)
                    If Not optItnFormatMSReport.Checked Or (mobjPNR.LastSegment.DepartureDate >= mItnFromDate And mobjPNR.LastSegment.DepartureDate <= mItnToDate) Then
                        If optItnFormatEuronav.Checked Then
                            pWebItn &= makeWebDoc()
                        Else
                            makeRTBDoc()
                        End If
                    Else
                        Dim pItnRTBDoc As New ItnRTBDoc(mobjPNR, mintMaxString, lstItnRemarks)
                        pPNRsOutsideRange.Append(pItnRTBDoc.MakeRTBMSReportOutsiderange)
                    End If
                End If
            Next
            If optItnFormatEuronav.Checked Then
                pWebItn &= makeWebClose()
                webItnDoc.DocumentText = pWebItn
            Else
                If pPNRsOutsideRange.Length > 0 Then
                    rtbItnDoc.Text &= vbCrLf & "OUTSIDE DATE RANGE" & vbCrLf
                    rtbItnDoc.Text &= pPNRsOutsideRange.ToString
                End If
            End If
        Catch ex As Exception
            Throw New Exception("ProcessRequestedPNRs(txtPNR)" & vbCrLf & ex.Message)
        End Try

    End Sub
    Private Sub makeRTBDoc()

        Dim pItnRTBDoc As New ItnRTBDoc(mobjPNR, mintMaxString, lstItnRemarks)

        If optItnFormatMSReport.Checked Then
            If optItnFormatMSReport.Checked AndAlso rtbItnDoc.TextLength = 0 Then
                rtbItnDoc.Text &= pItnRTBDoc.RTBMSReportHeader(mItnFromDate.ToShortDateString, mItnToDate.ToShortDateString)
            End If
            rtbItnDoc.Text &= pItnRTBDoc.MakeRTBMSReport.ToString
        Else
            Dim pFont As Font = rtbItnDoc.SelectionFont
            Dim pStart As Integer = rtbItnDoc.Text.Length + 1
            rtbItnDoc.Text &= pItnRTBDoc.RTBDocPassengers
            Dim pEnd As Integer = rtbItnDoc.Text.Length
            rtbItnDoc.Select(pStart, pEnd)
            rtbItnDoc.SelectionFont = New Font(pFont, FontStyle.Bold)
            rtbItnDoc.Text &= pItnRTBDoc.makeRTBDoc
        End If

    End Sub
    Private Sub PaxNamesToBold()

        Try
            Dim pFont As Font = rtbItnDoc.SelectionFont

            For i As Integer = 1 To mudtPaxNames.GetUpperBound(0)
                rtbItnDoc.Select(mudtPaxNames(i).StartPos - 1, mudtPaxNames(i).EndPos - mudtPaxNames(i).StartPos + 1)
                rtbItnDoc.SelectionFont = New Font(pFont, FontStyle.Bold)
            Next
        Catch ex As Exception
            Throw New Exception("PaxNamesToBold()" & vbCrLf & ex.Message)
        End Try

    End Sub
    Private Function makeWebDoc() As String

        makeWebDoc = ""
        Try
            Return makeWebDocBody()
        Catch ex As Exception

        End Try

    End Function
    Private Function makeWebHead() As String

        makeWebHead = "<head>
<style>
td {
    font-size:10.0pt;
    font-family:arial;
}
</style>
</head><body>"

    End Function
    Private Function makeWebDocBody() As String

        Dim pString As New System.Text.StringBuilder

        Try

            With mobjPNR
                pString.Clear()
                pString.AppendLine("")
                pString.AppendLine("<div>")
                pString.AppendLine("<span style='font-size:12.0pt;font-family:arial'>")
                pString.AppendLine("Flight Routing Information<br />")
                pString.AppendLine("</span>")
                pString.AppendLine("<span style='font-size:10.0pt;font-family:arial'>")
                pString.AppendLine(MySettings.FormalOfficeName & "<br />")
                pString.AppendLine("Flight routing information<br />")
                If .ClientName.Trim <> "" Then
                    pString.AppendLine("For: " & .ClientName)
                End If
                pString.AppendLine("<br /><br />")
                pString.AppendLine("Date: " & Format(Now, "dd/MM/yyyy") & "<br /><br />")
                If mobjPNR.VesselName <> "" Then
                    pString.AppendLine("<b><u>VESSEL:</u></b><br />" & mobjPNR.VesselName & "<br /><br />")
                End If

                If mobjPNR.Passengers.Count > 0 Then
                    pString.AppendLine("<b><u>")
                    If mobjPNR.Passengers.Count = 1 Then
                        pString.AppendLine("PASSENGER<br />")
                    Else
                        pString.AppendLine("PASSENGERS<br />")
                    End If
                    pString.AppendLine("</u></b>")
                    Dim iPaxCount As Integer = 0
                    For Each pobjPax In mobjPNR.Passengers.Values
                        iPaxCount = iPaxCount + 1
                        pString.AppendLine(pobjPax.ElementNo & " " & pobjPax.PaxName & " " & pobjPax.PaxID & "<br />")
                    Next pobjPax
                ElseIf mobjPNR.IsGroup Then
                    pString.AppendLine("<b><u>")
                    pString.AppendLine("GROUP<br />")
                    pString.AppendLine("</u></b>")
                    pString.AppendLine(mobjPNR.GroupName & " " & mobjPNR.GroupNamesCount & "<br />")
                Else
                    pString.AppendLine("PASSENGER INFORMATION NOT AVAILABLE")
                End If
                pString.AppendLine("<span style='font-size:12.0pt;font-family:arial'>")
                pString.AppendLine("<br />FLIGHT ROUTING<br />")
                pString.AppendLine("</span>")
                'pString.AppendLine("<div align=left>")
                pString.AppendLine("<span style='font-size:10.0pt;font-family:arial'>")
                pString.AppendLine("<table cellspacing=" & Chr(34) & "1" & Chr(34) & " cellpadding=" & Chr(34) & "2" & Chr(34) & " border=" & Chr(34) & "1" & Chr(34) & " >")
                pString.AppendLine("<tr bgcolor=" & Chr(34) & "#0Fffff" & Chr(34) & ">")
                pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>Airline</td>")
                pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>Flight</td>")
                pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>Date</td>")
                pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>Itinerary</td>")
                pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>Depart</td>")
                pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>Arrive</td>")
                pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>Airline Locator</td>")
                pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>Baggage Allowance</td>")
                pString.AppendLine("</tr>")

                Dim iSegCount As Integer = 0
                Dim pPrevOff As String = ""
                For Each pobjSeg In .Segments.Values
                    iSegCount = iSegCount + 1
                    If iSegCount > 1 And pPrevOff <> pobjSeg.BoardPoint Then
                        pString.AppendLine("<tr>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")

                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>** CHANGE OF AIRPORT **</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        pString.AppendLine("</tr>")
                    End If
                    pPrevOff = pobjSeg.OffPoint
                    pString.AppendLine("<tr>")
                    pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>" & iSegCount & "</td>")
                    pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>" & pobjSeg.Airline & "-" & pobjSeg.AirlineName)
                    If pobjSeg.OperatedBy <> "" Then
                        pString.AppendLine("<br /><span style='font-size:6.0pt;font-family:arial'>" & pobjSeg.OperatedBy & "</span>")
                    End If
                    pString.AppendLine("</td>")
                    pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>" & pobjSeg.FlightNo & "</td>")
                    pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>" & Format(pobjSeg.DepartureDate, "dd/MM/yyyy") & "<br><span style='font-size:6.0pt;font-family:arial'>" & pobjSeg.DepartureDay & "</span></td>")
                    pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>" & pobjSeg.BoardPoint & " " & pobjSeg.BoardCityName.PadRight(.MaxCityNameLength + 1, " "c).Substring(0, .MaxCityNameLength + 1) & " - " &
                    pobjSeg.OffPoint & " " & pobjSeg.OffPointCityName.PadRight(.MaxCityNameLength + 1, " "c).Substring(0, .MaxCityNameLength + 1) & "</td>")
                    If pobjSeg.Text.Length > 35 AndAlso pobjSeg.Text.Substring(35, 4) = "FLWN" Then
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>FLOWN</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                    Else
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>" & Format(pobjSeg.DepartTime, "HH:mm") & "</td>")
                        Dim pDateDiff As Long = DateDiff(DateInterval.Day, pobjSeg.DepartureDate, pobjSeg.ArrivalDate)
                        If pDateDiff = 0 Then
                            pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>" & Format(pobjSeg.ArriveTime, "HH:mm") & "</td>")
                        Else
                            pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>" & Format(pobjSeg.ArriveTime, "HH:mm") & " " & Format(pDateDiff, "+0;-0") & "</td>")
                        End If
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;" & pobjSeg.AirlineLocator & "</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;" & .AllowanceForSegment(pobjSeg.BoardPoint, pobjSeg.OffPoint, pobjSeg.Airline) & "</td>")

                    End If
                    pString.AppendLine("</tr>")
                    If pobjSeg.Stopovers <> "" Then
                        pString.AppendLine("<tr>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        If pobjSeg.Stopovers <> "" Then
                            pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>*INTERMEDIATE STOP*  " & pobjSeg.Stopovers.Trim & "</td>")
                        Else
                            pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        End If
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        pString.AppendLine("<td style='font-size:10.0pt;font-family:arial'>&nbsp;</td>")
                        pString.AppendLine("</tr>")
                    End If

                Next pobjSeg

                pString.AppendLine("</table>")
                pString.AppendLine("</span>")
                'pString.AppendLine("</div>")
                pString.AppendLine("<br />")
                pString.AppendLine("<br />")
                pString.AppendLine("<div>")
                pString.AppendLine("<span style='font-size:10.0pt;font-family:arial'><b><u>Booking Reference</u></b><br />")
                pString.AppendLine(.GDSAbbreviation & "/" & .RequestedPNR & "<br /><br />")
                pString.AppendLine("<b><u>Tickets</u></b><br />")

                For Each pobjPax In .Passengers.Values
                    If .Passengers.Values.Count > 1 Then
                        pString.AppendLine("<u>" & pobjPax.PaxName & "</u><br />")
                    End If

                    For Each tkt As GDSTickets.GDSTicketItem In .Tickets.Values
                        If tkt.Pax.Trim = pobjPax.PaxName.Trim Then
                            If tkt.TicketType = "PAX" Then
                                Dim pFF As String = mobjPNR.FrequentFlyerNumber(tkt.AirlineCode, tkt.Pax.Substring(0, tkt.Pax.Length - 2).Trim)
                                If pFF <> "" Then
                                    pFF = "Frequent Flyer Number: " & pFF
                                End If
                                pString.AppendLine(tkt.IssuingAirline & "-" & tkt.Document & " " & tkt.AirlineCode & " " & pFF & "<br />")
                            End If
                        End If
                    Next
                Next
                pString.AppendLine("<br />")

                pString.AppendLine("Kind Regards<br />")
                pString.AppendLine("</span>")
                pString.AppendLine("</div>")
                pString.AppendLine("</div>")

            End With
        Catch ex As Exception

        End Try

        Return pString.ToString

    End Function
    Private Function makeWebClose() As String

        makeWebClose = "</body></html>"

    End Function
    Private Sub cmdItnRead1ACurrent_Click(sender As Object, e As EventArgs) Handles cmdItn1AReadCurrent.Click
        Try
            mSelectedItnGDSCode = Utilities.EnumGDSCode.Amadeus
            ItnReadCurrentPNR()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub cmdItnRead1GCurrent_Click(sender As Object, e As EventArgs) Handles cmdItn1GReadCurrent.Click
        Try
            mSelectedItnGDSCode = Utilities.EnumGDSCode.Galileo
            ItnReadCurrentPNR()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub ItnReadCurrentPNR()
        Dim mGDSUser As New GDSUser(mSelectedItnGDSCode)
        InitSettings(mGDSUser)
        SetupPCCOptions()
        lblItnPNRCounter.Text = ""
        ReadPNRandCreateItn(False)
        cmdItnRefresh.Enabled = True
        cmdItnFormatOSMLoG.Enabled = True
    End Sub
    Private Sub cmdItnRefresh_Click(sender As Object, e As EventArgs) Handles cmdItnRefresh.Click

        Try
            ReadPNRandCreateItn(True)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub ReadPNRandCreateItn(ByVal RefreshOnly As Boolean)

        Try
            Cursor = System.Windows.Forms.Cursors.WaitCursor
            ProcessRequestedPNRs(RefreshOnly)
            CopyItinToClipboard()
            If Not RefreshOnly Then
                MessageBox.Show("Ready", "Read PNR", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            Throw New Exception("ReadPNRandCreateItn" & vbCrLf & ex.Message)
        Finally
            Cursor = Cursors.Default
        End Try

    End Sub
    Private Sub CopyItinToClipboard()

        Try
            If Not optItnFormatEuronav.Checked Then
                rtbItnDoc.SelectAll()
                Clipboard.Clear()
                Clipboard.SetText(rtbItnDoc.Rtf, TextDataFormat.Rtf)
                Clipboard.SetText(rtbItnDoc.SelectedText, TextDataFormat.Text)
            End If
        Catch ex As Exception
            ' ignore any error that occurs when copying to clipboard
        End Try

    End Sub
    Private Sub optItnAirportCode_CheckedChanged(sender As Object, e As EventArgs) Handles optItnAirportCode.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.AirportName = 0
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub optItnAirportname_CheckedChanged(sender As Object, e As EventArgs) Handles optItnAirportname.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.AirportName = 1
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub optItnAirportBoth_CheckedChanged(sender As Object, e As EventArgs) Handles optItnAirportBoth.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.AirportName = 2
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub optItnAirportCityName_CheckedChanged(sender As Object, e As EventArgs) Handles optItnAirportCityName.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.AirportName = 3
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub optItnAirportCityBoth_CheckedChanged(sender As Object, e As EventArgs) Handles optItnAirportCityBoth.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.AirportName = 4
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkItnVessel_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnVessel.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.ShowVessel = chkItnVessel.Checked
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkItnClass_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnClass.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.ShowClassOfService = chkItnClass.Checked
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkItnAirlineLocator_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnAirlineLocator.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.ShowAirlineLocator = chkItnAirlineLocator.Checked
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkItnTickets_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnTickets.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.ShowTickets = chkItnTickets.Checked
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkPaxSegPerTicket_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnPaxSegPerTicket.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.ShowPaxSegPerTkt = chkItnPaxSegPerTicket.Checked
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkSeating_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnSeating.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.ShowSeating = chkItnSeating.Checked
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkTerminal_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnTerminal.CheckedChanged
        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.ShowTerminal = chkItnTerminal.Checked
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkStopovers_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnStopovers.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.ShowStopovers = chkItnStopovers.Checked
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkItnFlyingTime_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnFlyingTime.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.ShowFlyingTime = chkItnFlyingTime.Checked
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkItnCostCentre_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnCostCentre.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.ShowCostCentre = chkItnCostCentre.Checked
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub chkItnElecItemsBan_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnElecItemsBan.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.ShowBanElectricalEquipment = chkItnElecItemsBan.Checked
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkBrazilText_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnBrazilText.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.ShowBrazilText = chkItnBrazilText.Checked
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkUSAText_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnUSAText.CheckedChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    MySettings.ShowUSAText = chkItnUSAText.Checked
                    MySettings.Save()
                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub txtPNR_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtItnPNR.TextChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    cmdItn1AReadPNR.Enabled = (txtItnPNR.Text.Trim.Length >= 6)
                    cmdItn1AReadQueue.Enabled = (txtItnPNR.Text.Trim.Length >= 2)
                    cmdItn1GReadPNR.Enabled = cmdItn1AReadPNR.Enabled
                    cmdItn1GReadQueue.Enabled = cmdItn1AReadQueue.Enabled
                    optItnFormatMSReport.Enabled = cmdItn1AReadQueue.Enabled
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub readGDS(ByVal RecordLocator As String)

        Try
            If RecordLocator = "" Then
                mobjPNR.CancelError = True
            Else
                mobjPNR.CancelError = False
            End If
            mobjPNR.Read(mSelectedItnGDSCode, RecordLocator, (optItnFormatMSReport.Checked))
        Catch ex As Exception
            Throw New Exception("readGDS()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub cmdCostCentre_Click(sender As Object, e As EventArgs) Handles cmdCostCentre.Click

        Try
            Dim pfrmcostCentre As New frmCostCentre
            Dim pResult As System.Windows.Forms.DialogResult
            mflgLoading = False
            pResult = pfrmcostCentre.ShowDialog()

            If pResult = Windows.Forms.DialogResult.OK Then
                txtCustomer.Text = pfrmcostCentre.CodeSelected
                txtVessel.Text = pfrmcostCentre.VesselSelected
                DisplayOldCustomProperty(cmbCostCentre, pfrmcostCentre.CostCentreSelected)
            End If
            pfrmcostCentre.Close()
        Catch ex As Exception
            MessageBox.Show("cmdCostCentre_Click()" & vbCrLf & ex.Message)
        End Try


    End Sub

    Private Sub cmdAveragePrice_Click(sender As Object, e As EventArgs) Handles cmdAveragePrice.Click

        Try
            With mobjAveragePrice
                If .Load() Then
                    lblAvPriceDetails.Text = "From: " & .FromDate & "  " & .Itinerary
                    lblAveragePrice.Text = .TicketCount & " tkts - Avge Price: " & Format(.AveragePrice, "#,##0 EUR")
                Else
                    lblAvPriceDetails.Text = "Cannot calculate round trip"
                    lblAveragePrice.Text = ""
                End If
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

#Region "OSM"

    Private Sub cmdOSMRefresh_Click(sender As Object, e As EventArgs) Handles cmdOSMRefresh.Click

        Try
            UtilitiesOSM.OSMRefreshVesselGroup(cmbOSMVesselGroup)
            UtilitiesOSM.OSMRefreshVessels(lstOSMVessels, chkOSMVesselInUse.Checked)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub cmbOSMVesselGroup_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbOSMVesselGroup.SelectedIndexChanged
        Try
            If Not mflgLoading Then
                If MySettings Is Nothing Then
                    InitSettings()
                End If
                Dim pSelectedItem As osmVessels.VesselGroupItem
                pSelectedItem = cmbOSMVesselGroup.SelectedItem
                MySettings.OSMVesselGroup = pSelectedItem.Id
                UtilitiesOSM.OSMRefreshVessels(lstOSMVessels, chkOSMVesselInUse.Checked)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub cmdOSMCopyTo_Click(sender As Object, e As EventArgs) Handles cmdOSMCopyTo.Click

        Try
            Dim pstrEmail As String = ""

            For Each pSelectedAgent As osmVessels.emailItem In lstOSMAgents.SelectedItems
                If pstrEmail <> "" Then
                    pstrEmail &= "; "
                End If
                pstrEmail &= pSelectedAgent.ToString
            Next

            For Each pEmailTO As osmVessels.emailItem In lstOSMToEmail.Items
                If pstrEmail <> "" Then
                    pstrEmail &= "; "
                End If
                pstrEmail &= pEmailTO.ToString
            Next
            Clipboard.Clear()
            Clipboard.SetText(pstrEmail)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmdOSMCopyCC_Click(sender As Object, e As EventArgs) Handles cmdOSMCopyCC.Click

        Try
            Dim pstrEmail As String = ""

            For Each pEmailTO As osmVessels.emailItem In lstOSMCCEmail.Items
                If pstrEmail <> "" Then
                    pstrEmail &= "; "
                End If
                pstrEmail &= pEmailTO.ToString
            Next
            Clipboard.Clear()
            Clipboard.SetText(pstrEmail)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmdOSMCopyDocument_Click(sender As Object, e As EventArgs) Handles cmdOSMCopyDocument.Click

        Try
            Dim dobj As New DataObject
            dobj.SetData(DataFormats.Html, webOSMDoc.DocumentStream)
            Clipboard.Clear()
            Clipboard.SetDataObject(dobj)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmdOSMCopyTextOnly_Click(sender As Object, e As EventArgs)

        Try
            Clipboard.Clear()
            Clipboard.SetText(webOSMDoc.DocumentText)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub lstOSMVessels_DrawItem(sender As Object, e As DrawItemEventArgs) Handles lstOSMVessels.DrawItem

        Try
            UtilitiesOSM.ListBox_DrawItem(sender, e)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub lstOSMVessels_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstOSMVessels.SelectedIndexChanged

        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    OSMShowSelectedVesselEmails()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub OSMShowSelectedVesselEmails()

        Try

            UtilitiesOSM.OSMDisplayEmails(lstOSMVessels, lstOSMToEmail, lstOSMCCEmail, lstOSMAgents)
            mOSMAgents.Load()
            mOSMAgentIndex = -1

            cmdOSMCopyTo.Enabled = (lstOSMToEmail.Items.Count > 0 Or lstOSMAgents.SelectedItems.Count > 0)
            cmdOSMCopyCC.Enabled = (lstOSMCCEmail.Items.Count > 0)

            lblOSMVessel.Text = ""
            txtOSMAgentsFilter.Clear()

            For Each pVessel As osmVessels.VesselItem In lstOSMVessels.SelectedItems
                If lblOSMVessel.Text <> "" Then
                    lblOSMVessel.Text &= " / "
                End If
                lblOSMVessel.Text &= pVessel.ToString
            Next
        Catch ex As Exception
            Throw New Exception("OSMShowSelectedVesselEmails()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub OSMWebCreate(ByVal ShowFullPaxDetails As Boolean)

        Try
            webOSMDoc.DocumentText = OSMWebHeader(ShowFullPaxDetails)
            cmdOSMCopyDocument.Enabled = True
        Catch ex As Exception
            Throw New Exception("OSMWebCreate()" & vbCrLf & ex.Message)
        End Try
    End Sub
    Private Function OSMWebHeader(ByVal ShowFullPaxDetails As Boolean) As String

        Try
            Dim xDoctext As String = "<html><head></head><body>"

            xDoctext &= "MESSAGE FROM :<br>"

            xDoctext &= "<b>ATPI GRIFFINSTONE GREECE</b><br><br>"

            For Each pSelectedAgent As osmVessels.emailItem In lstOSMAgents.SelectedItems
                xDoctext &= "TO         : " & pSelectedAgent.Name & " / " & pSelectedAgent.Details & "<br>"
            Next

            For Each pEmail As osmVessels.emailItem In lstOSMToEmail.Items
                If lstOSMVessels.SelectedItems.Count > 1 Then
                    xDoctext &= "TO         : " & pEmail.Name & If(pEmail.Details <> "", " / " & pEmail.Details, "") & If(pEmail.VesselName <> "", "(" & pEmail.VesselName & ")", "") & "<br>"
                Else
                    xDoctext &= "TO         : " & pEmail.Name & If(pEmail.Details <> "", " / " & pEmail.Details, "") & "<br>"
                End If
            Next

            xDoctext &= "<br>"
            xDoctext &= "CC         : OSM CYPRUS<br>"
            For Each pEmail As osmVessels.emailItem In lstOSMCCEmail.Items
                xDoctext &= "CC         : " & pEmail.Name & If(pEmail.Details <> "", " / " & pEmail.Details, "") & "<br>"
            Next
            xDoctext &= "CC         : 3rd party applicable<br>"
            xDoctext &= "<br>"
            xDoctext &= "<br>If more information is required please contact ATPI Greece and copy travel vessel IMO no@osm.biz<br><br>"
            xDoctext &= "DATE/REF   : " & Format(Now, "dd/MM/yyyy") & "<br><br><br>"
            Dim pTempSubject As String = ""

            For Each pVessel As osmVessels.VesselItem In lstOSMVessels.SelectedItems
                If pTempSubject <> "" Then
                    pTempSubject &= " / "
                End If
                pTempSubject &= pVessel.VesselName
            Next
            xDoctext &= "SUBJECT     : VSL " & pTempSubject & " CREW CHANGE AT PORT  <br>"

            xDoctext &= "<br><br>"
            xDoctext &= "PLEASE BE ADVISED OF THE FOLLOWING ARRANGEMENTS FOR EMBARKING / DISEMBARKING CREW.<br><br><br>"
            xDoctext &= "<font color=" & Chr(34) & "red" & Chr(34) & ">PLEASE CONFIRM RECEIPT OF BELOW :</font><br><br><br>"

            Dim pOnSigners As String = ""
            Dim pOnSignerNoVisa As String = ""
            Dim pOnSignerVisa As String = ""
            Dim pOnSignerOKTB As String = ""

            Dim pOffSigners As String = ""

            Dim pOther As String = ""

            For i As Integer = 0 To dgvOSMPax.Rows.Count - 1
                Dim pId As Integer = CInt(dgvOSMPax.Rows(i).Cells(0).Value)
                Dim pPax As osmPax.Pax = mOSMPax(pId)
                Select Case CStr(dgvOSMPax.Rows(i).Cells("JoinerLeaver").Value)
                    Case "ONSIGNER"
                        If ShowFullPaxDetails Then
                            pOnSigners &= "<pre>" & pPax.TextFullDetails & "</pre><br><br>"
                        Else
                            pOnSigners &= "<pre>" & pPax.Text & "</pre><br><br>"
                        End If
                    Case "OFFSIGNER"
                        If ShowFullPaxDetails Then
                            pOffSigners &= "<pre>" & pPax.TextFullDetails & "</pre><br><br>"
                        Else
                            pOffSigners &= "<pre>" & pPax.Text & "</pre><br><br>"
                        End If
                    Case Else
                        If ShowFullPaxDetails Then
                            pOther &= "<pre>" & pPax.TextFullDetails & "</pre><br><br>"
                        Else
                            pOther &= "<pre>" & pPax.Text & "</pre><br><br>"
                        End If
                End Select

                Select Case CStr(dgvOSMPax.Rows(i).Cells("VisaType").Value)
                    Case "OKTB"
                        pOnSignerOKTB &= dgvOSMPax.Rows(i).Cells("Lastname").Value.ToString & "/" & dgvOSMPax.Rows(i).Cells("Firstname").Value.ToString & "<br>"
                    Case "NO VISA"
                        pOnSignerNoVisa &= dgvOSMPax.Rows(i).Cells("Lastname").Value.ToString & "/" & dgvOSMPax.Rows(i).Cells("Firstname").Value.ToString & "<br>"
                    Case "VISA"
                        pOnSignerVisa &= dgvOSMPax.Rows(i).Cells("Lastname").Value.ToString & "/" & dgvOSMPax.Rows(i).Cells("Firstname").Value.ToString & "<br>"
                End Select
            Next

            If pOther <> "" Then
                xDoctext &= "PARTICULARS OF TRAVELLER AS FOLLOWS:</b><br><br>"
                xDoctext &= "<pre>" & pOther
                xDoctext &= "<font color=" & Chr(34) & "red" & Chr(34) & ">FLIGHT DETAILS: </font></pre><br><br>"
            End If
            If pOnSigners <> "" Then
                xDoctext &= "PARTICULARS OF JOINERS AS FOLLOWS:</b><br><br>"
                xDoctext &= "SIGNING ON<br><br>"
                xDoctext &= "<pre>" & pOnSigners
                xDoctext &= "<font color=" & Chr(34) & "red" & Chr(34) & "><b>FLIGHT DETAILS: </b></font><br><br>"
                If pOnSignerNoVisa <> "" Then
                    xDoctext &= "<hr width=30% align=left>" & pOnSignerNoVisa & "<br><br>"
                    xDoctext &= "<pre><b>NO VISA REQUIRED</b></pre><br><br>"
                End If
                If pOnSignerVisa <> "" Then
                    xDoctext &= "<hr width=30% align=left>" & pOnSignerVisa & "<br><br>"
                    xDoctext &= "<pre><b>VISA REQUIRED</b></pre><br><br>"
                    xDoctext &= "<b>CREW WILL TRAVEL WITH VALID VISA. <br><br>"
                    xDoctext &= "AGENT PLEASE ENSURE THAT ONSIGNER'S PASSPORT HAVE AN EXIT STAMP <br>"
                    xDoctext &= "FROM THE IMMIGRATION BEFORE THEY GO ON BOARD. </b><br><br>"
                End If
                If pOnSignerOKTB <> "" Then
                    xDoctext &= "<hr width=30% align=left>" & pOnSignerOKTB & "<br>"
                    xDoctext &= "<pre><b>OKTB</b></pre><br><br>"
                    xDoctext &= "<b>*****IMPORTANT******<br><br>"
                    xDoctext &= "PLS SEND -OK TO BOARD- TO ____ THROUGH NEAREST TOWNOFFICE/AIRPORT<br>"
                    xDoctext &= "OFFICE YOUR SIDE .<br><br>"
                    xDoctext &= "THE LETTER SHOULD CONTAIN THE FOLLOWING WORDINGS OK TO BOARD THAT<br>"
                    xDoctext &= "AIRLINE COUNTER IS REQUIRING FROM THE AGENT.<br><br>"
                    xDoctext &= "WE NEED YOUR FORMAL ACKNOWLEDGEMENT THAT YOU HAVE GIVEN THE -OK TO BOARD<br>"
                    xDoctext &= "PLS ALSO SEND COPY OF OK TO BOARD TO :<br><br>"
                    xDoctext &= "ATPI ATHENS  = E-MAIL : osmsmart.greece@ atpi.com<br>"
                    xDoctext &= "OSM ________ = E-MAIL : _________@_______<br><br>"
                    xDoctext &= "AGENT PLEASE ENSURE THAT ONSIGNER'S PASSPORT HAVE AN EXIT STAMP <br>"
                    xDoctext &= "FROM THE IMMIGRATION BEFORE THEY GO ON BOARD. </b><br><br>"
                End If
                xDoctext &= " "
                xDoctext &= "<br>AGENT PLS CHECK IF THIS ARRANGEMENT IS ACCEPTABLE WITH ETA/ETD<br><br>"
                xDoctext &= "PLS LIASE WITH MASTER, MEET CREW, ARRANGE ENTRY FORMALITIES AND SECURE <br>"
                xDoctext &= "SAFE JOINSHIP.</pre><br>"
            End If
            If pOffSigners <> "" Then
                xDoctext &= "<hr SIZE=2 COLOR=gray>"
                xDoctext &= "<b>TO MASTER/PORT AGENT</b><br><br>"
                xDoctext &= "FOLLOWING ROUTE ARE CONFIRMED FOR HOMEGOING CREW AS FLWS :<br><br>"
                xDoctext &= "OFFSIGNER:<br><br>"
                xDoctext &= "<pre>" & pOffSigners & "</pre><br>"
                xDoctext &= "<pre><font color=" & Chr(34) & "red" & Chr(34) & "><b>FLIGHT DETAILS: </b></font></pre><br><br>"
                xDoctext &= "<pre><b>AGENT – CREW TRAVELING ON E-TICKETS, AND MUST GO DIRECTLY TO CHECK-IN<br>"
                xDoctext &= "COUNTER AT AIRPORT WITH PASSPORT READY.</b></pre><br><br>"
                xDoctext &= "<br><pre>PLS LIASE WITH MASTER AND CONVEY CREW TO AIRPORT.</pre><br>"

            End If
            xDoctext &= "<br><pre>IF ANY PROBLEM REGARDING FLIGHT DETAILS, PLS CONTACT OUR OFFICE.</pre><br><br>"

            xDoctext &= "</body></html>"

            Return xDoctext
        Catch ex As Exception
            Throw New Exception("OSMWebHeader()" & vbCrLf & ex.Message)
        End Try
    End Function
    Private Sub txtOSMPax_KeyDown(sender As Object, e As KeyEventArgs) Handles txtOSMPax.KeyDown

        Try
            If e.Control And e.KeyCode = Keys.A Then
                txtOSMPax.SelectAll()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Sub txtOSMText_TextChanged(sender As Object, e As EventArgs) Handles txtOSMPax.TextChanged
        Try
            OSMAnalyzePax()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub OSMAnalyzePax()
        Try
            mOSMPax.Load(txtOSMPax.Text)
            dgvOSMPax.Rows.Clear()
            For Each iPax As osmPax.Pax In mOSMPax.Values
                Dim pId As New DataGridViewTextBoxCell
                Dim pLastName As New DataGridViewTextBoxCell
                Dim pFirstName As New DataGridViewTextBoxCell
                Dim pNationality As New DataGridViewTextBoxCell
                Dim pJoiner As New DataGridViewComboBoxCell
                Dim pVisaType As New DataGridViewComboBoxCell
                pId.Value = iPax.Id
                pLastName.Value = iPax.LastName
                pFirstName.Value = iPax.FirstName
                pNationality.Value = iPax.Nationality
                pJoiner.Items.AddRange({"ONSIGNER", "OFFSIGNER"})
                pVisaType.Items.AddRange({"OKTB", "VISA", "NO VISA"})
                If iPax.JoinerLeaver <> "" Then
                    pJoiner.Value = iPax.JoinerLeaver
                End If
                Dim pRow As New DataGridViewRow
                pRow.Cells.Add(pId)
                pRow.Cells.Add(pLastName)
                pRow.Cells.Add(pFirstName)
                pRow.Cells.Add(pNationality)
                pRow.Cells.Add(pJoiner)
                pRow.Cells.Add(pVisaType)
                dgvOSMPax.Rows.Add(pRow)
            Next
            dgvOSMPax.Columns(1).ReadOnly = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub cmdOSMPrepareDoc_Click(sender As Object, e As EventArgs) Handles cmdOSMPrepareDoc.Click
        Try
            OSMWebCreate(chkOSMFullPaxSDetails.Checked)
            cmdOSMCopyDocument.Enabled = True

        Catch ex As Exception
            MessageBox.Show("cmdOSMPrepareDoc_Click()" & vbCrLf & ex.Message)
        End Try
    End Sub
    Private Sub cmdOSMVesselsEdit_Click(sender As Object, e As EventArgs) Handles cmdOSMVesselsEdit.Click
        Try
            Dim pFrm As New frmOSMVessels
            pFrm.ShowDialog(Me)
            UtilitiesOSM.OSMRefreshVessels(lstOSMVessels, chkOSMVesselInUse.Checked)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub cmdOSMAgentEdit_Click(sender As Object, e As EventArgs) Handles cmdOSMAgentEdit.Click
        Try
            Dim pFrm As New frmOSMAgents
            If pFrm.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                UtilitiesOSM.OSMRefreshVessels(lstOSMVessels, chkOSMVesselInUse.Checked)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub dgvOSMPax_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvOSMPax.CellValueChanged
        Dim pflgLoading As Boolean = mflgLoading
        Try
            If Not mflgLoading Then
                mflgLoading = True
                If e.ColumnIndex = 5 Then
                    For i As Integer = 0 To dgvOSMPax.Rows.Count - 1
                        If i <> e.RowIndex AndAlso CStr(dgvOSMPax.Rows(i).Cells("JoinerLeaver").Value) = "ONSIGNER" AndAlso dgvOSMPax.Rows(i).Cells("VisaType").Value Is Nothing Then
                            dgvOSMPax.Rows(i).Cells("VisaType").Value = dgvOSMPax.Rows(e.RowIndex).Cells("VisaType").Value
                        End If
                    Next
                End If
            End If
        Catch ex As Exception

        Finally
            mflgLoading = pflgLoading
        End Try
    End Sub
    Private Sub cmdOSMEmailClear_Click(sender As Object, e As EventArgs) Handles cmdOSMEmailClear.Click
        Try
            txtOSMPax.Clear()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region
    Private Sub tabPNR_SelectedIndexChanged(sender As Object, e As EventArgs) Handles tabPNR.SelectedIndexChanged
        Try
            mflgLoading = True
            If tabPNR.SelectedIndex = 1 Then
                cmdItnFormatOSMLoG.Enabled = False
            ElseIf tabPNR.SelectedIndex = 2 Then
                UtilitiesOSM.OSMRefreshVesselGroup(cmbOSMVesselGroup)
                UtilitiesOSM.OSMRefreshVessels(lstOSMVessels, chkOSMVesselInUse.Checked)
                cmdOSMCopyDocument.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            mflgLoading = False
        End Try
    End Sub
    Private Sub SetITNEnabled(ByVal AllowOptions As Boolean)
        fraItnAirportName.Enabled = AllowOptions
        fraItnOptions.Enabled = AllowOptions
        lstItnRemarks.Enabled = AllowOptions
    End Sub
    Private Sub optItnFormatDefault_CheckedChanged(sender As Object, e As EventArgs) Handles optItnFormatDefault.CheckedChanged, optItnFormatPlain.CheckedChanged, optItnFormatSeaChefs.CheckedChanged, optItnFormatSeaChefsWith3LetterCode.CheckedChanged, optItnFormatMSReport.CheckedChanged, optItnFormatEuronav.CheckedChanged
        Try
            If Not mflgLoading Then
                If Not MySettings Is Nothing Then
                    ' optItnFormatDefault = 0
                    ' optItnFormatPlain = 1
                    ' optItnFormatSeaChefs = 2
                    ' chkItnSeaChefsWithCode = 3
                    ' optItnFormatEuronav = 4
                    If optItnFormatDefault.Checked Then
                        MySettings.FormatStyle = Utilities.EnumItnFormat.DefaultFormat
                    ElseIf optItnFormatPlain.Checked Then
                        MySettings.FormatStyle = Utilities.EnumItnFormat.Plain
                    ElseIf optItnFormatSeaChefs.Checked Then
                        MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefs
                    ElseIf optItnFormatSeaChefsWith3LetterCode.Checked Then
                        MySettings.FormatStyle = Utilities.EnumItnFormat.SeaChefsWithCode
                    ElseIf optItnFormatEuronav.Checked Then
                        MySettings.FormatStyle = Utilities.EnumItnFormat.Euronav
                    End If
                    MySettings.Save()

                    If cmdItnRefresh.Enabled Then
                        ReadPNRandCreateItn(True)
                    End If
                    If sender.Name = "optItnFormatDefault" Or sender.Name = "optItnFormatPlain" Then
                        SetITNEnabled(True)
                    Else
                        SetITNEnabled(False)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub cmdItnFormatOSMLoG_Click(sender As Object, e As EventArgs) Handles cmdItnFormatOSMLoG.Click
        Try
            If mobjPNR.Segments.Count > 0 And mobjPNR.Passengers.Count > 0 Then
                Dim pOSMLoG = New OsmLOG
                pOSMLoG.CreatePDF(MySettings.AgentName, mobjPNR)
            Else
                MessageBox.Show("PNR must have passengers and segments to produce a Letter of Guarantee")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub lstItnRemarks_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstItnRemarks.SelectedIndexChanged
        Try
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub cmdOSMClearSelected_Click(sender As Object, e As EventArgs) Handles cmdOSMClearSelected.Click
        Try
            mflgLoading = True
            For i As Integer = 0 To lstOSMVessels.Items.Count - 1
                lstOSMVessels.SetSelected(i, False)
            Next
            For i As Integer = 0 To lstOSMAgents.Items.Count - 1
                lstOSMAgents.SetSelected(i, False)
            Next
            mflgLoading = False
            OSMShowSelectedVesselEmails()
        Catch ex As Exception
            mflgLoading = False
            MessageBox.Show("cmdOSMClearSelected_Click()" & vbCrLf & ex.Message)
        End Try
    End Sub
    Private Sub lstOSMAgents_MouseMove(sender As Object, e As MouseEventArgs) Handles lstOSMAgents.MouseMove
        Try
            Dim pIndex As Integer = lstOSMAgents.IndexFromPoint(e.Location)
            If pIndex >= 0 And pIndex < lstOSMAgents.Items.Count And mOSMAgentIndex <> pIndex Then
                ttpToolTip.SetToolTip(lstOSMAgents, lstOSMAgents.Items(pIndex).ToString)
                mOSMAgentIndex = pIndex
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub lstOSMAgents_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstOSMAgents.SelectedIndexChanged
        Try
            cmdOSMCopyTo.Enabled = (lstOSMToEmail.Items.Count > 0 Or lstOSMAgents.SelectedItems.Count > 0)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub txtOSMAgentsFilter_TextChanged(sender As Object, e As EventArgs) Handles txtOSMAgentsFilter.TextChanged
        Try
            lstOSMAgents.Items.Clear()
            mOSMAgentIndex = -1
            If txtOSMAgentsFilter.Text.Trim = "" Then
                For Each pAgent As osmVessels.emailItem In mOSMAgents.Values
                    lstOSMAgents.Items.Add(pAgent)
                Next
            Else
                Dim pFilter() As String = txtOSMAgentsFilter.Text.ToUpper.Trim.Split({"|"}, StringSplitOptions.RemoveEmptyEntries)

                For Each pAgent As osmVessels.emailItem In mOSMAgents.Values
                    For i As Integer = 0 To pFilter.GetUpperBound(0)
                        If pAgent.ToString.ToUpper.IndexOf(pFilter(i).Trim) >= 0 Then
                            lstOSMAgents.Items.Add(pAgent)
                            Exit For
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub chkOSMVesselInUse_CheckedChanged(sender As Object, e As EventArgs) Handles chkOSMVesselInUse.CheckedChanged
        Try
            If Not mflgLoading And chkOSMVesselInUse.Visible Then
                UtilitiesOSM.OSMRefreshVessels(lstOSMVessels, chkOSMVesselInUse.Checked)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Public Function APISValidateDataRow(ByVal Row As DataGridViewRow) As Boolean
        Dim pdteDate As DateTime
        Dim pflgGenderFound As Boolean = False
        Dim pflgBirthDateOK As Boolean = False
        Dim pflgPassportNumberOK As Boolean = False
        Dim pstrErrorText As String = ""
        pflgPassportNumberOK = (Trim(Row.Cells("PassportNumber").Value).Length > 0)
        If Not Date.TryParse(Row.Cells("Birthdate").Value, pdteDate) Then
            pdteDate = Utilities.DateFromIATA(Row.Cells("Birthdate").Value)
            If pdteDate > Date.MinValue Then
                pflgBirthDateOK = True
            Else
                pflgBirthDateOK = False
            End If
        Else
            pflgBirthDateOK = True
        End If
        If Not Date.TryParse(CStr(Row.Cells("ExpiryDate").Value), pdteDate) Then
            pdteDate = Utilities.DateFromIATA(CStr(Row.Cells("ExpiryDate").Value))
        End If
        If pdteDate > Now Then
            mflgExpiryDateOK = True
        Else
            mflgExpiryDateOK = False
        End If
        pflgGenderFound = False
        For Each pGenderItem As PaxApisDB.ReferenceItem In mobjGender.Values
            If CStr(Row.Cells("Gender").Value) = pGenderItem.Code Then
                pflgGenderFound = True
                Exit For
            End If
        Next
        mflgAPISUpdate = mflgAPISUpdate Or (Not mobjPNR.SSRDocsExists And mobjPNR.SegmentsExist And pflgBirthDateOK And pflgGenderFound) '  And pflgPassportNumberOK)
        If Not pflgBirthDateOK Then
            pstrErrorText &= "Invalid birth date" & vbCrLf
        End If
        If Not pflgGenderFound Then
            pstrErrorText &= "Invalid gender" & vbCrLf
        End If
        If Not pflgPassportNumberOK Then
            pstrErrorText &= "Passport number missing" & vbCrLf
        End If
        If Not mflgExpiryDateOK Then
            pstrErrorText &= "Invalid expiry date" & vbCrLf
        End If
        If mobjPNR.SSRDocsExists Then
            lblSSRDocs.Text = "SSR DOCS already exist in the PNR"
            lblSSRDocs.BackColor = Color.Red
            cmdAPISEditPax.Enabled = False
        Else
            If mobjPNR.SegmentsExist Then
                lblSSRDocs.Text = "SSR DOCS"
                lblSSRDocs.BackColor = Color.Yellow
                cmdAPISEditPax.Enabled = True
            Else
                lblSSRDocs.Text = "SSR DOCS cannot be updated - No segments in PNR"
                lblSSRDocs.BackColor = Color.Red
                cmdAPISEditPax.Enabled = False
            End If
        End If
        Row.ErrorText = pstrErrorText
        SetEnabled()

    End Function
    Private Sub dgvApis_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvApis.CellValueChanged
        Try
            dgvApis.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = dgvApis.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString.ToUpper
        Catch ex As Exception

        End Try
        APISValidateDataRow(dgvApis.Rows(e.RowIndex))
    End Sub
    Private Sub dgvApis_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles dgvApis.CurrentCellDirtyStateChanged
        cmdPNROnlyDocs.Enabled = False
        cmdPNRWriteWithDocs.Enabled = False
    End Sub
    Private Sub dgvApis_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvApis.RowValidating
        APISValidateDataRow(dgvApis.Rows(e.RowIndex))
    End Sub
    Public Sub APISDisplayPax()
        If mobjPNR.SSRDocsExists Then
            txtPNRApis.Location = dgvApis.Location
            txtPNRApis.Size = dgvApis.Size
            txtPNRApis.Text = mobjPNR.SSRDocs
            txtPNRApis.BackColor = Color.Aqua
            txtPNRApis.ForeColor = Color.Blue
            txtPNRApis.Visible = True
            txtPNRApis.BringToFront()
            cmdAPISEditPax.Enabled = False
        Else
            txtPNRApis.Visible = False
            Dim pobjPaxApis As New PaxApisDB.Collection
            dgvApis.Rows.Clear()
            For Each pobjPax As GDSPax.GDSPaxItem In mobjPNR.Passengers.Values
                Dim pobjPaxItem As New PaxApisDB.Item(pobjPax.LastName, pobjPax.Initial)
                pobjPaxApis.Read(pobjPax.LastName, UtilitiesAPIS.APISModifyFirstName(pobjPax.Initial))
                If pobjPaxApis.Count = 0 Then
                    UtilitiesAPIS.APISAddRow(dgvApis, pobjPax.ElementNo, pobjPax.LastName, pobjPax.Initial, "", "", "", Date.MinValue, "", Date.MinValue)
                Else
                    If pobjPaxApis.Count > 1 Then
                        Dim pFrm As New frmAPISPaxSelect(pobjPax.ElementNo, pobjPax.LastName, pobjPax.Initial, pobjPaxApis)
                        If pFrm.ShowDialog(Me) = DialogResult.OK Then
                            pobjPaxItem = pFrm.SelectedPassenger
                        End If
                    Else
                        pobjPaxItem = pobjPaxApis.Values(0)
                    End If
                    UtilitiesAPIS.APISAddRow(dgvApis, pobjPax.ElementNo, pobjPax.LastName, pobjPax.Initial, pobjPaxItem.IssuingCountry, pobjPaxItem.PassportNumber, pobjPaxItem.Nationality, pobjPaxItem.BirthDate, pobjPaxItem.Gender, pobjPaxItem.ExpiryDate)
                End If
                APISValidateDataRow(dgvApis.Rows(dgvApis.RowCount - 1))
            Next
            cmdAPISEditPax.Enabled = True
        End If
    End Sub
    Private Sub MenuCopyItn_Click(sender As Object, e As EventArgs) Handles MenuCopyItn.Click
        Try
            rtbItnDoc.SelectAll()
            Clipboard.Clear()
            Clipboard.SetText(rtbItnDoc.Rtf, TextDataFormat.Rtf)
            Clipboard.SetText(rtbItnDoc.SelectedText, TextDataFormat.Text)
        Catch ex As Exception
            ' ignore any error that occurs when copying to clipboard
        End Try
    End Sub
    Private Sub cmdAdmin_Click(sender As Object, e As EventArgs) Handles cmdAdmin.Click
        Try
            Dim pfrmAdmin As New frmUser(Utilities.EnumGDSCode.Amadeus, "ATHG42100", "9044CN")
            MessageBox.Show(pfrmAdmin.ShowDialog(Me).ToString)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub webItnDoc_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles webItnDoc.DocumentCompleted

        Try
            If optItnFormatEuronav.Checked Then
                Dim dobj As New DataObject
                dobj.SetData(DataFormats.Text, webItnDoc.Document.Body.InnerText)
                dobj.SetData(DataFormats.Html, webItnDoc.DocumentStream)
                Clipboard.Clear()
                Clipboard.SetDataObject(dobj, True)
            End If
        Catch ex As Exception
            ' ignore any error that occurs when copying to clipboard
        End Try

    End Sub

    Private Sub cmdPNRRead1GPNR_Click(sender As Object, e As EventArgs) Handles cmdPNRRead1GPNR.Click
        Try
            mSelectedPNRGDSCode = Utilities.EnumGDSCode.Galileo
            ClearForm()
            ReadPNR(Utilities.EnumGDSCode.Galileo)
            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub cmdAPISEditPax_Click(sender As Object, e As EventArgs) Handles cmdAPISEditPax.Click

        Try
            Dim pFrm As New frmAPISPax
            If pFrm.ShowDialog(Me) = DialogResult.OK Then
                APISDisplayPax()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub cmdPriceOptimiser_Click(sender As Object, e As EventArgs) Handles cmdPriceOptimiser.Click

        ShowPriceOptimiser()

    End Sub
End Class