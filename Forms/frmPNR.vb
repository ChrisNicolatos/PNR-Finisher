Public Class frmPNR

    Const CtrlMask = 8

    Private mstrGenderIndicator() As String = {"M", "F", "MI", "FI", "U"}

    Private mstrSalutations() As String = {"MR", "MRS", "MS", "MISS", "MISTER"}

    Private Structure PaxNamesPos
        Dim StartPos As Integer
        Dim EndPos As Integer
    End Structure

    Private WithEvents mobjAmadeus As New gtmAmadeusPNR

    Private mobjReadPNR As New ReadPNR
    Private mflgReadPNR As Boolean
    'Private mflgNoPNR As Boolean
    Private mMaxString As Integer = 80

    Private mobjAirlinePoints As New AirlinePoints.Collection
    Private mobjAirlineNotes As New AirlineNotes.Collection

    Private mobjCustomerSelected As Customers.CustomerItem
    Private mobjSubDepartmentSelected As SubDepartments.Item
    Private mobjCRMSelected As CRM.Item
    Private mobjVesselSelected As Vessels.Item
    Private mobjAveragePrice As New PriceLookup.Collection
    Private mflgLoading As Boolean

    Private mobjAPISForm As New PaxApisDB.FormElements
    Private mobjCustomers As New Customers.CustomerCollection
    Private mPaxNames() As PaxNamesPos
    Private HeaderLength As Integer = 0
    Private mOSMPax As New osmPax.PaxCollection
    Private mOSMAgents As New osmVessels.emailCollection
    Private mOSMAgentIndex As Integer = -1

    Private mItnFromDate As Date
    Private mItnToDate As Date

    Private mflgExpiryDateOK As Boolean
    Private mflgAPISUpdate As Boolean

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click, cmdItnExit.Click

        Me.Close()

    End Sub

    Private Sub cmdReadPNR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReadPNR.Click

        Try
            ClearForm()
            ReadPNR()
            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub cmdNoPNR_Click(sender As Object, e As EventArgs)

        Try
            ClearForm()
            NoPNR()
            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

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
            txtAirlineEntries.Clear()

            lstVessels.Items.Clear()

            lstSubDepartments.Items.Clear()
            txtSubdepartment.Enabled = (lstSubDepartments.Items.Count > 0)

            lstCRM.Items.Clear()
            txtCRM.Enabled = (lstCRM.Items.Count > 0)

            txtReference.Clear()
            cmbDepartment.Items.Clear()
            cmbDepartment.Text = ""
            cmbBookedby.Items.Clear()
            cmbBookedby.Text = ""
            cmbReasonForTravel.Items.Clear()
            cmbReasonForTravel.Text = ""
            cmbCostCentre.Items.Clear()
            cmbCostCentre.Text = ""

            cmdPNRWrite.Enabled = False
            cmdPNRWriteWithDocs.Enabled = False
            cmdPNROnlyDocs.Enabled = False

            mobjReadPNR.ExistingElements.Clear()

            mflgAPISUpdate = False
            mflgExpiryDateOK = False

            APISPrepareGrid()

        Catch ex As Exception
            Throw New Exception("ClearForm()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub SetEnabled()

        Try
            ' read PNR and Exit are always enabled
            cmdReadPNR.Enabled = True
            cmdExit.Enabled = True

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

            ' Customer is always needed

            txtCustomer.BackColor = lstCustomers.BackColor
            txtSubdepartment.BackColor = lstCustomers.BackColor
            txtCRM.BackColor = lstCustomers.BackColor
            If Not mobjReadPNR.NewElements Is Nothing Then
                If mobjReadPNR.NewElements.CustomerCode.AmadeusCommand = "" Then
                    cmdPNRWrite.Enabled = False
                    txtCustomer.BackColor = Color.Red
                End If

                ' if subdepartments exist they are by default madatory
                If mobjReadPNR.NewElements.CustomerCode.AmadeusCommand <> "" And lstSubDepartments.Items.Count > 0 And mobjReadPNR.NewElements.SubDepartmentCode.AmadeusCommand = "" Then
                    cmdPNRWrite.Enabled = False
                    txtSubdepartment.BackColor = Color.Red
                End If

                ' the code above is complete validation but allow entry without CRM in any case
                If mobjReadPNR.NewElements.CustomerCode.AmadeusCommand <> "" And lstCRM.Items.Count > 0 And mobjReadPNR.NewElements.CRMCode.AmadeusCommand = "" Then
                    txtCRM.BackColor = Color.Pink
                End If

                If mobjReadPNR.NewElements.BookedBy.AmadeusCommand = "" And cmbBookedby.Enabled Then
                    cmdPNRWrite.Enabled = False
                End If
                If mobjReadPNR.NewElements.CostCentre.AmadeusCommand = "" And cmbCostCentre.Enabled Then
                    cmdPNRWrite.Enabled = False
                End If
                If mobjReadPNR.NewElements.ReasonForTravel.AmadeusCommand = "" And cmbReasonForTravel.Enabled Then
                    cmdPNRWrite.Enabled = False
                End If
            End If

            cmdPNRWriteWithDocs.Enabled = cmdPNRWrite.Enabled And mflgAPISUpdate
            cmdPNROnlyDocs.Enabled = mflgAPISUpdate And Not mobjReadPNR.NewPNR
            dgvApis.Enabled = False

            txtReference.Enabled = True

            lblBookedByHighlight.Enabled = (cmbBookedby.Enabled)
            lblDepartmentHighlight.Enabled = (cmbDepartment.Enabled)
            lblReasonForTravelHighLight.Enabled = (cmbReasonForTravel.Enabled)
            lblCostCentreHighlight.Enabled = (cmbCostCentre.Enabled)

            SetLabelColor(lblBookedByHighlight)
            SetLabelColor(lblDepartmentHighlight)
            SetLabelColor(lblReasonForTravelHighLight)
            SetLabelColor(lblCostCentreHighlight)
        Catch ex As Exception
            Throw New Exception("SetEnabled()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub SetLabelColor(ByRef TextLabel As Label)
        Try
            If TextLabel.Enabled Then
                TextLabel.BackColor = Color.FromArgb(255, 128, 128)
            Else
                TextLabel.BackColor = Color.Silver
            End If
        Catch ex As Exception
            Throw New Exception("SetLabelColor()" & vbCrLf & ex.Message)
        End Try
    End Sub
    Private Sub ReadPNR()

        Dim pDMI As String
        Try
            With mobjReadPNR
                mflgReadPNR = False
                pDMI = .Read
                If .NumberOfPax = 0 Then
                    Throw New Exception("Need passenger names")
                End If
                If pDMI <> "" Then
                    If MessageBox.Show("There is a problem with your itinerary. Do you want to cancel the PNR Finisher?" & vbCrLf & pDMI, "Itinerary Check", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                        Throw New Exception("PNR Finisher cancelled because of itinerary check")
                    End If
                End If
                mflgReadPNR = True
                lblPNR.Text = .PnrNumber
                lblPax.Text = .PaxName
                lblSegs.Text = .Itinerary

                Dim pFromDate As Date = DateAdd(DateInterval.Month, -3, Today)

                pFromDate = DateSerial(Year(pFromDate), Month(pFromDate), 1)

                mobjAveragePrice.SetValues(pFromDate, .Itinerary)
                PrepareAirlinePoints()
            End With
            DisplayCustomer()
            APISDisplayPax(dgvApis, mobjReadPNR.PNR)

        Catch ex As Exception
            Throw New Exception("ReadPNR()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub NoPNR()

        Try
            mflgReadPNR = False
            lblPNR.Text = ""
            lblPax.Text = ""
            lblSegs.Text = ""
            'mobjNewAmadeusElements = New AmadeusElements.NewItems(.OfficeOfResponsibility, .CreationDate, .DepartureDate, .NumberOfPax)
            PrepareAirlinePoints()
            DisplayCustomer()
        Catch ex As Exception
            Throw New Exception("NoPNR()" & vbCrLf & ex.Message)
        End Try
    End Sub
    Private Sub DisplayCustomer()

        Dim pstrCustomerCode As String
        Dim pintSubDepartment As Integer
        Dim pstrCRM As String
        Dim pstrVesselName As String
        Dim pstrVesselRegistration As String

        Try
            With mobjReadPNR.ExistingElements
                pstrCustomerCode = .CustomerCode.Key
                pintSubDepartment = If(IsNumeric(.SubDepartmentCode.Key), CInt(.SubDepartmentCode.Key), 0)
                pstrCRM = .CRMCode.Key
                pstrVesselName = .VesselName.Key
                pstrVesselRegistration = .VesselFlag.Key

                mobjReadPNR.NewElements.ClearCustomerElements()

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
                        mobjReadPNR.NewElements.VesselNameForPNR.Clear()
                        mobjReadPNR.NewElements.VesselFlagForPNR.Clear()
                        txtVessel.Text = pVessel.Name
                    Else
                        mobjReadPNR.NewElements.SetVesselForPNR(pstrVesselName, pstrVesselRegistration)
                        txtVessel.Text = mobjReadPNR.NewElements.VesselNameForPNR.TextRequested & " REG " & mobjReadPNR.NewElements.VesselFlagForPNR.TextRequested
                    End If
                End If

                DisplayOldCustomProperty(cmbBookedby, mobjReadPNR.ExistingElements.BookedBy)
                DisplayOldCustomProperty(cmbDepartment, mobjReadPNR.ExistingElements.Department)
                DisplayOldCustomProperty(cmbReasonForTravel, mobjReadPNR.ExistingElements.ReasonForTravel)
                DisplayOldCustomProperty(cmbCostCentre, mobjReadPNR.ExistingElements.CostCentre)

                txtReference.Text = mobjReadPNR.ExistingElements.Reference.Key
                PrepareAirlinePoints()
            End If
        Catch ex As Exception
            Throw New Exception("DisplayCustomer()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub DisplayOldCustomProperty(ByRef cmbList As ComboBox, ByVal Item As PNR_Finisher.AmadeusExisting.Item)

        Try
            If Item.Key <> "" Then

                If cmbList.DropDownStyle = ComboBoxStyle.DropDown Then
                    If Item.Key <> "" Then
                        cmbList.Text = Item.Key
                    End If
                Else
                    For i As Short = 0 To cmbList.Items.Count - 1
                        If Item.Key.ToUpper = cmbList.Items(i).ToString.ToUpper Then
                            cmbList.SelectedIndex = i
                            Exit For
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Throw New Exception("DisplayOldCustomProperty(ByRef cmbList As ComboBox, ByVal Item As PNR_Finisher.AmadeusExisting.Item)" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub DisplayOldCustomProperty(ByRef cmbList As ComboBox, ByVal Item As String)

        Try
            If Item <> "" Then

                If cmbList.DropDownStyle = ComboBoxStyle.DropDown Then
                    cmbList.Text = Item
                Else
                    For i As Short = 0 To cmbList.Items.Count - 1
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

            txtAirlineEntries.Clear()

            For Each pSeg As s1aPNR.AirSegment In mobjReadPNR.AirSegments
                mobjAirlinePoints.Load(mobjCustomerSelected.ID, pSeg.Airline)
                For Each pItem As AirlinePoints.Item In mobjAirlinePoints.Values
                    'Dim pFound As Boolean = False
                    If txtAirlineEntries.Text.IndexOf(pItem.PointsCommand) < 0 Then
                        txtAirlineEntries.AppendText(pItem.PointsCommand & vbCrLf)
                    End If
                Next
            Next
            If mflgReadPNR Then
                For Each pSeg As s1aPNR.AirSegment In mobjReadPNR.AirSegments
                    mobjAirlineNotes.Load(pSeg.Airline)
                    For Each pItem As AirlineNotes.Item In mobjAirlineNotes.Values
                        With pItem
                            If Not .Seaman Or Not mobjVesselSelected Is Nothing Then
                                Dim pAmadeusText As String = .AmadeusText
                                'Dim pFound As Boolean = False

                                If pAmadeusText.Contains("<?VESSEL NAME>") Then
                                    If Not mobjVesselSelected Is Nothing Then
                                        '                                    pAmadeusText = pAmadeusText.Replace("<?VESSEL NAME>", mobjVesselSelected.Name)
                                        If mobjVesselSelected.Name Is Nothing Then
                                            pAmadeusText = pAmadeusText.Replace("<?VESSEL NAME>", mobjVesselSelected.Name)
                                        Else
                                            pAmadeusText = pAmadeusText.Replace("<?VESSEL NAME>", mobjVesselSelected.Name.Replace("(", "-").Replace(")", "-").Replace("&", "-"))
                                        End If
                                        'pAmadeusText = pAmadeusText.Replace("<?VESSEL NAME>", mobjVesselSelected.Name.Replace("(", "-").Replace(")", "-").Replace("&", "-"))
                                        'Else
                                        '    pFound = True
                                    End If
                                End If

                                If pAmadeusText.Contains("<?VESSEL REGISTRATION>") Then
                                    If Not mobjVesselSelected Is Nothing Then
                                        If mobjVesselSelected.Flag Is Nothing Then
                                            pAmadeusText = pAmadeusText.Replace("<?VESSEL REGISTRATION>", mobjVesselSelected.Flag)
                                        Else
                                            pAmadeusText = pAmadeusText.Replace("<?VESSEL REGISTRATION>", mobjVesselSelected.Flag.Replace("(", "-").Replace(")", "-").Replace("&", "-"))
                                        End If
                                        'pAmadeusText = pAmadeusText.Replace("<?VESSEL REGISTRATION>", mobjVesselSelected.Flag.Replace("(", "-").Replace(")", "-").Replace("&", "-"))
                                        'Else
                                        '    pFound = True
                                    End If
                                End If

                                If pAmadeusText.Contains("<?NBR OF PSGRS>") Then
                                    pAmadeusText = pAmadeusText.Replace("<?NBR OF PSGRS>", mobjReadPNR.NumberOfPax)
                                End If

                                If pAmadeusText.Contains("<?Segment selection>") Then
                                    pAmadeusText = pAmadeusText.Replace("<?Segment selection>", pSeg.ElementNo)
                                End If

                                Dim pAmadeusCommand As String
                                If .AmadeusElement.StartsWith("R") Then
                                    pAmadeusCommand = .AmadeusElement & " " & .AirlineCode & " " & pAmadeusText
                                ElseIf .AmadeusElement.StartsWith("S") Then
                                    pAmadeusCommand = .AmadeusElement & "-" & pAmadeusText
                                Else
                                    pAmadeusCommand = .AmadeusElement & " " & pAmadeusText
                                End If
                                If txtAirlineEntries.Text.IndexOf(pAmadeusCommand) < 0 Then
                                    txtAirlineEntries.AppendText(pAmadeusCommand & vbCrLf)
                                End If

                            End If
                        End With
                    Next
                Next
            End If

        Catch ex As Exception
            Throw New Exception("PrepareAirlinePoints()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub frmPNR_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        mflgLoading = True
        dgvApis.VirtualMode = False

        Try
            Dim mAmadeusUser As New AmadeusUser
            InitSettings(mAmadeusUser)
            If MySettings.AmadeusPCC <> "" And MySettings.AmadeusUser <> "" Then
                Text = "Athens PNR Finisher (21/10/2017 15:34) " & MySettings.AmadeusPCC & " " & MySettings.AmadeusUser
            Else
                Throw New Exception("Please start Amadeus and restart the program")
            End If
            If CheckOptions() Then
                ' finisher tab
                mflgReadPNR = False
                ClearForm()
                SetEnabled()
                PrepareForm()
                APISPrepareGrid()

                ' itinerary tab
                LoadRemarks()
                If MySettings.AirportName = 0 Then
                    optItnAirportCode.Checked = True
                ElseIf MySettings.AirportName = 1 Then
                    optItnAirportname.Checked = True
                ElseIf MySettings.AirportName = 3 Then
                    optItnAirportCityName.Checked = True
                    optItnAirportBoth.Checked = True
                End If
                optItnFormatDefault.Checked = True
                chkItnVessel.Checked = MySettings.Vessel
                chkItnClass.Checked = MySettings.ClassOfService
                chkItnAirlineLocator.Checked = MySettings.AirlineLocator
                chkItnTickets.Checked = MySettings.Tickets
                chkItnPaxSegPerTicket.Checked = MySettings.PaxSegPerTkt
                chkItnSeating.Checked = MySettings.Seating
                chkItnStopovers.Checked = MySettings.ShowStopovers
                chkItnTerminal.Checked = MySettings.ShowTerminal
                chkFlyingTime.Checked = MySettings.FlyingTime
                chkItnCostCentre.Checked = MySettings.CostCentre

                chkElecItemsBan.Checked = MySettings.BanElectricalEquipment
                chkItnBrazilText.Checked = MySettings.BrazilText
                chkItnUSAText.Checked = MySettings.USAText

                cmdItnReadPNR.Enabled = False
                cmdItnReadQueue.Enabled = False
            Else
                Me.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.Close()
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

                    SelectCustomer(lstCustomers.Items(0))
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
                mobjReadPNR.NewElements.SetItem(mobjSubDepartmentSelected)
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
                mobjReadPNR.NewElements.SetItem(mobjCRMSelected)
            End If
            lstCRM.Items.Clear()

            If Not mobjCustomerSelected Is Nothing Then
                pobjCRM.Load(mobjCustomerSelected.ID)

                For Each pCRM As CRM.Item In pobjCRM.Values
                    If SearchString = "" Or pCRM.ToString.ToUpper.Contains(SearchString.ToUpper) Then
                        lstCRM.Items.Add(pCRM)
                    End If
                Next
                If mobjReadPNR.NewElements.CRMCode.TextRequested <> "" And lstCRM.Items.Count = 1 Then
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
                    If mobjReadPNR.NewElements.VesselName.TextRequested = "" Or pVessel.ToString.ToUpper.Contains(mobjReadPNR.NewElements.VesselName.TextRequested.ToUpper) Then
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

            If Not mobjCustomerSelected Is Nothing Then
                For Each pProp As CustomProperties.Item In mobjCustomerSelected.CustomerProperties.Values
                    If pProp.CustomPropertyID = CustomProperties.CustomPropertyIDValue.BookedBy Then
                        PrepareCustomProperty(cmbBookedby, pProp)
                    ElseIf pProp.CustomPropertyID = CustomProperties.CustomPropertyIDValue.Department Then
                        PrepareCustomProperty(cmbDepartment, pProp)
                    ElseIf pProp.CustomPropertyID = CustomProperties.CustomPropertyIDValue.ReasonFortravel Then
                        PrepareCustomProperty(cmbReasonForTravel, pProp)
                    ElseIf pProp.CustomPropertyID = CustomProperties.CustomPropertyIDValue.CostCentre Then
                        PrepareCustomProperty(cmbCostCentre, pProp)
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

    Private Sub txtCustomer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustomer.TextChanged

        Try
            If Not mflgLoading Then
                PopulateCustomerList(txtCustomer.Text)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub SelectCustomer(ByVal pCustomer As Customers.CustomerItem)

        Try
            'TODO
            mobjReadPNR.NewElements.ClearCustomerElements()
            mobjAirlinePoints.Clear()
            mobjAirlineNotes.Clear()
            mobjCustomerSelected = pCustomer
            txtCustomer.Text = pCustomer.ToString
            mobjReadPNR.NewElements.SetItem(mobjCustomerSelected)

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
        Catch ex As Exception
            Throw New Exception("SelectCustomer()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub txtSubdepartment_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSubdepartment.TextChanged

        Try
            If Not mflgLoading Then
                PopulateSubdepartmentsList(txtSubdepartment.Text)
            End If

            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub txtCRM_TextChanged(sender As Object, e As EventArgs) Handles txtCRM.TextChanged

        Try
            If Not mflgLoading Then
                PopulateCRMList(txtCRM.Text)
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
            mobjReadPNR.NewElements.SetItem(mobjSubDepartmentSelected)

            SetEnabled()
        Catch ex As Exception
            Throw New Exception("SelectSubDepartment()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub SelectCRM(ByVal pCRM As CRM.Item)

        Try
            mobjCRMSelected = pCRM
            txtCRM.Text = pCRM.ToString
            mobjReadPNR.NewElements.SetItem(mobjCRMSelected)

            SetEnabled()
        Catch ex As Exception
            Throw New Exception("SelectCRM()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub SelectVessel(ByVal pVessel As Vessels.Item)

        Try
            mobjVesselSelected = pVessel
            txtVessel.Text = pVessel.ToString
            mobjReadPNR.NewElements.SetItem(mobjVesselSelected)
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
                SelectCustomer(lstCustomers.SelectedItem)
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
                mobjReadPNR.NewElements.SetVesselForPNR("", "")
                mobjReadPNR.NewElements.VesselName.SetText(txtVessel.Text)
                PopulateVesselsList()
                'mobjReadPNR.NewElements.SetVesselForPNR("", "")
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
    Private Sub PNRWrite(ByVal WritePNR As Boolean, ByVal WriteDocs As Boolean)

        Try
            UpdatePNR(WritePNR, WriteDocs)
            mflgReadPNR = False
            ClearForm()
            SetEnabled()
        Catch ex As Exception
            Throw New Exception("PNRWrite(" & WritePNR & ", " & WriteDocs & ")" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub UpdatePNR(ByVal WritePNR As Boolean, ByVal WriteDocs As Boolean)

        Try
            Dim pPNR As New ReadPNR

            pPNR.Read()

            If pPNR.PnrNumber = mobjReadPNR.PnrNumber And pPNR.PaxName = mobjReadPNR.PaxName And pPNR.Itinerary = mobjReadPNR.Itinerary Then
                mobjReadPNR.SendNewAmadeusEntries(WritePNR, WriteDocs, mflgExpiryDateOK, dgvApis, txtAirlineEntries)
            Else
                Throw New Exception("PNR has been changed since read" & vbCrLf & "Please read again and re-enter data", New Exception("DifferentPNR"))
            End If
        Catch ex As Exception
            Throw New Exception("UpdatePNR()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub ShowOptionsForm()
        Try
            Dim pFrm As New frmOptions
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

    Private Sub llbTables_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbTables.LinkClicked

        Try
            Dim pFrm As New frmTables
            pFrm.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmdOneTimeVessel_Click(sender As Object, e As EventArgs) Handles cmdOneTimeVessel.Click

        Try
            Dim pFrm As New frmVesselForPNR

            If pFrm.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
                With mobjReadPNR.NewElements
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
                mobjReadPNR.NewElements.SetBookedBy(cmbBookedby.Text)
            End If

            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmbReasonForTravel_TextChanged(sender As Object, e As EventArgs) Handles cmbReasonForTravel.TextChanged

        Try
            If Not mflgLoading Then
                mobjReadPNR.NewElements.SetReasonForTravel(cmbReasonForTravel.Text)
            End If

            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmbCostCentre_TextChanged(sender As Object, e As EventArgs) Handles cmbCostCentre.TextChanged

        Try
            If Not mflgLoading Then
                mobjReadPNR.NewElements.SetCostCentre(cmbCostCentre.Text)
            End If

            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub txtReference_TextChanged(sender As Object, e As EventArgs) Handles txtReference.TextChanged

        Try
            If Not mflgLoading Then
                mobjReadPNR.NewElements.SetReference(txtReference.Text)
            End If

            SetEnabled()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmbDepartment_TextChanged(sender As Object, e As EventArgs) Handles cmbDepartment.TextChanged

        Try
            If Not mflgLoading Then
                mobjReadPNR.NewElements.SetDepartment(cmbDepartment.Text)
            End If

            SetEnabled()
        Catch ex As Exception
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmdItnReadPNR_Click(sender As Object, e As EventArgs) Handles cmdItnReadPNR.Click

        Try
            Cursor = System.Windows.Forms.Cursors.WaitCursor

            ProcessRequestedPNRs(txtItnPNR)

            rtbItnDoc.SelectAll()
            Clipboard.Clear()
            Clipboard.SetText(rtbItnDoc.Rtf, TextDataFormat.Rtf)
            cmdItnRefresh.Enabled = False
            Cursor = Cursors.Default
            MessageBox.Show("Ready", "Read PNR", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmdItnReadQueue_Click(sender As Object, e As EventArgs) Handles cmdItnReadQueue.Click

        Try
            If optItnFormatMSReport.Checked Then
                If ItnReadFromToDates() = Windows.Forms.DialogResult.Cancel Then
                    Exit Sub
                End If
            End If
            Cursor = System.Windows.Forms.Cursors.WaitCursor
            txtItnPNR.Text = mobjAmadeus.RetrievePNRsFromQueue(txtItnPNR.Text)

            ProcessRequestedPNRs(txtItnPNR)

            rtbItnDoc.SelectAll()
            Clipboard.Clear()
            Clipboard.SetText(rtbItnDoc.Rtf, TextDataFormat.Rtf)
            cmdItnRefresh.Enabled = False
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
            rtbItnDoc.Clear()
            If Not RefreshOnly Then
                ReDim mPaxNames(0)
                readAmadeus("")
            End If

            makeRTBDoc()
            PaxNamesToBold()
        Catch ex As Exception
            Throw New Exception("ProcessRequestedPNRs(RefreshOnly)" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub ProcessRequestedPNRs(ByVal txtPNR As TextBox)

        Try
            Dim pPNR() As String = txtPNR.Text.Split(vbCrLf)
            Dim pPNRsOutsideRange As New System.Text.StringBuilder
            pPNRsOutsideRange.Clear()

            rtbItnDoc.Clear()
            ReDim mPaxNames(0)

            For i As Integer = pPNR.GetLowerBound(0) To pPNR.GetUpperBound(0)
                If pPNR(i).Trim <> "" Then
                    readAmadeus(pPNR(i).Trim)
                    If Not optItnFormatMSReport.Checked Or (mobjAmadeus.LastSegment.DepartureDate >= mItnFromDate And mobjAmadeus.LastSegment.DepartureDate <= mItnToDate) Then
                        makeRTBDoc()
                    Else
                        pPNRsOutsideRange.Append(MakeRTBMSReportOutsiderange)
                    End If
                End If
            Next
            If pPNRsOutsideRange.Length > 0 Then
                rtbItnDoc.Text &= vbCrLf & "OUTSIDE DATE RANGE" & vbCrLf
                rtbItnDoc.Text &= pPNRsOutsideRange.ToString
            End If
            PaxNamesToBold()
        Catch ex As Exception
            Throw New Exception("ProcessRequestedPNRs(txtPNR)" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub makeRTBDoc()

        Dim pString As New System.Text.StringBuilder

        pString.Clear()
        mMaxString = 80

        Try
            If optItnFormatMSReport.Checked Then

                'TODO - Fix length of output line total 78 characters including spaces



                If optItnFormatMSReport.Checked AndAlso rtbItnDoc.TextLength = 0 Then
                    pString.AppendLine("FROM " & mItnFromDate.ToShortDateString & " : To " & mItnToDate.ToShortDateString)
                    pString.AppendLine("Last Name" & vbTab & "First Name" & vbTab & "ID No." & vbTab & "Department" & vbTab & "Vessel Name" & vbTab & "Date Of Travel" & vbTab & "Airline" & vbTab & "Flight No." & vbTab & "Dep.Time" & vbTab & "Dep.City" & vbTab & "Arr.Time" & vbTab & "Arr.City" & vbTab & "PNR" & vbTab & "PaxNo")
                End If
                pString.Append(MakeRTBMSReport)
            Else
                pString.Append(MakeRTBDocPart1)
                pString.Append(MakeRTBDocTickets)
                If Not optItnFormatSeaChefs.Checked And mMaxString > 0 Then
                    pString.AppendLine(StrDup(HeaderLength, "-"))
                End If
                pString.AppendLine()
                pString.Append(MakeRTBDocRemarks)
                pString.Append(MakeRTBDocCloseOff)
            End If
            rtbItnDoc.Text &= pString.ToString
        Catch ex As Exception
            Throw New Exception("makeRTBDoc()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Function MakeRTBDocCloseOff() As String

        Try
            Dim pString As New System.Text.StringBuilder

            pString.Clear()
            If MySettings.BrazilText Then
                pString.AppendLine(" ")
                pString.AppendLine("***Please be advised that all Seamen entering Brazil are required to have their joining letters, or letter of guarantee written in Portuguese.  These must be provided by their respective shipping companies.  Letters in English are no longer accepted.***")
                pString.AppendLine(" ")
            End If

            If MySettings.USAText Then
                pString.AppendLine(" ")
                pString.AppendLine("***Please note, all electronic equipment must be fully charged when travelling to/from the US.***")
                pString.AppendLine("**TSA SECURE FLIGHT PROGRAMME**")
                pString.AppendLine("**All passengers who intend to travel to the United States without a U.S. Visa under the terms of the Visa Waiver Program (VWP) must obtain an electronic preauthorisation or ESTA prior to boarding a flight to the U.S.**")
                pString.AppendLine("Passengers who do not obtain ESTA prior to travel are subject to denied boarding.")
                pString.AppendLine("A third party, such as a relative, friend or travel agent may submit an ESTA application on behalf of a VWP traveller.")
                pString.AppendLine("For more details on the Visa Waiver Program, a list of VWP eligible countries and the new ESTA process, please visit the ESTA website at http://www.cbp.gov/ESTA")
                pString.AppendLine(" ")
            End If

            If MySettings.BanElectricalEquipment Then
                pString.AppendLine("Important Security information")
                pString.AppendLine(" ")
                pString.AppendLine("UK and US authorities have imposed a ban on electrical items larger than mobile phones being carried in the cabin of inbound flights from specific countries.")
                pString.AppendLine("These items, including laptops, e-readers and tablets, must now be placed in your hold baggage.")
                pString.AppendLine("For more information please contact your ATPI consultant or refer to the airline web site.")
                pString.AppendLine(" ")
            End If

            Return pString.ToString
        Catch ex As Exception
            Throw New Exception("MakeRTBDocCloseOff()" & vbCrLf & ex.Message)
        End Try

    End Function
    Private Sub PaxNamesToBold()

        Try
            Dim pFont As Font = rtbItnDoc.SelectionFont

            For i As Integer = 1 To mPaxNames.GetUpperBound(0)
                rtbItnDoc.Select(mPaxNames(i).StartPos - 1, mPaxNames(i).EndPos - mPaxNames(i).StartPos + 1)
                rtbItnDoc.SelectionFont = New Font(pFont, FontStyle.Bold)
            Next
        Catch ex As Exception
            Throw New Exception("PaxNamesToBold()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Function MakeRTBMSReport() As String

        Try
            Dim pString As New System.Text.StringBuilder

            With mobjAmadeus
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
        Catch ex As Exception
            Throw New Exception("MakeRTBMSReport()" & vbCrLf & ex.Message)
        End Try

    End Function

    Private Function MakeRTBMSReportOutsiderange() As String

        Try
            Dim pString As New System.Text.StringBuilder

            With mobjAmadeus
                If .HasSegments Then
                    Dim pDepTime As String = ""
                    Dim pArrTime As String = ""
                    For Each pobjPax As gtmAmadeusPax In .Passengers.Values
                        For Each pSeg As gtmAmadeusSeg In .Segments.Values
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

    End Function
    Private Function MakeRTBDocPart1() As String

        Try
            Dim pString As New System.Text.StringBuilder
            Dim pAirlineLocator As String = ""

            Dim pobjSeg As gtmAmadeusSeg
            Dim pobjPax As gtmAmadeusPax

            pString.Clear()

            With mobjAmadeus

                Dim iPaxCount As Integer = 0
                If optItnFormatSeaChefs.Checked Then
                    pString.AppendLine("FOR PASSENGER" & If(.Passengers.Count > 1, "(S)", ""))
                End If
                For Each pobjPax In .Passengers.Values
                    iPaxCount = iPaxCount + 1
                    If optItnFormatSeaChefs.Checked Then
                        pString.AppendLine(pobjPax.PaxName)
                    Else
                        pString.AppendLine(pobjPax.ElementNo & " " & pobjPax.PaxName & " " & pobjPax.PaxID)
                    End If
                Next pobjPax
                If iPaxCount = 0 Then
                    pString.AppendLine("PASSENGER INFORMATION NOT AVAILABLE")
                End If

                ReDim Preserve mPaxNames(mPaxNames.GetUpperBound(0) + 1)
                mPaxNames(mPaxNames.GetUpperBound(0)).StartPos = rtbItnDoc.Text.Length + 1
                rtbItnDoc.Text &= pString.ToString
                mPaxNames(mPaxNames.GetUpperBound(0)).EndPos = rtbItnDoc.Text.Length

                pString.Clear()
                Dim pTemp As String = ""
                If Not optItnFormatSeaChefs.Checked And MySettings.Vessel And .VesselName <> "" Then
                    pTemp &= "VESSEL     : " & .VesselName
                End If
                If Not optItnFormatSeaChefs.Checked And MySettings.CostCentre And .CostCentre <> "" Then
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

                If optItnFormatDefault.Checked Then
                    pHeader.Append("Flight ")
                    If MySettings.ClassOfService Then
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
                            pHeader.Append("Origin " & StrDup(.MaxCityNameLength - 5, " ") & "Destination" & StrDup(.MaxCityNameLength - 9, " "))
                        Case 4
                            pHeader.Append("Origin " & StrDup(.MaxCityNameLength - 1, " ") & "Destination" & StrDup(.MaxCityNameLength - 5, " "))
                    End Select
                    'pHeader.Append("St ")
                    pHeader.Append("Dep   ")
                    pHeader.Append("Arr   ")
                    If MySettings.FlyingTime Then
                        pHeader.Append(" EFT  ")
                    End If
                    pHeader.Append("ArrDte ")
                    pHeader.Append(If(MySettings.AirlineLocator, "AL Locator", ""))
                    pHeader.Append(" - BagAl")

                    HeaderLength = pHeader.Length

                    pString.AppendLine(StrDup(HeaderLength, "-"))
                    pString.AppendLine(pHeader.ToString)
                    pString.AppendLine(StrDup(HeaderLength, "-"))
                ElseIf optItnFormatSeaChefs.Checked Then
                    pHeader.Append("Flight ")
                    pHeader.Append("Date  ")
                    If chkItnSeaChefsWithCode.Checked Then
                        pHeader.Append("Org    " & StrDup(.MaxAirportShortNameLength - 1, " ") & "Dest       " & StrDup(.MaxAirportShortNameLength - 5, " "))
                    Else
                        pHeader.Append("Org    " & StrDup(.MaxAirportShortNameLength - 5, " ") & "Dest       " & StrDup(.MaxAirportShortNameLength - 9, " "))
                    End If
                    pHeader.Append("Dep   ")
                    pHeader.Append("Arr   ")
                    pHeader.Append("Term   ")
                    pHeader.Append("Status")
                    pHeader.Append("   BagAl")
                    HeaderLength = pHeader.Length

                    pString.AppendLine(StrDup(HeaderLength, "-"))
                    pString.AppendLine(pHeader.ToString)
                    pString.AppendLine(StrDup(HeaderLength, "-"))
                End If

                Dim iSegCount As Integer = 0
                For Each pobjSeg In .Segments.Values
                    iSegCount = iSegCount + 1
                    Dim pSeg As New System.Text.StringBuilder

                    If optItnFormatSeaChefs.Checked Then
                        pSeg.Append(pobjSeg.Airline & pobjSeg.FlightNo.PadLeft(4) & " ")
                        pSeg.Append(pobjSeg.DepartureDateIATA & " ")
                        If chkItnSeaChefsWithCode.Checked Then
                            pSeg.Append(pobjSeg.BoardPoint & " " & pobjSeg.BoardAirportShortName.PadRight(.MaxAirportShortNameLength + 1, " ").Substring(0, .MaxAirportShortNameLength + 1) & " ")
                            pSeg.Append(pobjSeg.OffPoint & " " & pobjSeg.OffPointAirportShortName.PadRight(.MaxAirportShortNameLength + 1, " ").Substring(0, .MaxAirportShortNameLength + 1) & " ")
                        Else
                            pSeg.Append(pobjSeg.BoardAirportShortName.PadRight(.MaxAirportShortNameLength + 1, " ").Substring(0, .MaxAirportShortNameLength + 1) & " ")
                            pSeg.Append(pobjSeg.OffPointAirportShortName.PadRight(.MaxAirportShortNameLength + 1, " ").Substring(0, .MaxAirportShortNameLength + 1) & " ")
                        End If
                        If pobjSeg.Text.Substring(35, 4) = "FLWN" Then
                            pSeg.Append("FLWN")
                        Else
                            pSeg.Append(Format(pobjSeg.DepartTime, "HHmm") & "  ")
                            pSeg.Append(Format(pobjSeg.ArriveTime, "HHmm"))
                            If pobjSeg.ArrivalDate > pobjSeg.DepartureDate Then
                                pSeg.Append("+1 ")
                            ElseIf pobjSeg.ArrivalDate < pobjSeg.DepartureDate Then
                                pSeg.Append("-1 ")
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
                            pSeg.Append("    " & mobjAmadeus.AllowanceForSegment(pobjSeg.BoardPoint, pobjSeg.OffPoint, pobjSeg.Airline)) ', ""))
                            If pAirlineLocator.IndexOf(pobjSeg.AirlineLocator.Trim) = -1 Then
                                If pAirlineLocator <> "" Then
                                    pAirlineLocator &= " - "
                                End If
                                pAirlineLocator &= pobjSeg.AirlineLocator.Trim
                            End If
                        End If
                    Else
                        pSeg.Append(pobjSeg.Airline & pobjSeg.FlightNo.PadLeft(4) & " ")
                        If MySettings.ClassOfService Then
                            pSeg.Append(pobjSeg.ClassOfService & " ")
                        End If
                        pSeg.Append(pobjSeg.DepartureDateIATA & " ")
                        Select Case MySettings.AirportName
                            Case 0 'code
                                pSeg.Append(pobjSeg.BoardPoint & " " & pobjSeg.OffPoint & " ")
                            Case 1 'airport name
                                pSeg.Append(pobjSeg.BoardAirportName.PadRight(.MaxAirportNameLength + 1, " ").Substring(0, .MaxAirportNameLength + 1) & " " &
                                            pobjSeg.OffPointAirportName.PadRight(.MaxAirportNameLength + 1, " ").Substring(0, .MaxAirportNameLength + 1) & " ")
                            Case 2 'code and airport
                                pSeg.Append(pobjSeg.BoardPoint & " " & pobjSeg.BoardAirportName.PadRight(.MaxAirportNameLength + 1, " ").Substring(0, .MaxAirportNameLength + 1) & " " &
                                            pobjSeg.OffPoint & " " & pobjSeg.OffPointAirportName.PadRight(.MaxAirportNameLength + 1, " ").Substring(0, .MaxAirportNameLength + 1) & " ")
                            Case 3 'city name
                                pSeg.Append(pobjSeg.BoardCityName.PadRight(.MaxCityNameLength + 1, " ").Substring(0, .MaxCityNameLength + 1) & " " &
                                            pobjSeg.OffPointCityName.PadRight(.MaxCityNameLength + 1, " ").Substring(0, .MaxCityNameLength + 1) & " ")
                            Case 4 'code and city
                                pSeg.Append(pobjSeg.BoardPoint & " " & pobjSeg.BoardCityName.PadRight(.MaxCityNameLength + 1, " ").Substring(0, .MaxCityNameLength + 1) & " " &
                                            pobjSeg.OffPoint & " " & pobjSeg.OffPointCityName.PadRight(.MaxCityNameLength + 1, " ").Substring(0, .MaxCityNameLength + 1) & " ")
                        End Select
                        If pobjSeg.Text.Substring(35, 4) = "FLWN" Then
                            pSeg.Append("FLWN")
                        Else
                            'pSeg.Append(pobjSeg.Status.PadRight(3))
                            pSeg.Append(Format(pobjSeg.DepartTime, "HHmm") & "  ")
                            pSeg.Append(Format(pobjSeg.ArriveTime, "HHmm") & "  ")
                            If MySettings.FlyingTime Then
                                pSeg.Append(pobjSeg.EstimatedFlyingTime & " ")
                            End If
                            pSeg.Append(pobjSeg.ArrivalDateIATA & "   ")
                            pSeg.Append(If(MySettings.AirlineLocator, pobjSeg.AirlineLocator.PadRight(9, " "), ""))
                            pSeg.Append(" - " & mobjAmadeus.AllowanceForSegment(pobjSeg.BoardPoint, pobjSeg.OffPoint, pobjSeg.Airline)) ', ""))
                            If pobjSeg.Status = "HL" Then
                                pSeg.Append("   WAITLISTED")
                            End If
                            If MySettings.ShowTerminal And pobjSeg.DepartTerminal <> "" Then
                                pSeg.Append("   " & pobjSeg.DepartTerminal)
                            End If
                        End If
                    End If

                    pString.AppendLine(pSeg.ToString)

                    If Not optItnFormatPlain.Checked Then
                        If pobjSeg.OperatedBy <> "" Then
                            pString.AppendLine(StrDup(13, " ") & pobjSeg.OperatedBy)
                        End If
                        If (optItnFormatSeaChefs.Checked Or MySettings.ShowStopovers) And pobjSeg.Stopovers <> "" Then
                            pString.AppendLine("             *INTERMEDIATE STOP*  " & pobjSeg.Stopovers)
                        End If
                    End If

                    If pSeg.ToString.Length > mMaxString Then
                        mMaxString = pSeg.ToString.Length
                    End If
                Next pobjSeg

                If iSegCount = 0 Then
                    pString.AppendLine("ROUTING INFORMATION NOT AVAILABLE")
                End If

                If .RequestedPNR <> "" Then
                    pString.AppendLine(" ")
                    If optItnFormatSeaChefs.Checked Then
                        pString.AppendLine("ATPI REF: " & .RequestedPNR)
                        If pAirlineLocator <> "" Then
                            pString.AppendLine("AIRLINE REF: " & pAirlineLocator)
                        End If
                    Else
                        pString.AppendLine("ATPI Booking Reference: " & .RequestedPNR)
                    End If
                End If

            End With

            Return pString.ToString
        Catch ex As Exception
            Throw New Exception("MakeRTBDocPart1()" & vbCrLf & ex.Message)
        End Try

    End Function

    Private Function MakeRTBDocRemarks() As String

        Try
            Dim pString As New System.Text.StringBuilder
            pString.Clear()

            For iRem As Integer = 0 To lstItnRemarks.CheckedItems.Count - 1
                pString.AppendLine(lstItnRemarks.CheckedItems(iRem).ToString)
            Next

            Return pString.ToString
        Catch ex As Exception
            Throw New Exception("MakeRTBDocRemarks()" & vbCrLf & ex.Message)
        End Try

    End Function

    Private Function MakeRTBDocTickets() As String

        Try

            Dim pString As New System.Text.StringBuilder
            pString.Clear()

            With mobjAmadeus
                If (optItnFormatSeaChefs.Checked Or MySettings.Tickets) And .Tickets.Count >= 1 Then
                    If optItnFormatDefault.Checked Then
                        pString.AppendLine(StrDup(HeaderLength, "-"))
                    ElseIf optItnFormatPlain.Checked Then
                        pString.AppendLine()
                    End If
                    If Not optItnFormatSeaChefs.Checked Then
                        Dim pHeader As String = "Ticket Number   "
                        If MySettings.PaxSegPerTkt Then
                            pHeader &= "Routing      Passenger"
                        End If
                        pString.AppendLine(pHeader)
                        If Not optItnFormatPlain.Checked Then
                            pString.AppendLine(StrDup(HeaderLength, "-"))
                        End If
                    End If

                    If optItnFormatSeaChefs.Checked Then
                        For Each pobjPax In .Passengers.Values
                            pString.AppendLine()
                            pString.AppendLine(pobjPax.PaxName)
                            For Each tkt As gtmTicket In .Tickets.Values
                                If tkt.Pax.Trim = pobjPax.PaxName.Trim Then
                                    Dim pFF As String = mobjAmadeus.FrequentFlyerNumber(tkt.AirlineCode, tkt.Pax.Substring(0, tkt.Pax.Length - 2).Trim)
                                    If pFF <> "" Then
                                        pFF = "Frequent Flyer Number: " & pFF
                                    End If
                                    pString.AppendLine(If(tkt.TicketType <> "PAX", tkt.TicketType & " ", "ETICKET NUMBER: ") & tkt.IssuingAirline & "-" & tkt.Document & " " & tkt.AirlineCode & " " & pFF)
                                End If
                            Next
                        Next
                    Else
                        For Each tkt As gtmTicket In .Tickets.Values
                            If tkt.eTicket Then
                                If MySettings.PaxSegPerTkt Then

                                    'todo - Issuing airline is code, we need airline 2 letter code for frequent flyer or maybe ff element has airline number code?
                                    Dim pFF As String = mobjAmadeus.FrequentFlyerNumber(tkt.AirlineCode, tkt.Pax.Substring(0, tkt.Pax.Length - 2).Trim)
                                    If pFF <> "" Then
                                        pFF = "Frequent Flyer Number: " & pFF
                                    End If
                                    pString.AppendLine(If(tkt.TicketType <> "PAX", tkt.TicketType & " ", "") & tkt.IssuingAirline & "-" & tkt.Document & "  " & tkt.Segs.Substring(0, 10) & "   " & tkt.Pax.Substring(0, tkt.Pax.Length - 2) & "  " & pFF)
                                    For i As Integer = 12 To tkt.Segs.Length - 10 Step 12
                                        pString.AppendLine(If(tkt.TicketType <> "PAX", "    ", "") & StrDup(16, " ") & tkt.Segs.Substring(i, 10))
                                    Next
                                Else
                                    pString.AppendLine(If(tkt.TicketType <> "PAX", tkt.TicketType & " ", "") & tkt.IssuingAirline & "-" & tkt.Document)
                                End If
                            End If
                        Next
                    End If

                End If

                If optItnFormatSeaChefs.Checked Or MySettings.Seating Then
                    If .Seats <> "" Then
                        If Not optItnFormatPlain.Checked Then
                            pString.AppendLine(StrDup(HeaderLength, "-"))
                        End If
                        pString.AppendLine("Seat Assignment")
                        If Not optItnFormatPlain.Checked Then
                            pString.AppendLine(StrDup(HeaderLength, "-"))
                        End If
                        pString.AppendLine(.Seats & vbCrLf)
                    End If
                End If

            End With

            Return pString.ToString
        Catch ex As Exception
            Throw New Exception("MakeRTBDocTickets()" & vbCrLf & ex.Message)
        End Try

    End Function

    Private Sub cmdItnReadCurrent_Click(sender As Object, e As EventArgs) Handles cmdItnReadCurrent.Click

        Try
            ReadPNRandCreateItn(False)
            cmdItnRefresh.Enabled = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

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
            rtbItnDoc.SelectAll()

            Clipboard.Clear()
            Clipboard.SetText(rtbItnDoc.Rtf, TextDataFormat.Rtf)
            Clipboard.SetText(rtbItnDoc.SelectedText, TextDataFormat.Text)

            Cursor = Cursors.Default
            If Not RefreshOnly Then
                MessageBox.Show("Ready", "Read PNR", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            Throw New Exception("ReadPNRandCreateItn" & vbCrLf & ex.Message)
        End Try

    End Sub
    Private Sub optAirportCode_CheckedChanged(sender As Object, e As EventArgs) Handles optItnAirportCode.CheckedChanged

        Try
            MySettings.AirportName = 0
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub optAirportname_CheckedChanged(sender As Object, e As EventArgs) Handles optItnAirportname.CheckedChanged

        Try
            MySettings.AirportName = 1
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub optAirportBoth_CheckedChanged(sender As Object, e As EventArgs) Handles optItnAirportBoth.CheckedChanged

        Try
            MySettings.AirportName = 2
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub optAirportCityName_CheckedChanged(sender As Object, e As EventArgs) Handles optItnAirportCityName.CheckedChanged

        Try
            MySettings.AirportName = 3
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub optAirportCityBoth_CheckedChanged(sender As Object, e As EventArgs) Handles optItnAirportCityBoth.CheckedChanged

        Try
            MySettings.AirportName = 4
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkVessel_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnVessel.CheckedChanged

        Try
            MySettings.Vessel = chkItnVessel.Checked
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkClass_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnClass.CheckedChanged

        Try
            MySettings.ClassOfService = chkItnClass.Checked
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkAirlineLocator_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnAirlineLocator.CheckedChanged

        Try
            MySettings.AirlineLocator = chkItnAirlineLocator.Checked
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkTickets_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnTickets.CheckedChanged

        Try
            MySettings.Tickets = chkItnTickets.Checked
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkPaxSegPerTicket_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnPaxSegPerTicket.CheckedChanged

        Try
            MySettings.PaxSegPerTkt = chkItnPaxSegPerTicket.Checked
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkSeating_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnSeating.CheckedChanged

        Try
            MySettings.Seating = chkItnSeating.Checked
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkTerminal_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnTerminal.CheckedChanged

        MySettings.ShowTerminal = chkItnTerminal.Checked
        MySettings.Save()
        If cmdItnRefresh.Enabled Then
            ReadPNRandCreateItn(True)
        End If

    End Sub

    Private Sub chkStopovers_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnStopovers.CheckedChanged

        Try
            MySettings.ShowStopovers = chkItnStopovers.Checked
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkFlyingTime_CheckedChanged(sender As Object, e As EventArgs) Handles chkFlyingTime.CheckedChanged

        Try
            MySettings.FlyingTime = chkFlyingTime.Checked
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkItnCostCentre_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnCostCentre.CheckedChanged

        Try
            MySettings.CostCentre = chkItnCostCentre.Checked
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub chkElecItemsBan_CheckedChanged(sender As Object, e As EventArgs) Handles chkElecItemsBan.CheckedChanged

        Try
            MySettings.BanElectricalEquipment = chkElecItemsBan.Checked
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkBrazilText_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnBrazilText.CheckedChanged

        Try
            MySettings.BrazilText = chkItnBrazilText.Checked
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub chkUSAText_CheckedChanged(sender As Object, e As EventArgs) Handles chkItnUSAText.CheckedChanged

        Try
            MySettings.USAText = chkItnUSAText.Checked
            MySettings.Save()
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub txtPNR_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtItnPNR.TextChanged

        Try
            cmdItnReadPNR.Enabled = (txtItnPNR.Text.Trim.Length >= 6)
            cmdItnReadQueue.Enabled = (txtItnPNR.Text.Trim.Length >= 2)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub readAmadeus(ByVal RecordLocator As String)

        Try
            If RecordLocator = "" Then
                mobjAmadeus.CancelError = True
            Else
                mobjAmadeus.CancelError = False
            End If
            mobjAmadeus.ReadPNR(RecordLocator, optItnFormatMSReport.Checked)
        Catch ex As Exception
            Throw New Exception("readAmadeus()" & vbCrLf & ex.Message)
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
            OSMRefreshVessels(lstOSMVessels, chkOSMVesselInUse.Checked)
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
            ListBox_DrawItem(sender, e)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub lstOSMVessels_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstOSMVessels.SelectedIndexChanged

        Try
            If Not mflgLoading Then
                OSMShowSelectedVesselEmails()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub OSMShowSelectedVesselEmails()

        Try

            OSMDisplayEmails(lstOSMVessels, lstOSMToEmail, lstOSMCCEmail, lstOSMAgents)
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

    Private Sub OSMWebCreate()

        Try
            webOSMDoc.DocumentText = OSMWebHeader()
            cmdOSMCopyDocument.Enabled = True
        Catch ex As Exception
            Throw New Exception("OSMWebCreate()" & vbCrLf & ex.Message)
        End Try
    End Sub
    Private Function OSMWebHeader() As String

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


            'Dim pJoinerText As String = ""

            Dim pOnSigners As String = ""
            Dim pOnSignerNoVisa As String = ""
            Dim pOnSignerVisa As String = ""
            Dim pOnSignerOKTB As String = ""

            Dim pOffSigners As String = ""

            Dim pOther As String = ""


            For i As Integer = 0 To dgvOSMPax.Rows.Count - 1
                Dim pId = dgvOSMPax.Rows(i).Cells(0).Value
                Dim pPax As osmPax.Pax = mOSMPax(pId)
                Select Case dgvOSMPax.Rows(i).Cells("JoinerLeaver").Value
                    Case "ONSIGNER"
                        pOnSigners &= "<pre>" & pPax.Text & "</pre><br><br>"
                    Case "OFFSIGNER"
                        pOffSigners &= "<pre>" & pPax.Text & "</pre><br><br>"
                    Case Else
                        pOther &= "<pre>" & pPax.Text & "</pre><br><br>"
                End Select
                'pJoinerText = dgvOSMPax.Rows(i).Cells("JoinerLeaver").Value

                Select Case dgvOSMPax.Rows(i).Cells("VisaType").Value
                    Case "OKTB"
                        pOnSignerOKTB &= dgvOSMPax.Rows(i).Cells("Lastname").Value & "/" & dgvOSMPax.Rows(i).Cells("Firstname").Value & "<br>"
                    Case "NO VISA"
                        pOnSignerNoVisa &= dgvOSMPax.Rows(i).Cells("Lastname").Value & "/" & dgvOSMPax.Rows(i).Cells("Firstname").Value & "<br>"
                    Case "VISA"
                        pOnSignerVisa &= dgvOSMPax.Rows(i).Cells("Lastname").Value & "/" & dgvOSMPax.Rows(i).Cells("Firstname").Value & "<br>"
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
                'pVisaType.Value = 0
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
            OSMWebCreate()
            cmdOSMCopyDocument.Enabled = True

        Catch ex As Exception
            MessageBox.Show("cmdOSMPrepareDoc_Click()" & vbCrLf & ex.Message)
        End Try

    End Sub

   
    Private Sub cmdOSMVesselsEdit_Click(sender As Object, e As EventArgs) Handles cmdOSMVesselsEdit.Click

        Try
            Dim pFrm As New frmOSMVessels

            pFrm.ShowDialog(Me)
            OSMRefreshVessels(lstOSMVessels, chkOSMVesselInUse.Checked)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub cmdOSMAgentEdit_Click(sender As Object, e As EventArgs) Handles cmdOSMAgentEdit.Click
        Try
            Dim pFrm As New frmOSMAgents

            If pFrm.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                OSMRefreshVessels(lstOSMVessels, chkOSMVesselInUse.Checked)
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
                        If i <> e.RowIndex AndAlso dgvOSMPax.Rows(i).Cells("JoinerLeaver").Value = "ONSIGNER" AndAlso dgvOSMPax.Rows(i).Cells("VisaType").Value Is Nothing Then
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
            If tabPNR.SelectedIndex = 2 Then
                OSMRefreshVessels(lstOSMVessels, chkOSMVesselInUse.Checked)
                cmdOSMCopyDocument.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub optItnFormatDefault_CheckedChanged(sender As Object, e As EventArgs) Handles optItnFormatDefault.CheckedChanged, optItnFormatPlain.CheckedChanged, optItnFormatSeaChefs.CheckedChanged, chkItnSeaChefsWithCode.CheckedChanged

        Try
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub optItnFormatMSReport_CheckedChanged(sender As Object, e As EventArgs) Handles optItnFormatMSReport.CheckedChanged
        Try
            If cmdItnRefresh.Enabled Then
                ReadPNRandCreateItn(True)
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
                OSMRefreshVessels(lstOSMVessels, chkOSMVesselInUse.Checked)
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
            pdteDate = APISDateFromIATA(Row.Cells("Birthdate").Value)
            If pdteDate > Date.MinValue Then
                pflgBirthDateOK = True
            Else
                pflgBirthDateOK = False
            End If
        Else
            pflgBirthDateOK = True
        End If

        If Not Date.TryParse(Row.Cells("ExpiryDate").Value, pdteDate) Then
            pdteDate = APISDateFromIATA(Row.Cells("ExpiryDate").Value)
        End If
        If pdteDate > Now Then
            mflgExpiryDateOK = True
        Else
            mflgExpiryDateOK = False
        End If

        pflgGenderFound = False
        For i As Integer = 0 To mstrGenderIndicator.GetUpperBound(0)
            If Row.Cells("Gender").Value = mstrGenderIndicator(i) Then
                pflgGenderFound = True
                Exit For
            End If
        Next

        mflgAPISUpdate = mobjReadPNR.SegmentsExist And pflgBirthDateOK And pflgGenderFound And pflgPassportNumberOK

        If Not mflgAPISUpdate Then
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
        End If
        If mobjReadPNR.SegmentsExist Then
            lblSSRDocs.Text = "SSR DOCS"
            lblSSRDocs.BackColor = Color.Yellow
        Else
            lblSSRDocs.Text = "SSR DOCS cannot be updated - No segments in PNR"
            lblSSRDocs.BackColor = Color.Red
        End If
        Row.ErrorText = pstrErrorText

        SetEnabled()

    End Function

    Private Sub dgvApis_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvApis.CellValueChanged

        APISValidateDataRow(dgvApis.Rows(e.RowIndex))

    End Sub

    Private Sub dgvApis_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvApis.RowValidating

        APISValidateDataRow(dgvApis.Rows(e.RowIndex))

    End Sub

    Public Sub APISDisplayPax(ByRef dgvApis As Windows.Forms.DataGridView, ByVal mobjPNR As s1aPNR.PNR)

        Dim pobjPax As s1aPNR.NameElement
        Dim pobjPaxApis As New PaxApisDB.Collection
        Dim pobjPaxItem As PaxApisDB.Item

        dgvApis.Rows.Clear()

        For Each pobjPax In mobjPNR.NameElements
            pobjPaxItem = pobjPaxApis.Read(pobjPax.LastName, APISModifyFirstName(pobjPax.Initial))
            Dim dgvRow As New DataGridViewRow With {
                .DefaultCellStyle = dgvApis.RowsDefaultCellStyle
            }
            If pobjPaxApis.Count = 0 Then
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(0).Value = pobjPax.ElementNo
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(1).Value = pobjPax.LastName
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(2).Value = APISModifyFirstName(pobjPax.Initial)
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(3).Value = "" ' Issuing Country
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(4).Value = "" ' Passport Number
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(5).Value = "" ' Nationality
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(6).Value = 0 ' Birth date
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(7).Value = "M" ' Gender
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(8).Value = 0 ' Expiry Date
                'dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                'dgvRow.Cells(9).Value = "" ' QR Frequent flyer
            Else
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(0).Value = pobjPax.ElementNo
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(1).Value = pobjPax.LastName
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(2).Value = APISModifyFirstName(pobjPax.Initial)

                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(3).Value = pobjPaxItem.IssuingCountry ' Issuing Country
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(4).Value = pobjPaxItem.PassportNumber ' Passport Number
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(5).Value = pobjPaxItem.Nationality ' Nationality
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(6).Value = APISDateToIATA(pobjPaxItem.BirthDate) ' Birth date
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(7).Value = pobjPaxItem.Gender ' Gender
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                If pobjPaxItem.ExpiryDate > Date.MinValue Then
                    dgvRow.Cells(8).Value = APISDateToIATA(pobjPaxItem.ExpiryDate) ' Expiry Date
                End If

                'dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                'dgvRow.Cells(9).Value = pobjPaxItem.QRFreqFlyer  ' QR Frequent flyer
            End If

            dgvApis.Rows.Add(dgvRow)
            APISValidateDataRow(dgvApis.Rows(dgvApis.RowCount - 1))
        Next

    End Sub
    Public Sub APISDisplayPax(ByRef dgvApis As DataGridView, ByRef mobjPNR As s1aPNR.PNR, cmdAPISUpdate As Button)

        Dim pobjPax As s1aPNR.NameElement
        Dim pobjPaxApis As New PaxApisDB.Collection
        Dim pobjPaxItem As PaxApisDB.Item

        cmdAPISUpdate.Enabled = False

        APISPrepareGrid()
        dgvApis.Rows.Clear()
        cmdAPISUpdate.Enabled = False

        For Each pobjPax In mobjPNR.NameElements
            pobjPaxItem = pobjPaxApis.Read(pobjPax.LastName, APISModifyFirstName(pobjPax.Initial))
            Dim dgvRow As New DataGridViewRow With {
                .DefaultCellStyle = dgvApis.RowsDefaultCellStyle
            }
            If pobjPaxApis.Count = 0 Then
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(0).Value = pobjPax.ElementNo
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(1).Value = pobjPax.LastName
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(2).Value = APISModifyFirstName(pobjPax.Initial)
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(3).Value = "" ' Issuing Country
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(4).Value = "" ' Passport Number
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(5).Value = "" ' Nationality
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(6).Value = 0 ' Birth date
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(7).Value = "M" ' Gender
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(8).Value = 0 ' Expiry Date
                'dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                'dgvRow.Cells(9).Value = "" ' QR Frequent flyer
            Else
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(0).Value = pobjPax.ElementNo
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(1).Value = pobjPax.LastName
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(2).Value = APISModifyFirstName(pobjPax.Initial)

                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(3).Value = pobjPaxItem.IssuingCountry ' Issuing Country
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(4).Value = pobjPaxItem.PassportNumber ' Passport Number
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(5).Value = pobjPaxItem.Nationality ' Nationality
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(6).Value = APISDateToIATA(pobjPaxItem.BirthDate) ' Birth date
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                dgvRow.Cells(7).Value = pobjPaxItem.Gender ' Gender
                dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                If pobjPaxItem.ExpiryDate > Date.MinValue Then
                    dgvRow.Cells(8).Value = APISDateToIATA(pobjPaxItem.ExpiryDate) ' Expiry Date
                End If

                'dgvRow.Cells.Add(New DataGridViewTextBoxCell)
                'dgvRow.Cells(9).Value = pobjPaxItem.QRFreqFlyer  ' QR Frequent flyer
            End If

            dgvApis.Rows.Add(dgvRow)
            cmdAPISUpdate.Enabled = APISValidateDataRow(dgvApis.Rows(dgvApis.RowCount - 1))
        Next

    End Sub

    Private Function APISModifyFirstName(ByVal FirstName As String) As String

        Dim pintFindPos As Integer

        FirstName = Trim(FirstName)

        For i As Short = 0 To mstrSalutations.GetUpperBound(0)
            pintFindPos = FirstName.IndexOf(mstrSalutations(i))
            If pintFindPos > 0 And pintFindPos = FirstName.Length - mstrSalutations(i).Length Then
                FirstName = FirstName.Substring(0, pintFindPos).Trim
            End If
        Next

        Return FirstName

    End Function

    Private Sub APISPrepareGrid()

        dgvApis.Columns.Clear()
        dgvApis.Columns.Add("Id", "Id")
        dgvApis.Columns.Add("Surname", "Surname")
        dgvApis.Columns.Add("FirstName", "First Name")
        dgvApis.Columns.Add("IssuingCountry", "Issuing Country")
        dgvApis.Columns.Add("Passportnumber", "Passport number")
        dgvApis.Columns.Add("Nationality", "Nationality")
        dgvApis.Columns.Add("BirthDate", "Birth Date")
        dgvApis.Columns.Add("Gender", "Gender")
        dgvApis.Columns.Add("ExpiryDate", "Expiry Date")

    End Sub

End Class