Friend Class frmOSMLoG
    Private mflgLoading As Boolean
    Private mobjAddresses As New OSMAddressCollection
    Private mobjAddressSelected As OSMAddressItem
    Private mOSMAgents As New OSMEmailCollection
    Private mobjAgent As OSMEmailItem
    Private mobjPNR As GDSReadPNR

    Friend Sub New(ByRef pPNR As GDSReadPNR)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        mobjPNR = pPNR

    End Sub
    Public ReadOnly Property PortAgent As OSMEmailItem
        Get
            Return mobjAgent
        End Get
    End Property
    Public ReadOnly Property NoPortAgent As Boolean
        Get
            Return chkNoPortAgent.Checked
        End Get
    End Property
    Public ReadOnly Property AddressItem As OSMAddressItem
        Get
            Return mobjAddressSelected
        End Get
    End Property
    Private Sub frmOSMLoG_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            mflgLoading = True
            If MySettings.OSMLoGPerPax Then
                optPerPax.Checked = True
            Else
                optPerPNR.Checked = True
            End If
            If MySettings.OSMLoGOnSigner Then
                optOnSigners.Checked = True
            Else
                optOffSigners.Checked = True
            End If
            If System.IO.Directory.Exists(MySettings.OSMLoGPath) Then
                txtFileDestination.Text = MySettings.OSMLoGPath
            Else
                txtFileDestination.Text = ""
            End If
            LoadAgents()
            LoadAddresses()
            ShowPNRDetails()
            EnableSelection()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.Close()
        Finally
            mflgLoading = False
        End Try

    End Sub
    Private Sub LoadAddresses()
        lstOfficeAddress.Items.Clear()
        mobjAddresses.Load
        For Each pItem As OSMAddressItem In mobjAddresses.Values
            lstOfficeAddress.Items.Add(pItem)
        Next
    End Sub
    Private Sub ShowPNRDetails()

        With mobjPNR
            lblPNR.Text = "PNR: " & .RequestedPNR & vbCrLf & "Client Code: " & .ClientCode & vbCrLf & .ClientName & vbCrLf & .VesselName
            If .Passengers.Count > 1 Then
                lblPax.Text = .Passengers.Count & " Passengers" & vbCrLf
            Else
                lblPax.Text = .Passengers.Count & " Passenger" & vbCrLf
            End If
            For Each pPax As GDSPaxItem In .Passengers.Values
                lblPax.Text &= pPax.ElementNo & ". " & pPax.PaxName & vbCrLf
            Next

            lblSegs.Text = ""
            For Each pSeg As GDSSegItem In .Segments.Values
                With pSeg
                    lblSegs.Text &= .Airline & " " & .FlightNo.PadLeft(5) & " " & .DepartureDateIATA.PadLeft(6) & " " & .BoardPoint & " " & .OffPoint & " " & Format(.DepartTime, "HHmm") & vbCrLf
                End With
            Next
            If .BookedBy.IndexOf("-") > 0 Then
                txtSignedBy.Text = .BookedBy.Substring(0, .BookedBy.IndexOf("-"))
            Else
                txtSignedBy.Text = .BookedBy
            End If

        End With
    End Sub
    Private Sub LoadAgents()

        mOSMAgents.Load()
        lstPortAgent.Items.Clear()
        For Each pAgent As OSMEmailItem In mOSMAgents.Values
            lstPortAgent.Items.Add(pAgent)
        Next

    End Sub
    Private Sub lstPortAgent_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstPortAgent.SelectedIndexChanged

        If Not lstPortAgent.SelectedItem Is Nothing Then
            mobjAgent = lstPortAgent.SelectedItem
        End If
        EnableSelection()
    End Sub
    Private Sub EnableSelection()

        cmdCreatePDF.Enabled = ((optPerPax.Checked Or optPerPNR.Checked) And (optOnSigners.Checked Or optOffSigners.Checked Or optOnSignersBrazil.Checked) And txtFileDestination.Text <> "" And (Not mobjAgent Is Nothing Or chkNoPortAgent.Checked))
        If txtSignedBy.Text = "" Then
            cmdCreatePDF.Enabled = False
            txtSignedBy.BackColor = Color.Red
        Else
            txtSignedBy.BackColor = Color.FromKnownColor(KnownColor.Window)
        End If

    End Sub

    Private Sub Option_CheckedChanged(sender As Object, e As EventArgs) Handles optPerPax.CheckedChanged, optPerPNR.CheckedChanged, optOnSigners.CheckedChanged, optOffSigners.CheckedChanged, optOnSignersBrazil.CheckedChanged, txtFileDestination.TextChanged, chkNoPortAgent.CheckedChanged

        If Not mflgLoading Then
            MySettings.OSMLoGPerPax = optPerPax.Checked
            MySettings.OSMLoGOnSigner = optOnSigners.Checked Or optOnSignersBrazil.Checked
            If optOnSignersBrazil.Checked Then
                MySettings.OSMLoGLanguage = EnumLoGLanguage.Brazil
            Else
                MySettings.OSMLoGLanguage = EnumLoGLanguage.English
            End If
            MySettings.OSMLoGPath = txtFileDestination.Text
            MySettings.Save()
            EnableSelection()
        End If

    End Sub

    Private Sub cmdFileDestination_Click(sender As Object, e As EventArgs) Handles cmdFileDestination.Click
        Try
            fileBrowser.SelectedPath = MySettings.OSMLoGPath
            If fileBrowser.ShowDialog(Me) = DialogResult.OK Then
                txtFileDestination.Text = fileBrowser.SelectedPath
            End If
            EnableSelection()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub cmdCreatePDF_Click(sender As Object, e As EventArgs) Handles cmdCreatePDF.Click
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub
    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
    Private Sub txtOSMAgentsFilter_TextChanged(sender As Object, e As EventArgs) Handles txtOSMAgentsFilter.TextChanged
        Try
            lstPortAgent.Items.Clear()
            If txtOSMAgentsFilter.Text.Trim = "" Then
                For Each pAgent As OSMEmailItem In mOSMAgents.Values
                    lstPortAgent.Items.Add(pAgent)
                Next
            Else
                Dim pFilter() As String = txtOSMAgentsFilter.Text.ToUpper.Trim.Split({"|"}, StringSplitOptions.RemoveEmptyEntries)

                For Each pAgent As OSMEmailItem In mOSMAgents.Values
                    For i As Integer = 0 To pFilter.GetUpperBound(0)
                        If pAgent.ToString.ToUpper.IndexOf(pFilter(i).Trim) >= 0 Then
                            lstPortAgent.Items.Add(pAgent)
                            Exit For
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub txtSignedBy_TextChanged(sender As Object, e As EventArgs) Handles txtSignedBy.TextChanged
        If Not mobjAddressSelected Is Nothing Then
            mobjAddressSelected.SignedByName = txtSignedBy.Text
        End If
        EnableSelection()
    End Sub

    Private Sub lstOfficeAddress_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstOfficeAddress.SelectedIndexChanged
        If Not lstOfficeAddress Is Nothing Then
            mobjAddressSelected = lstOfficeAddress.SelectedItem
            DisplayAddress()
        End If
        EnableSelection()
    End Sub
    Private Sub DisplayAddress()
        With mobjAddressSelected
            If .SignedByName = "" Then
                If mobjPNR.BookedBy.IndexOf("-") > 0 Then
                    txtSignedBy.Text = mobjPNR.BookedBy.Substring(0, mobjPNR.BookedBy.IndexOf("-"))
                Else
                    txtSignedBy.Text = mobjPNR.BookedBy
                End If
            Else
                txtSignedBy.Text = .SignedByName
            End If
            txtTitle.Text = .Title
            txtCompanyName.Text = .CompanyName
            txtAddress.Text = .Address
            txtPCArea.Text = .PCArea
            txtCountry.Text = .Country
            txtTelephone.Text = .Telephone
            If .LogoImage_fk = 0 Then
                picLogo.Visible = False
            Else
                picLogo.Image = .LogoImage
                picLogo.Visible = True
            End If
            If .SignatureImage_fk = 0 Then
                picSignature.Visible = False
            Else
                picSignature.Image = .SignatureImage
                picSignature.Visible = True
            End If
        End With
    End Sub
End Class