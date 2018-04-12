Friend Class frmOSMLoG
    Private mflgLoading As Boolean
    Private mobjAgent As osmVessels.emailItem
    Private mobjPNR As GDSReadPNR

    Friend Sub New(ByRef pPNR As GDSReadPNR)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        mobjPNR = pPNR

    End Sub
    Public ReadOnly Property PortAgent As osmVessels.emailItem
        Get
            PortAgent = mobjAgent
        End Get
    End Property
    Public ReadOnly Property NoPortAgent As Boolean
        Get
            NoPortAgent = chkNoPortAgent.Checked
        End Get
    End Property
    Public ReadOnly Property SignedBy As String
        Get
            SignedBy = txtSignedBy.Text
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
            ShowPNRDetails()
            EnableSelection()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.Close()
        Finally
            mflgLoading = False
        End Try

    End Sub
    Private Sub ShowPNRDetails()

        With mobjPNR
            If .Passengers.Count > 1 Then
                lblPax.Text = .Passengers.Count & " Passengers" & vbCrLf
            Else
                lblPax.Text = .Passengers.Count & " Passenger" & vbCrLf
            End If
            For Each pPax As GDSPax.GDSPaxItem In .Passengers.Values
                lblPax.Text &= pPax.ElementNo & ". " & pPax.PaxName & vbCrLf
            Next

            lblSegs.Text = ""
            For Each pSeg As GDSSeg.GDSSegItem In .Segments.Values
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

        Dim pAgents As New osmVessels.EmailCollection

        pAgents.Load()

        lstPortAgent.Items.Clear()
        For Each pAgent As osmVessels.emailItem In pAgents.Values
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

        cmdCreatePDF.Enabled = ((optPerPax.Checked Or optPerPNR.Checked) And (optOnSigners.Checked Or optOffSigners.Checked) And txtFileDestination.Text <> "" And (Not mobjAgent Is Nothing Or chkNoPortAgent.Checked))

    End Sub

    Private Sub Option_CheckedChanged(sender As Object, e As EventArgs) Handles optPerPax.CheckedChanged, optPerPNR.CheckedChanged, optOnSigners.CheckedChanged, optOffSigners.CheckedChanged, txtFileDestination.TextChanged, chkNoPortAgent.CheckedChanged

        If Not mflgLoading Then
            MySettings.OSMLoGPerPax = optPerPax.Checked
            MySettings.OSMLoGOnSigner = optOnSigners.Checked
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
End Class