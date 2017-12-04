Public Class frmOptions

    Private mflgIsDirty As Boolean

    Private Sub frmOptions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        DisplayOptions()
        mflgIsDirty = False

    End Sub

    Private Sub DisplayOptions()

        With MySettings

            txtOfficePCCAmadeus.Enabled = .Administrator
            txtOfficeCityCode.Enabled = .Administrator
            txtCountryCode.Enabled = .Administrator
            txtOfficeName.Enabled = .Administrator
            txtCityName.Enabled = .Administrator
            txtOfficePhone.Enabled = .Administrator
            txtAOHPhone.Enabled = .Administrator


            txtAgentIDAmadeus.Enabled = .Administrator
            txtAgentQueueAmadeus.Enabled = True
            txtAgentOPQueueAmadeus.Enabled = True
            txtAgentName.Enabled = True
            txtAgentEmail.Enabled = True

            txtOfficePCCAmadeus.Text = .AmadeusPCC
            txtOfficeCityCode.Text = .OfficeCityCode
            txtCountryCode.Text = .CountryCode
            txtOfficeName.Text = .OfficeName
            txtCityName.Text = .CityName
            txtOfficePhone.Text = .Phone
            txtAOHPhone.Text = .AOHPhone


            txtAgentIDAmadeus.Text = .AmadeusUser
            txtAgentQueueAmadeus.Text = .AgentQueue
            txtAgentOPQueueAmadeus.Text = .AgentOPQueue
            txtAgentName.Text = .AgentName
            txtAgentEmail.Text = .AgentEmail

            lblDBConnectionFile.Text = DBConnectionsFile
            lblSQLServer.Text = ConnectionStringPNR
        End With

    End Sub

    'Private Sub SetIsDirty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAgentEmail.TextChanged, txtAgentIDAmadeus.TextChanged, txtAgentName.TextChanged, _
    '                                    txtAgentOPQueueAmadeus.TextChanged, txtAgentQueueAmadeus.TextChanged, _
    '                                    txtAOHPhone.TextChanged, txtCityName.TextChanged, txtCountryCode.TextChanged, _
    '                                    txtOfficeCityCode.TextChanged, txtOfficePCCAmadeus.TextChanged, _
    '                                    txtOfficePhone.TextChanged, txtOfficeName.TextChanged

    '    mflgIsDirty = True

    'End Sub


    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        If MySettings.IsDirty Then
            SaveSettings()
        End If
        Me.Close()

    End Sub

    Private Sub SaveSettings()

        MySettings.Save()
        mflgIsDirty = False

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click

        Dim mResult As DialogResult

        If mflgIsDirty Then
            mResult = MessageBox.Show("Do you want to save your changes?", "Options", MessageBoxButtons.YesNoCancel)
            If mResult <> Windows.Forms.DialogResult.Cancel Then
                If mResult = Windows.Forms.DialogResult.Yes Then
                    SaveSettings()
                End If
                Me.Close()
            End If
        Else
            Me.Close()
        End If

    End Sub

    Private Sub txtAgentQueueAmadeus_TextChanged(sender As Object, e As EventArgs) Handles txtAgentQueueAmadeus.TextChanged

        MySettings.AgentQueue = txtAgentQueueAmadeus.Text

    End Sub

    Private Sub txtAgentOPQueueAmadeus_TextChanged(sender As Object, e As EventArgs) Handles txtAgentOPQueueAmadeus.TextChanged

        MySettings.AgentOPQueue = txtAgentOPQueueAmadeus.Text

    End Sub

    Private Sub txtAgentName_TextChanged(sender As Object, e As EventArgs) Handles txtAgentName.TextChanged

        MySettings.AgentName = txtAgentName.Text

    End Sub

    Private Sub txtAgentEmail_TextChanged(sender As Object, e As EventArgs) Handles txtAgentEmail.TextChanged

        MySettings.AgentEmail = txtAgentEmail.Text

    End Sub
End Class