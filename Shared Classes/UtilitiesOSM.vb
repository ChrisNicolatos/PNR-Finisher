<CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1812:AvoidUninstantiatedInternalClasses")>
Friend Class UtilitiesOSM
    Public Shared Sub OSMRefreshVessels(ByRef lstListBox As ListBox, ByVal InUseOnly As Boolean)

        Dim pOSMVessels As New osmVessels.VesselCollection

        If Not MySettings Is Nothing Then
            pOSMVessels.Load(MySettings.OSMVesselGroup)
        Else
            pOSMVessels.Load()
        End If
        lstListBox.Items.Clear()

        For Each pVessels As osmVessels.VesselItem In pOSMVessels.Values
            If Not InUseOnly Or pVessels.InUse Then
                lstListBox.Items.Add(pVessels)
            End If
        Next

    End Sub
    Public Shared Sub OSMRefreshVessels(ByRef lstListBox As ListBox)

        Dim pOSMVessels As New osmVessels.VesselCollection

        pOSMVessels.Load()
        lstListBox.Items.Clear()

        For Each pVessels As osmVessels.VesselItem In pOSMVessels.Values
            lstListBox.Items.Add(pVessels)
        Next

    End Sub
    Public Shared Sub OSMRefreshVesselGroup(ByRef cmbComboBox As ComboBox)

        Dim pOSMVesselGroup As New osmVessels.VesselGroupCollection

        pOSMVesselGroup.Load()
        cmbComboBox.Items.Clear()

        For Each PVesselGroup As osmVessels.VesselGroupItem In pOSMVesselGroup.Values
            cmbComboBox.Items.Add(PVesselGroup)
            If Not MySettings Is Nothing Then
                If PVesselGroup.Id = MySettings.OSMVesselGroup Then
                    cmbComboBox.SelectedItem = PVesselGroup
                End If
            End If
        Next

    End Sub
    Public Shared Sub OSMDisplayVesselGroups(ByRef lstListBox As CheckedListBox, ByVal pobjGroups As osmVessels.Vessel_VesselGroupCollection)

        lstListBox.Items.Clear()
        For Each Vessel_VesselGroup As osmVessels.Vessel_VesselGroupItem In pobjGroups.Values
            lstListBox.Items.Add(Vessel_VesselGroup, Vessel_VesselGroup.Exists)
        Next

    End Sub
    Public Shared Sub OSMDisplayEmails(ByVal VesselList As ListBox, ByRef lstToEmail As ListBox, ByRef lstCCEmail As ListBox, ByRef lstAgentsEmail As ListBox)

        Dim pobjEmails As New osmVessels.EmailCollection

        lstToEmail.Items.Clear()
        lstCCEmail.Items.Clear()
        lstAgentsEmail.Items.Clear()
        lstAgentsEmail.Items.Add("")

        For Each SelectedVessel As osmVessels.VesselItem In VesselList.SelectedItems

            pobjEmails.Load(SelectedVessel.Id)

            For Each pEmail As osmVessels.emailItem In pobjEmails.Values
                With pEmail
                    If .EmailType = "TO" Then
                        Dim pFound As Boolean = False
                        For Each pItem As osmVessels.emailItem In lstToEmail.Items
                            If pEmail.ToString = pItem.ToString Then
                                pFound = True
                                Exit For
                            End If
                        Next
                        If Not pFound Then
                            lstToEmail.Items.Add(pEmail)
                        End If
                    ElseIf .EmailType = "CC" Then
                        Dim pFound As Boolean = False
                        For Each pItem As osmVessels.emailItem In lstCCEmail.Items
                            If pEmail.ToString = pItem.ToString Then
                                pFound = True
                                Exit For
                            End If
                        Next
                        If Not pFound Then
                            lstCCEmail.Items.Add(pEmail)
                        End If

                    ElseIf .EmailType = "AGENT" Then
                        lstAgentsEmail.Items.Add(pEmail)
                    End If
                End With
            Next
        Next

    End Sub
    Public Shared Sub OSMDisplayEmails(ByVal VesselId As Integer, ByRef lstToEmail As ListBox, ByRef lstCCEmail As ListBox)

        Dim pobjEmails As New osmVessels.EmailCollection

        pobjEmails.Load(VesselId)

        lstToEmail.Items.Clear()
        lstCCEmail.Items.Clear()

        For Each pEmail As osmVessels.emailItem In pobjEmails.Values
            With pEmail
                If .EmailType = "TO" Then
                    lstToEmail.Items.Add(pEmail)
                ElseIf .EmailType = "CC" Then
                    lstCCEmail.Items.Add(pEmail)
                End If
            End With
        Next
    End Sub
    Public Shared Sub OSMDisplayEmails(ByRef lstAgents As ListBox)

        Dim pobjEmails As New osmVessels.EmailCollection

        pobjEmails.Load()

        lstAgents.Items.Clear()

        For Each pEmail As osmVessels.emailItem In pobjEmails.Values
            With pEmail
                If .EmailType = "AGENT" Then
                    lstAgents.Items.Add(pEmail)
                End If
            End With
        Next
    End Sub

    Public Shared Sub ListBox_DrawItem(sender As Object, e As DrawItemEventArgs)

        If e.Index >= 0 And e.Index < sender.Items.Count Then
            Dim stringToDraw As String = sender.Items(e.Index).ToString
            Dim VesselToDraw As osmVessels.VesselItem = sender.Items(e.Index)
            Dim C As Color
            If Not VesselToDraw.InUse Then
                C = Color.Red
            Else
                C = sender.ForeColor
            End If

            e.DrawBackground()
            e.DrawFocusRectangle()
            e.Graphics.DrawString(stringToDraw, e.Font, New SolidBrush(C), e.Bounds)
        End If

    End Sub
End Class
