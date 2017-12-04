Imports System.IO
Imports System.Text

Module modPNR
    '
    ' Prepares the SQL connection string for the Travel Force Cosmos database
    '
    ' The SQL connection string for the Travel Force Cosmos database
    ' Returns the connection string to be used for SQLConnection
    ' The options for the connection string are bound to the application and cannot be modified by the user
    ' 

    Private Const MONTH_NAMES As String = "JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC"
    Private Const mstrDBConnectionsFileGT As String = "\\192.168.102.223\Common\Click-Once Applications\PNR Finisher ATH\Config\PNRFinisher.txt"
    Private Const mstrDBConnectionsFile As String = "\\ath2-svrdc1\PNR Finisher ATH Config\PNRFinisher.txt"
    Private mstrDBConnectionFileActual As String

    Private mMySettings As Config
    Private mAccDataSource As String = ""
    Private mAccDataCatalog As String = ""
    Private mAccUserName As String = ""
    Private mAccPassword As String = ""

    Private mPnrDataSource As String = ""
    Private mPnrDataCatalog As String = ""
    Private mPnrUserName As String = ""
    Private mPnrPassword As String = ""

    Public Sub InitSettings(ByVal mAmadeusUser As AmadeusUser)
        mMySettings = New Config(mAmadeusUser)
    End Sub
    Public ReadOnly Property MySettings As Config
        Get
            MySettings = mMySettings
        End Get

    End Property
    Public ReadOnly Property DBConnectionsFile As String
        Get
            DBConnectionsFile = mstrDBConnectionFileActual
        End Get
    End Property
    Private Sub ReadDBConnections()

        Dim pFileExists As Boolean = False

        If File.Exists(mstrDBConnectionsFile) Then
            mstrDBConnectionFileActual = mstrDBConnectionsFile
            pFileExists = True
        ElseIf File.Exists(mstrDBConnectionsFileGT) Then
            mstrDBConnectionFileActual = mstrDBConnectionsFileGT
            pFileExists = True
        End If

        If pFileExists Then
            Dim GDSData As StreamReader = File.OpenText(mstrDBConnectionFileActual)
            Dim xLine() As String = Split(GDSData.ReadToEnd, vbCrLf)
            GDSData.Close()

            If IsArray(xLine) Then
                For i As Integer = xLine.GetLowerBound(0) To xLine.GetUpperBound(0)
                    Dim pValues() As String = xLine(i).Split("=")
                    Select Case pValues(0).Trim.ToUpper
                        Case "DATASOURCEACC"
                            mAccDataSource = pValues(1).Trim
                        Case "DATACATALOGACC"
                            mAccDataCatalog = pValues(1).Trim
                        Case "DATAUSERNAMEACC"
                            mAccUserName = pValues(1).Trim
                        Case "DATAUSERPASSWORDACC"
                            mAccPassword = pValues(1).Trim

                        Case "DATASOURCEPNR"
                            mPnrDataSource = pValues(1).Trim
                        Case "DATACATALOGPNR"
                            mPnrDataCatalog = pValues(1).Trim
                        Case "DATAUSERNAMEPNR"
                            mPnrUserName = pValues(1).Trim
                        Case "DATAUSERPASSWORDPNR"
                            mPnrPassword = pValues(1).Trim
                    End Select
                Next
            Else
                Throw New Exception("Settings File Error" & vbCrLf & mstrDBConnectionFileActual)
            End If
        Else
            Throw New Exception("DB Connection file does not exist. Please contact you system administrator" & vbCrLf & mstrDBConnectionFileActual)
        End If

    End Sub

    Public ReadOnly Property ConnectionStringACC() As String
        Get
            If mAccDataSource = "" Then
                ReadDBConnections()
            End If
            ConnectionStringACC = "Data Source=" & mAccDataSource & _
                                  ";Initial Catalog=" & mAccDataCatalog & _
                                  ";User ID=" & mAccUserName & _
                                  ";Password=" & mAccPassword
        End Get
    End Property

    '
    ' Prepares the SQL connection string for the Amadeus Reports database
    '
    ' The SQL connection string for the Amadeus Reports database
    ' Returns the connection string to be used for SQLConnection
    ' The options for the connection string are bound to the application and cannot be modified by the user. The Amadeus Reports database contains tables that are not part of the Travel Force Cosmos database
    Public ReadOnly Property ConnectionStringPNR() As String
        Get
            If mPnrDataSource = "" Then
                ReadDBConnections()
            End If
            ConnectionStringPNR = "Data Source=" & mPnrDataSource & _
                                  ";Initial Catalog=" & mPnrDataCatalog & _
                                  ";User ID=" & mPnrUserName & _
                                  ";Password=" & mPnrPassword
        End Get
    End Property

    Public Function myCurr(ByVal StringToParse As String) As Decimal

        Dim i As Integer
        Dim pintPoint As Short
        Dim pintComma As Short

        Do While Not IsNumeric(Right(StringToParse, 1)) And Len(StringToParse) > 0
            StringToParse = Left(StringToParse, Len(StringToParse) - 1)
        Loop
        StringToParse = Trim(StringToParse)
        pintPoint = InStr(StringToParse, My.Application.Culture.NumberFormat.CurrencyGroupSeparator)
        pintComma = InStr(StringToParse, My.Application.Culture.NumberFormat.CurrencyDecimalSeparator)
        If pintPoint > pintComma Then
            If Len(StringToParse) > 2 Then
                If Mid(StringToParse, Len(StringToParse) - 2, 1) = My.Application.Culture.NumberFormat.CurrencyGroupSeparator Then
                    Mid(StringToParse, Len(StringToParse) - 2, 1) = My.Application.Culture.NumberFormat.CurrencyDecimalSeparator
                End If
            End If
        End If

        If IsDBNull(StringToParse) Then
            StringToParse = ""
        End If
        If IsNumeric(StringToParse) Then
            myCurr = CDec(StringToParse)
        Else
            myCurr = 0
            For i = 1 To Len(StringToParse)
                If IsNumeric(Mid(StringToParse, 1, i)) Then
                    myCurr = CDec(Mid(StringToParse, 1, i))
                Else
                    Exit For
                End If
            Next i
        End If

    End Function

    Public Sub OSMRefreshVessels(ByRef lstListBox As ListBox, ByVal InUseOnly As Boolean)

        Dim pOSMVessels As New osmVessels.VesselCollection

        pOSMVessels.Load()
        lstListBox.Items.Clear()

        For Each pVessels As osmVessels.VesselItem In pOSMVessels.Values
            If Not InUseOnly Or pVessels.InUse Then
                lstListBox.Items.Add(pVessels)
            End If
        Next

    End Sub

    <CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")>
    Public Sub OSMDisplayEmails(ByVal VesselId As Integer, ByRef lstToEmail As ListBox, ByRef lstCCEmail As ListBox, ByRef lstAgentsEmail As ListBox)

        Dim pobjEmails As New osmVessels.emailCollection

        pobjEmails.Load(VesselId)

        lstToEmail.Items.Clear()
        lstCCEmail.Items.Clear()
        lstAgentsEmail.Items.Clear()
        lstAgentsEmail.Items.Add("")

        For Each pEmail As osmVessels.emailItem In pobjEmails.Values
            With pEmail
                If .EmailType = "TO" Then
                    lstToEmail.Items.Add(pEmail)
                ElseIf .EmailType = "CC" Then
                    lstCCEmail.Items.Add(pEmail)
                ElseIf .EmailType = "AGENT" Then
                    lstAgentsEmail.Items.Add(pEmail)
                End If
            End With
        Next

    End Sub
    Public Sub OSMDisplayEmails(ByVal VesselList As ListBox, ByRef lstToEmail As ListBox, ByRef lstCCEmail As ListBox, ByRef lstAgentsEmail As ListBox)

        Dim pobjEmails As New osmVessels.emailCollection

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


    Public Sub OSMDisplayEmails(ByVal VesselId As Integer, ByRef lstToEmail As ListBox, ByRef lstCCEmail As ListBox)

        Dim pobjEmails As New osmVessels.emailCollection

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
    Public Sub OSMDisplayEmails(ByRef lstAgents As ListBox)

        Dim pobjEmails As New osmVessels.emailCollection

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

    Public Sub ListBox_DrawItem(sender As Object, e As DrawItemEventArgs)

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

    End Sub

    Public Function APISDateFromIATA(ByVal InDate As String) As Date

        Dim pintDay As Integer
        Dim pintMonth As Integer
        Dim pintYear As Integer

        Try
            If Not Date.TryParse(InDate, APISDateFromIATA) Then
                APISDateFromIATA = Date.MinValue
                pintDay = InDate.Substring(0, 2)
                pintMonth = (MONTH_NAMES.IndexOf(InDate.Substring(3, 3)) + 2) / 3
                pintYear = InDate.Substring(5)

                If pintMonth >= 1 Then
                    APISDateFromIATA = DateSerial(pintYear, pintMonth, pintDay)
                End If
            End If
        Catch ex As Exception

        End Try

    End Function

    Public Function APISDateToIATA(ByVal InDate As Date) As String

        APISDateToIATA = Format(InDate.Day, "00") & MONTH_NAMES.Substring(InDate.Month * 3 - 3, 3) & Format(InDate.Year Mod 100, "00")

    End Function
End Module
