Option Strict Off
Option Explicit On
Imports s1aPNR

Public Class GDS1GReadPNR
    ' Implements IGDSReadPNR

    Public ReadOnly Property AirSegments As Object ' Implements IGDSReadPNR.AirSegments
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property CreationDate As Date ' Implements IGDSReadPNR.CreationDate
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property DepartureDate As Date ' Implements IGDSReadPNR.DepartureDate
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property ExistingElements As GDSExisting.Collection ' Implements IGDSReadPNR.ExistingElements
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property GroupName As String ' Implements IGDSReadPNR.GroupName
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property GroupNamesCount As Integer ' Implements IGDSReadPNR.GroupNamesCount
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property HasQRSegment As Boolean ' Implements IGDSReadPNR.HasQRSegment
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property IsGroup As Boolean ' Implements IGDSReadPNR.IsGroup
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property Itinerary As String ' Implements IGDSReadPNR.Itinerary
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property NewElements As GDSNew.Collection ' Implements IGDSReadPNR.NewElements
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property NewPNR As Boolean ' Implements IGDSReadPNR.NewPNR
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property NumberOfPax As Integer ' Implements IGDSReadPNR.NumberOfPax
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property OfficeOfResponsibility As String ' Implements IGDSReadPNR.OfficeOfResponsibility
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property PaxName As String ' Implements IGDSReadPNR.PaxName
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property PNR As PNR ' Implements IGDSReadPNR.PNR
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property PnrNumber As String ' Implements IGDSReadPNR.PnrNumber
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property SegmentsExist As Boolean ' Implements IGDSReadPNR.SegmentsExist
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property SSRDocsExists As Boolean ' Implements IGDSReadPNR.SSRDocsExists
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property Tickets As GDSTickets.Collection ' Implements IGDSReadPNR.Tickets
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Sub PrepareNewGDSElements() ' Implements IGDSReadPNR.PrepareNewGDSElements
        Throw New NotImplementedException()
    End Sub

    Public Sub SendNewGDSEntries(GDSEntry As String) ' Implements IGDSReadPNR.SendNewGDSEntries
        Throw New NotImplementedException()
    End Sub

    Public Sub SendNewGDSEntries(WritePNR As Boolean, WriteDocs As Boolean, mflgExpiryDateOK As Boolean, dgvApis As DataGridView, AirlineEntries As TextBox) ' Implements IGDSReadPNR.SendNewGDSEntries
        Throw New NotImplementedException()
    End Sub

    Public Function CopyGDSEntries(AirlineNotes As CheckedListBox, AirlinePoints As CheckedListBox) As String ' Implements IGDSReadPNR.CopyGDSEntries
        Throw New NotImplementedException()
    End Function

    Public Function Read() As String ' Implements IGDSReadPNR.Read

        Dim Session As New Travelport.TravelData.Factory.GalileoDesktopFactory("SPG720", "MYCONNECTION", False, True, "SMRT")

        Read = ""

        Try
            Session.SendTerminalCommand("QXI+I")
            mudtProps.RequestedPNR = PNR
            Read = RetrievePNR1G()


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
        Catch ex As Exception
            Throw New Exception("Read()" & vbCrLf & ex.Message)
        End Try

        Try


        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function
    Private Function RetrievePNR1G() As Boolean

        Dim Session As New Travelport.TravelData.Factory.GalileoDesktopFactory("SPG720", "MYCONNECTION", False, True, "SMRT")
        mobjTickets = New Ticket.TicketColl
        mstrVesselName = ""
        mstrBookedBy = ""
        mstrCC = ""
        mstrCLA = ""
        mstrCLN = ""

        With mudtProps

            If .RequestedPNR = "" Then
                Dim pErr As Integer = 1

                Do While pErr < 10
                    Try
                        mobjPNR1G = Session.RetrieveCurrentBookingFile
                        pErr = 99
                    Catch ex As Exception
                        System.Threading.Thread.Sleep(2000)
                        pErr += 1
                    End Try
                Loop
                If pErr < 99 Then
                    Throw New Exception("Galileo communication problem. Please try again or contact your system adminbistrator")
                End If
            Else
                mobjPNR1G = Session.RetrieveBookingFile(.RequestedPNR)
            End If
            If mobjPNR1G.IsEmpty Then
                Throw New Exception("No B.F. to display")
            End If
            .PNRCreationdate = Today

            .RequestedPNR = setRecordLocator1G()

            'GetTQT1G()
            'GetGroup1G()
            GetPax1G()
                GetSegs1G(ForReportOnly)
                'GetAutoTickets1G()
                GetOtherServiceElements1G()
                'GetSSRElements1G()
                GetRMElements1G()
            RetrievePNR1G = True
        End With

    End Function
End Class
