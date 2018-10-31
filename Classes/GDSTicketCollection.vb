Public Class GDSTicketCollection
    Inherits Collections.Generic.Dictionary(Of String, GDSTicketItem)

    Private mintCount As Integer

    Private mobjTickets() As GDSTicketItem
    Private mobjPNR As s1aPNR.PNR
    Public Sub New()
        MyBase.Clear()
    End Sub
    Public Sub New(ByVal pPnr As s1aPNR.PNR)
        mobjPNR = pPnr
        ReadTickets()
    End Sub
    Public Sub addTicket(ByVal pGDSLine As String, ByVal pTicketType As Integer, ByVal pTicketNumber As Decimal, ByVal pTicketCount As Integer, ByVal IssuingAirline As String, ByVal AirlineCode As String, ByVal eTicket As Boolean, ByVal Segs As String, ByVal Pax As String, ByVal TicketType As String)

        Dim pobjTicket As GDSTicketItem

        Try
            If pTicketNumber > 0 Then
                pobjTicket = New GDSTicketItem

                mintCount = mintCount + 1
                pobjTicket.SetValues(pGDSLine, pTicketType, pTicketNumber, pTicketCount, IssuingAirline, AirlineCode, eTicket, Segs, Pax, TicketType)
                MyBase.Add(Format(mintCount), pobjTicket)
            End If
        Catch ex As Exception
            Throw New Exception("addTicket()" & vbCrLf & Err.Description)
        End Try

    End Sub

    Private Sub ReadTickets()
        ' example : 12 OSI YY ATH VCHR 97893469 AL.M/SG2-3   
        ReDim mobjTickets(0)
        For Each objOSI As s1aPNR.OtherServiceElement In mobjPNR.OtherServiceElements
            If objOSI.Text.IndexOf("ATH VCHR") > 0 Then
                ReDim Preserve mobjTickets(mobjTickets.GetUpperBound(0) + 1)
                mobjTickets(mobjTickets.GetUpperBound(0)) = New GDSTicketItem
                mobjTickets(mobjTickets.GetUpperBound(0)).SetElement(objOSI.Text, Utilities.EnumTicketDocType.VCHR, "", "") ', "") ', "")
            End If
        Next
        For Each objFA As s1aPNR.FareAutoTktElement In mobjPNR.FareAutoTktElements
            If objFA.Text.Replace(" ", "").Contains(MySettings.GDSPcc) Then
                ReDim Preserve mobjTickets(mobjTickets.GetUpperBound(0) + 1)
                mobjTickets(mobjTickets.GetUpperBound(0)) = New GDSTicketItem
                mobjTickets(mobjTickets.GetUpperBound(0)).SetElement(objFA.Text, Utilities.EnumTicketDocType.ETKT, BuildPaxname(objFA.Associations.Passengers), BuildSegments(objFA.Associations.Segments)) ', "") ', "")
            Else
                ReDim Preserve mobjTickets(mobjTickets.GetUpperBound(0) + 1)
                mobjTickets(mobjTickets.GetUpperBound(0)) = New GDSTicketItem
                mobjTickets(mobjTickets.GetUpperBound(0)).SetElement(objFA.Text, Utilities.EnumTicketDocType.INTR, BuildPaxname(objFA.Associations.Passengers), BuildSegments(objFA.Associations.Segments)) ', "") ', "")
            End If
        Next

        For i As Integer = 1 To mobjTickets.GetUpperBound(0)
            If mobjTickets(i).DocType = Utilities.EnumTicketDocType.VCHR Then
                For j = i + 1 To mobjTickets.GetUpperBound(0)
                    If mobjTickets(j).SegsElementNo.StartsWith(mobjTickets(i).SegsElementNo) Then
                        mobjTickets(i).SegsElementNo = mobjTickets(j).SegsElementNo
                        Exit For
                    End If
                Next
            End If
        Next

    End Sub
    Public ReadOnly Property GetUpperBound As Integer
        Get
            GetUpperBound = mobjTickets.GetUpperBound(0)
        End Get
    End Property
    Public ReadOnly Property Tickets(ByVal Index As Integer) As GDSTicketItem
        Get
            If Index >= 1 And Index <= mobjTickets.GetUpperBound(0) Then
                Tickets = mobjTickets(Index)
            Else
                Throw New Exception("Invalid Index")
            End If
        End Get
    End Property

    Private Function BuildPaxname(ByVal PaxElements As Object) As String

        Dim PaxCount As Integer = 0
        BuildPaxname = ""

        For Each Pax As s1aPNR.NameElement In PaxElements
            BuildPaxname &= MakePaxNameString(Pax)
            PaxCount += 1
        Next
        If PaxCount = 0 Then
            For Each Pax As s1aPNR.NameElement In mobjPNR.NameElements
                BuildPaxname &= MakePaxNameString(Pax)
            Next
        End If

    End Function

    Private Function MakePaxNameString(ByVal Pax As s1aPNR.NameElement) As String

        MakePaxNameString = Pax.ElementNo & ". " & Pax.LastName & "/" & Pax.Initial
        If Pax.PassengerType <> "" Then
            MakePaxNameString &= " (" & Pax.PassengerType & ")"
        End If
        If Pax.ID <> "" Then
            MakePaxNameString &= " (" & Pax.ID & ")"
        End If
        MakePaxNameString &= vbCrLf

    End Function

    Private Function BuildSegments(ByVal SegElements As Object) As String

        Dim SegCount As Integer = 0
        Dim SegNo As String = ""
        Dim SegItn As String = ""
        Dim SegClass As String = ""
        Dim PrevSeg As String = ""
        Dim ElementNo As Integer = 0

        BuildSegments = ""

        Try

            For Each Seg As Object In SegElements
                ElementNo = Seg.ElementNo
                ' the first segment no goes automatically into the list as the start
                If SegNo = "" Then
                    SegNo &= ElementNo
                Else
                    ' compare the segment no with the last one in the list
                    Dim a2 As Integer = CInt(SegNo.Substring(SegNo.Length - 1, 1))
                    ' if it is in sequence
                    If ElementNo = a2 + 1 Then
                        ' and it is the second one, put it in the list preceded by a hyphen
                        If SegNo.Length = 1 Then
                            SegNo &= "-" & ElementNo
                        Else
                            ' otherwise replace the last no with this one
                            SegNo = SegNo.Substring(0, SegNo.Length - 1) & ElementNo
                        End If
                    Else
                        ' if it is not in sequence, precede it with a comma
                        SegNo &= "," & ElementNo
                    End If
                End If
                If PrevSeg = Seg.BoardPoint Then
                    SegItn &= "-" & Seg.OffPoint
                    SegClass &= "-" & Seg.Class
                Else
                    If SegItn <> "" Then
                        SegItn &= "-***-"
                        SegClass &= "-"
                    End If
                    SegItn &= Seg.BoardPoint & "-" & Seg.OffPoint
                    SegClass &= Seg.Class
                End If
                PrevSeg = Seg.OffPoint
                SegCount += 1
            Next
            If SegCount = 0 Then
                For Each Seg As Object In mobjPNR.AirSegments
                    SegNo &= Seg.ElementNo
                    If PrevSeg = Seg.BoardPoint Then
                        SegItn &= "-" & Seg.OffPoint
                        SegClass &= "-" & Seg.Class
                    Else
                        If SegItn <> "" Then
                            SegItn &= "-***-"
                            SegClass &= "-*-"
                        End If
                        SegItn &= Seg.BoardPoint & "-" & Seg.OffPoint
                        SegClass &= Seg.Class
                    End If
                    PrevSeg = Seg.OffPoint
                Next
            End If

            BuildSegments = "S" & SegNo & ": " & SegItn & ": " & SegClass
        Catch ex As MissingMemberException
            If ElementNo <> 0 Then
                BuildSegments = mobjPNR.AllElements.Item(ElementNo).text
            Else
                BuildSegments = ""
            End If
        Catch ex As Exception
            BuildSegments = ""
        End Try

    End Function

End Class
