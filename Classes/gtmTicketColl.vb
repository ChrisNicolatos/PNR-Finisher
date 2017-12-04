Option Strict Off
Option Explicit On
Friend Class gtmTicketColl
    Inherits Collections.Generic.Dictionary(Of String, gtmTicket)
	
    Private mintCount As Short
	
    Public Sub addTicket(ByVal pAmadeusLine As String, ByVal pTicketType As Short, ByVal pTicketNumber As Decimal, ByVal pTicketCount As Short, ByVal IssuingAirline As String, ByVal AirlineCode As String, ByVal eTicket As Boolean, ByVal Segs As String, ByVal Pax As String, ByVal TicketType As String)

        Dim pobjTicket As gtmTicket

        Try
            pobjTicket = New gtmTicket

            mintCount = mintCount + 1
            pobjTicket.SetValues(pAmadeusLine, pTicketType, pTicketNumber, pTicketCount, IssuingAirline, AirlineCode, eTicket, Segs, Pax, TicketType)
            MyBase.Add(Format(mintCount), pobjTicket)
        Catch ex As Exception
            Throw New Exception("addTicket()" & vbCrLf & Err.Description)
        End Try

    End Sub
End Class