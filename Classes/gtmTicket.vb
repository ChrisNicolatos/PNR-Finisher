Option Strict Off
Option Explicit On
Friend Class gtmTicket

	Private Structure ClassProps
		Dim AmadeusLine As String
		Dim StockType As Short
		Dim Document As Decimal
		Dim Books As Short
        Dim IssuingAirline As String
        Dim AirlineCode As String
        Dim eTicket As Boolean
        Dim Segs As String
        Dim Pax As String
        Dim TicketType As String
	End Structure
	Private mudtProps As ClassProps

    'Public ReadOnly Property AmadeusLine() As String
    '	Get

    '		AmadeusLine = Trim(mudtProps.AmadeusLine)

    '	End Get
    'End Property
    'Public ReadOnly Property StockType() As Short
    '	Get

    '		StockType = mudtProps.StockType

    '	End Get
    'End Property
    Public ReadOnly Property Document() As Decimal
		Get
			
			Document = mudtProps.Document
			
		End Get
	End Property
    'Public ReadOnly Property Books() As Short
    '	Get

    '		Books = mudtProps.Books

    '	End Get
    'End Property
    Public ReadOnly Property IssuingAirline() As String
		Get
			
			IssuingAirline = Trim(mudtProps.IssuingAirline)
			
		End Get
    End Property
    Public ReadOnly Property AirlineCode As String
        Get
            AirlineCode = Trim(mudtProps.AirlineCode)
        End Get
    End Property
	Public ReadOnly Property eTicket() As Boolean
		Get
			
			eTicket = mudtProps.eTicket
			
		End Get
    End Property
    Public ReadOnly Property Segs As String
        Get
            Segs = mudtProps.Segs
        End Get
    End Property
    Public ReadOnly Property Pax As String
        Get
            Pax = mudtProps.Pax
        End Get
    End Property
    Public ReadOnly Property TicketType As String
        Get
            TicketType = mudtProps.TicketType
        End Get
    End Property
    Friend Sub SetValues(ByRef pAmadeusLine As String, ByRef pStockType As Short, ByRef pDocument As Decimal, ByRef pBooks As Short, ByRef pIssuingAirline As String, ByVal AirlineCode As String, ByRef peTicket As Boolean, pSegs As String, pPax As String, pTicketType As String)

        With mudtProps
            .AmadeusLine = pAmadeusLine
            .StockType = pStockType
            .Document = pDocument
            .Books = pBooks
            .IssuingAirline = pIssuingAirline
            .AirlineCode = AirlineCode
            .eTicket = peTicket
            .Segs = pSegs
            .Pax = pPax
            .TicketType = pTicketType
        End With

    End Sub
End Class