Public Class GDSTicketItem
    Private Structure ClassProps
        Dim DocType As Utilities.EnumTicketDocType ' 1=ETKT 2= VCHR 3=Interoffice ticket
        Dim ID As String
        Dim PaxType As String
        Dim TicketNumber As String
        Dim IssuingAirline As String
        Dim Price As String
        Dim IssueDate As String
        Dim PCC As String
        Dim IATA As String
        Dim PaxID As String
        Dim SegsElementNo As String
        Dim SegsDescription As String
        Dim ClassAir As String
        Dim ClassCust As String
        Dim RawText As String
        Dim SellingPrice As Decimal

        Dim GDSLine As String
        Dim StockType As Integer
        Dim Document As Decimal
        Dim Books As Integer
        Dim AirlineCode As String
        Dim eTicket As Boolean
        Dim Segs As String
        Dim Pax As String
        Dim TicketType As String

    End Structure
    Private mudtProps As ClassProps

    Public ReadOnly Property DocType As Utilities.EnumTicketDocType
        Get
            DocType = mudtProps.DocType
        End Get
    End Property

    Public ReadOnly Property TicketNumber As String
        Get
            TicketNumber = mudtProps.TicketNumber
        End Get
    End Property
    Public ReadOnly Property Document() As Decimal
        Get
            Document = mudtProps.Document
        End Get
    End Property
    Public ReadOnly Property LastDocument As Decimal
        Get
            LastDocument = mudtProps.Document + mudtProps.Books - 1
        End Get
    End Property
    Public ReadOnly Property Conjunction As String
        Get
            Dim pTemp As String = LastDocument.ToString
            If pTemp.Length = 10 Then
                Return "-" & pTemp.Substring(7)
            Else
                Return ""
            End If
        End Get
    End Property
    Public ReadOnly Property Books As Integer
        Get
            Books = mudtProps.Books
        End Get
    End Property
    Public ReadOnly Property AirlineCode As String
        Get
            AirlineCode = mudtProps.AirlineCode.Trim
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
    Friend Sub SetValues(ByRef pGDSLine As String, ByRef pStockType As Integer, ByRef pDocument As Decimal, ByRef pBooks As Integer, ByRef pIssuingAirline As String, ByVal AirlineCode As String, ByRef peTicket As Boolean, pSegs As String, pPax As String, pTicketType As String)

        With mudtProps
            .GDSLine = pGDSLine
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
    Public ReadOnly Property ID As String
        Get
            ID = mudtProps.ID
        End Get
    End Property
    Public ReadOnly Property PaxType As String
        Get
            PaxType = mudtProps.PaxType
        End Get
    End Property
    Public ReadOnly Property IssuingAirline As String
        Get
            IssuingAirline = mudtProps.IssuingAirline
        End Get
    End Property
    Public ReadOnly Property Price As String
        Get
            Price = mudtProps.Price
        End Get
    End Property
    Public ReadOnly Property IssueDate As String
        Get
            IssueDate = mudtProps.IssueDate
        End Get
    End Property
    Public ReadOnly Property PCC As String
        Get
            PCC = mudtProps.PCC
        End Get
    End Property
    Public ReadOnly Property IATA As String
        Get
            IATA = mudtProps.IATA
        End Get
    End Property
    Public ReadOnly Property RawText As String
        Get
            RawText = mudtProps.RawText
        End Get
    End Property
    Public ReadOnly Property PaxID() As String
        Get
            PaxID = mudtProps.PaxID
        End Get
    End Property
    Public Property SegsElementNo As String
        Get
            SegsElementNo = mudtProps.SegsElementNo
        End Get
        Set(value As String)
            mudtProps.SegsElementNo = value
        End Set
    End Property
    Public Property SegsDescription As String
        Get
            SegsDescription = mudtProps.SegsDescription
        End Get
        Set(value As String)
            mudtProps.SegsDescription = value
        End Set
    End Property
    Public Property ClassAir As String
        Get
            ClassAir = mudtProps.ClassAir
        End Get
        Set(value As String)
            mudtProps.ClassAir = value
        End Set
    End Property
    Public Property ClassCust As String
        Get
            ClassCust = mudtProps.ClassCust
        End Get
        Set(value As String)
            mudtProps.ClassCust = value
        End Set
    End Property
    Public Property SellingPrice As Decimal
        Get
            SellingPrice = mudtProps.SellingPrice
        End Get
        Set(value As Decimal)
            mudtProps.SellingPrice = value
        End Set
    End Property
    Public Sub SetElement(ByVal RawText As String, ByVal DocType As Utilities.EnumTicketDocType, ByVal PaxID As String, ByVal SegsElementNo As String) ', ByVal SegsDescription As String) ', ByVal ClassAir As String)

        With mudtProps
            ' 2 examples of RawText
            ' 30 FA PAX 724-4175946315/ETLX/EUR369.39/13SEP13/ATHG42100/27280                
            '       573/S5-8/P4
            '28 FA PAX 157-4175946329/ETQR/13SEP13/ATHG42100/27280573                       
            '       /S6-7/P1 
            ' after the split we can have:
            '(0) - 30FAPAX724-4175946315
            '(1) - ETLX
            '(2) - EUR369.39
            '(3) - 13SEP13
            '(4) - ATHG42100'
            '(5) - 27280573
            '(6) - S5-8
            '(7) - P4
            '
            ' or
            '
            '(0) - 28FAPAX157-4175946329
            '(1) - ETQR
            '(2) - 13SEP13
            '(3) - ATHG42100
            '(4) - 27280573
            '(5) - S6-7
            '(6) - P1 
            '
            ' or for a voucher
            '
            ' 21 OSI YY ATH VCHR 9783035 AL.O/SG4   
            '


            .RawText = RawText
            .DocType = DocType
            .PaxID = PaxID
            .ClassCust = ""
            Dim pSegs() As String = Split(SegsElementNo, ":")
            If pSegs.GetUpperBound(0) = 2 Then
                .SegsElementNo = pSegs(0).Trim
                .SegsDescription = pSegs(1).Trim
                .ClassAir = pSegs(2).Trim
            Else
                .SegsElementNo = ""
                .SegsDescription = ""
                .ClassAir = ""
            End If

            If DocType = Utilities.EnumTicketDocType.VCHR Then
                Dim iVchrFrom As Integer = -1
                Dim iALFrom As Integer = -1
                Dim iSGFrom As Integer = -1
                .PaxType = "Voucher"
                iVchrFrom = .RawText.IndexOf("VCHR")
                If iVchrFrom > 0 And iVchrFrom < .RawText.Length + 5 Then
                    iALFrom = .RawText.IndexOf("AL", iVchrFrom + 4)
                    If iALFrom > 0 And iALFrom < .RawText.Length + 5 Then
                        iSGFrom = .RawText.IndexOf("/SG", iALFrom + 4)
                    Else
                        iSGFrom = .RawText.IndexOf("/SG", iVchrFrom + 4)
                    End If
                    If iALFrom = -1 Then
                        iALFrom = .RawText.Length
                    End If
                    If iSGFrom = -1 Then
                        iSGFrom = .RawText.Length
                    End If
                    If iALFrom <= iSGFrom Then
                        .TicketNumber = .RawText.Substring(iVchrFrom + 4, iALFrom - iVchrFrom - 5)
                    Else
                        .TicketNumber = .RawText.Substring(iVchrFrom + 4, iSGFrom - iVchrFrom - 5)
                    End If
                    If iALFrom < .RawText.Length - 4 Then
                        .IssuingAirline = .RawText.Substring(iALFrom + 2, iSGFrom - iALFrom - 2)
                    End If
                    If iSGFrom < .RawText.Length - 3 Then
                        .SegsElementNo = "S" & .RawText.Substring(iSGFrom + 3, .RawText.Length - iSGFrom - 3)
                    End If
                Else
                    .TicketNumber = .RawText
                End If

            Else
                Dim pItems() As String = Split(RawText.Replace(" ", ""), "/")
                If pItems.GetUpperBound(0) >= 4 Then
                    .TicketNumber = pItems(0)
                    Dim i1 As Integer = pItems(0).IndexOf("FA")
                    If i1 > 0 Then
                        .ID = pItems(0).Substring(0, i1)
                        .PaxType = pItems(0).Substring(i1 + 2, 3)
                        .TicketNumber = pItems(0).Substring(i1 + 5)
                    End If
                    .IssuingAirline = pItems(1).Substring(2)
                    Dim pPriceIndex As Integer = 2
                    If IsNumeric(("0" & pItems(2)).Substring(0, 1)) Then
                        pPriceIndex = 1
                        .Price = ""
                    Else
                        .Price = pItems(pPriceIndex)
                    End If
                    .IssueDate = pItems(pPriceIndex + 1)
                    .PCC = pItems(pPriceIndex + 2)
                    .IATA = pItems(pPriceIndex + 3)
                Else
                    .TicketNumber = RawText
                    .ID = PaxID
                    .PaxType = ""
                    .IssuingAirline = ""
                    .Price = ""
                    .IssueDate = ""
                    .PCC = ""
                    .IATA = ""
                End If
            End If
        End With

    End Sub
End Class