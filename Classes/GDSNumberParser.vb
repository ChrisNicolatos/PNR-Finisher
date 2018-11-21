Option Strict On
Option Explicit On
Friend Class GDSNumberParser

    Private Structure ParserProps

        Dim ParseType As Integer ' 1 = Ticket number text 2 = Segment number association

        ' Ticket number text properties
        Dim TicketNumberText As String ' Input
        Dim StockType As Integer
        Dim DocumentNumber As Decimal
        Dim Books As Integer
        Dim Airline As String
        Dim AirlineNumber As String

        ' Segment number association text properties
        Dim SegmentAssociationText As String
        Dim SegmentCount As Integer
        Dim SegmentElements() As Integer
        Dim isValid As Boolean
    End Structure

    Private mudtProps As ParserProps

    Public Function TicketNumberText(ByVal Value As String) As Boolean

        With mudtProps
            .ParseType = 1
            .TicketNumberText = Value.Trim
            .isValid = False
        End With

        TicketNumberText = parseTicketNumber()

    End Function

    Public ReadOnly Property StockType() As Integer
        Get

            StockType = mudtProps.StockType

        End Get
    End Property
    Public ReadOnly Property DocumentNumber() As Decimal
        Get

            DocumentNumber = mudtProps.DocumentNumber

        End Get
    End Property
    Public ReadOnly Property Books() As Integer
        Get

            Books = mudtProps.Books

        End Get
    End Property
    Public ReadOnly Property AirlineNumber() As String
        Get

            AirlineNumber = mudtProps.AirlineNumber

        End Get
    End Property
    Private Function parseTicketNumber() As Boolean
        ' acceptable strings are:
        ' TTTTTTTTTT 10 digit ticket number
        ' AAATTTTTTTTTT 3 digit airline + 10 digit ticket number
        ' AAA-TTTTTTTTTT 3 digit airline + 10 digit ticket number separated by hyphen
        ' all the above can have conjunction tickets with last digits of last ticket separated by hyphen
        ' the number of last digits is variable
        ' TTTTTTTTTT-XX 10 digit ticket number
        ' AAATTTTTTTTTT-XX 3 digit airline + 10 digit ticket number
        ' AAA-TTTTTTTTTT-XX 3 digit airline + 10 digit ticket number separated by hyphen

        Dim i As Integer
        Dim pstrTicket As String = ""
        Dim pintAirlineFrom As Integer
        Dim pintAirlineTo As Integer
        Dim pintTicketFrom As Integer
        Dim pintTicketTo As Integer
        Dim pintConjFrom As Integer
        Dim pintConjTo As Integer
        Dim pintEndOfString As Integer

        Dim pstrTemp As String
        Dim pstrTemp2 As String
        Dim pcurrDoc2 As Decimal
        Dim pflgOK As Boolean

        Try
            With mudtProps
                .Airline = ""
                .AirlineNumber = ""
                .Books = 0
                .DocumentNumber = 0
                .StockType = 1

                pintEndOfString = .TicketNumberText.Length
                For i = 1 To .TicketNumberText.Length
                    pflgOK = False
                    pstrTemp = .TicketNumberText.Substring(i - 1, 1)
                    If pstrTemp >= "0" And pstrTemp <= "9" Then
                        pflgOK = True
                    ElseIf pstrTemp = "-" Then
                        Select Case i
                            Case 4
                                pintAirlineFrom = 1
                                pintAirlineTo = 3
                                pintTicketFrom = 5
                                pintTicketTo = 14
                                pflgOK = True
                            Case 11
                                If pintTicketFrom <> 0 Then
                                    pintEndOfString = 0
                                    Exit For
                                End If
                                pintAirlineFrom = 0
                                pintAirlineTo = 0
                                pintTicketFrom = 1
                                pintTicketTo = 10
                                pintConjFrom = 12
                                pflgOK = True
                            Case 14
                                If pintConjFrom <> 0 Then
                                    pintEndOfString = 0
                                    Exit For
                                End If
                                pintAirlineFrom = 1
                                pintAirlineTo = 3
                                pintTicketFrom = 4
                                pintTicketTo = 13
                                pintConjFrom = 15
                                pflgOK = True
                            Case 15
                                If pintConjFrom <> 0 Then
                                    pintEndOfString = 0
                                    Exit For
                                End If
                                pintAirlineFrom = 1
                                pintAirlineTo = 3
                                pintTicketFrom = 5
                                pintTicketTo = 14
                                pintConjFrom = 16
                                pflgOK = True
                        End Select
                    End If
                    If Not pflgOK Then
                        pintEndOfString = i - 1
                        Exit For
                    End If
                Next i
                If pintConjFrom > 0 Then
                    pintConjTo = pintEndOfString
                End If

                If pintEndOfString < 10 Or (pintConjFrom > 0 And pintConjTo <= pintConjFrom) Then
                    .isValid = False
                ElseIf pintEndOfString = 10 And pintTicketFrom = 0 Then
                    pintAirlineFrom = 0
                    pintAirlineTo = 0
                    pintTicketFrom = 1
                    pintTicketTo = 10
                ElseIf pintEndOfString = 13 And pintTicketFrom = 0 Then
                    pintAirlineFrom = 1
                    pintAirlineTo = 3
                    pintTicketFrom = 4
                    pintTicketTo = 13
                End If

                If pintAirlineFrom > 0 And pintAirlineTo > pintAirlineFrom Then
                    .AirlineNumber = Format(myCurr(Mid(.TicketNumberText, pintAirlineFrom, pintAirlineTo - pintAirlineFrom + 1)), "000")
                End If

                If pintTicketFrom > 0 And pintTicketTo > pintTicketFrom Then
                    pstrTicket = .TicketNumberText.Substring(pintTicketFrom - 1, pintTicketTo - pintTicketFrom + 1)
                    .DocumentNumber = myCurr(pstrTicket)
                    .Books = 1
                End If

                If pintConjFrom > 0 And pintConjTo >= pintConjFrom Then
                    pstrTemp = .TicketNumberText.Substring(pintConjFrom - 1, pintConjTo - pintConjFrom + 1)
                    pstrTemp2 = pstrTicket
                    If pstrTemp.Length <= pstrTemp2.Length Then
                        Mid(pstrTemp2, pstrTemp2.Length - pstrTemp.Length + 1, pstrTemp.Length) = pstrTemp
                        pcurrDoc2 = myCurr(pstrTemp2)
                        If pcurrDoc2 > .DocumentNumber Then
                            .Books = CInt(pcurrDoc2 - .DocumentNumber + 1)
                        End If
                    End If
                End If

                If .DocumentNumber > 0 Then
                    .isValid = True
                Else
                    .isValid = False
                End If

                parseTicketNumber = .isValid
            End With
        Catch ex As Exception
            With mudtProps
                .Airline = ""
                .AirlineNumber = ""
                .Books = 0
                .DocumentNumber = 0
                .StockType = 0
                .isValid = False
            End With

            parseTicketNumber = False
        End Try

    End Function
End Class