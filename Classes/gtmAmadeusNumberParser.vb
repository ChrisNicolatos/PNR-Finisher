Option Strict Off
Option Explicit On
Friend Class gtmAmadeusNumberParser
	
	Private Structure ParserProps

        Dim ParseType As Short ' 1 = Ticket number text 2 = Segment number association

		' Ticket number text properties
		Dim TicketNumberText As String ' Input
		Dim StockType As Short
		Dim DocumentNumber As Decimal
		Dim Books As Short
		Dim Airline As String
        Dim AirlineNumber As String

		' Segment number association text properties
		Dim SegmentAssociationText As String
		Dim SegmentCount As Short
		Dim SegmentElements() As Short
		Dim isValid As Boolean
	End Structure
	
    Private mudtProps As ParserProps
	
    Public Function TicketNumberText(ByVal Value As String) As Boolean

        With mudtProps
            .ParseType = 1
            .TicketNumberText = Trim(Value)
            .isValid = False
        End With

        TicketNumberText = parseTicketNumber

    End Function
	
	Public ReadOnly Property StockType() As Short
		Get
			
			StockType = mudtProps.StockType
			
		End Get
	End Property
	Public ReadOnly Property DocumentNumber() As Decimal
		Get
			
			DocumentNumber = mudtProps.DocumentNumber
			
		End Get
	End Property
	Public ReadOnly Property Books() As Short
		Get
			
			Books = mudtProps.Books
			
		End Get
	End Property
    'Public ReadOnly Property Airline() As String
    '	Get

    '		Airline = mudtProps.Airline

    '	End Get
    'End Property
    Public ReadOnly Property AirlineNumber() As String
		Get
			
			AirlineNumber = mudtProps.AirlineNumber
			
		End Get
	End Property
    'Public ReadOnly Property isValid() As Boolean
    '	Get

    '		isValid = mudtProps.isValid

    '	End Get
    'End Property

    '   Public ReadOnly Property SegmentCount() As Short
    '	Get

    '		SegmentCount = mudtProps.SegmentCount

    '	End Get
    'End Property

    '   Public ReadOnly Property SegmentElement(ByVal Index As Short) As Short
    '	Get

    '		With mudtProps
    '			If Index >= 1 And Index <= .SegmentCount Then
    '				SegmentElement = .SegmentElements(Index)
    '			End If
    '		End With

    '	End Get
    'End Property

    'Public Function SegmentAssociationText(ByVal Value As String) As Boolean

    '	With mudtProps
    '		.ParseType = 2
    '		.SegmentAssociationText = Value
    '		.isValid = False
    '	End With

    '	SegmentAssociationText = parseSegments

    'End Function

    'Private Function parseSegments() As Boolean
    '	' acceptable strings are:
    '	' 1           element 1
    '	' 1-3         elements 1,2,3
    '	' 1,4         elements 1,4
    '	' 1,3-4       elements 1,3,4
    '	' 1-2, 4-5    elements 1,2,4,5

    '	Dim i As Short
    '	Dim k As Short
    '       Dim pstrSplit() As String
    '	Dim pstrSplit2() As String
    '	Dim pintFrom As Short
    '	Dim pintTo As Short

    '       Try
    '           ReDim pstrSplit(0)
    '           With mudtProps
    '               .SegmentCount = 0
    '               ReDim .SegmentElements(0)

    '               pstrSplit2 = Split(.SegmentAssociationText, "/")
    '               If IsArray(pstrSplit2) Then
    '                   For i = LBound(pstrSplit2) To UBound(pstrSplit2)
    '                       If pstrSplit2(i) <> "" Then
    '                           pstrSplit = Split(pstrSplit2(i), ",")
    '                           Exit For
    '                       End If
    '                   Next i
    '               Else
    '                   pstrSplit = Split(.SegmentAssociationText, ",")
    '               End If

    '               If IsArray(pstrSplit) Then
    '                   If Left(pstrSplit(0), 2) = "SG" Then
    '                       pstrSplit(0) = Mid(pstrSplit(0), 3)
    '                   ElseIf Left(pstrSplit(0), 1) = "S" Then
    '                       pstrSplit(0) = Mid(pstrSplit(0), 2)
    '                   End If

    '                   For i = LBound(pstrSplit) To UBound(pstrSplit)
    '                       pstrSplit2 = Split(pstrSplit(i), "-")
    '                       If IsArray(pstrSplit2) Then
    '                           If UBound(pstrSplit2) > 1 Then
    '                               Throw New Exception("Segment association contains spurious separator " & pstrSplit(i))
    '                           End If
    '                           pintFrom = myCurr(pstrSplit2(0))
    '                           If UBound(pstrSplit2) = 1 Then
    '                               pintTo = myCurr(pstrSplit2(1))
    '                           Else
    '                               pintTo = pintFrom
    '                           End If
    '                           If pintFrom <> 0 And pintTo >= pintFrom Then
    '                               For k = pintFrom To pintTo
    '                                   addToSegments(k)
    '                               Next k
    '                           End If
    '                       End If
    '                   Next i
    '               End If

    '               parseSegments = True

    '           End With
    '       Catch ex As Exception
    '           With mudtProps
    '               .SegmentCount = 0
    '               ReDim .SegmentElements(0)
    '           End With

    '           parseSegments = False
    '       End Try

    'End Function

    '   Private Sub addToSegments(ByVal Element As Short)

    '	Dim i As Short

    '	If Element < 1 Then
    '           Throw New Exception("Invalid segment number")
    '	End If
    '	With mudtProps
    '		For i = 1 To .SegmentCount
    '			If Element = .SegmentElements(i) Then
    '                   Throw New Exception("Duplicate segment association")
    '			End If
    '		Next i
    '		.SegmentCount = .SegmentCount + 1
    '		ReDim Preserve .SegmentElements(.SegmentCount)
    '		.SegmentElements(.SegmentCount) = Element
    '	End With

    'End Sub
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
		
		Dim i As Short
        Dim pstrTicket As String = ""
		Dim pintAirlineFrom As Short
		Dim pintAirlineTo As Short
		Dim pintTicketFrom As Short
		Dim pintTicketTo As Short
		Dim pintConjFrom As Short
		Dim pintConjTo As Short
		Dim pintEndOfString As Short
		
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

                pintEndOfString = Len(.TicketNumberText)
                For i = 1 To Len(.TicketNumberText)
                    pflgOK = False
                    pstrTemp = Mid(.TicketNumberText, i, 1)
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
                    pstrTicket = Mid(.TicketNumberText, pintTicketFrom, pintTicketTo - pintTicketFrom + 1)
                    .DocumentNumber = myCurr(pstrTicket)
                    .Books = 1
                End If

                If pintConjFrom > 0 And pintConjTo >= pintConjFrom Then
                    pstrTemp = Mid(.TicketNumberText, pintConjFrom, pintConjTo - pintConjFrom + 1)
                    pstrTemp2 = pstrTicket
                    If Len(pstrTemp) <= Len(pstrTemp2) Then
                        Mid(pstrTemp2, Len(pstrTemp2) - Len(pstrTemp) + 1, Len(pstrTemp)) = pstrTemp
                        pcurrDoc2 = myCurr(pstrTemp2)
                        If pcurrDoc2 > .DocumentNumber Then
                            .Books = pcurrDoc2 - .DocumentNumber + 1
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