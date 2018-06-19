Public NotInheritable Class Utilities
    Private Const MONTH_NAMES As String = "JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC"
    Public Enum EnumItnFormat
        DefaultFormat = 0
        Plain = 1
        SeaChefs = 2
        SeaChefsWithCode = 3
        Euronav = 4
    End Enum
    Public Enum EnumGDSCode
        Unknown = 0
        Amadeus = 1
        Galileo = 2
    End Enum
    Public Enum EnumCustomPropertyID As Integer
        None = 0
        BookedBy = 1
        Department = 2
        ReasonFortravel = 4
        CostCentre = 5
        Savings = 6
        Losses = 7
        SavingsLossesReason = 8
        TravelDefinition = 9
        VesselCostCentre = 10
        RequisitionNumber = 11
        PassengerID = 12
        OPT = 13
        TRId = 14
    End Enum
    Public Enum EnumTicketDocType
        NONE = 0
        ETKT = 1
        VCHR = 2
        INTR = 3
    End Enum
    Public Enum EnumLoGLanguage
        English = 0
        Brazil = 1
    End Enum
    Public Enum CustomPropertyRequiredType
        PropertyOptional = 613
        PropertyReqToSave = 614
        PropertyReqToInv = 615
    End Enum
    Private Sub New()
    End Sub

    Public Shared Function MyMonthName(ByVal pDate As Date, ByVal Language As EnumLoGLanguage) As String
        Static Dim pNamesLang1() As String = {"janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"}


        If Language = EnumLoGLanguage.Brazil Then
            Return pDate.Day & " de " & pNamesLang1(pDate.Month - 1) & " de " & pDate.Year
        Else
            Return pDate.Day & " " & MonthName(pDate.Month) & " " & pDate.Year
        End If

    End Function
    Public Shared Function myCurr(ByVal StringToParse As String) As Decimal

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
    Public Shared Function DateFromIATA(ByVal InDate As String) As Date

        Dim pintDay As Integer
        Dim pintMonth As Integer
        Dim pintYear As Integer

        Try
            If Not IsNothing(InDate) Then
                If Not Date.TryParse(InDate, DateFromIATA) Then
                    DateFromIATA = Date.MinValue
                    If InDate.Length > 5 Then
                        pintDay = InDate.Substring(0, 2)
                        pintMonth = (MONTH_NAMES.IndexOf(InDate.Substring(3, 3)) + 2) / 3
                        pintYear = InDate.Substring(5)

                        If pintMonth >= 1 Then
                            DateFromIATA = DateSerial(pintYear, pintMonth, pintDay)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception

        End Try

    End Function

    Public Shared Function DateToIATA(ByVal InDate As Date) As String

        DateToIATA = Format(InDate.Day, "00") & MONTH_NAMES.Substring(InDate.Month * 3 - 3, 3) & Format(InDate.Year Mod 100, "00")

    End Function
End Class

