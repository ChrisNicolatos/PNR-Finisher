Option Strict On
Option Explicit On
Imports System.IO
Imports System.Text

Module modPNR
    Public Const MONTH_NAMES As String = "JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC"
    Public Enum EnumLoGLanguage
        English = 0
        Brazil = 1
    End Enum

    Private mMySettings As Config
    Private mHomeSettings As Config
    Private mHomeSettingsExist As Boolean
    Private mstrRequestedPCC As String = ""
    Private mstrRequestedUser As String = ""
    Private mClassOfService As New ClassOfServiceCollection

    Public Sub InitSettings()
        Try
            mstrRequestedPCC = ""
            mstrRequestedUser = ""

            mMySettings = New Config
            If Not mHomeSettingsExist Then
                mHomeSettings = mMySettings
                mHomeSettingsExist = True
            End If
        Catch ex As Exception
            If mHomeSettingsExist Then
                mMySettings = mHomeSettings
            Else
                Throw New Exception(ex.Message)
            End If
        End Try


    End Sub
    Public ReadOnly Property GetClassOfService(ByVal pCarrier As String, ByVal pOrigin As String, ByVal pDestination As String, ByVal pClassOfService As String) As ClassOfServiceItem
        Get
            Return mClassOfService.Load(pCarrier, pOrigin, pDestination, pClassOfService)
        End Get
    End Property
    Public Sub InitSettings(ByVal mGDSUser As GDSUser)
        Try
            mstrRequestedPCC = mGDSUser.PCC
            mstrRequestedUser = mGDSUser.User

            mMySettings = New Config(mGDSUser)
            If Not mHomeSettingsExist Then
                mHomeSettings = mMySettings
                mHomeSettingsExist = True
            End If
        Catch ex As Exception
            If mHomeSettingsExist Then
                mMySettings = mHomeSettings
            Else
                Throw New Exception(ex.Message)
            End If
        End Try

    End Sub
    Public ReadOnly Property MySettings As Config
        Get
            MySettings = mMySettings
        End Get

    End Property
    Public ReadOnly Property MyHomeSettings As Config
        Get
            MyHomeSettings = mHomeSettings
        End Get
    End Property
    Public ReadOnly Property RequestedPCC As String
        Get
            RequestedPCC = mstrRequestedPCC
        End Get
    End Property
    Public ReadOnly Property RequestedUser As String
        Get
            RequestedUser = mstrRequestedUser
        End Get
    End Property
    Public Function MyMonthName(ByVal pDate As Date, ByVal Language As EnumLoGLanguage) As String
        Static Dim pNamesLang1() As String = {"janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"}

        If Language = EnumLoGLanguage.Brazil Then
            Return pDate.Day & " de " & pNamesLang1(pDate.Month - 1) & " de " & pDate.Year
        Else
            Return pDate.Day & " " & MonthName(pDate.Month) & " " & pDate.Year
        End If

    End Function

    Public Function myCurr(ByVal StringToParse As String) As Decimal

        Dim pintPoint As Integer
        Dim pintComma As Integer

        Do While Not IsNumeric(Right(StringToParse, 1)) And StringToParse.Length > 0
            StringToParse = Left(StringToParse, StringToParse.Length - 1)
        Loop
        StringToParse = StringToParse.Trim
        pintPoint = InStr(StringToParse, My.Application.Culture.NumberFormat.CurrencyGroupSeparator)
        pintComma = InStr(StringToParse, My.Application.Culture.NumberFormat.CurrencyDecimalSeparator)
        If pintPoint > pintComma Then
            If StringToParse.Length > 2 Then
                If StringToParse.Substring(StringToParse.Length - 3, 1) = My.Application.Culture.NumberFormat.CurrencyGroupSeparator Then
                    Mid(StringToParse, StringToParse.Length - 2, 1) = My.Application.Culture.NumberFormat.CurrencyDecimalSeparator
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
            For i As Integer = 1 To StringToParse.Length
                If IsNumeric(Mid(StringToParse, 1, i)) Then
                    myCurr = CDec(Mid(StringToParse, 1, i))
                Else
                    Exit For
                End If
            Next i
        End If

    End Function

    Public Function DateFromIATA(ByVal InDate As String) As Date

        Dim pintDay As Integer
        Dim pintMonth As Integer
        Dim pintYear As Integer

        Try
            If Not IsNothing(InDate) Then
                If Not Date.TryParse(InDate, DateFromIATA) Then
                    DateFromIATA = Date.MinValue
                    If InDate.Length > 5 Then
                        pintDay = CInt(InDate.Substring(0, 2))
                        pintMonth = CInt((MONTH_NAMES.IndexOf(InDate.Substring(3, 3)) + 2) / 3)
                        pintYear = CInt(InDate.Substring(5))

                        If pintMonth >= 1 Then
                            DateFromIATA = DateSerial(pintYear, pintMonth, pintDay)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception

        End Try

    End Function
    Public Function DateToIATA(ByVal InDate As Date) As String

        Return Format(InDate.Day, "00") & MONTH_NAMES.Substring(InDate.Month * 3 - 3, 3) & Format(InDate.Year Mod 100, "00")

    End Function
End Module
