Option Strict On
Option Explicit On
Public Class AlertsItem
    Private Structure ClassProps
        Dim BackOfficeId As Integer
        Dim ClientCode As String
        Dim AlertForFinisher As String
        Dim OriginCountry As String
        Dim DestinationCountry As String
        Dim Airline As String
        Dim AmadeusQueue As String
        Dim GalileoQueue As String
        Dim AlertForDownsell As String
    End Structure
    Dim mudtprops As ClassProps
    Public ReadOnly Property BackOfficeID As Integer
        Get
            Return mudtprops.BackOfficeId
        End Get
    End Property
    Public ReadOnly Property ClientCode As String
        Get
            Return mudtprops.ClientCode
        End Get
    End Property
    Public ReadOnly Property AlertForFinisher() As String
        Get
            Return mudtprops.AlertForFinisher
        End Get
    End Property
    Public ReadOnly Property OriginCountry As String
        Get
            Return mudtprops.OriginCountry
        End Get
    End Property
    Public ReadOnly Property DestinationCountry As String
        Get
            Return mudtprops.DestinationCountry
        End Get
    End Property
    Public ReadOnly Property Airline As String
        Get
            Return mudtprops.Airline
        End Get
    End Property
    Public ReadOnly Property AmadeusQueue As String
        Get
            Return mudtprops.AmadeusQueue
        End Get
    End Property
    Public ReadOnly Property GalileoQueue As String
        Get
            Return mudtprops.GalileoQueue
        End Get
    End Property
    Public ReadOnly Property AlertForDownsell As String
        Get
            Return mudtprops.AlertForDownsell
        End Get
    End Property
    Friend Sub SetValues(ByVal pBackOfficeID As Integer, ByVal pClientCode As String, ByVal pAlertForFinisher As String, ByVal pOriginCountry As String, ByVal pDestinationCountry As String, ByVal pAirline As String, ByVal pAmadeusQueue As String, ByVal pGalileoQueue As String, ByVal pAlertForDownsell As String)
        With mudtprops
            .BackOfficeId = pBackOfficeID
            .ClientCode = pClientCode
            .AlertForFinisher = pAlertForFinisher
            .OriginCountry = pOriginCountry
            .DestinationCountry = pDestinationCountry
            .Airline = pAirline
            .AmadeusQueue = pAmadeusQueue
            .GalileoQueue = pGalileoQueue
            .AlertForDownsell = pAlertForDownsell
        End With
    End Sub

End Class
