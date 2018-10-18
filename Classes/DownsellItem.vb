Public Class DownsellItem
    Private Structure ClassProps
        Dim OwnPNR As Integer
        Dim PCC As String
        Dim GDS As String
        Dim PNR As String
        Dim UserGdsId As String
        Dim DownsellDecision As String
        Dim PaxName As String
        Dim Itinerary As String
        Dim Total As Decimal
        Dim DownsellTotal As Decimal
        Dim FareBasis As String
        Dim DownsellFareBasis As String
        Dim GDSCommand As String
    End Structure
    Private mudtProps As ClassProps
    Public ReadOnly Property OwnPNR As Integer
        Get
            Return mudtProps.OwnPNR
        End Get
    End Property
    Public ReadOnly Property PCC As String
        Get
            Return mudtProps.PCC
        End Get
    End Property
    Public ReadOnly Property GDS As String
        Get
            Return mudtProps.GDS
        End Get
    End Property
    Public ReadOnly Property PNR As String
        Get
            Return mudtProps.PNR
        End Get
    End Property
    Public ReadOnly Property UserGdsId As String
        Get
            Return mudtProps.UserGdsId
        End Get
    End Property
    Public ReadOnly Property DownsellDecision As String
        Get
            Return mudtProps.DownsellDecision
        End Get
    End Property
    Public ReadOnly Property PaxName As String
        Get
            Return mudtProps.PaxName
        End Get
    End Property
    Public ReadOnly Property Itinerary As String
        Get
            Return mudtProps.Itinerary
        End Get
    End Property
    Public ReadOnly Property Total As Decimal
        Get
            Return mudtProps.Total
        End Get
    End Property
    Public ReadOnly Property DownsellTotal As Decimal
        Get
            Return mudtProps.DownsellTotal
        End Get
    End Property
    Public ReadOnly Property FareBasis As String
        Get
            Return mudtProps.FareBasis
        End Get
    End Property
    Public ReadOnly Property DownsellFareBasis As String
        Get
            Return mudtProps.DownsellFareBasis
        End Get
    End Property
    Public ReadOnly Property GDSCommand As String
        Get
            Return mudtProps.GDSCommand
        End Get
    End Property
    Public Sub SetValues(ByVal pOwnPNR As Integer, ByVal pPCC As String, ByVal pGDS As String, ByVal pPNR As String _
                          , ByVal pUserGdsId As String, ByVal pDownsellDecision As String _
                          , ByVal pPaxName As String, ByVal pItinerary As String _
                          , ByVal pTotal As Decimal, ByVal pDownsellTotal As Decimal, ByVal pFareBasis As String, ByVal pDownsellFareBasis As String _
                          , ByVal pGDSCommand As String)
        With mudtProps
            .OwnPNR = pOwnPNR
            .PCC = pPCC
            .GDS = pGDS
            .PNR = pPNR
            .UserGdsId = pUserGdsId
            .DownsellDecision = pDownsellDecision
            .PaxName = pPaxName
            .Itinerary = pItinerary
            .Total = pTotal
            .DownsellTotal = pDownsellTotal
            .FareBasis = pFareBasis
            .DownsellFareBasis = pDownsellFareBasis
            .GDSCommand = pGDSCommand
        End With
    End Sub

End Class
