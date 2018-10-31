Option Strict On
Option Explicit On
Public Class AirlineNotesItem
    Private Structure ClassProps
        Dim ID As Integer
        Dim AirlineCode As String
        Dim FlightType As String
        Dim Seaman As Boolean
        Dim SeqNo As Integer
        Dim GDSElement As String
        Dim GDSText As String
    End Structure
    Private mudtProps As ClassProps

    Public ReadOnly Property ID() As Integer
        Get
            ID = mudtProps.ID
        End Get
    End Property
    Public ReadOnly Property AirlineCode() As String
        Get
            AirlineCode = mudtProps.AirlineCode
        End Get
    End Property
    Public ReadOnly Property FlightType() As String
        Get
            FlightType = mudtProps.FlightType
        End Get
    End Property
    Public ReadOnly Property Seaman() As Boolean
        Get
            Seaman = mudtProps.Seaman
        End Get
    End Property
    Public ReadOnly Property SeqNo() As Integer
        Get
            SeqNo = mudtProps.SeqNo
        End Get
    End Property
    Public ReadOnly Property GDSElement() As String
        Get
            GDSElement = mudtProps.GDSElement
        End Get
    End Property
    Public ReadOnly Property GDSText() As String
        Get
            GDSText = mudtProps.GDSText
        End Get
    End Property

    Friend Sub SetValues(ByVal pID As Integer, ByVal pAirlineCode As String, ByVal pFlightType As String, ByVal pSeaman As Boolean,
                         ByVal pSeqNo As Integer, ByVal pGDSElement As String, ByVal pGDSText As String)
        With mudtProps
            .ID = pID
            .AirlineCode = pAirlineCode
            .FlightType = pFlightType
            .Seaman = pSeaman
            .SeqNo = pSeqNo
            .GDSElement = pGDSElement
            .GDSText = pGDSText
        End With
    End Sub
End Class
