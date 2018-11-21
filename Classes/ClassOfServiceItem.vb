Public Class ClassOfServiceItem
    Private Structure ClassProps
        Dim Carrier As String
        Dim Origin As String
        Dim Destination As String
        Dim ClassOfService As String
        Dim CabinDescription As String
        Dim ClassDescription As String
    End Structure
    Private mudtProps As ClassProps
    Public ReadOnly Property Key As String
        Get
            Return Carrier & "|" & Origin & "|" & Destination & "|" & ClassOfService
        End Get
    End Property
    Public ReadOnly Property Carrier As String
        Get
            Return mudtProps.Carrier
        End Get
    End Property
    Public ReadOnly Property Origin As String
        Get
            Return mudtProps.Origin
        End Get
    End Property
    Public ReadOnly Property Destination As String
        Get
            Return mudtProps.Destination
        End Get
    End Property
    Public ReadOnly Property ClassOfService As String
        Get
            Return mudtProps.ClassOfService
        End Get
    End Property
    Public ReadOnly Property CabinDescription As String
        Get
            Return mudtProps.CabinDescription
        End Get
    End Property
    Public ReadOnly Property ClassDescription As String
        Get
            Return mudtProps.ClassDescription
        End Get
    End Property
    Public Sub SetValues(ByVal pCarrier As String, ByVal pOrigin As String, ByVal pDestination As String, ByVal pClassOfService As String, ByVal pCabinDescription As String, ByVal pClassDescription As String)
        With mudtProps
            .Carrier = pCarrier
            .Origin = pOrigin
            .Destination = pDestination
            .ClassOfService = pClassOfService
            .CabinDescription = pCabinDescription
            .ClassDescription = pClassDescription
        End With
    End Sub
End Class
