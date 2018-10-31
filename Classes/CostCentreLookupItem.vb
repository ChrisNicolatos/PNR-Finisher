Public Class CostCentreLookupItem
    Private Structure ClassProps
        Dim Id As Integer
        Dim Code As String
        Dim OldCode As String
        Dim ClientName As String
        Dim ClientLogo As String
        Dim VesselName As String
        Dim CostCentre As String
    End Structure
    Dim mudtProps As ClassProps

    Friend Sub New(ByVal Id As Integer, ByVal Code As String, ByVal OldCode As String, ByVal ClientName As String, ByVal ClientLogo As String, ByVal VesselName As String, ByVal CostCentre As String)

        With mudtProps
            .Id = Id
            .Code = Code
            .OldCode = OldCode
            .ClientName = ClientName
            .ClientLogo = ClientLogo
            .VesselName = VesselName
            .CostCentre = CostCentre
        End With
    End Sub
    Public ReadOnly Property Id As Integer
        Get
            Return mudtProps.Id
        End Get
    End Property
    Public ReadOnly Property Code As String
        Get
            Return mudtProps.Code
        End Get
    End Property
    Public ReadOnly Property OldCode As String
        Get
            Return mudtProps.OldCode
        End Get
    End Property
    Public ReadOnly Property ClientName As String
        Get
            Return mudtProps.ClientName
        End Get
    End Property
    Public ReadOnly Property ClientLogo As String
        Get
            Return mudtProps.ClientLogo
        End Get
    End Property
    Public ReadOnly Property VesselName As String
        Get
            Return mudtProps.VesselName
        End Get
    End Property
    Public ReadOnly Property CostCentre As String
        Get
            Return mudtProps.CostCentre
        End Get
    End Property
End Class
