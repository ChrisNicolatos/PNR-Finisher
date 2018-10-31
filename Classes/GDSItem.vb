Public Class GDSItem
    Private Structure ClassProps
        Dim Id As Integer
        Dim GDSName As String
        Dim GDSCode As String
    End Structure
    Private mudtProps As ClassProps
    Public ReadOnly Property Id As Integer
        Get
            Id = mudtProps.Id
        End Get
    End Property
    Public ReadOnly Property GDSName As String
        Get
            GDSName = mudtProps.GDSName
        End Get
    End Property
    Public ReadOnly Property GDSCode As String
        Get
            GDSCode = mudtProps.GDSCode
        End Get
    End Property
    Friend Sub SetValues(ByVal pId As Integer, ByVal pGDSName As String, ByVal pGDSCode As String)
        With mudtProps
            .Id = pId
            .GDSName = pGDSName
            .GDSCode = pGDSCode
        End With
    End Sub
End Class
