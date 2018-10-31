Public Class BackOfficeItem
    Private Structure ClassProps
        Dim Id As Integer
        Dim BackOfficeName As String
    End Structure
    Private mudtProps As ClassProps

    Public ReadOnly Property Id As Integer
        Get
            Id = mudtProps.Id
        End Get
    End Property
    Public ReadOnly Property BackOfficeName As String
        Get
            BackOfficeName = mudtProps.BackOfficeName
        End Get
    End Property
    Friend Sub SetValues(ByVal pId As Integer, ByVal pBackOfficeName As String)
        With mudtProps
            .Id = pId
            .BackOfficeName = pBackOfficeName
        End With
    End Sub
End Class
