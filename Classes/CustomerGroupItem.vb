Public Class CustomerGroupItem
    Private Structure ClassProps
        Dim ID As Integer
        Dim Name As String
    End Structure
    Private mudtProps As ClassProps
    Public Overrides Function ToString() As String
        With mudtProps
            Return .Name
        End With
    End Function
    Public ReadOnly Property ID() As Integer
        Get
            ID = mudtProps.ID
        End Get
    End Property
    Public ReadOnly Property Name() As String
        Get
            Name = mudtProps.Name
        End Get
    End Property
    Friend Sub SetValues(ByVal pID As Integer, ByVal pName As String)
        With mudtProps
            .ID = pID
            .Name = pName
        End With
    End Sub
End Class