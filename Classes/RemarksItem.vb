Public Class RemarksItem
    Private Structure ClassProps
        Dim Id As Integer
        Dim Title As String
        Dim Remark As String
        Dim InUse As Boolean
    End Structure
    Private mudtProps As ClassProps
    Public Overrides Function ToString() As String
        Return Title
    End Function
    Public ReadOnly Property Id As Integer
        Get
            Return mudtProps.Id
        End Get
    End Property
    Public ReadOnly Property Title As String
        Get
            Return mudtProps.Title
        End Get
    End Property
    Public ReadOnly Property Remark As String
        Get
            Return mudtProps.Remark
        End Get
    End Property
    Public ReadOnly Property InUse As Boolean
        Get
            Return mudtProps.InUse
        End Get
    End Property
    Public Sub SetValues(ByVal pId As Integer, ByVal pTitle As String, ByVal pRemark As String, ByVal pInUse As Boolean)
        With mudtProps
            .Id = pId
            .Title = pTitle
            .Remark = pRemark
            .InUse = pInUse
        End With
    End Sub
End Class
