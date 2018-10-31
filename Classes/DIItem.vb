Public Class DIItem
    Private Structure ClassProps
        Dim ElementNo As Short
        Dim Category As String
        Dim CategoryDescription As String
        Dim Remark As String
    End Structure
    Private mudtProps As ClassProps
    Public ReadOnly Property ElementNo As Short
        Get
            Return mudtProps.ElementNo
        End Get
    End Property
    Public ReadOnly Property Category As String
        Get
            Return mudtProps.Category
        End Get
    End Property
    Public ReadOnly Property CategoryDescription As String
        Get
            Return mudtProps.CategoryDescription
        End Get
    End Property

    Public ReadOnly Property Remark As String
        Get
            Return mudtProps.Remark
        End Get
    End Property
    Friend Sub SetValues(ByVal pElementNo As Short, ByVal pCategory As String, ByVal pRemark As String)
        With mudtProps
            .ElementNo = pElementNo
            .CategoryDescription = pCategory
            Select Case pCategory
                Case "FREE TEXT"
                    .Category = "FT"
                Case Else
                    .Category = pCategory
            End Select
            .Remark = pRemark
        End With
    End Sub
End Class
