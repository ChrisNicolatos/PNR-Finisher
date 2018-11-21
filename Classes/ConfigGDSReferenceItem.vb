Public Class ConfigGDSReferenceItem
    Private Structure ClassProps
        Dim Id As Integer
        Dim Key As String
        Dim Value As String
        Dim GDSKey As Integer
        Dim BOKey As Integer
        Dim Element As String
        Dim RefId As String
        Dim RefDetail As String
    End Structure
    Private mudtProps As ClassProps
    Public ReadOnly Property Key As String
        Get
            Return mudtProps.Key
        End Get
    End Property
    Public ReadOnly Property Value As String
        Get
            Return mudtProps.Value
        End Get
    End Property
    Public ReadOnly Property Element As String
        Get
            Return mudtProps.Element
        End Get
    End Property
    Public Sub SetValues(pId As Integer, pKey As String, pValue As String, pGDSKey As Integer, pBOKey As Integer, pElement As String, pRefId As String, pRefDetail As String)
        With mudtProps
            .Id = pId
            .Key = pKey
            .Value = pValue
            .GDSKey = pGDSKey
            .BOKey = pBOKey
            .Element = pElement
            .RefId = pRefId
            .RefDetail = pRefDetail
        End With
    End Sub
End Class