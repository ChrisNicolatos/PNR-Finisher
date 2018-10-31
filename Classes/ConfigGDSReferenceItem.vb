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
    Public ReadOnly Property Id As Integer
        Get
            Id = mudtProps.Id
        End Get
    End Property
    Public ReadOnly Property Key As String
        Get
            Key = mudtProps.Key
        End Get
    End Property
    Public ReadOnly Property Value As String
        Get
            Value = mudtProps.Value
        End Get
    End Property
    Public ReadOnly Property GDSKey As Integer
        Get
            GDSKey = mudtProps.GDSKey
        End Get
    End Property
    Public ReadOnly Property BOKey As Integer
        Get
            BOKey = mudtProps.BOKey
        End Get
    End Property
    Public ReadOnly Property Element As String
        Get
            Element = mudtProps.Element
        End Get
    End Property
    Public ReadOnly Property RefId As String
        Get
            RefId = mudtProps.RefId
        End Get
    End Property
    Public ReadOnly Property RefDetail As String
        Get
            RefDetail = mudtProps.RefDetail
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