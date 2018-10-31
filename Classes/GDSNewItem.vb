Public Class GDSNewItem
    Private Structure NewItemClass
        Dim GDSCommand As String
        Dim TextRequested As String
        Friend Sub Clear()
            GDSCommand = ""
            TextRequested = ""
        End Sub

    End Structure
    Private mudtProps As NewItemClass

    Public ReadOnly Property GDSCommand As String
        Get
            GDSCommand = mudtProps.GDSCommand
        End Get
    End Property
    Public ReadOnly Property TextRequested As String
        Get
            TextRequested = mudtProps.TextRequested
        End Get
    End Property

    Friend Sub SetText(ByVal Value As String)
        mudtProps.TextRequested = Value
    End Sub

    Friend Sub SetText(ByVal Value As String, ByVal pGDSCommand As String)
        mudtProps.TextRequested = Value
        mudtProps.GDSCommand = pGDSCommand
    End Sub

    Friend Sub Clear()
        mudtProps.Clear()
    End Sub

End Class
