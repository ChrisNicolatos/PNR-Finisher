Public Class CloseOffEntriesItem

    Dim mEntry As String

    Public ReadOnly Property CloseOffEntry As String
        Get
            CloseOffEntry = MySettings.ConvertGDSValue(mEntry)
        End Get
    End Property
    Friend Sub SetValues(ByVal CloseOffEntry As String)
        mEntry = CloseOffEntry
    End Sub
End Class