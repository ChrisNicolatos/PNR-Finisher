Public Class EmailCollection
    Inherits Collections.Generic.Dictionary(Of Short, EmailItem)
    Private mFromAddress As String
    Public Sub SetFromAddress(ByVal pFromAddress As String)
        mFromAddress = pFromAddress
    End Sub
    Public ReadOnly Property FromAddress As String
        Get
            FromAddress = mFromAddress
        End Get
    End Property
    Public Sub AddItem(ByVal pElementNo As Short, ByVal pEmailAddress As String, ByVal pEmailComment As String)
        Dim pobjClass As New EmailItem
        pobjClass.SetValues(pElementNo, pEmailAddress, pEmailComment)
        MyBase.Add(pobjClass.ElementNo, pobjClass)
    End Sub
End Class