Public Class DICollection
    Inherits Collections.Generic.Dictionary(Of Short, DIItem)
    Public Sub AddItem(ByVal pElementNo As Short, ByVal pCategory As String, ByVal pRemark As String)
        Dim pobjClass As New DIItem
        pobjClass.SetValues(pElementNo, pCategory, pRemark)
        MyBase.Add(pobjClass.ElementNo, pobjClass)
    End Sub
End Class
