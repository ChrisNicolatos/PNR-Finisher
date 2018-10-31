Public Class OpenSegmentCollection
    Inherits Collections.Generic.Dictionary(Of Short, OpenSegmentItem)
    Public Sub AddItem(ByVal pElementNo As Short, ByVal pSegmentType As String, ByVal pRemarkType As String, ByVal pRemarkDate As Date, ByVal pRemark As String)
        Dim pobjClass As New OpenSegmentItem
        pobjClass.SetValues(pElementNo, pSegmentType, pRemarkType, pRemarkDate, pRemark)
        MyBase.Add(pobjClass.ElementNo, pobjClass)
    End Sub
End Class