Public Class OptionQueueCollection
    Inherits Collections.Generic.Dictionary(Of Short, OptionQueueItem)
    Public Sub AddItem(ByVal pElementNo As Short, ByVal pPCC As String, ByVal pActionDateTime As Date, ByVal pQueueNumber As String, ByVal pRemark As String)
        Dim pobjClass As New OptionQueueItem
        pobjClass.SetValues(pElementNo, pPCC, pActionDateTime, pQueueNumber, pRemark)
        MyBase.Add(pobjClass.ElementNo, pobjClass)
    End Sub

End Class
