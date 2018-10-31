Public Class FrequentFlyerCollection
    Inherits Collections.Generic.Dictionary(Of String, FrequentFlyerItem)
    Friend Sub AddItem(ByVal pPaxName As String, ByVal pAirline As String, ByVal pFrequentTravelerNo As String, ByVal pCrossAccrual As String)
        If Not MyBase.ContainsKey(pPaxName) Then
            Dim pobjClass As FrequentFlyerItem

            pobjClass = New FrequentFlyerItem

            pobjClass.SetValues(pPaxName, pAirline, pFrequentTravelerNo, pCrossAccrual)
            MyBase.Add(pPaxName, pobjClass)
        End If

    End Sub
End Class
