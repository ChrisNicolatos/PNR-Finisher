Option Strict Off
Option Explicit On
Public Class FrequentFlyerColl
    Inherits Collections.Generic.Dictionary(Of String, FrequentFlyer)

    Friend Sub AddItem(ByVal pPaxName As String, ByVal pAirline As String, ByVal pFrequentTravelerNo As String)
        ' Friend Sub AddItem(ByVal pElementNo As Integer, ByVal pPaxName As String, ByVal pAirline As String, ByVal pFrequentTravelerNo As String)
        If Not MyBase.ContainsKey(pPaxName) Then
            Dim pobjClass As FrequentFlyer

            pobjClass = New FrequentFlyer

            pobjClass.SetValues(pPaxName, pAirline, pFrequentTravelerNo)
            MyBase.Add(pPaxName, pobjClass)
        End If

    End Sub
End Class
