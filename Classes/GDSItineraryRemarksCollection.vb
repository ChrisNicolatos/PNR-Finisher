Public Class GDSItineraryRemarksCollection
    Inherits Collections.Generic.List(Of GDSItineraryRemarksItem)
    Public Sub Load1A(ByRef pItinRemarks As s1aPNR.PNR)
        MyBase.Clear()

        For Each pItinRem As s1aPNR.ItineraryRemarkElement In pItinRemarks.ItineraryRemarkElements
            Dim pItem As New GDSItineraryRemarksItem(pItinRem)
            MyBase.Add(pItem)
        Next

    End Sub
    Public Sub Load1G(ByRef pRI() As String)
        MyBase.Clear()
        If IsArray(pRI) Then
            Dim pFound As Boolean = False
            For i As Integer = 0 To pRI.GetUpperBound(0)
                If Not pFound Then
                    If pRI(i).StartsWith("UNASSOCIATED ITINERARY REMARKS") Then
                        pFound = True
                    End If
                Else
                    Dim pItem As New GDSItineraryRemarksItem(pRI(i))
                    MyBase.Add(pItem)
                End If
            Next
        End If
    End Sub
End Class
