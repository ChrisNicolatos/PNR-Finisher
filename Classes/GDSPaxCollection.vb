Public Class GDSPaxColl
    Inherits Collections.Generic.Dictionary(Of String, GDSPaxItem)

    Friend Sub AddItem(ByVal pElementNo As Short, ByVal pInitial As String, ByVal pLastName As String, ByVal pID As String)

        Dim pobjClass As GDSPaxItem

        pobjClass = New GDSPaxItem

        pobjClass.SetValues(pElementNo, pInitial, pLastName, pID)
        MyBase.Add(Format(pElementNo), pobjClass)
    End Sub
    Public ReadOnly Property LeadName As String
        Get
            LeadName = MyBase.ElementAt(0).Value.LastName & "/" & MyBase.ElementAt(0).Value.Initial
            If MyBase.Count > 1 Then
                LeadName &= " x " & MyBase.Count
            End If
        End Get
    End Property
    Public ReadOnly Property AllPassengers As String
        Get
            AllPassengers = ""
            For Each pPax As GDSPaxItem In MyBase.Values
                AllPassengers &= pPax.ElementNo & "." & pPax.PaxName & vbCrLf
            Next
        End Get
    End Property
End Class