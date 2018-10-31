Public Class ReferenceGenderCollection
    Inherits Collections.Generic.Dictionary(Of Integer, ReferenceItem)
    Private mstrGenderIndicator(,) As String = {{"M", "MALE"}, {"F", "FEMALE"}, {"MI", "MALE INFANT"}, {"FI", "FEMALE INFANT"}, {"U", "UNDISCLOSED GENDER"}}
    Public Sub New()
        For i As Integer = 0 To mstrGenderIndicator.GetUpperBound(0)
            Dim pItem As New ReferenceItem
            pItem.SetValues(mstrGenderIndicator(i, 0), mstrGenderIndicator(i, 1))
            MyBase.Add(i, pItem)
        Next
    End Sub
End Class
