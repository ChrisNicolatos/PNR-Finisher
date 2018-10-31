Public Class SSRCollection
    Inherits Collections.Generic.Dictionary(Of Short, SSRitem)
    Public Sub AddItem(ByVal pElementNo As Short, ByVal pSSRType As String, ByVal pSSRCode As String, ByVal pCarrierCode As String _
                         , ByVal pStatusCode As String, ByVal pText As String, ByVal pLastName As String, ByVal pFirstname As String _
                         , ByVal pDateOfBirth As Date, ByVal pPassportNumber As String)
        Dim pobjClass As New SSRitem
        pobjClass.SetValues(pElementNo, pSSRType, pSSRCode, pCarrierCode _
                         , pStatusCode, pText, pLastName, pFirstname _
                         , pDateOfBirth, pPassportNumber)
        MyBase.Add(pobjClass.ElementNo, pobjClass)
    End Sub
End Class
