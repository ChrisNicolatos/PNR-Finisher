Public Class GDSAssociationsCollection
    Dim mPaxCollection As New Collections.Generic.List(Of Integer)
    Dim mSegCollection As New Collections.Generic.List(Of Integer)
    Public Sub LoadAssociations(ByVal PNRElement As Object)
        mPaxCollection.Clear()
        mSegCollection.Clear()
        Try
            'If PNRElement.Count > 0 Then
            '    For Each pobjSSR In PNRElement
            If PNRElement.Associations.Passengers.Count > 0 Then
                For Each objPax In PNRElement.Associations.Passengers
                    mPaxCollection.Add(objPax.ElementNo)
                Next
            End If
            If PNRElement.Associations.Segments.Count > 0 Then
                For Each objPax In PNRElement.Associations.Segments
                    mSegCollection.Add(objPax.ElementNo)
                Next
            End If

            '    Next
            'End If
        Catch ex As Exception
        End Try
    End Sub
    Public ReadOnly Property PaxAssociations As Collections.Generic.List(Of Integer)
        Get
            Return mPaxCollection
        End Get
    End Property
    Public ReadOnly Property SegAssociation As Collections.Generic.List(Of Integer)
        Get
            Return mSegCollection
        End Get
    End Property
End Class
