Public Class CustomerGroupCollection
    Inherits Collections.Generic.Dictionary(Of Integer, CustomerGroupItem)
    Private mAllCustomer As New CustomerGroupCollectionAll

    Public Sub Load(ByVal SearchString As String)

        Try
            If mAllCustomer.Count = 0 Then
                Cursor.Current = Cursors.WaitCursor
                mAllCustomer.Load()
            End If

            MyBase.Clear()

            Dim pItem As CustomerGroupItem

            For Each pItem In mAllCustomer.Values
                If pItem.Name.ToUpper.IndexOf(SearchString.ToUpper) >= 0 Then
                    MyBase.Add(pItem.ID, pItem)
                End If
            Next


        Catch ex As Exception
            Throw New Exception("CustomerGroupCollection.Load()" & vbCrLf & ex.Message)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

End Class