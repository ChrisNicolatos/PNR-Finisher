Public Class GDSItineraryRemarksItem
    Private Structure ClassProps
        Dim ElementNo As Integer
        Dim FreeFlow As String
        Dim PassengerType As String
        Dim Text As String
        Dim Association As GDSAssociationsCollection
    End Structure
    Private mudtProps As ClassProps
    Public Sub New(ByVal pItem As s1aPNR.ItineraryRemarkElement)
        With mudtProps
            .ElementNo = pItem.ElementNo
            .FreeFlow = pItem.FreeFlow
            .PassengerType = pItem.PassengerType
            .Text = pItem.Text
            .Association = New GDSAssociationsCollection
            .Association.LoadAssociations(pItem)
        End With
    End Sub
    Public Sub New(ByVal pRI As String)
        With mudtProps
            If pRI.IndexOf(".") > 0 Then
                If pRI.IndexOf(".") < pRI.Length - 1 AndAlso IsNumeric(pRI.Substring(0, pRI.IndexOf("."))) Then
                    .ElementNo = pRI.Substring(0, pRI.IndexOf("."))
                    .FreeFlow = pRI.Substring(pRI.IndexOf(".") + 1)
                    .Text = pRI
                    .PassengerType = ""
                    .Association = New GDSAssociationsCollection
                End If
            End If
        End With
    End Sub
    Public ReadOnly Property ElementNo As Integer
        Get
            Return mudtProps.ElementNo
        End Get
    End Property
    Public ReadOnly Property FreeFlow As String
        Get
            Return mudtProps.FreeFlow
        End Get
    End Property
    Public ReadOnly Property PassengerType As String
        Get
            Return mudtProps.PassengerType
        End Get
    End Property
    Public ReadOnly Property Text As String
        Get
            Return mudtProps.Text
        End Get
    End Property
    Public ReadOnly Property Association As GDSAssociationsCollection
        Get
            Return mudtProps.Association
        End Get
    End Property
    Public Sub SetValues(ByVal pElementNo As Integer, ByVal pFreeFlow As String, ByVal pPassengerType As String, ByVal pText As String, ByVal pAssociation As GDSAssociationsCollection)

        With mudtProps
            .ElementNo = pElementNo
            .FreeFlow = pFreeFlow
            .PassengerType = pPassengerType
            .Text = pText
            .Association = pAssociation
        End With
    End Sub

End Class
