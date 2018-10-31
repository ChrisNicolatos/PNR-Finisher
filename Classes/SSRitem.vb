Public Class SSRitem
    Private Structure ClassProps
        Dim ElementNo As Short
        Dim SSRType As String
        Dim SSRCode As String
        Dim CarrierCode As String
        Dim StatusCode As String
        Dim Text As String
        Dim LastName As String
        Dim FirstName As String
        Dim DateOfBirth As Date
        Dim PassportNumber As String
    End Structure
    Private mudtProps As ClassProps
    Public ReadOnly Property ElementNo As Short
        Get
            Return mudtProps.ElementNo
        End Get
    End Property
    Public ReadOnly Property SSRType As String
        Get
            Return mudtProps.SSRType
        End Get
    End Property
    Public ReadOnly Property SSRCode As String
        Get
            Return mudtProps.SSRCode
        End Get
    End Property
    Public ReadOnly Property CarrierCode As String
        Get
            Return mudtProps.CarrierCode
        End Get
    End Property
    Public ReadOnly Property StatusCode As String
        Get
            Return mudtProps.StatusCode
        End Get
    End Property
    Public ReadOnly Property Text As String
        Get
            Return mudtProps.Text
        End Get
    End Property
    Public ReadOnly Property LastName As String
        Get
            Return mudtProps.LastName
        End Get
    End Property
    Public ReadOnly Property FirstName As String
        Get
            Return mudtProps.FirstName
        End Get
    End Property
    Public ReadOnly Property DateOfBirth As Date
        Get
            Return mudtProps.DateOfBirth
        End Get
    End Property
    Public ReadOnly Property PassportNumber As String
        Get
            Return mudtProps.PassportNumber
        End Get
    End Property

    Friend Sub SetValues(ByVal pElementNo As Short, ByVal pSSRType As String, ByVal pSSRCode As String, ByVal pCarrierCode As String _
                         , ByVal pStatusCode As String, ByVal pText As String, ByVal pLastName As String, ByVal pFirstname As String _
                         , ByVal pDateOfBirth As Date, ByVal pPassportNumber As String)
        With mudtProps
            .ElementNo = pElementNo
            .SSRType = pSSRType
            .SSRCode = pSSRCode
            .CarrierCode = pCarrierCode
            .StatusCode = pStatusCode
            .Text = pText
            .LastName = pLastName
            .FirstName = pFirstname
            .DateOfBirth = pDateOfBirth
            .PassportNumber = pPassportNumber
        End With
    End Sub
End Class
