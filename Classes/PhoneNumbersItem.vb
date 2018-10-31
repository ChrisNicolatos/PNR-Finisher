Public Class PhoneNumbersItem
    Private Structure ClassProps
        Dim ElementNo As Short
        Dim CityCode As String
        Dim PhoneType As String
        Dim PhoneNumber As String
    End Structure
    Dim mudtProps As ClassProps
    Public ReadOnly Property ElementNo As Short
        Get
            Return mudtProps.ElementNo
        End Get
    End Property
    Public ReadOnly Property CityCode As String
        Get
            Return mudtProps.CityCode
        End Get
    End Property
    Public ReadOnly Property PhoneType As String
        Get
            Return mudtProps.PhoneType
        End Get
    End Property
    Public ReadOnly Property PhoneNumber As String
        Get
            Return mudtProps.PhoneNumber
        End Get
    End Property
    Friend Sub SetValues(ByVal pElementNo As Short, ByVal pCityCode As String, ByVal pPhoneType As String, ByVal pPhoneNumber As String)
        With mudtProps
            .ElementNo = pElementNo
            .CityCode = pCityCode
            .PhoneType = pPhoneType
            .PhoneNumber = pPhoneNumber
        End With
    End Sub
End Class
