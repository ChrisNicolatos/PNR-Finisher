Option Strict On
Option Explicit On
Public Class ApisPaxItem
    Event Valid(IsValid As Boolean)
    Private Structure ClassProps
        Friend Id As Integer
        Friend Surname As String
        Friend FirstName As String
        Friend Birthdate As Date
        Friend Gender As String
        Friend IssuingCountry As String
        Friend PassportNumber As String
        Friend ExpiryDate As Date
        Friend Nationality As String
        Friend IsValid As Boolean
    End Structure
    Private mudtProps As ClassProps
    Public Sub New()
        With mudtProps
            .Id = 0
            .Surname = ""
            .FirstName = ""
            .Birthdate = Date.MinValue
            .Gender = "M"
            .IssuingCountry = ""
            .PassportNumber = ""
            .ExpiryDate = Date.MinValue
            .Nationality = ""
        End With
        SetValid()
    End Sub
    Public Sub New(ByVal pSurname As String, ByVal pFirstName As String)
        With mudtProps
            .Id = 0
            .Surname = pSurname
            .FirstName = pFirstName
            .Birthdate = Date.MinValue
            .Gender = "M"
            .IssuingCountry = ""
            .PassportNumber = ""
            .ExpiryDate = Date.MinValue
            .Nationality = ""
        End With
        SetValid()
    End Sub
    Public Sub New(ByVal pId As Integer, ByVal pSurname As String, ByVal pFirstName As String, ByVal pBirthDate As Date,
                       ByVal pGender As String, ByVal pIssuingCountry As String, ByVal pPassportNumber As String,
                       ByVal pExpiryDate As Date, ByVal pNationality As String)
        With mudtProps
            .Id = pId
            .Surname = pSurname
            .FirstName = pFirstName
            .Birthdate = pBirthDate
            .Gender = pGender
            .IssuingCountry = pIssuingCountry
            .PassportNumber = pPassportNumber
            .ExpiryDate = pExpiryDate
            .Nationality = pNationality
        End With
        SetValid()
    End Sub
    Public Sub New(ByVal pId As Integer, ByVal pSSRDocs As String)
        Dim pItems() As String = pSSRDocs.Split("/"c)
        If pItems.GetUpperBound(0) >= 8 Then
            With mudtProps
                .Id = pId
                .Surname = pItems(7)
                .FirstName = pItems(8)

                If IsDate(pItems(4)) Then
                    .Birthdate = CDate(pItems(4))
                Else
                    .Birthdate = Date.MinValue
                End If
                .Gender = pItems(5)
                .IssuingCountry = pItems(1)
                .PassportNumber = pItems(2)
                If IsDate(pItems(6)) Then
                    .ExpiryDate = CDate(pItems(6))
                Else
                    .ExpiryDate = Date.MinValue
                End If
                .Nationality = pItems(3)
            End With
        Else
            With mudtProps
                .Id = 0
                .Surname = ""
                .FirstName = ""
                .Birthdate = Date.MinValue
                .Gender = "M"
                .IssuingCountry = ""
                .PassportNumber = ""
                .ExpiryDate = Date.MinValue
                .Nationality = ""
            End With
        End If

        SetValid()
    End Sub
    Private Sub SetValid()

        mudtProps.IsValid = (Surname <> "" And FirstName <> "" And Gender <> "" And BirthDate > Date.MinValue)
        RaiseEvent Valid(mudtProps.IsValid)

    End Sub
    Public ReadOnly Property IsValid As Boolean
        Get
            IsValid = mudtProps.IsValid
        End Get
    End Property
    Public ReadOnly Property Id() As Integer
        Get
            Id = mudtProps.Id
        End Get
    End Property
    Public Property Surname() As String
        Get
            Surname = mudtProps.Surname.Trim.ToUpper
        End Get
        Set(ByVal value As String)
            mudtProps.Surname = value.Trim.ToUpper
            SetValid()
        End Set
    End Property
    Public Property FirstName() As String
        Get
            FirstName = mudtProps.FirstName.Trim.ToUpper
        End Get
        Set(ByVal value As String)
            mudtProps.FirstName = value.Trim.ToUpper
            SetValid()
        End Set
    End Property
    Public Property BirthDate() As Date
        Get
            BirthDate = mudtProps.Birthdate
        End Get
        Set(ByVal value As Date)
            mudtProps.Birthdate = value
            SetValid()
        End Set
    End Property
    Public Property Gender() As String
        Get
            Gender = mudtProps.Gender.Trim.ToUpper
        End Get
        Set(ByVal value As String)
            mudtProps.Gender = value.Trim.ToUpper
            SetValid()
        End Set
    End Property
    Public Property IssuingCountry() As String
        Get
            IssuingCountry = mudtProps.IssuingCountry.Trim.ToUpper
        End Get
        Set(ByVal value As String)
            mudtProps.IssuingCountry = value.Trim.ToUpper
            SetValid()
        End Set
    End Property
    Public Property PassportNumber() As String
        Get
            PassportNumber = mudtProps.PassportNumber.Trim.ToUpper
        End Get
        Set(ByVal value As String)
            mudtProps.PassportNumber = value.Trim.ToUpper
            SetValid()
        End Set
    End Property
    Public Property ExpiryDate() As Date
        Get
            ExpiryDate = mudtProps.ExpiryDate
        End Get
        Set(ByVal value As Date)
            mudtProps.ExpiryDate = value
            SetValid()
        End Set
    End Property
    Public Property Nationality() As String
        Get
            Nationality = mudtProps.Nationality.Trim.ToUpper
        End Get
        Set(ByVal value As String)
            mudtProps.Nationality = value.Trim.ToUpper
            SetValid()
        End Set
    End Property

    Public Sub Update(ByVal ExpiryDateOK As Boolean)

        SetValid()

        If IsValid Then
            Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR)
            Dim pobjComm As New SqlClient.SqlCommand

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.StoredProcedure
                .CommandText = "PaxApisInformationUpdate"
                .Parameters.Add("@ppId", SqlDbType.Int).Value = mudtProps.Id
                .Parameters.Add("@ppSurname", SqlDbType.NVarChar, 30).Value = mudtProps.Surname
                .Parameters.Add("@ppFirstName", SqlDbType.NVarChar, 30).Value = mudtProps.FirstName
                .Parameters.Add("@ppBirthDate", SqlDbType.DateTime).Value = mudtProps.Birthdate
                .Parameters.Add("@ppGender", SqlDbType.NVarChar, 10).Value = mudtProps.Gender
                .Parameters.Add("@ppDocIssuingCountry", SqlDbType.NVarChar, 3).Value = mudtProps.IssuingCountry
                .Parameters.Add("@ppDocnumber", SqlDbType.NVarChar, 15).Value = mudtProps.PassportNumber
                .Parameters.Add("@ppNationality", SqlDbType.NVarChar, 3).Value = mudtProps.Nationality
                If ExpiryDateOK And mudtProps.ExpiryDate > Now Then
                    .Parameters.Add("@ppDocExpiryDate", SqlDbType.DateTime).Value = mudtProps.ExpiryDate
                Else
                    .Parameters.Add("@ppDocExpiryDate", SqlDbType.DateTime).Value = DateSerial(1902, 12, 31)
                End If
                '.Parameters.Add("@ppQRFreqFlyer", SqlDbType.NChar, 30).Value = False
                .ExecuteNonQuery()
            End With
        Else
            'Throw New Exception("PaxApisDB().Update" & vbCrLf & "Pax not updated")
        End If

    End Sub
    Public Sub Insert()

        If IsValid Then
            Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR)
            Dim pobjComm As New SqlClient.SqlCommand

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.StoredProcedure
                .CommandText = "PaxApisInformation_Insert"
                .Parameters.Add("@ppSurname", SqlDbType.NVarChar, 30).Value = mudtProps.Surname
                .Parameters.Add("@ppFirstName", SqlDbType.NVarChar, 30).Value = mudtProps.FirstName
                .Parameters.Add("@ppBirthDate", SqlDbType.DateTime).Value = mudtProps.Birthdate
                .Parameters.Add("@ppGender", SqlDbType.NVarChar, 10).Value = mudtProps.Gender
                .Parameters.Add("@ppDocIssuingCountry", SqlDbType.NVarChar, 3).Value = mudtProps.IssuingCountry
                .Parameters.Add("@ppDocnumber", SqlDbType.NVarChar, 15).Value = mudtProps.PassportNumber
                .Parameters.Add("@ppNationality", SqlDbType.NVarChar, 3).Value = mudtProps.Nationality
                If mudtProps.ExpiryDate > Now Then
                    .Parameters.Add("@ppDocExpiryDate", SqlDbType.DateTime).Value = mudtProps.ExpiryDate
                Else
                    .Parameters.Add("@ppDocExpiryDate", SqlDbType.DateTime).Value = DateSerial(1902, 12, 31)
                End If
                '.Parameters.Add("@ppQRFreqFlyer", SqlDbType.NChar, 30).Value = False
                .ExecuteNonQuery()
            End With
        Else
            Throw New Exception("PaxApidDB().Update" & vbCrLf & "Pax not updated")
        End If

    End Sub

End Class
