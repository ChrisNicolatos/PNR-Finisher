Option Strict On
Option Explicit On
Namespace PaxApisDB
    'Friend Class ApisPaxItem
    '    Event Valid(IsValid As Boolean)
    '    Private Structure ClassProps
    '        Friend Id As Integer
    '        Friend Surname As String
    '        Friend FirstName As String
    '        Friend Birthdate As Date
    '        Friend Gender As String
    '        Friend IssuingCountry As String
    '        Friend PassportNumber As String
    '        Friend ExpiryDate As Date
    '        Friend Nationality As String
    '        Friend IsValid As Boolean
    '    End Structure
    '    Private mudtProps As ClassProps
    '    Public Sub New()
    '        With mudtProps
    '            .Id = 0
    '            .Surname = ""
    '            .FirstName = ""
    '            .Birthdate = Date.MinValue
    '            .Gender = "M"
    '            .IssuingCountry = ""
    '            .PassportNumber = ""
    '            .ExpiryDate = Date.MinValue
    '            .Nationality = ""
    '        End With
    '        SetValid()
    '    End Sub
    '    Public Sub New(ByVal pSurname As String, ByVal pFirstName As String)
    '        With mudtProps
    '            .Id = 0
    '            .Surname = pSurname
    '            .FirstName = pFirstName
    '            .Birthdate = Date.MinValue
    '            .Gender = "M"
    '            .IssuingCountry = ""
    '            .PassportNumber = ""
    '            .ExpiryDate = Date.MinValue
    '            .Nationality = ""
    '        End With
    '        SetValid()
    '    End Sub
    '    Public Sub New(ByVal pId As Integer, ByVal pSurname As String, ByVal pFirstName As String, ByVal pBirthDate As Date,
    '                   ByVal pGender As String, ByVal pIssuingCountry As String, ByVal pPassportNumber As String,
    '                   ByVal pExpiryDate As Date, ByVal pNationality As String)
    '        With mudtProps
    '            .Id = pId
    '            .Surname = pSurname
    '            .FirstName = pFirstName
    '            .Birthdate = pBirthDate
    '            .Gender = pGender
    '            .IssuingCountry = pIssuingCountry
    '            .PassportNumber = pPassportNumber
    '            .ExpiryDate = pExpiryDate
    '            .Nationality = pNationality
    '        End With
    '        SetValid()
    '    End Sub
    '    Public Sub New(ByVal pId As Integer, ByVal pSSRDocs As String)
    '        Dim pItems() As String = pSSRDocs.Split("/"c)
    '        If pItems.GetUpperBound(0) >= 8 Then
    '            With mudtProps
    '                .Id = pId
    '                .Surname = pItems(7)
    '                .FirstName = pItems(8)

    '                If IsDate(pItems(4)) Then
    '                    .Birthdate = CDate(pItems(4))
    '                Else
    '                    .Birthdate = Date.MinValue
    '                End If
    '                .Gender = pItems(5)
    '                .IssuingCountry = pItems(1)
    '                .PassportNumber = pItems(2)
    '                If IsDate(pItems(6)) Then
    '                    .ExpiryDate = CDate(pItems(6))
    '                Else
    '                    .ExpiryDate = Date.MinValue
    '                End If
    '                .Nationality = pItems(3)
    '            End With
    '        Else
    '            With mudtProps
    '                .Id = 0
    '                .Surname = ""
    '                .FirstName = ""
    '                .Birthdate = Date.MinValue
    '                .Gender = "M"
    '                .IssuingCountry = ""
    '                .PassportNumber = ""
    '                .ExpiryDate = Date.MinValue
    '                .Nationality = ""
    '            End With
    '        End If

    '        SetValid()
    '    End Sub
    '    Private Sub SetValid()

    '        mudtProps.IsValid = (Surname <> "" And FirstName <> "" And Gender <> "" And BirthDate > Date.MinValue)
    '        RaiseEvent Valid(mudtProps.IsValid)

    '    End Sub
    '    Public ReadOnly Property IsValid As Boolean
    '        Get
    '            IsValid = mudtProps.IsValid
    '        End Get
    '    End Property
    '    Public ReadOnly Property Id() As Integer
    '        Get
    '            Id = mudtProps.Id
    '        End Get
    '    End Property
    '    Public Property Surname() As String
    '        Get
    '            Surname = mudtProps.Surname.Trim.ToUpper
    '        End Get
    '        Set(ByVal value As String)
    '            mudtProps.Surname = value.Trim.ToUpper
    '            SetValid()
    '        End Set
    '    End Property
    '    Public Property FirstName() As String
    '        Get
    '            FirstName = mudtProps.FirstName.Trim.ToUpper
    '        End Get
    '        Set(ByVal value As String)
    '            mudtProps.FirstName = value.Trim.ToUpper
    '            SetValid()
    '        End Set
    '    End Property
    '    Public Property BirthDate() As Date
    '        Get
    '            BirthDate = mudtProps.Birthdate
    '        End Get
    '        Set(ByVal value As Date)
    '            mudtProps.Birthdate = value
    '            SetValid()
    '        End Set
    '    End Property
    '    Public Property Gender() As String
    '        Get
    '            Gender = mudtProps.Gender.Trim.ToUpper
    '        End Get
    '        Set(ByVal value As String)
    '            mudtProps.Gender = value.Trim.ToUpper
    '            SetValid()
    '        End Set
    '    End Property
    '    Public Property IssuingCountry() As String
    '        Get
    '            IssuingCountry = mudtProps.IssuingCountry.Trim.ToUpper
    '        End Get
    '        Set(ByVal value As String)
    '            mudtProps.IssuingCountry = value.Trim.ToUpper
    '            SetValid()
    '        End Set
    '    End Property
    '    Public Property PassportNumber() As String
    '        Get
    '            PassportNumber = mudtProps.PassportNumber.Trim.ToUpper
    '        End Get
    '        Set(ByVal value As String)
    '            mudtProps.PassportNumber = value.Trim.ToUpper
    '            SetValid()
    '        End Set
    '    End Property
    '    Public Property ExpiryDate() As Date
    '        Get
    '            ExpiryDate = mudtProps.ExpiryDate
    '        End Get
    '        Set(ByVal value As Date)
    '            mudtProps.ExpiryDate = value
    '            SetValid()
    '        End Set
    '    End Property
    '    Public Property Nationality() As String
    '        Get
    '            Nationality = mudtProps.Nationality.Trim.ToUpper
    '        End Get
    '        Set(ByVal value As String)
    '            mudtProps.Nationality = value.Trim.ToUpper
    '            SetValid()
    '        End Set
    '    End Property

    '    Public Sub Update(ByVal ExpiryDateOK As Boolean)

    '        SetValid()

    '        If IsValid Then
    '            Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR)
    '            Dim pobjComm As New SqlClient.SqlCommand

    '            pobjConn.Open()
    '            pobjComm = pobjConn.CreateCommand

    '            With pobjComm
    '                .CommandType = CommandType.StoredProcedure
    '                .CommandText = "PaxApisInformationUpdate"
    '                .Parameters.Add("@ppId", SqlDbType.Int).Value = mudtProps.Id
    '                .Parameters.Add("@ppSurname", SqlDbType.NVarChar, 30).Value = mudtProps.Surname
    '                .Parameters.Add("@ppFirstName", SqlDbType.NVarChar, 30).Value = mudtProps.FirstName
    '                .Parameters.Add("@ppBirthDate", SqlDbType.DateTime).Value = mudtProps.Birthdate
    '                .Parameters.Add("@ppGender", SqlDbType.NVarChar, 10).Value = mudtProps.Gender
    '                .Parameters.Add("@ppDocIssuingCountry", SqlDbType.NVarChar, 3).Value = mudtProps.IssuingCountry
    '                .Parameters.Add("@ppDocnumber", SqlDbType.NVarChar, 15).Value = mudtProps.PassportNumber
    '                .Parameters.Add("@ppNationality", SqlDbType.NVarChar, 3).Value = mudtProps.Nationality
    '                If ExpiryDateOK And mudtProps.ExpiryDate > Now Then
    '                    .Parameters.Add("@ppDocExpiryDate", SqlDbType.DateTime).Value = mudtProps.ExpiryDate
    '                Else
    '                    .Parameters.Add("@ppDocExpiryDate", SqlDbType.DateTime).Value = DateSerial(1902, 12, 31)
    '                End If
    '                '.Parameters.Add("@ppQRFreqFlyer", SqlDbType.NChar, 30).Value = False
    '                .ExecuteNonQuery()
    '            End With
    '        Else
    '            'Throw New Exception("PaxApisDB().Update" & vbCrLf & "Pax not updated")
    '        End If

    '    End Sub
    '    Public Sub Insert()

    '        If IsValid Then
    '            Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR)
    '            Dim pobjComm As New SqlClient.SqlCommand

    '            pobjConn.Open()
    '            pobjComm = pobjConn.CreateCommand

    '            With pobjComm
    '                .CommandType = CommandType.StoredProcedure
    '                .CommandText = "PaxApisInformation_Insert"
    '                .Parameters.Add("@ppSurname", SqlDbType.NVarChar, 30).Value = mudtProps.Surname
    '                .Parameters.Add("@ppFirstName", SqlDbType.NVarChar, 30).Value = mudtProps.FirstName
    '                .Parameters.Add("@ppBirthDate", SqlDbType.DateTime).Value = mudtProps.Birthdate
    '                .Parameters.Add("@ppGender", SqlDbType.NVarChar, 10).Value = mudtProps.Gender
    '                .Parameters.Add("@ppDocIssuingCountry", SqlDbType.NVarChar, 3).Value = mudtProps.IssuingCountry
    '                .Parameters.Add("@ppDocnumber", SqlDbType.NVarChar, 15).Value = mudtProps.PassportNumber
    '                .Parameters.Add("@ppNationality", SqlDbType.NVarChar, 3).Value = mudtProps.Nationality
    '                If mudtProps.ExpiryDate > Now Then
    '                    .Parameters.Add("@ppDocExpiryDate", SqlDbType.DateTime).Value = mudtProps.ExpiryDate
    '                Else
    '                    .Parameters.Add("@ppDocExpiryDate", SqlDbType.DateTime).Value = DateSerial(1902, 12, 31)
    '                End If
    '                '.Parameters.Add("@ppQRFreqFlyer", SqlDbType.NChar, 30).Value = False
    '                .ExecuteNonQuery()
    '            End With
    '        Else
    '            Throw New Exception("PaxApidDB().Update" & vbCrLf & "Pax not updated")
    '        End If

    '    End Sub
    'End Class

    'Friend Class ApisPaxCollection
    '    Inherits Collections.Generic.Dictionary(Of Integer, ApisPaxItem)

    '    Public Sub Read(ByVal Surname As String, ByVal FirstName As String)
    '        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR)
    '        Dim pobjComm As New SqlClient.SqlCommand
    '        Dim pobjReader As SqlClient.SqlDataReader
    '        Dim pobjItem As ApisPaxItem

    '        pobjConn.Open()
    '        pobjComm = pobjConn.CreateCommand

    '        With pobjComm
    '            .CommandType = CommandType.StoredProcedure
    '            .CommandText = "PaxApisInformationAllSelect"
    '            .Parameters.Add("@ppSurname", SqlDbType.NVarChar, 30).Value = Surname
    '            .Parameters.Add("@ppFirstName", SqlDbType.NVarChar, 30).Value = FirstName
    '            pobjReader = .ExecuteReader
    '        End With

    '        Clear()

    '        With pobjReader
    '            Do While .Read
    '                pobjItem = New ApisPaxItem(CInt(.Item("ppId")), Surname, FirstName, CDate(If(IsDBNull(.Item("ppBirthdate")), Date.MinValue, .Item("ppBirthdate"))), CStr(.Item("ppGender")),
    '                                    CStr(.Item("ppDocIssuingCountry")), CStr(.Item("ppDocnumber")), CDate(If(IsDBNull(.Item("ppDocExpiryDate")), Date.MinValue, .Item("ppDocExpiryDate"))),
    '                                    CStr(.Item("ppNationality")))
    '                MyBase.Add(pobjItem.Id, pobjItem)
    '                'Exit Do
    '            Loop
    '            .Close()
    '        End With
    '        pobjConn.Close()
    '    End Sub
    '    Public Sub AddSSRDocsItem(ByVal Id As Integer, ByVal pSSRDocs As String)
    '        Dim pItem As New ApisPaxItem(Id, pSSRDocs)
    '        If pItem.Id > 0 AndAlso Not MyBase.ContainsKey(pItem.Id) Then
    '            MyBase.Add(pItem.Id, pItem)

    '        End If
    '    End Sub
    'End Class
    'Friend Class ReferenceItem
    '    Private Structure ClassProps
    '        Dim Code As String
    '        Dim Description As String
    '    End Structure
    '    Private mudtProps As ClassProps
    '    Public ReadOnly Property Code As String
    '        Get
    '            Code = mudtProps.Code
    '        End Get
    '    End Property
    '    Public ReadOnly Property Description As String
    '        Get
    '            Description = mudtProps.Description
    '        End Get
    '    End Property
    '    Public Overrides Function ToString() As String
    '        Return mudtProps.Code & If(Description = "", "", " - " & mudtProps.Description)
    '    End Function
    '    Friend Sub SetValues(ByVal pCode As String, ByVal pDescription As String)
    '        With mudtProps
    '            .Code = pCode
    '            .Description = pDescription
    '        End With
    '    End Sub
    'End Class
    'Friend Class ReferenceGenderCollection
    '    Inherits Collections.Generic.Dictionary(Of Integer, ReferenceItem)
    '    Private mstrGenderIndicator(,) As String = {{"M", "MALE"}, {"F", "FEMALE"}, {"MI", "MALE INFANT"}, {"FI", "FEMALE INFANT"}, {"U", "UNDISCLOSED GENDER"}}
    '    Public Sub New()
    '        For i As Integer = 0 To mstrGenderIndicator.GetUpperBound(0)
    '            Dim pItem As New ReferenceItem
    '            pItem.SetValues(mstrGenderIndicator(i, 0), mstrGenderIndicator(i, 1))
    '            MyBase.Add(i, pItem)
    '        Next
    '    End Sub
    'End Class
    'Friend Class ReferenceSalutationsCollection
    '    Inherits Collections.Generic.Dictionary(Of Integer, ReferenceItem)
    '    Private mstrSalutations() As String = {"MR", "MRS", "MS", "MISS", "MISTER"}
    '    Public Sub New()
    '        For i As Integer = 0 To mstrSalutations.GetUpperBound(0)
    '            Dim pItem As New ReferenceItem
    '            pItem.SetValues(mstrSalutations(i), "")
    '            MyBase.Add(i, pItem)
    '        Next
    '    End Sub
    'End Class
    'Friend Class ReferenceCountriesCollection
    '    Inherits Collections.Generic.Dictionary(Of String, ReferenceItem)
    '    Public Sub New()

    '        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
    '        Dim pobjComm As New SqlClient.SqlCommand
    '        Dim pobjReader As SqlClient.SqlDataReader
    '        Dim pobjClass As ReferenceItem

    '        pobjConn.Open()
    '        pobjComm = pobjConn.CreateCommand

    '        With pobjComm
    '            .CommandType = CommandType.Text
    '            .CommandText = "SELECT countryISO3Code " &
    '                            "      ,countryName " &
    '                            " FROM AmadeusReports.dbo.zzCountries " &
    '                            " WHERE LEN(CountryCode) = 2 AND countryISO3Code IS NOT NULL " &
    '                            " ORDER BY countryname "
    '            pobjReader = .ExecuteReader
    '        End With

    '        With pobjReader
    '            Do While .Read
    '                pobjClass = New ReferenceItem
    '                pobjClass.SetValues(CStr(.Item("countryISO3Code")), CStr(.Item("countryName")))
    '                MyBase.Add(pobjClass.Code, pobjClass)
    '            Loop
    '            .Close()
    '        End With
    '        pobjConn.Close()
    '    End Sub
    'End Class
End Namespace

