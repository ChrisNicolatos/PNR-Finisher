Namespace PaxApisDB

   Public Class Item
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
      End Sub
        Public Sub New(ByVal pId As Integer, ByVal pSurname As String, ByVal pFirstName As String, ByVal pBirthDate As Date, _
                       ByVal pGender As String, ByVal pIssuingCountry As String, ByVal pPassportNumber As String, _
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
        End Sub
      Public ReadOnly Property Id() As Integer
         Get
            Id = mudtProps.Id
         End Get
      End Property
      Public Property Surname() As String
         Get
            Surname = mudtProps.Surname
         End Get
         Set(ByVal value As String)
            mudtProps.Surname = value
         End Set
      End Property
      Public Property FirstName() As String
         Get
            FirstName = mudtProps.FirstName
         End Get
         Set(ByVal value As String)
            mudtProps.FirstName = value
         End Set
      End Property
      Public Property BirthDate() As Date
         Get
            BirthDate = mudtProps.Birthdate
         End Get
         Set(ByVal value As Date)
            mudtProps.Birthdate = value
         End Set
      End Property
      Public Property Gender() As String
         Get
            Gender = mudtProps.Gender
         End Get
         Set(ByVal value As String)
            mudtProps.Gender = value
         End Set
      End Property
      Public Property IssuingCountry() As String
         Get
            IssuingCountry = mudtProps.IssuingCountry
         End Get
         Set(ByVal value As String)
            mudtProps.IssuingCountry = value
         End Set
      End Property
      Public Property PassportNumber() As String
         Get
            PassportNumber = mudtProps.PassportNumber
         End Get
         Set(ByVal value As String)
            mudtProps.PassportNumber = value
         End Set
      End Property
      Public Property ExpiryDate() As Date
         Get
            ExpiryDate = mudtProps.ExpiryDate
         End Get
         Set(ByVal value As Date)
            mudtProps.ExpiryDate = value
         End Set
      End Property
      Public Property Nationality() As String
         Get
            Nationality = mudtProps.Nationality
         End Get
         Set(ByVal value As String)
            mudtProps.Nationality = value
         End Set
      End Property
     
      Public Sub Update(ByVal ExpiryDateOK As Boolean)
            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR)    ' (My.Settings.AmadeusReportsConnectionString)
         Dim pobjComm As New SqlClient.SqlCommand

         pobjConn.Open()
         pobjComm = pobjConn.CreateCommand

         With pobjComm
            .CommandType = CommandType.StoredProcedure
            .CommandText = "PaxApisAllInformationInsert"
            .Parameters.Add("@ppSurname", SqlDbType.NVarChar, 30).Value = mudtProps.Surname
            .Parameters.Add("@ppFirstName", SqlDbType.NVarChar, 30).Value = mudtProps.FirstName
            .Parameters.Add("@ppBirthDate", SqlDbType.DateTime).Value = mudtProps.Birthdate
            .Parameters.Add("@ppGender", SqlDbType.NVarChar, 10).Value = mudtProps.Gender
            .Parameters.Add("@ppDocIssuingCountry", SqlDbType.NVarChar, 3).Value = mudtProps.IssuingCountry
            .Parameters.Add("@ppDocnumber", SqlDbType.NVarChar, 15).Value = mudtProps.PassportNumber
            .Parameters.Add("@ppNationality", SqlDbType.NVarChar, 3).Value = mudtProps.Nationality
            If ExpiryDateOK And mudtProps.ExpiryDate > Date.MinValue Then
               .Parameters.Add("@ppDocExpiryDate", SqlDbType.DateTime).Value = mudtProps.ExpiryDate
            Else
               .Parameters.Add("@ppDocExpiryDate", SqlDbType.DateTime).Value = DateSerial(1902, 12, 31)
            End If
                .Parameters.Add("@ppQRFreqFlyer", SqlDbType.NChar, 30).Value = False
            .ExecuteNonQuery()
         End With

      End Sub
   End Class

   Public Class Collection
      Inherits Collections.Generic.Dictionary(Of Integer, Item)

      Public Function Read(ByVal Surname As String, ByVal FirstName As String) As Item
            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR)    ' (My.Settings.AmadeusReportsConnectionString)(My.Settings.AmadeusReportsConnectionString)
         Dim pobjComm As New SqlClient.SqlCommand
         Dim pobjReader As SqlClient.SqlDataReader
         Dim pobjItem As Item

         pobjConn.Open()
         pobjComm = pobjConn.CreateCommand
         Read = New Item

         With pobjComm
            .CommandType = CommandType.StoredProcedure
            .CommandText = "PaxApisInformationAllSelect"
            .Parameters.Add("@ppSurname", SqlDbType.NVarChar, 30).Value = Surname
            .Parameters.Add("@ppFirstName", SqlDbType.NVarChar, 30).Value = FirstName
            pobjReader = .ExecuteReader
         End With

         Clear()

         With pobjReader
            Do While .Read
                    pobjItem = New Item(.Item("ppId"), Surname, FirstName, If(IsDBNull(.Item("ppBirthdate")), Date.MinValue, .Item("ppBirthdate")), .Item("ppGender"), _
                                        .Item("ppDocIssuingCountry"), .Item("ppDocnumber"), If(IsDBNull(.Item("ppDocExpiryDate")), Date.MinValue, .Item("ppDocExpiryDate")), _
                                        .Item("ppNationality"))
               MyBase.Add(pobjItem.Id, pobjItem)
               Read = pobjItem
            Loop
            .Close()
         End With
         pobjConn.Close()

      End Function

   End Class

    Public Class FormElements

        Private Const MONTH_NAMES As String = "JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC"
        Private mstrSalutations() As String = {"MR", "MRS", "MS", "MISS", "MISTER"}
        Private mstrGenderIndicator() As String = {"M", "F", "MI", "FI", "U"}
        Private mflgExpiryDateOK As Boolean

       
       
        Private Function APISModifyFirstName(ByVal FirstName As String) As String

            Dim pintFindPos As Integer

            FirstName = Trim(FirstName)

            For i As Short = 0 To mstrSalutations.GetUpperBound(0)
                pintFindPos = FirstName.IndexOf(mstrSalutations(i))
                If pintFindPos > 0 And pintFindPos = FirstName.Length - mstrSalutations(i).Length Then
                    FirstName = FirstName.Substring(0, pintFindPos).Trim
                End If
            Next

            Return FirstName

        End Function
        Private Function APISDateToIATA(ByVal InDate As Date) As String

            APISDateToIATA = Format(InDate.Day, "00") & MONTH_NAMES.Substring(InDate.Month * 3 - 3, 3) & Format(InDate.Year Mod 100, "00")

        End Function
        Private Function APISDateFromIATA(ByVal InDate As String) As Date

            Dim pintDay As Integer
            Dim pintMonth As Integer
            Dim pintYear As Integer

            Try
                If Not Date.TryParse(InDate, APISDateFromIATA) Then
                    APISDateFromIATA = Date.MinValue
                    pintDay = InDate.Substring(0, 2)
                    pintMonth = (MONTH_NAMES.IndexOf(InDate.Substring(3, 3)) + 2) / 3
                    pintYear = InDate.Substring(5)

                    If pintMonth >= 1 Then
                        APISDateFromIATA = DateSerial(pintYear, pintMonth, pintDay)
                    End If
                End If
            Catch ex As Exception

            End Try

        End Function
        Private Function APISValidateDataRow(ByVal Row As DataGridViewRow) As Boolean

            Dim pdteDate As DateTime
            Dim pflgGenderFound As Boolean = False
            Dim pflgBirthDateOK As Boolean = False
            Dim pflgPassportNumberOK As Boolean = False

            Dim pstrErrorText As String = ""

            pflgPassportNumberOK = (Trim(Row.Cells("PassportNumber").Value).Length > 0)

            If Not Date.TryParse(Row.Cells("Birthdate").Value, pdteDate) Then
                pdteDate = APISDateFromIATA(Row.Cells("Birthdate").Value)
                If pdteDate > Date.MinValue Then
                    pflgBirthDateOK = True
                Else
                    pflgBirthDateOK = False
                End If
            Else
                pflgBirthDateOK = True
            End If

            If Not Date.TryParse(Row.Cells("ExpiryDate").Value, pdteDate) Then
                pdteDate = APISDateFromIATA(Row.Cells("ExpiryDate").Value)
            End If
            If pdteDate > Now Then
                mflgExpiryDateOK = True
            Else
                mflgExpiryDateOK = False
            End If

            pflgGenderFound = False
            For i As Integer = 0 To mstrGenderIndicator.GetUpperBound(0)
                If Row.Cells("Gender").Value = mstrGenderIndicator(i) Then
                    pflgGenderFound = True
                    Exit For
                End If
            Next

            'cmdAPISUpdate.Enabled = pflgBirthDateOK ' And pflgGenderFound And pflgExpiryDateOK And pflgPassportNumberOK

            If Not pflgBirthDateOK Then
                If Not pflgBirthDateOK Then
                    pstrErrorText &= "Invalid birth date" & vbCrLf
                End If
                If Not pflgGenderFound Then
                    pstrErrorText &= "Invalid gender" & vbCrLf
                End If
                If Not pflgPassportNumberOK Then
                    pstrErrorText &= "Passport number missing" & vbCrLf
                End If
                If Not mflgExpiryDateOK Then
                    pstrErrorText &= "Invalid expiry date" & vbCrLf
                End If
            End If
            Row.ErrorText = pstrErrorText

            APISValidateDataRow = pflgBirthDateOK

        End Function
    End Class
End Namespace

