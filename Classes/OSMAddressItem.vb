Option Strict Off
Option Explicit On
Imports System.IO
Public Class OSMAddressItem
    Private Structure ClassProps
        Dim Id As Integer
        Dim Office As String
        Dim SignedByName As String
        Dim Title As String
        Dim CompanyName As String
        Dim Address As String
        Dim PCArea As String
        Dim Country As String
        Dim Telephone As String
        Dim LogoImage_fk As Integer
        Dim SignatureImage_fk As Integer
        Dim LogoImage() As Byte
        Dim SignatureImage() As Byte
    End Structure
    Private mudtProps As ClassProps
    Public Overrides Function ToString() As String
        Return Office
    End Function
    Public Property Id As Integer
        Get
            Return mudtProps.Id
        End Get
        Set(value As Integer)
            mudtProps.Id = value
        End Set
    End Property
    Public Property Office As String
        Get
            Return mudtProps.Office
        End Get
        Set(value As String)
            mudtProps.Office = value
        End Set
    End Property
    Public Property SignedByName As String
        Get
            Return mudtProps.SignedByName
        End Get
        Set(value As String)
            mudtProps.SignedByName = value
        End Set
    End Property
    Public Property Title As String
        Get
            Return mudtProps.Title
        End Get
        Set(value As String)
            mudtProps.Title = value
        End Set
    End Property
    Public Property CompanyName As String
        Get
            Return mudtProps.CompanyName
        End Get
        Set(value As String)
            mudtProps.CompanyName = value
        End Set
    End Property
    Public Property Address As String
        Get
            Return mudtProps.Address
        End Get
        Set(value As String)
            mudtProps.Address = value
        End Set
    End Property
    Public Property PCArea As String
        Get
            Return mudtProps.PCArea
        End Get
        Set(value As String)
            mudtProps.PCArea = value
        End Set
    End Property
    Public Property Country As String
        Get
            Return mudtProps.Country
        End Get
        Set(value As String)
            mudtProps.Country = value
        End Set
    End Property
    Public Property Telephone As String
        Get
            Return mudtProps.Telephone
        End Get
        Set(value As String)
            mudtProps.Telephone = value
        End Set
    End Property
    Public Property LogoImage_fk As Integer
        Get
            Return mudtProps.LogoImage_fk
        End Get
        Set(value As Integer)
            mudtProps.LogoImage_fk = value
        End Set
    End Property
    Public Property SignatureImage_fk As Integer
        Get
            Return mudtProps.SignatureImage_fk
        End Get
        Set(value As Integer)
            mudtProps.SignatureImage_fk = value
        End Set
    End Property
    Public ReadOnly Property LogoImage As Image
        Get
            Return Image.FromStream(New MemoryStream(mudtProps.LogoImage))
        End Get
    End Property
    Public ReadOnly Property LogoImageByteArray() As Byte()
        Get
            Return mudtProps.LogoImage
        End Get
    End Property
    Public ReadOnly Property SignatureImage As Image
        Get
            Return Image.FromStream(New MemoryStream(mudtProps.SignatureImage))
        End Get
    End Property
    Public ReadOnly Property SignatureImageByteArray As Byte()
        Get
            Return mudtProps.SignatureImage
        End Get
    End Property
    Public Sub SetValues(ByVal pId As Integer, ByVal pOffice As String, ByVal pSignedByName As String, ByVal pTitle As String _
                       , ByVal pCompanyName As String, ByVal pAddress As String, ByVal pPCArea As String, ByVal pCountry As String _
                       , ByVal pTelephone As String, ByVal pLogoImage_fk As Integer, ByVal pSignatureImage_fk As Integer)
        With mudtProps
            .Id = pId
            .Office = pOffice
            .SignedByName = pSignedByName
            .Title = pTitle
            .CompanyName = pCompanyName
            .Address = pAddress
            .PCArea = pPCArea
            .Country = pCountry
            .Telephone = pTelephone
            .LogoImage_fk = pLogoImage_fk
            .SignatureImage_fk = pSignatureImage_fk
            If LogoImage_fk <> 0 Then
                .LogoImage = LoadImage(LogoImage_fk)
            End If
            If SignatureImage_fk <> 0 Then
                .SignatureImage = LoadImage(SignatureImage_fk)
            End If
        End With

    End Sub
    Private Function LoadImage(ByVal pImage_FK As Integer) As Byte()
        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand
        With pobjComm
            .CommandType = CommandType.Text
            .Parameters.Add("@Id", SqlDbType.Int).Value = pImage_FK
            .CommandText = "SELECT imageData FROM [AmadeusReports].[dbo].[Images] WHERE imageId = @Id"
            LoadImage = .ExecuteScalar
        End With

        pobjConn.Close()
    End Function
End Class
