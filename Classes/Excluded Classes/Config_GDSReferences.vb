Option Strict On
Option Explicit On
Namespace Config_GDSReferences
    'Public Class ConfigGDSReferenceItem
    '    Private Structure ClassProps
    '        Dim Id As Integer
    '        Dim Key As String
    '        Dim Value As String
    '        Dim GDSKey As Integer
    '        Dim BOKey As Integer
    '        Dim Element As String
    '        Dim RefId As String
    '        Dim RefDetail As String
    '    End Structure
    '    Private mudtProps As ClassProps
    '    Public ReadOnly Property Id As Integer
    '        Get
    '            Id = mudtProps.Id
    '        End Get
    '    End Property
    '    Public ReadOnly Property Key As String
    '        Get
    '            Key = mudtProps.Key
    '        End Get
    '    End Property
    '    Public ReadOnly Property Value As String
    '        Get
    '            Value = mudtProps.Value
    '        End Get
    '    End Property
    '    Public ReadOnly Property GDSKey As Integer
    '        Get
    '            GDSKey = mudtProps.GDSKey
    '        End Get
    '    End Property
    '    Public ReadOnly Property BOKey As Integer
    '        Get
    '            BOKey = mudtProps.BOKey
    '        End Get
    '    End Property
    '    Public ReadOnly Property Element As String
    '        Get
    '            Element = mudtProps.Element
    '        End Get
    '    End Property
    '    Public ReadOnly Property RefId As String
    '        Get
    '            RefId = mudtProps.RefId
    '        End Get
    '    End Property
    '    Public ReadOnly Property RefDetail As String
    '        Get
    '            RefDetail = mudtProps.RefDetail
    '        End Get
    '    End Property
    '    Public Sub SetValues(pId As Integer, pKey As String, pValue As String, pGDSKey As Integer, pBOKey As Integer, pElement As String, pRefId As String, pRefDetail As String)
    '        With mudtProps
    '            .Id = pId
    '            .Key = pKey
    '            .Value = pValue
    '            .GDSKey = pGDSKey
    '            .BOKey = pBOKey
    '            .Element = pElement
    '            .RefId = pRefId
    '            .RefDetail = pRefDetail
    '        End With
    '    End Sub
    'End Class
    'Public Class ConfigGDSReferenceCollection
    '    Inherits Collections.Generic.Dictionary(Of String, ConfigGDSReferenceItem)
    '    Public Sub Read(ByVal BackOffice As Integer, ByVal GDSCode As Utilities.EnumGDSCode)

    '        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
    '        Dim pobjComm As New SqlClient.SqlCommand
    '        Dim pobjReader As SqlClient.SqlDataReader

    '        pobjConn.Open()
    '        pobjComm = pobjConn.CreateCommand

    '        With pobjComm
    '            .CommandType = CommandType.Text
    '            .Parameters.Add("@PCCBackOffice", SqlDbType.BigInt).Value = BackOffice
    '            .Parameters.Add("@GDS", SqlDbType.BigInt).Value = GDSCode
    '            .CommandText = " SELECT pfrID " &
    '                           " , ISNULL(pfrKey,'') AS pfrKey " &
    '                           " , ISNULL(pfrValue,'') AS pfrValue " &
    '                           " , ISNULL(pfrGDS_fkey,0) AS pfrGDS_fkey " &
    '                           " , ISNULL(pfrBO_fkey,0) AS pfrBO_fkey " &
    '                           " , ISNULL(pfrGDSElement,'') AS pfrGDSElement " &
    '                           " , ISNULL(pfrReferenceIdentifier,'') AS pfrReferenceIdentifier " &
    '                           " , ISNULL(pfrReferenceDetail,'') AS pfrReferenceDetail " &
    '                           " FROM [AmadeusReports].[dbo].[PNRFinisherGDS_BOReferences] " &
    '                           " WHERE pfrGDS_fkey = @GDS AND pfrBO_fkey = @PCCBackOffice"
    '            pobjReader = .ExecuteReader
    '        End With

    '        MyBase.Clear()

    '        With pobjReader
    '            While pobjReader.Read
    '                Dim pItem As New ConfigGDSReferenceItem
    '                pItem.SetValues(CInt(.Item("pfrID")), CStr(.Item("pfrKey")), CStr(.Item("pfrValue")), CInt(.Item("pfrGDS_fkey")), CInt(.Item("pfrBO_fkey")), CStr(.Item("pfrGDSElement")), CStr(.Item("pfrReferenceIdentifier")), CStr(.Item("pfrReferenceDetail")))
    '                MyBase.Add(pItem.Key, pItem)
    '            End While
    '            .Close()
    '        End With
    '        pobjConn.Close()

    '    End Sub
    'End Class
End Namespace
