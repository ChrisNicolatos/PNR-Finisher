Option Strict On
Option Explicit On
Namespace ConditionalGDSEntry
    'Public Class ConditionalGDSEntryItem

    '    Dim m1AEntry As String
    '    Dim m1GEntry As String
    '    Public ReadOnly Property ConditionalEntry1A As String
    '        Get
    '            ConditionalEntry1A = MySettings.ConvertGDSValue(m1AEntry)
    '        End Get
    '    End Property
    '    Public ReadOnly Property ConditionalEntry1G As String
    '        Get
    '            ConditionalEntry1G = MySettings.ConvertGDSValue(m1GEntry)
    '        End Get
    '    End Property
    '    Friend Sub SetValues(ByVal p1AEntry As String, ByVal p1GEntry As String)
    '        m1AEntry = p1AEntry
    '        m1GEntry = p1GEntry
    '    End Sub
    'End Class
    'Public Class ConditionalGDSEntryCollection
    '    Inherits Collections.Generic.Dictionary(Of Integer, ConditionalGDSEntryItem)
    '    Public Sub Load(ByVal BOFkey As Integer, ByVal ClientId As Integer, ByVal Vesselname As String)

    '        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
    '        Dim pobjComm As New SqlClient.SqlCommand
    '        Dim pobjReader As SqlClient.SqlDataReader

    '        pobjConn.Open()
    '        pobjComm = pobjConn.CreateCommand

    '        With pobjComm
    '            .CommandType = CommandType.Text
    '            .Parameters.Add("@BOKey", SqlDbType.BigInt).Value = BOFkey
    '            .Parameters.Add("@ClientId", SqlDbType.BigInt).Value = ClientId
    '            .Parameters.Add("@VesselName", SqlDbType.NVarChar, 254).Value = Vesselname
    '            .CommandText = "SELECT pfcAmadeusEntry, pfcGalileoEntry " &
    '            "  FROM AmadeusReports.dbo.PNRFinisherConditionalGDSEntry " &
    '            "  WHERE pfcBO_fkey = @BOKey AND pfcClientId_fkey = @ClientId AND pfcVesselName = @VesselName "
    '            pobjReader = .ExecuteReader
    '        End With

    '        MyBase.Clear()

    '        Dim pIndex As Integer = 0
    '        With pobjReader
    '            While pobjReader.Read
    '                Dim pItem As New ConditionalGDSEntryItem
    '                pIndex += 1
    '                pItem.SetValues(CStr(.Item("pfcAmadeusEntry")), CStr(.Item("pfcGalileoEntry")))
    '                MyBase.Add(pIndex, pItem)
    '            End While
    '            .Close()
    '        End With
    '        pobjConn.Close()

    '    End Sub
    'End Class
End Namespace
