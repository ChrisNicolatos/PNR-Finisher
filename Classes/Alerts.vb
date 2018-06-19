Option Strict Off
Option Explicit On
Namespace Alerts
    Friend Class AlertItem
        Private Structure ClassProps
            Dim BackOfficeId As Integer
            Dim ClientCode As String
            Dim Alert As String
            Dim OriginCountry As String
            Dim DestinationCountry As String
        End Structure
        Dim mudtprops As ClassProps
        Public ReadOnly Property BackOfficeID As Integer
            Get
                BackOfficeID = mudtprops.BackOfficeId
            End Get
        End Property
        Public ReadOnly Property ClientCode As String
            Get
                ClientCode = mudtprops.ClientCode
            End Get
        End Property
        Public ReadOnly Property Alert() As String
            Get
                Alert = mudtprops.Alert
            End Get
        End Property
        Public ReadOnly Property OriginCountry As String
            Get
                Return mudtprops.OriginCountry
            End Get
        End Property
        Public ReadOnly Property DestinationCountry As String
            Get
                Return mudtprops.DestinationCountry
            End Get
        End Property
        Friend Sub SetValues(ByVal pBackOfficeID As Integer, ByVal pClientCode As String, ByVal pAlert As String, ByVal pOriginCountry As String, ByVal pDestinationCountry As String)
            With mudtprops
                .BackOfficeId = pBackOfficeID
                .ClientCode = pClientCode
                .Alert = pAlert
                .OriginCountry = pOriginCountry
                .DestinationCountry = pDestinationCountry
            End With
        End Sub
    End Class

    Friend Class Collection
        Inherits Collections.Generic.Dictionary(Of String, AlertItem)
        Public Sub Load()
            Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader
            Dim pobjClass As AlertItem

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            MyBase.Clear()

            With pobjComm
                .CommandType = CommandType.Text
                .CommandText = "SELECT pnaID " &
                               "     , ISNULL(pnaBOId_fkey, 0) AS pnaBOId_fkey " &
                               "     , ISNULL(pnaClientCode, '') AS pnaClientCode " &
                               "     , pnaAlert " &
                               "     , ISNULL(pnaOriginCountry, '') AS pnaOriginCountry " &
                               "     , ISNULL(pnaDestinationCountry, '') AS pnaDestinationCountry " &
                               "FROM [AmadeusReports].[dbo].[PNRFinisherAlerts]"
                pobjReader = .ExecuteReader
            End With

            With pobjReader
                Do While .Read
                    pobjClass = New AlertItem
                    pobjClass.SetValues(.Item("pnaBOId_fkey"), .Item("pnaClientCode"), .Item("pnaAlert"), .Item("pnaOriginCountry"), .Item("pnaDestinationCountry"))
                    MyBase.Add(.Item("pnaID").ToString, pobjClass)
                Loop
            End With
        End Sub
        Public ReadOnly Property Alert(ByVal pBackOfficeId As Integer, ByVal pClientCode As String) As String
            Get
                Alert = ""
                For Each pItem As AlertItem In MyBase.Values
                    If pItem.BackOfficeID = pBackOfficeId And pClientCode = pItem.ClientCode Then
                        Alert = pItem.Alert
                        Exit For
                    End If
                Next
            End Get
        End Property
        Public ReadOnly Property Alert(ByVal pOriginCountry As String, ByVal pDestinationCountry As String) As String
            Get
                Alert = ""
                For Each pItem As AlertItem In MyBase.Values
                    If (pItem.OriginCountry = pOriginCountry And pItem.DestinationCountry = pDestinationCountry) _
                        Or (pItem.OriginCountry = pOriginCountry And pItem.DestinationCountry = "") _
                        Or (pItem.DestinationCountry = pDestinationCountry And pItem.OriginCountry = "") Then
                        Alert &= pItem.Alert & vbCrLf
                    End If
                Next
            End Get
        End Property
    End Class
End Namespace