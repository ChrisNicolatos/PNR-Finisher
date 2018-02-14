Option Strict Off
Option Explicit On
Namespace AirlineNotes
    Public Class Item
        Private Structure ClassProps
            Dim ID As Integer
            Dim AirlineCode As String
            Dim FlightType As String
            Dim Seaman As Boolean
            Dim SeqNo As Integer
            Dim AmadeusElement As String
            Dim AmadeusText As String
        End Structure
        Private mudtProps As ClassProps

        Public ReadOnly Property ID() As Integer
            Get
                ID = mudtProps.ID
            End Get
        End Property
        Public ReadOnly Property AirlineCode() As String
            Get
                AirlineCode = mudtProps.AirlineCode
            End Get
        End Property
        Public ReadOnly Property FlightType() As String
            Get
                FlightType = mudtProps.FlightType
            End Get
        End Property
        Public ReadOnly Property Seaman() As Boolean
            Get
                Seaman = mudtProps.Seaman
            End Get
        End Property
        Public ReadOnly Property SeqNo() As Integer
            Get
                SeqNo = mudtProps.SeqNo
            End Get
        End Property
        Public ReadOnly Property AmadeusElement() As String
            Get
                AmadeusElement = mudtProps.AmadeusElement
            End Get
        End Property
        Public ReadOnly Property AmadeusText() As String
            Get
                AmadeusText = mudtProps.AmadeusText
            End Get
        End Property

        Friend Sub SetValues(ByVal pID As Integer, ByVal pAirlineCode As String, ByVal pFlightType As String, ByVal pSeaman As Boolean,
                             ByVal pSeqNo As Integer, ByVal pAmadeusElement As String, ByVal pAmadeusText As String)
            With mudtProps
                .ID = pID
                .AirlineCode = pAirlineCode
                .FlightType = pFlightType
                .Seaman = pSeaman
                .SeqNo = pSeqNo
                .AmadeusElement = pAmadeusElement
                .AmadeusText = pAmadeusText
            End With
        End Sub
    End Class

    Public Class Collection
        Inherits System.Collections.Generic.Dictionary(Of Integer, Item)

        Public Sub Load(ByVal pIATACode As String)

            Dim pCommandText As String
            pCommandText = "SELECT anID, " &
                            " anAirlineCode, " &
                            " anFlightType, " &
                            " ISNULL(anSeaman, 0) AS anSeaman, " &
                            " anSeqNo, " &
                            " anAmadeusElement, " &
                            " anAmadeusText " &
                            " FROM AmadeusReports.dbo.AirlineNotes " &
                            " WHERE anAirlineCode = '" & pIATACode & "' " &
                            " ORDER BY anSeqNo"
            ReadFromDB(pCommandText)

        End Sub

        Public Sub Load()

            Dim pCommandText As String
            pCommandText = "SELECT anID, " &
                            " anAirlineCode, " &
                            " anFlightType, " &
                            " ISNULL(anSeaman, 0) AS anSeaman, " &
                            " anSeqNo, " &
                            " anAmadeusElement, " &
                            " anAmadeusText " &
                            " FROM AmadeusReports.dbo.AirlineNotes " &
                            " ORDER BY anAirlineCode, anSeqNo"

            ReadFromDB(pCommandText)

        End Sub
        Private Sub ReadFromDB(ByVal CommandText As String)

            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader
            Dim pobjClass As Item
            Dim pID As Integer = 0

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand
            MyBase.Clear()
            With pobjComm
                .CommandType = CommandType.Text
                .CommandText = CommandText
                pobjReader = .ExecuteReader
            End With

            With pobjReader
                Do While .Read
                    pID += 1
                    pobjClass = New Item
                    pobjClass.SetValues(.Item("anID"), .Item("anAirlineCode"), .Item("anFlightType"), .Item("anSeaman"),
                                        .Item("anSeqNo"), .Item("anAmadeusElement"), .Item("anAmadeusText"))
                    MyBase.Add(pID, pobjClass)
                Loop
                .Close()
            End With
            pobjConn.Close()

        End Sub

    End Class

End Namespace
