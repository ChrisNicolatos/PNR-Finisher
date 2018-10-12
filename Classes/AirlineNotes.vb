﻿Option Strict On
Option Explicit On
Namespace AirlineNotes
    Friend Class Item
        Private Structure ClassProps
            Dim ID As Integer
            Dim AirlineCode As String
            Dim FlightType As String
            Dim Seaman As Boolean
            Dim SeqNo As Integer
            Dim GDSElement As String
            Dim GDSText As String
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
        Public ReadOnly Property GDSElement() As String
            Get
                GDSElement = mudtProps.GDSElement
            End Get
        End Property
        Public ReadOnly Property GDSText() As String
            Get
                GDSText = mudtProps.GDSText
            End Get
        End Property

        Friend Sub SetValues(ByVal pID As Integer, ByVal pAirlineCode As String, ByVal pFlightType As String, ByVal pSeaman As Boolean,
                             ByVal pSeqNo As Integer, ByVal pGDSElement As String, ByVal pGDSText As String)
            With mudtProps
                .ID = pID
                .AirlineCode = pAirlineCode
                .FlightType = pFlightType
                .Seaman = pSeaman
                .SeqNo = pSeqNo
                .GDSElement = pGDSElement
                .GDSText = pGDSText
            End With
        End Sub
    End Class

    Friend Class Collection
        Inherits System.Collections.Generic.Dictionary(Of Integer, Item)

        Public Sub Load(ByVal pAirlineCode As String, ByVal GDSCode As Utilities.EnumGDSCode)

            Dim pCommandText As String
            If GDSCode = Utilities.EnumGDSCode.Amadeus Then
                pCommandText = "SELECT anID, " &
                            " anAirlineCode, " &
                            " anFlightType, " &
                            " ISNULL(anSeaman, 0) AS anSeaman, " &
                            " anSeqNo, " &
                            " anAmadeusElement AS GDSElement, " &
                            " anAmadeusText AS GDSText " &
                            " FROM AmadeusReports.dbo.AirlineNotes " &
                            " WHERE anAirlineCode = @AirlineCode " &
                            " ORDER BY anSeqNo"
            ElseIf GDSCode = Utilities.EnumGDSCode.Galileo Then
                pCommandText = "SELECT anID, " &
                            " anAirlineCode, " &
                            " anFlightType, " &
                            " ISNULL(anSeaman, 0) AS anSeaman, " &
                            " anSeqNo, " &
                            " '' AS GDSElement, " &
                            " anGalileoEntry AS GDSText " &
                            " FROM AmadeusReports.dbo.AirlineNotes " &
                            " WHERE anAirlineCode = @AirlineCode " &
                            " ORDER BY anSeqNo"
            Else
                Throw New Exception("AirlineNotes.Collection.Load()" & vbCrLf & "GDS is not selected")
            End If
            ReadFromDB(pCommandText, pAirlineCode)

        End Sub
        Private Sub ReadFromDB(ByVal CommandText As String, ByVal pAirlineCode As String)

            Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader
            Dim pobjClass As Item
            Dim pID As Integer = 0

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand
            MyBase.Clear()
            With pobjComm
                .CommandType = CommandType.Text
                .Parameters.Add("@AirlineCode", SqlDbType.NVarChar, 10).Value = pAirlineCode
                .CommandText = CommandText
                pobjReader = .ExecuteReader
            End With

            With pobjReader
                Do While .Read
                    pID += 1
                    pobjClass = New Item
                    pobjClass.SetValues(CInt(.Item("anID")), CStr(.Item("anAirlineCode")), CStr(.Item("anFlightType")), CBool(.Item("anSeaman")),
                                        CInt(.Item("anSeqNo")), CStr(.Item("GDSElement")), CStr(.Item("GDSText")))
                    MyBase.Add(pID, pobjClass)
                Loop
                .Close()
            End With
            pobjConn.Close()

        End Sub

    End Class

End Namespace
