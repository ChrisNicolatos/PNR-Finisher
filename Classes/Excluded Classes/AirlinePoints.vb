﻿'Option Strict On
'Option Explicit On
'Namespace AirlinePoints
'    Friend Class AirlinePointsItem
'        Private Structure ClassProps
'            Dim PointsCommand As String
'        End Structure
'        Private mudtProps As ClassProps
'        Public ReadOnly Property PointsCommand() As String
'            Get
'                Return mudtProps.PointsCommand
'            End Get
'        End Property
'        Public Overrides Function ToString() As String
'            Return PointsCommand
'        End Function

'        Friend Sub SetValues(ByVal pPointsCommand As String)
'            With mudtProps
'                .PointsCommand = pPointsCommand
'            End With
'        End Sub
'    End Class

'    Friend Class AirlinePointsCollection
'        Inherits System.Collections.Generic.Dictionary(Of Integer, AirlinePointsItem)

'        Public Sub Load(ByVal pCustID As Integer, ByVal pIATACode As String, ByVal GDSCode As Utilities.EnumGDSCode)

'            Dim pCommandText As String
'            Select Case MySettings.PCCBackOffice
'                Case 1
'                    If GDSCode = Utilities.EnumGDSCode.Amadeus Then
'                        pCommandText = "SELECT TravelForceCosmos.dbo.FrequentFlyerCards.Remarks " &
'                                   " FROM TravelForceCosmos.dbo.FrequentFlyerCards  " &
'                                   " 	LEFT OUTER JOIN TravelForceCosmos.dbo.Airlines  " &
'                                   " 		ON TravelForceCosmos.dbo.FrequentFlyerCards.AirlineID = TravelForceCosmos.dbo.Airlines.Id " &
'                                   " WHERE (TravelForceCosmos.dbo.FrequentFlyerCards.TFEntityID = " & pCustID & ")  " &
'                                   " 			AND (TravelForceCosmos.dbo.Airlines.IATACode = '" & pIATACode & "')"
'                    ElseIf GDSCode = Utilities.EnumGDSCode.Galileo Then
'                        pCommandText = "SELECT FrequentFlyerCards_1G.ff1GRemark  AS Remarks " &
'                                        " FROM TravelForceCosmos.dbo.FrequentFlyerCards " &
'                                        " LEFT OUTER JOIN TravelForceCosmos.dbo.Airlines " &
'                                        " ON TravelForceCosmos.dbo.FrequentFlyerCards.AirlineID = TravelForceCosmos.dbo.Airlines.Id " &
'                                        " LEFT JOIN [EUDC-CLSSQL14.ATPI.PRI].AmadeusReports.dbo.FrequentFlyerCards_1G " &
'                                        " ON FrequentFlyerCards.Remarks = FrequentFlyerCards_1G.ffTFCRemark " &
'                                        " WHERE (TravelForceCosmos.dbo.FrequentFlyerCards.TFEntityID = " & pCustID & ")  " &
'                                        " AND (TravelForceCosmos.dbo.Airlines.IATACode = '" & pIATACode & "')"
'                    Else
'                        Throw New Exception("AirlinePoints.Collection.Load()" & vbCrLf & "GDS is not selected")
'                    End If
'                    ReadFromDB(pCommandText, UtilitiesDB.ConnectionStringACC)
'                Case 2
'                    If GDSCode = Utilities.EnumGDSCode.Amadeus Then
'                        pCommandText = "SELECT pnfAmadeusEntry AS Remarks " &
'                                   "  FROM AmadeusReports.dbo.PNRFinisherCorporateDeals " &
'                                   "  WHERE pnfClientId_fkey = " & pCustID & " AND pnfAirlineCode = '" & pIATACode & "' "
'                    ElseIf GDSCode = Utilities.EnumGDSCode.Galileo Then
'                        pCommandText = "SELECT pnfGalileoEntry AS Remarks " &
'                                   "  FROM AmadeusReports.dbo.PNRFinisherCorporateDeals " &
'                                   "  WHERE pnfClientId_fkey = " & pCustID & " AND pnfAirlineCode = '" & pIATACode & "' "
'                    Else
'                        Throw New Exception("FrequentFlyer.Collection.Load()" & vbCrLf & "GDS is not selected")
'                    End If
'                    ReadFromDB(pCommandText, UtilitiesDB.ConnectionStringPNR)
'            End Select

'        End Sub

'        Private Sub ReadFromDB(ByVal CommandText As String, ByVal ConnectionString As String)

'            Dim pobjConn As New SqlClient.SqlConnection(ConnectionString) ' ActiveConnection)
'            Dim pobjComm As New SqlClient.SqlCommand
'            Dim pobjReader As SqlClient.SqlDataReader
'            Dim pobjClass As AirlinePointsItem
'            Dim pID As Integer = 0

'            pobjConn.Open()
'            pobjComm = pobjConn.CreateCommand
'            MyBase.Clear()
'            With pobjComm
'                .CommandType = CommandType.Text
'                .CommandText = CommandText
'                pobjReader = .ExecuteReader
'            End With

'            With pobjReader
'                Do While .Read
'                    pID += 1
'                    pobjClass = New AirlinePointsItem
'                    pobjClass.SetValues(CStr(.Item("Remarks")))
'                    MyBase.Add(pID, pobjClass)
'                Loop
'                .Close()
'            End With
'            pobjConn.Close()

'        End Sub

'    End Class

'End Namespace
