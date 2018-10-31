Public Class CustomPropertiesXMLValues

    Inherits Collections.Generic.List(Of String)

    Private mstrID As String

    Public Sub ReadValues(ByVal pCustomPropertyID As Integer, ByVal pTfEntityID As Integer)

        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringACC) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        MyBase.Clear()
        mstrID = ""

        If MySettings.PCCBackOffice = 1 Then

            Do While pTfEntityID <> 0 And mstrID.IndexOf("," & pTfEntityID & ",") < 0
                With pobjComm
                    .CommandType = CommandType.Text
                    .CommandText = "SELECT LookUpValues, ISNULL(RelatedEntityID, 0) AS RelatedEntityID " &
                               " FROM TravelForceCosmos.dbo.ClientCustomProperties " &
                               " LEFT JOIN TravelForceCosmos.dbo.TFEntities " &
                               " 	ON TFEntityID=TFEntities.Id " &
                               " WHERE CustomPropertyID = " & pCustomPropertyID & " And TFEntityID = " & pTfEntityID
                    pobjReader = .ExecuteReader
                End With
                With pobjReader
                    Do While .Read
                        ParseXML(CStr(.Item("LookUpValues")))
                        mstrID &= "," & pTfEntityID & ","
                        pTfEntityID = CInt(.Item("RelatedEntityID"))
                    Loop
                End With
                pobjReader.Close()
            Loop
            MyBase.Sort()
        ElseIf MySettings.PCCBackOffice = 2 Then
            With pobjComm
                .CommandType = CommandType.Text
                Select Case pCustomPropertyID
                    Case 1 ' booked by
                        .CommandText = "SELECT Child_Value AS Name " &
                                       " From Disco_Instone_EU.dbo.Costcen " &
                                       "  LEFT JOIN Company " &
                                       "  ON Costcen.Account_Id=Company.Account_Id " &
                                       "  WHERE CostCen.Account_id = " & pTfEntityID & " AND Child_Name = 'BBY' " &
                                       " ORDER BY Child_Value"
                    Case 4 ' reason for travel
                        .CommandText = "SELECT Child_Value AS Name " &
                                      " From Disco_Instone_EU.dbo.Costcen " &
                                      "  LEFT JOIN Company " &
                                      "  ON Costcen.Account_Id=Company.Account_Id " &
                                      "  WHERE CostCen.Account_id = " & pTfEntityID & " AND Child_Name = 'REF2' " &
                                      " ORDER BY Child_Value"
                End Select

                pobjReader = .ExecuteReader
            End With
            With pobjReader
                Do While .Read
                    If Not MyBase.Contains(CStr(.Item("Name"))) Then
                        MyBase.Add(CStr(.Item("Name")))
                    End If
                Loop
            End With
        End If

    End Sub
    Private Sub ParseXML(ByVal pXMLString As String)

        Try
            Dim xmlString As String = pXMLString
            Dim sr As New System.IO.StringReader(xmlString)
            Dim doc As New Xml.XmlDocument
            doc.Load(sr)
            'or just in this case doc.LoadXML(xmlString)
            Dim reader As New Xml.XmlNodeReader(doc)

            While reader.Read()
                Select Case reader.NodeType
                    Case Xml.XmlNodeType.Element
                        If reader.Name = "CustomPropertyLookupValue" Then
                            'Dim pFound As Boolean = False
                            Dim pText As String = ""
                            pText = reader.GetAttribute("Value").ToUpper.Trim
                            If reader.GetAttribute("Description").ToUpper.Trim <> "" And reader.GetAttribute("Description").ToUpper.Trim <> pText Then
                                If pText <> "" Then
                                    pText &= "/"
                                End If
                                pText &= reader.GetAttribute("Description").ToUpper.Trim
                            End If
                            If pText <> "" Then
                                If Not MyBase.Contains(pText) Then
                                    MyBase.Add(pText)
                                End If
                            End If
                        End If
                End Select
            End While
        Catch ex As Exception

        End Try

    End Sub

End Class