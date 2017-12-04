Imports System.Xml
Namespace CustomProperties

    Public Enum CustomPropertyIDValue As Integer
        BookedBy = 1
        Department = 2
        ReasonFortravel = 4
        CostCentre = 5
    End Enum
    Public Class Item
        Private Structure ClassProps
            Dim ID As Long
            Dim CustomPropertyID As CustomPropertyIDValue
            Dim LookUpValues As String
            Dim LimitToLookup As Boolean
            Dim Label As String
            Dim TFEntityID As Long
            Dim Values() As String
        End Structure
        Private mudtProps As ClassProps

        Public ReadOnly Property ID() As Long
            Get
                ID = mudtProps.ID
            End Get
        End Property

        Public ReadOnly Property CustomPropertyID() As CustomPropertyIDValue
            Get
                CustomPropertyID = mudtProps.CustomPropertyID
            End Get
        End Property

        Public ReadOnly Property LookUpValues() As String
            Get
                LookUpValues = mudtProps.LookUpValues
            End Get
        End Property

        Public ReadOnly Property LimitToLookup() As Boolean
            Get
                LimitToLookup = mudtProps.LimitToLookup
            End Get
        End Property

        Public ReadOnly Property Label() As String
            Get
                Label = mudtProps.Label
            End Get
        End Property

        Public ReadOnly Property TFEntityID() As Long
            Get
                TFEntityID = mudtProps.TFEntityID
            End Get
        End Property

        Public ReadOnly Property ValuesCount As Integer
            Get
                ValuesCount = mudtProps.Values.Length
            End Get
        End Property

        Public ReadOnly Property Value(ByVal Index As Integer) As String
            Get
                If Index >= 0 And Index <= mudtProps.Values.GetUpperBound(0) Then
                    Value = mudtProps.Values(Index)
                Else
                    Throw New Exception("Index out of bounds")
                End If
            End Get
        End Property

        Friend Sub SetValues(ByVal pID As Long, ByVal pCustomPropertyID As CustomPropertyIDValue, ByVal pLookUpValues As String, ByVal pLimitToLookup As Boolean, ByVal pLabel As String, ByVal pTFEntityID As Long)
            With mudtProps
                .ID = pID
                'Debug.Assert(pID <> 113)
                .CustomPropertyID = pCustomPropertyID
                .LookUpValues = pLookUpValues
                .LimitToLookup = pLimitToLookup
                .Label = pLabel
                .TFEntityID = pTFEntityID
                ReDim .Values(0)
                If .LimitToLookup Then
                    ReadXML(pCustomPropertyID, pTFEntityID)
                Else
                    ReadLookUpValues()
                End If
            End With
        End Sub

        Private Sub ReadXML(ByVal pCustomPropertyID As Long, ByVal pTfEntityID As Long)

            Dim pobjXMLValues As New XMLValues
            pobjXMLValues.ReadValues(pCustomPropertyID, pTfEntityID)
            mudtProps.Values = pobjXMLValues.ToArray

        End Sub

        Private Sub ReadLookUpValues()

            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringACC) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .CommandText = "SELECT [Value] " & _
                               " FROM [TravelForceCosmos].[dbo].[CustomPropertyValues] " & _
                               " WHERE CustomPropertyID = " & mudtProps.CustomPropertyID & " And TFEntityID = " & mudtProps.TFEntityID & _
                               " GROUP BY Value " & _
                               " ORDER BY Value"
                pobjReader = .ExecuteReader
            End With
            mudtProps.Values(0) = ""
            With pobjReader
                Dim iCount As Integer = 0
                Do While .Read
                    iCount += 1
                    ReDim Preserve mudtProps.Values(iCount - 1)
                    mudtProps.Values(iCount - 1) = .Item("Value")
                Loop
                .Close()
            End With
            pobjConn.Close()
        End Sub

    End Class

    Public Class XMLValues
        Inherits Collections.Generic.List(Of String)

        Private mstrID As String

        Public Sub ReadValues(ByVal pCustomPropertyID As Long, ByVal pTfEntityID As Long)

            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringACC) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            MyBase.Clear()
            mstrID = ""

            Do While pTfEntityID <> 0 And mstrID.IndexOf("," & pTfEntityID & ",") < 0
                With pobjComm
                    .CommandType = CommandType.Text
                    .CommandText = "SELECT LookUpValues, ISNULL(RelatedEntityID, 0) AS RelatedEntityID " & _
                                   " FROM [TravelForceCosmos].[dbo].[ClientCustomProperties] " & _
                                   " LEFT JOIN TravelForceCosmos.dbo.TFEntities " &
                                   " 	ON TFEntityID=TFEntities.Id " &
                                   " WHERE CustomPropertyID = " & pCustomPropertyID & " And TFEntityID = " & pTfEntityID
                    pobjReader = .ExecuteReader
                End With
                With pobjReader
                    Do While .Read
                        ParseXML(.Item("LookUpValues"))
                        mstrID &= "," & pTfEntityID & ","
                        pTfEntityID = .Item("RelatedEntityID")
                    Loop
                End With
                pobjReader.Close()
            Loop
            MyBase.Sort()

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

    Public Class Values
        Private Structure ClassProps
            Dim Description As String
            Dim Value As String
            Dim IsDefault As Boolean
        End Structure
        Private mudtProps As ClassProps

        Public ReadOnly Property Description() As String
            Get
                Description = mudtProps.Description
            End Get
        End Property

        Public ReadOnly Property Value() As String
            Get
                Value = mudtProps.Value
            End Get
        End Property

        Public ReadOnly Property IsDefault() As Boolean
            Get
                IsDefault = mudtProps.IsDefault
            End Get
        End Property

        'Friend Sub SetValues(ByVal pDescription As String, ByVal pValue As String, ByVal pIsDefault As Boolean)
        '    With mudtProps
        '        .Description = pDescription
        '        .Value = pValue
        '        .IsDefault = pIsDefault
        '    End With
        'End Sub
    End Class

    Public Class Collection
        Inherits Collections.Generic.Dictionary(Of String, Item)

        Private Const MyXMLString As String = "<?xml version='1.0' encoding='utf-8'?><LookUpValues xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema'><CustomPropertyLookupValue Description='Crew' Value='Crew' IsDefault='false' /><CustomPropertyLookupValue Description='Technical' Value='Technical' IsDefault='false' /><CustomPropertyLookupValue Description='Marine' Value='Marine' IsDefault='false' /><CustomPropertyLookupValue Description='HSQE' Value='HSQE' IsDefault='false' /><CustomPropertyLookupValue Description='Finance' Value='Finance' IsDefault='false' /></LookUpValues>"
        Private mflgBookedBy As Boolean
        Private mflgDepartment As Boolean
        Private mflgReasonForTravel As Boolean
        Private mflgCostCentre As Boolean

        Public Sub Load(ByVal pEntityID As Long)

            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringACC) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader
            Dim pobjClass As Item

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .CommandText = " SELECT [Id] " & _
                               "       ,[CustomPropertyID] " & _
                               "       ,[LookUpValues] " & _
                               "       ,[LimitToLookUp] " & _
                               "       ,[Label] " & _
                               "       ,[TFEntityID] " & _
                               "   FROM [TravelForceCosmos].[dbo].[ClientCustomProperties] " & _
                               "   WHERE TFEntityID = '" & pEntityID & "'   " & _
                               "   AND IsDisabled = 0"

                pobjReader = .ExecuteReader
            End With

            mflgBookedBy = False
            mflgDepartment = False
            mflgReasonForTravel = False
            mflgCostCentre = False

            With pobjReader
                Do While .Read
                    pobjClass = New Item
                    pobjClass.SetValues(.Item("Id"), .Item("CustomPropertyID"), .Item("LookUpValues"), .Item("LimitToLookUp"), .Item("Label"), .Item("TFEntityID"))
                    MyBase.Add(pobjClass.ID, pobjClass)
                    If pobjClass.CustomPropertyID = CustomPropertyIDValue.BookedBy Then
                        mflgBookedBy = True
                    ElseIf pobjClass.CustomPropertyID = CustomPropertyIDValue.Department Then
                        mflgDepartment = True
                    ElseIf pobjClass.CustomPropertyID = CustomPropertyIDValue.ReasonFortravel Then
                        mflgReasonForTravel = True
                    ElseIf pobjClass.CustomPropertyID = CustomPropertyIDValue.CostCentre Then
                        mflgCostCentre = True
                    End If
                Loop
                .Close()
            End With
            pobjConn.Close()
        End Sub
        Public ReadOnly Property BookedBy As Boolean
            Get
                BookedBy = mflgBookedBy
            End Get
        End Property
        Public ReadOnly Property Department As Boolean
            Get
                Department = mflgDepartment
            End Get
        End Property
        Public ReadOnly Property ReasonForTravel As Boolean
            Get
                ReasonForTravel = mflgReasonForTravel
            End Get
        End Property
        Public ReadOnly Property CostCentre As Boolean
            Get
                CostCentre = mflgCostCentre
            End Get
        End Property

    End Class

    Public Class CostCentreLookupItem
        Private Structure ClassProps
            Dim Id As Integer
            Dim Code As String
            Dim OldCode As String
            Dim ClientName As String
            Dim ClientLogo As String
            Dim VesselName As String
            Dim CostCentre As String
        End Structure
        Dim mudtProps As ClassProps

        Friend Sub New(ByVal Id As Integer, ByVal Code As String, ByVal OldCode As String, ByVal ClientName As String, ByVal ClientLogo As String, ByVal VesselName As String, ByVal CostCentre As String)

            With mudtProps
                .Id = Id
                .Code = Code
                .OldCode = OldCode
                .ClientName = ClientName
                .ClientLogo = ClientLogo
                .VesselName = VesselName
                .CostCentre = CostCentre
            End With
        End Sub
        Public ReadOnly Property Id As Integer
            Get
                Id = mudtProps.Id
            End Get
        End Property
        Public ReadOnly Property Code As String
            Get
                Code = mudtProps.Code
            End Get
        End Property
        Public ReadOnly Property OldCode As String
            Get
                OldCode = mudtProps.OldCode
            End Get
        End Property
        Public ReadOnly Property ClientName As String
            Get
                ClientName = mudtProps.ClientName
            End Get
        End Property
        Public ReadOnly Property ClientLogo As String
            Get
                ClientLogo = mudtProps.ClientLogo
            End Get
        End Property
        Public ReadOnly Property VesselName As String
            Get
                VesselName = mudtProps.VesselName
            End Get
        End Property
        Public ReadOnly Property CostCentre As String
            Get
                CostCentre = mudtProps.CostCentre
            End Get
        End Property
    End Class

    Public Class CostCentreLookupCollection
        Inherits Collections.Generic.Dictionary(Of String, CostCentreLookupItem)

        Public Sub LoadCustomerGroup(ByVal CustomerGroup As Integer)
            Load(True, CustomerGroup)
        End Sub
        Public Sub LoadCustomer(ByVal CustomerID As Integer)
            Load(False, CustomerID)
        End Sub

        Private Sub Load(ByVal byGroup As Boolean, ByVal Id As Integer)
            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringACC) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader
            Dim pobjClass As CostCentreLookupItem
            Dim pCommandText As String = ""

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                'pCommandText = " USE TravelForceCosmos " & _
                '            " DECLARE @TagSeaChefsMarine INT = 154 " & _
                '            " DECLARE @TagSeaChefsCorporate INT = 155 " & _
                '            " If(OBJECT_ID('tempdb..#TempTable') Is Not Null) " & _
                '            " Begin " & _
                '            "     Drop Table #TempTable " & _
                '            " End " & _
                '            " SELECT ClientCustomProperties.TFEntityID " & _
                '            " ,  CAST(REPLACE(REPLACE(ClientCustomProperties.LookUpValues, 'utf-8', 'utf-16'), ' xmlns:xsd=" & Chr(34) & "http://www.w3.org/2001/XMLSchema" & Chr(34) & " xmlns:xsi=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & "', '') AS XML) AS LookUpValues " & _
                '            " ,  CAST(REPLACE(REPLACE(ClientCustomProperties.DependsOnLookUpValues, 'utf-8', 'utf-16'), ' xmlns:xsd=" & Chr(34) & "http://www.w3.org/2001/XMLSchema" & Chr(34) & " xmlns:xsi=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & "', '') AS XML) AS DependsOnLookUpValues " & _
                '            " INTO #TempTable " & _
                '            " FROM ClientCustomProperties " & _
                '            "   WHERE CustomPropertyID=5 "
                'If byGroup Then
                '    pCommandText &= " 		AND TFEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID=" & Id & ") "
                'Else
                '    pCommandText &= " 		AND TFEntityID =" & Id & " "
                'End If

                'pCommandText &= " If(OBJECT_ID('tempdb..#TempTable1') Is Not Null) " & _
                '        " Begin " & _
                '        "     Drop Table #TempTable1 " & _
                '        " End " & _
                '        " SELECT DISTINCT " & _
                '        " 	TFEntityID " & _
                '        " 	,CustomProperties.CustProps.value('../@MasterLookupValue[1]','VARCHAR(1000)') AS Vessel " & _
                '        " 	,CustomProperties.CustProps.value('.','VARCHAR(1000)') AS CostCentre " & _
                '        " INTO #TempTable1	 " & _
                '        " FROM #TempTable  " & _
                '        "    CROSS APPLY DependsOnLookUpValues.nodes('/DependsOnValues/CustomPropertyDependsOnValue/DependentLookupValues')  CustomProperties(CustProps) " & _
                '        " LEFT JOIN TFEntities " & _
                '        " 	ON TFEntities.Id = TFEntityID " & _
                '        " ORDER BY  Vessel,CostCentre, TFEntityID    " & _
                '        " SELECT DISTINCT " & _
                '        " 	#TempTable.TFEntityID " & _
                '        " 	, Code " & _
                '        " 	,Remarks " & _
                '        " 	, Name " & _
                '        " 	, Logo " & _
                '        " 	,CustomProperties.CustProps.value('@Value[1]','VARCHAR(1000)') AS CostCentre " & _
                '        " 	,#TempTable1.Vessel AS ActualVessel " & _
                '        " FROM #TempTable  " & _
                '        "    CROSS APPLY LookUpValues.nodes('/LookUpValues/CustomPropertyLookupValue')  CustomProperties(CustProps) " & _
                '        " LEFT JOIN TFEntities " & _
                '        " 	ON TFEntities.Id = #TempTable.TFEntityID " & _
                '        " LEFT JOIN #TempTable1 " & _
                '        " 	ON #TempTable.TFEntityID=#TempTable1.TFEntityID " & _
                '        " 		AND CustomProperties.CustProps.value('@Value[1]','VARCHAR(1000)') = #TempTable1.CostCentre " & _
                '        " WHERE #TempTable1.Vessel IS NOT NULL		 " & _
                '        " UNION " & _
                '        " SELECT DISTINCT TFEntities.Id " & _
                '        " 				, TFEntities.Code " & _
                '        " 				, '' AS Remarks " & _
                '        " 				, TFEntities.Name " & _
                '        " 				, TFEntities.Logo " & _
                '        " 				, '' AS CostCentre " & _
                '        " 				, TFEntityDepartments.Name AS ActualVessel " & _
                '        " FROM TFEntities " & _
                '        " LEFT JOIN TFEntityDepartments ON TFEntityDepartments.EntityID=TFEntities.Id  " & _
                '        " 		  AND TFEntityDepartments.InUse=1 " & _
                '        " 		  AND (SELECT COUNT(*) FROM #TempTable1 WHERE TFEntityDepartments.Name = #TempTable1.Vessel) = 0 "
                'If byGroup Then
                '    pCommandText &= " 		WHERE TFEntities.Id IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID=" & Id & ") "
                'Else
                '    pCommandText &= " 		WHERE TFEntities.Id =" & Id & " "
                'End If
                'pCommandText &= " ORDER BY  CostCentre, ActualVessel, Code    " & _
                '        " If(OBJECT_ID('tempdb..#TempTable') Is Not Null) " & _
                '        " Begin " & _
                '        "     Drop Table #TempTable " & _
                '        " End " & _
                '        " If(OBJECT_ID('tempdb..#TempTable1') Is Not Null) " & _
                '        " Begin " & _
                '        "     Drop Table #TempTable1 " & _
                '        " End "
                pCommandText = "USE TravelForceCosmos   " & _
                " If(OBJECT_ID('tempdb..#TempTable') Is Not Null)   " & _
                " Begin       " & _
                " Drop Table #TempTable   " & _
                " End   " & _
                " SELECT ClientCustomProperties.TFEntityID   " & _
                " 		, CAST(REPLACE(REPLACE(ClientCustomProperties.LookUpValues, 'utf-8', 'utf-16'), ' xmlns:xsd=" & Chr(34) & "http://www.w3.org/2001/XMLSchema" & Chr(34) & " xmlns:xsi=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & "', '') AS XML) AS LookUpValues   " & _
                " 		, CAST(REPLACE(REPLACE(ClientCustomProperties.DependsOnLookUpValues, 'utf-8', 'utf-16'), ' xmlns:xsd=" & Chr(34) & "http://www.w3.org/2001/XMLSchema" & Chr(34) & " xmlns:xsi=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & "', '') AS XML) AS DependsOnLookUpValues   " & _
                " 		INTO #TempTable   " & _
                " FROM ClientCustomProperties     " & _
                " WHERE CustomPropertyID=5     "
                If byGroup Then
                    pCommandText &= " 		AND TFEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID=" & Id & ") "
                Else
                    pCommandText &= " 		AND TFEntityID =" & Id & " "
                End If
                pCommandText &= " If(OBJECT_ID('tempdb..#TempTable1') Is Not Null)   " & _
                " Begin " & _
                " Drop Table #TempTable1 " & _
                " End " & _
                " SELECT DISTINCT   #TempTable.TFEntityID " & _
                " 				  ,CustomProperties.CustProps.value('../@MasterLookupValue[1]','VARCHAR(1000)') AS Vessel    " & _
                " 				  ,CustomProperties.CustProps.value('.','VARCHAR(1000)') AS CostCentre   " & _
                " 				  INTO #TempTable1    " & _
                " FROM #TempTable       " & _
                " CROSS APPLY DependsOnLookUpValues.nodes('/DependsOnValues/CustomPropertyDependsOnValue/DependentLookupValues')  CustomProperties(CustProps)   " & _
                " LEFT JOIN TFEntities   ON TFEntities.Id = #TempTable.TFEntityID   " & _
                " LEFT JOIN TFEntityDepartments ON TFEntityDepartments.EntityID = TFEntities.Id  " & _
                " 								 AND CustomProperties.CustProps.value('../@MasterLookupValue[1]','VARCHAR(1000)') = TFEntityDepartments.Name  " & _
                " 								 AND TFEntityDepartments.InUse=1 " & _
                " ORDER BY  Vessel,CostCentre, TFEntityID      " & _
                " SELECT DISTINCT   #TempTable.TFEntityID    " & _
                " 				  , Code    " & _
                " 				  , '' AS Remarks    " & _
                " 				  , Name    " & _
                " 				  , Logo    " & _
                " 				  ,CustomProperties.CustProps.value('@Value[1]','VARCHAR(1000)') AS CostCentre    " & _
                " 				  ,#TempTable1.Vessel AS ActualVessel   " & _
                " FROM #TempTable       " & _
                " CROSS APPLY LookUpValues.nodes('/LookUpValues/CustomPropertyLookupValue')  CustomProperties(CustProps)   " & _
                " LEFT JOIN TFEntities   ON TFEntities.Id = #TempTable.TFEntityID   " & _
                " LEFT JOIN #TempTable1  ON #TempTable.TFEntityID=#TempTable1.TFEntityID     " & _
                "                           AND CustomProperties.CustProps.value('@Value[1]','VARCHAR(1000)') = #TempTable1.CostCentre   " & _
                " WHERE #TempTable1.Vessel IS NOT NULL     " & _
                " UNION " & _
                " SELECT DISTINCT TFEntities.Id " & _
                " 				, TFEntities.Code " & _
                " 				, '' AS Remarks " & _
                " 				, TFEntities.Name " & _
                " 				, TFEntities.Logo " & _
                " 				, '' AS CostCentre " & _
                " 				, TFEntityDepartments.Name AS ActualVessel " & _
                " FROM TFEntities " & _
                " LEFT JOIN TFEntityDepartments ON TFEntityDepartments.EntityID=TFEntities.Id  " & _
                " 		  AND TFEntityDepartments.InUse=1 " & _
                " 		  AND (SELECT COUNT(*) FROM #TempTable1 WHERE TFEntityDepartments.Name = #TempTable1.Vessel) = 0 " & _
                " WHERE TFEntityDepartments.Name IS NOT NULL "
                If byGroup Then
                    pCommandText &= " 		AND TFEntities.Id IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID=" & Id & ") "
                Else
                    pCommandText &= " 		AND TFEntities.Id =" & Id & " "
                End If
                pCommandText &= " ORDER BY   ActualVessel,CostCentre, Code    " & _
                " If(OBJECT_ID('tempdb..#TempTable') Is Not Null)   " & _
                " Begin       " & _
                " Drop Table #TempTable   " & _
                " End   " & _
                " If(OBJECT_ID('tempdb..#TempTable1') Is Not Null)   " & _
                " Begin       " & _
                " Drop Table #TempTable1   " & _
                " End  "


                .CommandText = pCommandText
                pobjReader = .ExecuteReader
            End With

            Dim pId As Integer = 0
            MyBase.Clear()
            With pobjReader
                Do While .Read
                    pId = pId + 1
                    pobjClass = New CostCentreLookupItem(pId, .Item("Code"), .Item("Remarks"), .Item("Name"), .Item("Logo"), .Item("ActualVessel"), .Item("CostCentre"))
                    MyBase.Add(pobjClass.Id, pobjClass)
                Loop
                .Close()
            End With
            pobjConn.Close()

        End Sub
    End Class
End Namespace
