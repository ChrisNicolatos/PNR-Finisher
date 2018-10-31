Public Class CostCentreLookupCollection
    Inherits Collections.Generic.Dictionary(Of Integer, CostCentreLookupItem)

    Public Sub LoadCustomerGroup(ByVal CustomerGroup As Integer)
        Load(True, CustomerGroup)
    End Sub
    Public Sub LoadCustomer(ByVal CustomerID As Integer)
        Load(False, CustomerID)
    End Sub

    Private Sub Load(ByVal byGroup As Boolean, ByVal Id As Integer)

        If MySettings.PCCBackOffice = 1 Then

            Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringACC) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader
            Dim pobjClass As CostCentreLookupItem
            Dim pCommandText As String = ""

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text

                pCommandText = "USE TravelForceCosmos   " &
            " If(OBJECT_ID('tempdb..#TempTable') Is Not Null)   " &
            " Begin       " &
            " Drop Table #TempTable   " &
            " End   " &
            " SELECT ClientCustomProperties.TFEntityID   " &
            " 		, CAST(REPLACE(REPLACE(ClientCustomProperties.LookUpValues, 'utf-8', 'utf-16'), ' xmlns:xsd=" & Chr(34) & "http://www.w3.org/2001/XMLSchema" & Chr(34) & " xmlns:xsi=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & "', '') AS XML) AS LookUpValues   " &
            " 		, CAST(REPLACE(REPLACE(ClientCustomProperties.DependsOnLookUpValues, 'utf-8', 'utf-16'), ' xmlns:xsd=" & Chr(34) & "http://www.w3.org/2001/XMLSchema" & Chr(34) & " xmlns:xsi=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & "', '') AS XML) AS DependsOnLookUpValues   " &
            " 		INTO #TempTable   " &
            " FROM ClientCustomProperties     " &
            " LEFT JOIN TFEntities ON TFEntityId = TFEntities.Id" &
            " WHERE CustomPropertyID=5 AND TFEntities.IsActive = 1  "
                If byGroup Then
                    pCommandText &= " 		AND TFEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID=" & Id & ") "
                Else
                    pCommandText &= " 		AND TFEntityID =" & Id & " "
                End If
                pCommandText &= " If(OBJECT_ID('tempdb..#TempTable1') Is Not Null)   " &
            " Begin " &
            " Drop Table #TempTable1 " &
            " End " &
            " SELECT DISTINCT   #TempTable.TFEntityID " &
            " 				  ,CustomProperties.CustProps.value('../@MasterLookupValue[1]','VARCHAR(1000)') AS Vessel    " &
            " 				  ,CustomProperties.CustProps.value('.','VARCHAR(1000)') AS CostCentre   " &
            " 				  INTO #TempTable1    " &
            " FROM #TempTable       " &
            " CROSS APPLY DependsOnLookUpValues.nodes('/DependsOnValues/CustomPropertyDependsOnValue/DependentLookupValues')  CustomProperties(CustProps)   " &
            " LEFT JOIN TFEntities   ON TFEntities.Id = #TempTable.TFEntityID   " &
            " LEFT JOIN TFEntityDepartments ON TFEntityDepartments.EntityID = TFEntities.Id  " &
            " 								 AND CustomProperties.CustProps.value('../@MasterLookupValue[1]','VARCHAR(1000)') = TFEntityDepartments.Name  " &
            " 								 AND TFEntityDepartments.InUse=1 " &
            " ORDER BY  Vessel,CostCentre, TFEntityID      " &
            " SELECT DISTINCT   #TempTable.TFEntityID    " &
            " 				  , Code    " &
            " 				  , '' AS Remarks    " &
            " 				  , Name    " &
            " 				  , Logo    " &
            " 				  ,CustomProperties.CustProps.value('@Value[1]','VARCHAR(1000)') AS CostCentre    " &
            " 				  ,#TempTable1.Vessel AS ActualVessel   " &
            " FROM #TempTable       " &
            " CROSS APPLY LookUpValues.nodes('/LookUpValues/CustomPropertyLookupValue')  CustomProperties(CustProps)   " &
            " LEFT JOIN TFEntities   ON TFEntities.Id = #TempTable.TFEntityID   " &
            " LEFT JOIN #TempTable1  ON #TempTable.TFEntityID=#TempTable1.TFEntityID     " &
            "                           AND CustomProperties.CustProps.value('@Value[1]','VARCHAR(1000)') = #TempTable1.CostCentre   " &
            " WHERE #TempTable1.Vessel IS NOT NULL AND TFEntities.IsActive = 1     " &
            " UNION " &
            " SELECT DISTINCT TFEntities.Id " &
            " 				, TFEntities.Code " &
            " 				, '' AS Remarks " &
            " 				, TFEntities.Name " &
            " 				, TFEntities.Logo " &
            " 				, '' AS CostCentre " &
            " 				, TFEntityDepartments.Name AS ActualVessel " &
            " FROM TFEntities " &
            " LEFT JOIN TFEntityDepartments ON TFEntityDepartments.EntityID=TFEntities.Id  " &
            " 		  AND TFEntityDepartments.InUse=1 " &
            " 		  AND (SELECT COUNT(*) FROM #TempTable1 WHERE TFEntityDepartments.Name = #TempTable1.Vessel) = 0 " &
            " WHERE TFEntityDepartments.Name IS NOT NULL AND TFEntities.IsActive = 1 "
                If byGroup Then
                    pCommandText &= " 		AND TFEntities.Id IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID=" & Id & ") "
                Else
                    pCommandText &= " 		AND TFEntities.Id =" & Id & " "
                End If
                pCommandText &= " ORDER BY   ActualVessel,CostCentre, Code    " &
            " If(OBJECT_ID('tempdb..#TempTable') Is Not Null)   " &
            " Begin       " &
            " Drop Table #TempTable   " &
            " End   " &
            " If(OBJECT_ID('tempdb..#TempTable1') Is Not Null)   " &
            " Begin       " &
            " Drop Table #TempTable1   " &
            " End  "


                .CommandText = pCommandText
                pobjReader = .ExecuteReader
            End With

            Dim pId As Integer = 0
            MyBase.Clear()
            With pobjReader
                Do While .Read
                    pId = pId + 1
                    pobjClass = New CostCentreLookupItem(pId, CStr(.Item("Code")), CStr(.Item("Remarks")), CStr(.Item("Name")), CStr(.Item("Logo")), CStr(.Item("ActualVessel")), CStr(.Item("CostCentre")))
                    MyBase.Add(pobjClass.Id, pobjClass)
                Loop
                .Close()
            End With
            pobjConn.Close()
        End If

    End Sub
End Class
