Public Class VesselCollection
    Inherits Collections.Generic.Dictionary(Of String, VesselItem)
    Private mlngEntityID As Integer

    Public Sub Load(ByVal pEntityID As Integer)
        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringACC) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader
        Dim pobjClass As VesselItem

        mlngEntityID = pEntityID

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = PrepareVesselSelectCommand(mlngEntityID)
            pobjReader = .ExecuteReader
        End With

        With pobjReader
            Do While .Read
                pobjClass = New VesselItem
                pobjClass.SetValues(CStr(.Item("Name")), CStr(.Item("Flag")))
                If pobjClass.ToString <> "" And Not MyBase.ContainsKey(pobjClass.ToString) Then
                    MyBase.Add(pobjClass.ToString, pobjClass)
                End If
            Loop
            .Close()
        End With
        pobjConn.Close()
    End Sub
    Private Function PrepareVesselSelectCommand(ByVal pEntityID As Integer) As String

        Select Case MySettings.PCCBackOffice
            Case 1
                PrepareVesselSelectCommand = " SELECT DISTINCT " &
                           " RTRIM(LTRIM(Name)) AS Name " &
                           " ,ISNULL(RTRIM(LTRIM(Flag)), '') AS Flag " &
                           " FROM [TravelForceCosmos].[dbo].[TFEntityDepartments] " &
                           " WHERE InUse = 1 " &
                           " AND RTRIM(LTRIM(Name)) <> '' AND EntityID = " & pEntityID & " " &
                           " ORDER BY Name "
            Case 2
                PrepareVesselSelectCommand = " SELECT [Child_Value] AS Name " &
                                             ",'' AS Flag " &
                                             "  FROM [Disco_Instone_EU].[dbo].[Costcen] " &
                                             "  WHERE Child_Value IS NOT NULL AND Child_Name = 'CC1' AND CostCen.Account_id =  " & pEntityID
            Case Else
                PrepareVesselSelectCommand = ""
        End Select

    End Function
End Class