Public Class CustomPropertiesItem
    Private Structure ClassProps
        Dim ID As Integer
        Dim CustomPropertyID As Utilities.EnumCustomPropertyID
        Dim LookUpValues As String
        Dim LimitToLookup As Boolean
        Dim RequiredType As Utilities.CustomPropertyRequiredType
        Dim Label As String
        Dim TFEntityID As Long
        Dim Values() As String
    End Structure
    Private mudtProps As ClassProps

    Public ReadOnly Property ID() As Integer
        Get
            Return mudtProps.ID
        End Get
    End Property

    Public ReadOnly Property CustomPropertyID() As Utilities.EnumCustomPropertyID
        Get
            Return mudtProps.CustomPropertyID
        End Get
    End Property

    Public ReadOnly Property LookUpValues() As String
        Get
            Return mudtProps.LookUpValues
        End Get
    End Property

    Public ReadOnly Property LimitToLookup() As Boolean
        Get
            Return mudtProps.LimitToLookup
        End Get
    End Property
    Public ReadOnly Property RequiredType As Utilities.CustomPropertyRequiredType
        Get
            Return mudtProps.RequiredType
        End Get
    End Property
    Public ReadOnly Property Label() As String
        Get
            Return mudtProps.Label
        End Get
    End Property

    Public ReadOnly Property TFEntityID() As Long
        Get
            Return mudtProps.TFEntityID
        End Get
    End Property

    Public ReadOnly Property ValuesCount As Integer
        Get
            Return mudtProps.Values.Length
        End Get
    End Property

    Public ReadOnly Property Value(ByVal Index As Integer) As String
        Get
            If Index >= 0 And Index <= mudtProps.Values.GetUpperBound(0) Then
                Return mudtProps.Values(Index)
            Else
                Throw New Exception("Index out of bounds")
            End If
        End Get
    End Property

    Friend Sub SetValues(ByVal pID As Integer, ByVal pCustomPropertyID As Utilities.EnumCustomPropertyID, ByVal pLookUpValues As String, ByVal pLimitToLookup As Boolean, ByVal pRequiredType As Utilities.CustomPropertyRequiredType, ByVal pLabel As String, ByVal pTFEntityID As Integer)
        With mudtProps
            .ID = pID
            .CustomPropertyID = pCustomPropertyID
            .LookUpValues = pLookUpValues
            .RequiredType = pRequiredType
            .LimitToLookup = pLimitToLookup
            .Label = pLabel
            .TFEntityID = pTFEntityID
            ReDim .Values(0)
            If .LookUpValues.IndexOf("<") >= 0 Or MySettings.PCCBackOffice = 2 Then
                ReadXML(pCustomPropertyID, pTFEntityID)
            Else
                ReadLookUpValues()
            End If
        End With
    End Sub

    Private Sub ReadXML(ByVal pCustomPropertyID As Integer, ByVal pTfEntityID As Integer)

        Dim pobjXMLValues As New CustomPropertiesXMLValues
        pobjXMLValues.ReadValues(pCustomPropertyID, pTfEntityID)
        mudtProps.Values = pobjXMLValues.ToArray

    End Sub

    Private Sub ReadLookUpValues()

        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringACC) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand
        Dim pobjReader As SqlClient.SqlDataReader

        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand

        With pobjComm
            .CommandType = CommandType.Text
            .CommandText = "SELECT Value " &
                           " FROM TravelForceCosmos.dbo.CustomPropertyValues " &
                           " WHERE CustomPropertyID = " & mudtProps.CustomPropertyID & " And TFEntityID = " & mudtProps.TFEntityID &
                           " GROUP BY Value " &
                           " ORDER BY Value"
            pobjReader = .ExecuteReader
        End With
        mudtProps.Values(0) = ""
        With pobjReader
            Dim iCount As Integer = 0
            Do While .Read
                iCount += 1
                ReDim Preserve mudtProps.Values(iCount - 1)
                mudtProps.Values(iCount - 1) = CStr(.Item("Value"))
            Loop
            .Close()
        End With
        pobjConn.Close()
    End Sub

End Class