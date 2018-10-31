Public Class CustomPropertiesCollection
    Inherits Collections.Generic.Dictionary(Of Integer, CustomPropertiesItem)

    Private Const MyXMLString As String = "<?xml version='1.0' encoding='utf-8'?><LookUpValues xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema'><CustomPropertyLookupValue Description='Crew' Value='Crew' IsDefault='false' /><CustomPropertyLookupValue Description='Technical' Value='Technical' IsDefault='false' /><CustomPropertyLookupValue Description='Marine' Value='Marine' IsDefault='false' /><CustomPropertyLookupValue Description='HSQE' Value='HSQE' IsDefault='false' /><CustomPropertyLookupValue Description='Finance' Value='Finance' IsDefault='false' /></LookUpValues>"
    Private mflgBookedBy As Boolean
    Private mflgDepartment As Boolean
    Private mflgReasonForTravel As Boolean
    Private mflgCostCentre As Boolean

    Public Sub Load(ByVal pEntityID As Integer)

        If MySettings.PCCBackOffice = 1 Then
            Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringACC) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader
            Dim pobjClass As CustomPropertiesItem

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .CommandText = " SELECT Id " &
                           "       ,CustomPropertyID " &
                           "       ,LookUpValues " &
                           "       ,LimitToLookUp " &
                           "       ,Label " &
                           "       ,TFEntityID " &
                           "       ,LTRequiredTypeID " &
                           "   FROM TravelForceCosmos.dbo.ClientCustomProperties " &
                           "   WHERE TFEntityID = '" & pEntityID & "'   " &
                           "   AND IsDisabled = 0"

                pobjReader = .ExecuteReader
            End With

            mflgBookedBy = False
            mflgDepartment = False
            mflgReasonForTravel = False
            mflgCostCentre = False

            With pobjReader
                Do While .Read
                    pobjClass = New CustomPropertiesItem
                    pobjClass.SetValues(CInt(.Item("Id")), CType(.Item("CustomPropertyID"), Utilities.EnumCustomPropertyID), CStr(.Item("LookUpValues")), CBool(.Item("LimitToLookUp")), CType(.Item("LTRequiredTypeID"), Utilities.CustomPropertyRequiredType), CStr(.Item("Label")), CInt(.Item("TFEntityID")))
                    MyBase.Add(pobjClass.ID, pobjClass)
                    If pobjClass.CustomPropertyID = Utilities.EnumCustomPropertyID.BookedBy Then
                        mflgBookedBy = True
                    ElseIf pobjClass.CustomPropertyID = Utilities.EnumCustomPropertyID.Department Then
                        mflgDepartment = True
                    ElseIf pobjClass.CustomPropertyID = Utilities.EnumCustomPropertyID.ReasonFortravel Then
                        mflgReasonForTravel = True
                    ElseIf pobjClass.CustomPropertyID = Utilities.EnumCustomPropertyID.CostCentre Then
                        mflgCostCentre = True
                    End If
                Loop
                .Close()
            End With
            pobjConn.Close()
        Else
            Dim pobjClass As CustomPropertiesItem
            pobjClass = New CustomPropertiesItem
            pobjClass.SetValues(1, Utilities.EnumCustomPropertyID.BookedBy, "", True, Utilities.CustomPropertyRequiredType.PropertyReqToSave, "BookedBy", pEntityID)
            If pobjClass.ValuesCount > 0 Then
                MyBase.Add(pobjClass.ID, pobjClass)
                mflgBookedBy = True
            Else
                mflgBookedBy = False
            End If
            pobjClass = New CustomPropertiesItem
            pobjClass.SetValues(2, Utilities.EnumCustomPropertyID.ReasonFortravel, "", True, Utilities.CustomPropertyRequiredType.PropertyReqToSave, "ReasonFortravel", pEntityID)
            If pobjClass.ValuesCount > 0 Then
                MyBase.Add(pobjClass.ID, pobjClass)
                mflgReasonForTravel = True
            Else
                mflgReasonForTravel = False
            End If

            mflgDepartment = False
            mflgCostCentre = False
        End If

    End Sub
    Public ReadOnly Property BookedBy As Boolean
        Get
            Return mflgBookedBy
        End Get
    End Property
    Public ReadOnly Property Department As Boolean
        Get
            Return mflgDepartment
        End Get
    End Property
    Public ReadOnly Property ReasonForTravel As Boolean
        Get
            Return mflgReasonForTravel
        End Get
    End Property
    Public ReadOnly Property CostCentre As Boolean
        Get
            Return mflgCostCentre
        End Get
    End Property

End Class
