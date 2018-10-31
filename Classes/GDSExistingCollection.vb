Public Class GDSExistingCollection
    Private mobjOpenSegment As New GDSExistingItem
    Private mobjPhoneElement As New GDSExistingItem
    Private mobjAgentElement As New GDSExistingItem
    Private mobjEmailElement As New GDSExistingItem
    Private mobjTicketElement As New GDSExistingItem
    Private mobjOptionQueueElement As New GDSExistingItem
    Private mobjAOH As New GDSExistingItem
    Private mobjAgentID As New GDSExistingItem
    Private mobjSavingsElement As New GDSExistingItem
    Private mobjLossElement As New GDSExistingItem

    Private mobjCustomerCodeAI As New GDSExistingItem
    Private mobjCustomerCode As New GDSExistingItem
    Private mobjCustomerName As New GDSExistingItem
    Private mobjSubDepartmentCode As New GDSExistingItem
    Private mobjSubDepartmentName As New GDSExistingItem
    Private mobjCRMCode As New GDSExistingItem
    Private mobjCRMName As New GDSExistingItem
    Private mobjVesselName As New GDSExistingItem
    Private mobjVesselFlag As New GDSExistingItem
    Private mobjVesselOSI As New GDSExistingItem
    Private mobjReference As New GDSExistingItem
    Private mobjBookedBy As New GDSExistingItem
    Private mobjDepartment As New GDSExistingItem
    Private mobjReasonForTravel As New GDSExistingItem
    Private mobjCostCentre As New GDSExistingItem
    Private mobjTRId As New GDSExistingItem

    Public ReadOnly Property OpenSegment As GDSExistingItem
        Get
            OpenSegment = mobjOpenSegment
        End Get
    End Property
    Public ReadOnly Property PhoneElement As GDSExistingItem
        Get
            PhoneElement = mobjPhoneElement
        End Get
    End Property
    Public ReadOnly Property AgentElement As GDSExistingItem
        Get
            AgentElement = mobjAgentElement
        End Get
    End Property
    Public ReadOnly Property EmailElement As GDSExistingItem
        Get
            EmailElement = mobjEmailElement
        End Get
    End Property
    Public ReadOnly Property TicketElement As GDSExistingItem
        Get
            TicketElement = mobjTicketElement
        End Get
    End Property
    Public ReadOnly Property OptionQueueElement As GDSExistingItem
        Get
            OptionQueueElement = mobjOptionQueueElement
        End Get
    End Property
    Public ReadOnly Property AOH As GDSExistingItem
        Get
            AOH = mobjAOH
        End Get
    End Property
    Public ReadOnly Property AgentID As GDSExistingItem
        Get
            AgentID = mobjAgentID
        End Get
    End Property
    Public ReadOnly Property SavingsElement As GDSExistingItem
        Get
            SavingsElement = mobjSavingsElement
        End Get
    End Property
    Public ReadOnly Property LossElement As GDSExistingItem
        Get
            LossElement = mobjLossElement
        End Get
    End Property
    Public ReadOnly Property CustomerCodeAI As GDSExistingItem
        Get
            Return mobjCustomerCodeAI
        End Get
    End Property
    Public ReadOnly Property CustomerCode As GDSExistingItem
        Get
            CustomerCode = mobjCustomerCode
        End Get
    End Property
    Public ReadOnly Property CustomerName As GDSExistingItem
        Get
            CustomerName = mobjCustomerName
        End Get
    End Property
    Public ReadOnly Property SubDepartmentCode As GDSExistingItem
        Get
            SubDepartmentCode = mobjSubDepartmentCode
        End Get
    End Property
    Public ReadOnly Property SubDepartmentName As GDSExistingItem
        Get
            SubDepartmentName = mobjSubDepartmentName
        End Get
    End Property
    Public ReadOnly Property CRMCode As GDSExistingItem
        Get
            CRMCode = mobjCRMCode
        End Get
    End Property
    Public ReadOnly Property CRMName As GDSExistingItem
        Get
            CRMName = mobjCRMName
        End Get
    End Property
    Public ReadOnly Property VesselName As GDSExistingItem
        Get
            VesselName = mobjVesselName
        End Get
    End Property
    Public ReadOnly Property VesselFlag As GDSExistingItem
        Get
            VesselFlag = mobjVesselFlag
        End Get
    End Property
    Public ReadOnly Property VesselOSI As GDSExistingItem
        Get
            VesselOSI = mobjVesselOSI
        End Get
    End Property
    Public ReadOnly Property Reference As GDSExistingItem
        Get
            Reference = mobjReference
        End Get
    End Property
    Public ReadOnly Property BookedBy As GDSExistingItem
        Get
            BookedBy = mobjBookedBy
        End Get
    End Property
    Public ReadOnly Property Department As GDSExistingItem
        Get
            Department = mobjDepartment
        End Get
    End Property
    Public ReadOnly Property ReasonForTravel As GDSExistingItem
        Get
            ReasonForTravel = mobjReasonForTravel
        End Get
    End Property
    Public ReadOnly Property CostCentre As GDSExistingItem
        Get
            CostCentre = mobjCostCentre
        End Get
    End Property
    Public ReadOnly Property TRId As GDSExistingItem
        Get
            TRId = mobjTRId
        End Get
    End Property
    Public Sub Clear()
        mobjOpenSegment.Clear()
        mobjPhoneElement.Clear()
        mobjAgentElement.Clear()
        mobjEmailElement.Clear()
        mobjTicketElement.Clear()
        mobjOptionQueueElement.Clear()
        mobjAOH.Clear()
        mobjAgentID.Clear()
        mobjSavingsElement.Clear()
        mobjLossElement.Clear()

        mobjCustomerCodeAI.Clear()
        mobjCustomerCode.Clear()
        mobjCustomerName.Clear()
        mobjSubDepartmentCode.Clear()
        mobjSubDepartmentName.Clear()
        mobjCRMCode.Clear()
        mobjCRMName.Clear()
        mobjVesselName.Clear()
        mobjVesselFlag.Clear()
        mobjVesselOSI.Clear()
        mobjReference.Clear()
        mobjBookedBy.Clear()
        mobjDepartment.Clear()
        mobjReasonForTravel.Clear()
        mobjCostCentre.Clear()
    End Sub

End Class