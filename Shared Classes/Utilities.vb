Public NotInheritable Class Utilities

    Public Enum EnumItnFormat
        DefaultFormat = 0
        Plain = 1
        SeaChefs = 2
        SeaChefsWithCode = 3
        Euronav = 4
    End Enum
    Public Enum EnumGDSCode
        Unknown = 0
        Amadeus = 1
        Galileo = 2
    End Enum
    Public Enum EnumCustomPropertyID As Integer
        None = 0
        BookedBy = 1
        Department = 2
        ReasonFortravel = 4
        CostCentre = 5
        Savings = 6
        Losses = 7
        SavingsLossesReason = 8
        TravelDefinition = 9
        VesselCostCentre = 10
        RequisitionNumber = 11
        PassengerID = 12
        OPT = 13
        TRId = 14
    End Enum
    Public Enum EnumTicketDocType
        NONE = 0
        ETKT = 1
        VCHR = 2
        INTR = 3
    End Enum
    Public Enum CustomPropertyRequiredType
        PropertyNone = 0
        PropertyOptional = 613
        PropertyReqToSave = 614
        PropertyReqToInv = 615
    End Enum
    Private Sub New()
    End Sub



End Class

