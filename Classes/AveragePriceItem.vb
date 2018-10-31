Option Strict On
Option Explicit On
Public Class AveragePriceItem
    Private Structure ClassProps
        Dim Id As Integer
        Dim CustomerName As String
        Dim Airline As String
        Dim ClassOfService As String
        Dim TicketCount As Integer
        Dim AveragePrice As Decimal
        Dim CustomerNameNull As Boolean
        Dim AirlineNull As Boolean
        Dim ClassOfServiceNull As Boolean
    End Structure
    Private mudtProps As ClassProps
    Public ReadOnly Property Id As Integer
        Get
            Id = mudtProps.Id
        End Get
    End Property
    Public ReadOnly Property CustomerName As String
        Get
            CustomerName = mudtProps.CustomerName
        End Get
    End Property
    Public ReadOnly Property Airline As String
        Get
            Airline = mudtProps.Airline
        End Get
    End Property
    Public ReadOnly Property ClassOfService As String
        Get
            ClassOfService = mudtProps.ClassOfService
        End Get
    End Property
    Public ReadOnly Property CustomerNameNull As Boolean
        Get
            CustomerNameNull = mudtProps.CustomerNameNull
        End Get
    End Property
    Public ReadOnly Property AirlineNull As Boolean
        Get
            AirlineNull = mudtProps.AirlineNull
        End Get
    End Property
    Public ReadOnly Property ClassOfServiceNull As Boolean
        Get
            ClassOfServiceNull = mudtProps.ClassOfServiceNull
        End Get
    End Property
    Public ReadOnly Property TicketCount As Integer
        Get
            TicketCount = mudtProps.TicketCount
        End Get
    End Property
    Public ReadOnly Property AveragePrice As Decimal
        Get
            AveragePrice = mudtProps.AveragePrice
        End Get
    End Property
    Friend Sub SetValues(ByVal pId As Integer, ByVal pCustomerName As Object, ByVal pAirline As Object, ByVal pClassOfService As Object, ByVal pTicketCount As Object, ByVal pAveragePrice As Object)
        With mudtProps
            .Id = pId
            If IsDBNull(pCustomerName) Then
                .CustomerNameNull = True
                .CustomerName = ""
            Else
                .CustomerNameNull = False
                .CustomerName = CStr(pCustomerName)
            End If
            If IsDBNull(pAirline) Then
                .AirlineNull = True
                .Airline = ""
            Else
                .AirlineNull = False
                .Airline = CStr(pAirline)
            End If
            If IsDBNull(pClassOfService) Then
                .ClassOfServiceNull = True
                .ClassOfService = ""
            Else
                .ClassOfServiceNull = False
                .ClassOfService = CStr(pClassOfService)
            End If
            If IsDBNull(pTicketCount) Then
                .TicketCount = 0
            Else
                .TicketCount = CInt(pTicketCount)
            End If
            If IsDBNull(pAveragePrice) Then
                .AveragePrice = 0
            Else
                .AveragePrice = CDec(pAveragePrice)
            End If
        End With
    End Sub

End Class
