Option Strict On
Option Explicit On
Public Class AirlinePointsItem
    Private Structure ClassProps
        Dim CustomerID As Integer
        Dim CustomerCode As String
        Dim CustomerName As String
        Dim AirlineCode As String
        Dim AirlineName As String
        Dim PointsCommand As String
    End Structure
    Private mudtProps As ClassProps

    Public ReadOnly Property CustomerID() As Long
        Get
            CustomerID = mudtProps.CustomerID
        End Get
    End Property
    Public ReadOnly Property CustomerCode() As String
        Get
            CustomerCode = mudtProps.CustomerCode
        End Get
    End Property
    Public ReadOnly Property CustomerName() As String
        Get
            CustomerName = mudtProps.CustomerName
        End Get
    End Property
    Public ReadOnly Property AirlineCode() As String
        Get
            AirlineCode = mudtProps.AirlineCode
        End Get
    End Property
    Public ReadOnly Property AirlineName() As String
        Get
            AirlineName = mudtProps.AirlineName
        End Get
    End Property
    Public ReadOnly Property PointsCommand() As String
        Get
            PointsCommand = mudtProps.PointsCommand
        End Get
    End Property

    Friend Sub SetValues(ByVal pCustID As Integer, ByVal pCustCode As String, ByVal pCustName As String,
                             ByVal pAirlineCode As String, ByVal pAirlineName As String, ByVal pPointsCommand As String)
        With mudtProps
            .CustomerID = pCustID
            .CustomerCode = pCustCode
            .CustomerName = pCustName
            .AirlineCode = pAirlineCode
            .AirlineName = pAirlineName
            .PointsCommand = pPointsCommand
        End With
    End Sub
    Friend Sub SetValues(ByVal pPointsCommand As String)
        With mudtProps
            .CustomerID = 0
            .CustomerCode = ""
            .CustomerName = ""
            .AirlineCode = ""
            .AirlineName = ""
            .PointsCommand = pPointsCommand
        End With
    End Sub

End Class
