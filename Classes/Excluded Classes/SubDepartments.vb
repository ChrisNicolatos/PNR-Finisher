Option Strict On
Option Explicit On
Namespace SubDepartments
    'Public Class SubDepartmentItem
    '    Private Structure ClassProps
    '        Dim ID As Integer
    '        Dim Code As String
    '        Dim Name As String
    '    End Structure
    '    Private mudtProps As ClassProps
    '    Public Overrides Function ToString() As String
    '        With mudtProps
    '            Return .Code & " " & .Name
    '        End With
    '    End Function
    '    Public ReadOnly Property ID() As Integer
    '        Get
    '            ID = mudtProps.ID
    '        End Get
    '    End Property
    '    Public ReadOnly Property Code() As String
    '        Get
    '            Code = mudtProps.Code
    '        End Get
    '    End Property
    '    Public ReadOnly Property Name() As String
    '        Get
    '            Name = mudtProps.Name
    '        End Get
    '    End Property
    '    Friend Sub SetValues(ByVal pID As Integer, ByVal pCode As String, ByVal pName As String)
    '        With mudtProps
    '            .ID = pID
    '            .Code = pCode
    '            .Name = pName
    '        End With
    '    End Sub

    '    Public Sub Load(ByVal pSubID As Integer)
    '        If MySettings.PCCBackOffice = 1 Then

    '            Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringACC) ' ActiveConnection)
    '            Dim pobjComm As New SqlClient.SqlCommand
    '            Dim pobjReader As SqlClient.SqlDataReader

    '            pobjConn.Open()
    '            pobjComm = pobjConn.CreateCommand

    '            With pobjComm
    '                .CommandType = CommandType.Text
    '                .CommandText = " SELECT [Id] " &
    '                           " ,[Code] " &
    '                           " ,[Name] " &
    '                           " FROM [TravelForceCosmos].[dbo].[TFEntitySubdepartments] " &
    '                           " WHERE ID = " & pSubID & "  " &
    '                           " ORDER BY Name "


    '                pobjReader = .ExecuteReader
    '            End With

    '            With pobjReader
    '                Do While .Read
    '                    SetValues(CInt(.Item("Id")), CStr(.Item("Code")), CStr(.Item("Name")))
    '                Loop
    '                .Close()
    '            End With
    '            pobjConn.Close()
    '        End If
    '    End Sub
    'End Class
    'Friend Class SubDepartmentCollection

    '    Inherits Collections.Generic.Dictionary(Of Integer, SubDepartmentItem)
    '    Private mlngEntityID As Long

    '    Public Sub Load(ByVal pEntityID As Long)
    '        If MySettings.PCCBackOffice = 1 Then

    '            Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringACC) ' ActiveConnection)
    '            Dim pobjComm As New SqlClient.SqlCommand
    '            Dim pobjReader As SqlClient.SqlDataReader
    '            Dim pobjClass As SubDepartmentItem

    '            mlngEntityID = pEntityID

    '            pobjConn.Open()
    '            pobjComm = pobjConn.CreateCommand

    '            With pobjComm
    '                .CommandType = CommandType.Text
    '                .CommandText = " SELECT [Id] " &
    '                           " ,[Code] " &
    '                           " ,[Name] " &
    '                           " FROM [TravelForceCosmos].[dbo].[TFEntitySubdepartments] " &
    '                           " WHERE EntityID = " & mlngEntityID & "  AND InUse = 1 " &
    '                           " ORDER BY Name "


    '                pobjReader = .ExecuteReader
    '            End With

    '            With pobjReader
    '                Do While .Read
    '                    pobjClass = New SubDepartmentItem
    '                    pobjClass.SetValues(CInt(.Item("Id")), CStr(.Item("Code")), CStr(.Item("Name")))
    '                    MyBase.Add(pobjClass.ID, pobjClass)
    '                Loop
    '                .Close()
    '            End With
    '            pobjConn.Close()
    '        End If
    '    End Sub

    'End Class
End Namespace
