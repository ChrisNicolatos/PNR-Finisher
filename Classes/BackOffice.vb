﻿Option Strict On
Option Explicit On
Namespace BackOffice
    'Public Class BackOfficeItem
    '    Private Structure ClassProps
    '        Dim Id As Integer
    '        Dim BackOfficeName As String
    '    End Structure
    '    Private mudtProps As ClassProps

    '    Public ReadOnly Property Id As Integer
    '        Get
    '            Id = mudtProps.Id
    '        End Get
    '    End Property
    '    Public ReadOnly Property BackOfficeName As String
    '        Get
    '            BackOfficeName = mudtProps.BackOfficeName
    '        End Get
    '    End Property
    '    Friend Sub SetValues(ByVal pId As Integer, ByVal pBackOfficeName As String)
    '        With mudtProps
    '            .Id = pId
    '            .BackOfficeName = pBackOfficeName
    '        End With
    '    End Sub
    'End Class
    'Public Class BackOfficeCollection
    '    Inherits Collections.Generic.Dictionary(Of Integer, BackOfficeItem)

    '    Public Sub Load()

    '        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR)
    '        Dim pobjComm As New SqlClient.SqlCommand
    '        Dim pobjReader As SqlClient.SqlDataReader
    '        Dim pobjClass As BackOfficeItem
    '        pobjConn.Open()
    '        pobjComm = pobjConn.CreateCommand

    '        With pobjComm
    '            .CommandType = CommandType.Text
    '            .CommandText = " SELECT  pfrBOId, pfrBOName " &
    '                           " FROM AmadeusReports.dbo.PNRFinisherBackOffice " &
    '                           " Order By pfrBOId"
    '            pobjReader = .ExecuteReader
    '        End With

    '        MyBase.Clear()

    '        With pobjReader
    '            Do While .Read
    '                pobjClass = New BackOfficeItem
    '                pobjClass.SetValues(CInt(.Item("pfrBOId")), CStr(.Item("pfrBOName")))
    '                MyBase.Add(pobjClass.Id, pobjClass)
    '            Loop
    '            .Close()
    '        End With
    '        pobjConn.Close()

    '    End Sub
    'End Class
End Namespace

