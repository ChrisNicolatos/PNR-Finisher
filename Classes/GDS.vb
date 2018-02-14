Namespace GDS
    Public Class GDSItem
        Private Structure ClassProps
            Dim Id As Integer
            Dim GDSName As String
        End Structure
        Private mudtProps As ClassProps

        Public ReadOnly Property Id As Integer
            Get
                Id = mudtProps.Id
            End Get
        End Property
        Public ReadOnly Property GDSName As String
            Get
                GDSName = mudtProps.GDSName
            End Get
        End Property
        Friend Sub SetValues(ByVal pId As Integer, ByVal pGDSName As String)
            With mudtProps
                .Id = pId
                .GDSName = pGDSName
            End With
        End Sub
    End Class
    Public Class GDSCollection
        Inherits Collections.Generic.Dictionary(Of String, GDSItem)
        Public Sub Load()

            Dim pobjConn As New SqlClient.SqlConnection(ConnectionStringPNR)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader
            Dim pobjClass As GDSItem
            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .CommandText = " SELECT  pfrGDSId, pfrGDSName " &
                               " FROM AmadeusReports.dbo.PNRFinisherGDS " &
                               " Order By pfrGDSId"
                pobjReader = .ExecuteReader
            End With

            MyBase.Clear()

            With pobjReader
                Do While .Read
                    pobjClass = New GDSItem
                    pobjClass.SetValues(.Item("pfrGDSId"), .Item("pfrGDSName"))
                    MyBase.Add(pobjClass.ID, pobjClass)
                Loop
                .Close()
            End With
            pobjConn.Close()

        End Sub
    End Class
End Namespace

