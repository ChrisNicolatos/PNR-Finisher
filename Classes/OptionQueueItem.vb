Public Class OptionQueueItem
    Private Structure ClassProps
        Dim ElementNo As Short
        Dim PCC As String
        Dim ActionDateTime As Date
        Dim QueueNumber As String
        Dim Remark As String
    End Structure
    Private mudtProps As ClassProps
    Public ReadOnly Property ElementNo As Short
        Get
            Return mudtProps.ElementNo
        End Get
    End Property
    Public ReadOnly Property PCC As String
        Get
            Return mudtProps.PCC
        End Get
    End Property
    Public ReadOnly Property ActionDateTime As Date
        Get
            Return mudtProps.ActionDateTime
        End Get
    End Property
    Public ReadOnly Property QueueNumber As String
        Get
            Return mudtProps.QueueNumber
        End Get
    End Property
    Public ReadOnly Property Remark As String
        Get
            Return mudtProps.Remark
        End Get
    End Property
    Friend Sub SetValues(ByVal pElementNo As Short, ByVal pPCC As String, ByVal pActionDateTime As Date, ByVal pQueueNumber As String, ByVal pRemark As String)
        With mudtProps
            .ElementNo = pElementNo
            .PCC = pPCC
            .ActionDateTime = pActionDateTime
            .QueueNumber = pQueueNumber
            .Remark = pRemark
        End With
    End Sub
End Class

