Public Class frmPriceOptimiser
    Private mstrPCC As String
    Private mstrUserID As String
    Public Sub New(ByVal pPCC As String, ByVal pUserId As String)


        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        mstrPCC = pPCC
        mstrUserID = pUserId

    End Sub
End Class