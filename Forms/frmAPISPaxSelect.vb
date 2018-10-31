Friend Class frmAPISPaxSelect
    Dim mobjPaxApis As New ApisPaxCollection
    Dim mobjPaxItem As ApisPaxItem

    Public Sub New(ByVal PaxId As Short, ByVal PaxSurname As String, ByVal PaxFirstName As String, ByVal PaxCollection As ApisPaxCollection)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        txtNumber.Text = PaxId
        txtSurname.Text = PaxSurname
        txtFirstName.Text = PaxFirstName
        mobjPaxApis = PaxCollection


    End Sub
    Private Sub frmAPISPaxSelect_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        UtilitiesAPIS.APISPrepareGrid(dgvPax)
        dgvPax.Rows.Clear()

        For Each pPax As ApisPaxItem In mobjPaxApis.Values
            UtilitiesAPIS.APISAddRow(dgvPax, pPax)
        Next
        EnableSelection()
    End Sub
    Public ReadOnly Property SelectedPassenger As ApisPaxItem
        Get
            SelectedPassenger = mobjPaxItem
        End Get
    End Property
    Private Sub dgvPax_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPax.CellContentClick, dgvPax.RowEnter
        Try
            Dim Id As Integer = e.RowIndex
            mobjPaxItem = mobjPaxApis.Item(dgvPax.Rows(Id).Cells(0).Value)
        Catch ex As Exception
            mobjPaxItem = Nothing
        End Try
        EnableSelection()
    End Sub
    Private Sub EnableSelection()
        cmdSelect.Enabled = Not IsNothing(mobjPaxItem)
    End Sub

    Private Sub cmdSelect_Click(sender As Object, e As EventArgs) Handles cmdSelect.Click
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class