Public Class OSMPaxItem
    Private Structure ClassProps
        Dim Id As Integer
        Dim LastName As String
        Dim FirstName As String
        Dim Nationality As String
        Dim Text As String
        Dim TextFullDetails As String
        Dim JoinerLeaver As String
    End Structure
    Dim mudtProps As ClassProps
    Public ReadOnly Property Id As Integer
        Get
            Id = mudtProps.Id
        End Get
    End Property
    Public ReadOnly Property LastName As String
        Get
            LastName = mudtProps.LastName
        End Get
    End Property
    Public ReadOnly Property FirstName As String
        Get
            FirstName = mudtProps.FirstName
        End Get
    End Property
    Public ReadOnly Property Nationality As String
        Get
            Nationality = mudtProps.Nationality
        End Get
    End Property
    Public ReadOnly Property Text As String
        Get
            Text = mudtProps.Text
        End Get
    End Property
    Public ReadOnly Property TextFullDetails As String
        Get
            Return mudtProps.TextFullDetails
        End Get
    End Property
    Public Property JoinerLeaver As String
        Get
            JoinerLeaver = mudtProps.JoinerLeaver
        End Get
        Set(value As String)
            mudtProps.JoinerLeaver = value
        End Set
    End Property
    Friend Sub SetData(ByVal pId As Integer, ByVal pJoiner As String, ByVal pText As String)

        mudtProps.Id = pId
        mudtProps.TextFullDetails = pText
        mudtProps.Text = ""
        Dim pLines() As String = pText.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
        For i As Integer = 0 To pLines.GetUpperBound(0)
            If pLines(i).ToUpper.StartsWith("LAST NAME") Then
                mudtProps.Text &= pLines(i) & vbCrLf
                Dim pN() As String = pLines(i).Split(":"c)
                mudtProps.LastName = pN(1).Trim
            ElseIf pLines(i).ToUpper.StartsWith("FIRST AND MIDDLE NAMES") Then
                mudtProps.Text &= pLines(i) & vbCrLf
                Dim pN() As String = pLines(i).Split(":"c)
                mudtProps.FirstName = pN(1).Trim
            ElseIf pLines(i).ToUpper.StartsWith("NATIONALITY") Then
                Dim pN() As String = pLines(i).Split(":"c)
                mudtProps.Nationality = pN(1).Trim
            ElseIf pLines(i).ToUpper.StartsWith("POSITION") Then
                mudtProps.Text &= pLines(i) & vbCrLf & vbCrLf
            End If
        Next
        mudtProps.JoinerLeaver = pJoiner
    End Sub
End Class