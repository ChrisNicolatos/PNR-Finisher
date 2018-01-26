Public Class AmadeusUser

    Private Structure UserProps
        Dim ID As Integer
        Dim PCC As String
        Dim User As String
        Dim AgentQueue As String
        Dim AgentOPQueue As String
        Dim AgentName As String
        Dim AgentEmail As String
        Dim AirportName As Integer
        Dim AirlineLocator As Boolean
        Dim ClassOfService As Boolean
        Dim BanElectricalEquipment As Boolean
        Dim BrazilText As Boolean
        Dim USAText As Boolean
        Dim Tickets As Boolean
        Dim PaxSegPerTkt As Boolean
        Dim ShowStopovers As Boolean
        Dim ShowTerminal As Boolean
        Dim FlyingTime As Boolean
        Dim CostCentre As Boolean
        Dim Seating As Boolean
        Dim Vessel As Boolean
        Dim PlainFormat As Boolean
    End Structure

    Private WithEvents mobjSession As k1aHostToolKit.HostSession
    Private mstrResponse As String
    Private mstrPCC As String = ""
    Private mstrUser As String = ""

    Public Sub New()

        Dim Sessions As k1aHostToolKit.HostSessions

        Try

            ' To be able to retrieve the PNR that have been created we need to link our '
            ' application to the current session of the FOS                             '
            Sessions = New k1aHostToolKit.HostSessions

            If Sessions.Count > 0 Then
                ' There is at least one session opened.                    '
                ' We link our application to the active session of the FOS '
                mobjSession = Sessions.UIActiveSession
                mobjSession.Send("JGD/C")
                Dim pLines() As String = mstrResponse.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
                For i As Integer = 0 To pLines.GetUpperBound(0)
                    If pLines(i).Trim.StartsWith("OFFICE") Then
                        mstrPCC = pLines(i).Substring(pLines(i).IndexOf("-") + 1).Trim
                    ElseIf pLines(i).Trim.StartsWith("SIGN ") Then
                        mstrUser = pLines(i).Substring(pLines(i).IndexOf("-") + 1).Trim
                    End If
                Next

            End If
            If mstrPCC = "" Or mstrUser = "" Then
                Throw New Exception("Please start Amadeus and restart the program")
            End If
        Catch ex As Exception
            Throw New Exception("Amadeus Error" & vbCrLf & ex.Message)
        End Try
    End Sub

    Public ReadOnly Property PCC As String
        Get
            PCC = mstrPCC
        End Get
    End Property
    Public ReadOnly Property User As String
        Get
            User = mstrUser
        End Get
    End Property
    Private Sub mobjSession_ReceivedResponse(ByRef newResponse As k1aHostToolKit.CHostResponse) Handles mobjSession.ReceivedResponse

        mstrResponse = newResponse.Text

    End Sub


End Class
