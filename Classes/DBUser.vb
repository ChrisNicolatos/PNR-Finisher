Public Class DBUser
    Public Event UserValid(ByVal isValid As Boolean, ByVal isUserNameValid As Boolean, ByVal isEmailValid As Boolean, ByVal isQueueNumberValid As Boolean, ByVal isOPQueueValid As Boolean)
    Private Structure ClassProps
        Dim GDS As Utilities.EnumGDSCode
        Dim PCC As String
        Dim UserID As String
        Dim Username As String
        Dim Email As String
        Dim QueueNumber As String
        Dim OPQueue As String
        Dim IsNew As Boolean
        Dim IsDirty As Boolean
        Dim IsValid As Boolean
        Dim IsUserNameValid As Boolean
        Dim IsEmailValid As Boolean
        Dim IsQueueNumberValid As Boolean
        Dim IsOPQueueValid As Boolean
    End Structure
    Private mudtProps As ClassProps
    Public Sub New(ByVal pGDS As Utilities.EnumGDSCode, ByVal pPCC As String, ByVal pUserId As String)
        With mudtProps
            .GDS = pGDS
            .PCC = pPCC
            .UserID = pUserId
            .Username = ""
            .Email = ""
            .QueueNumber = ""
            .OPQueue = ""
            .IsNew = True
            .IsDirty = False
        End With
        CheckValid()
    End Sub
    Public ReadOnly Property GDS As Utilities.EnumGDSCode
        Get
            GDS = mudtProps.GDS
        End Get
    End Property
    Public ReadOnly Property PCC As String
        Get
            PCC = mudtProps.PCC
        End Get
    End Property
    Public ReadOnly Property UserID As String
        Get
            UserID = mudtProps.UserID
        End Get
    End Property
    Public ReadOnly Property isValid As Boolean
        Get
            isValid = mudtProps.IsValid
        End Get
    End Property
    Public ReadOnly Property isUserNameValid As Boolean
        Get
            isUserNameValid = mudtProps.IsUserNameValid
        End Get
    End Property
    Public ReadOnly Property isEmailValid As Boolean
        Get
            isEmailValid = mudtProps.IsEmailValid
        End Get
    End Property
    Public ReadOnly Property isQueueNumberValid As Boolean
        Get
            isQueueNumberValid = mudtProps.IsQueueNumberValid
        End Get
    End Property
    Public ReadOnly Property isOPQueueValid As Boolean
        Get
            isOPQueueValid = mudtProps.IsOPQueueValid
        End Get
    End Property
    Public Property Username As String
        Get
            Username = mudtProps.Username
        End Get
        Set(value As String)
            mudtProps.Username = value.Trim
            CheckValid()
        End Set
    End Property
    Public Property Email As String
        Get
            Email = mudtProps.Email
        End Get
        Set(value As String)
            mudtProps.Email = value.Trim
            CheckValid()
        End Set
    End Property
    Public Property QueueNumber As String
        Get
            QueueNumber = mudtProps.QueueNumber
        End Get
        Set(value As String)
            mudtProps.QueueNumber = value.Trim
            CheckValid()
        End Set
    End Property
    Public Property OPQueue As String
        Get
            OPQueue = mudtProps.OPQueue
        End Get
        Set(value As String)
            mudtProps.OPQueue = value.Trim
            CheckValid()
        End Set
    End Property
    Private Sub CheckValid()
        With mudtProps
            .IsUserNameValid = (.Username <> "")
            .IsEmailValid = System.Text.RegularExpressions.Regex.IsMatch(.Email, "(?i)^(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+\/=\?\^`{}|~\w])*)(?<=[0-9a-z])@)(([0-9a-z][-0-9a-z]*[0-9a-z]*\.)+[a-z0-9][-a-z0-9]{0,22}[a-z0-9])$")
            .IsQueueNumberValid = False
            .IsOPQueueValid = False
            If .GDS = Utilities.EnumGDSCode.Amadeus Then
                .IsQueueNumberValid = System.Text.RegularExpressions.Regex.IsMatch(.QueueNumber, "(?i)^[0-9]{1,2}(C[0-9]{1}|C[0-9]{1,2}|C[01][0-9][0-9]|C2[0-4][0-9]|C25[0-5])?$")
                .IsOPQueueValid = System.Text.RegularExpressions.Regex.IsMatch(.OPQueue, "(?i)^[0-9]{1,2}(C[0-9]{1}|C[0-9]{1,2}|C[01][0-9][0-9]|C2[0-4][0-9]|C25[0-5])?$")
            ElseIf .GDS = Utilities.EnumGDSCode.Galileo Then
                .IsQueueNumberValid = System.Text.RegularExpressions.Regex.IsMatch(.QueueNumber, "(?i)^[0-9]{1,2}(\*C[0-9]{1}|\*C[0-9]{1,2}|\*C[01][0-9][0-9]|\*C2[0-4][0-9]|\*C25[0-5])?$")
                .IsOPQueueValid = System.Text.RegularExpressions.Regex.IsMatch(.OPQueue, "(?i)^[0-9]{1,2}(\*C[0-9]{1}|\*C[0-9]{1,2}|\*C[01][0-9][0-9]|\*C2[0-4][0-9]|\*C25[0-5])?$")
            End If
            .IsValid = .IsUserNameValid And .IsEmailValid And .IsQueueNumberValid And .IsOPQueueValid
            RaiseEvent UserValid(isValid, isUserNameValid, isEmailValid, isQueueNumberValid, isOPQueueValid)
        End With
    End Sub
    Public Sub Update()

        Dim pErrorMessage As String = ""
        If isValid Then
            Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand

            With pobjComm
                .CommandType = CommandType.Text
                .Parameters.Add("@pfPCC", SqlDbType.NVarChar, 9).Value = mudtProps.PCC
                .Parameters.Add("@pfUser", SqlDbType.NVarChar, 20).Value = mudtProps.UserID
                .Parameters.Add("@pfAgentQueue", SqlDbType.NVarChar, 20).Value = mudtProps.QueueNumber
                .Parameters.Add("@pfAgentOPQueue", SqlDbType.NVarChar, 20).Value = mudtProps.OPQueue
                .Parameters.Add("@pfAgentName", SqlDbType.NVarChar, 50).Value = mudtProps.Username
                .Parameters.Add("@pfAgentEmail", SqlDbType.NVarChar, 255).Value = mudtProps.Email
                .Parameters.Add("@pfAirportName", SqlDbType.Int).Value = 0
                .Parameters.Add("@pfAirlineLocator", SqlDbType.Bit).Value = True
                .Parameters.Add("@pfClassOfService", SqlDbType.Bit).Value = False
                .Parameters.Add("@pfBanElectricalEquipment", SqlDbType.Bit).Value = False
                .Parameters.Add("@pfBrazilText", SqlDbType.Bit).Value = False
                .Parameters.Add("@pfUSAText", SqlDbType.Bit).Value = False
                .Parameters.Add("@pfTickets", SqlDbType.Bit).Value = True
                .Parameters.Add("@pfPaxSegPerTkt", SqlDbType.Bit).Value = True
                .Parameters.Add("@pfShowStopovers", SqlDbType.Bit).Value = True
                .Parameters.Add("@pfShowTerminal", SqlDbType.Bit).Value = True
                .Parameters.Add("@pfFlyingTime", SqlDbType.Bit).Value = True
                .Parameters.Add("@pfCostCentre", SqlDbType.Bit).Value = True
                .Parameters.Add("@pfSeating", SqlDbType.Bit).Value = True
                .Parameters.Add("@pfVessel", SqlDbType.Bit).Value = True
                .Parameters.Add("@pfPlainFormat", SqlDbType.Bit).Value = False
                .Parameters.Add("@pfAdministrator", SqlDbType.Bit).Value = False
                .Parameters.Add("@pfFormatStyle", SqlDbType.Int).Value = 1
                .Parameters.Add("@pfOSMVesselGroup", SqlDbType.Int).Value = 0
                .Parameters.Add("@pfOSMLOGPerPax", SqlDbType.Bit).Value = 0
                .Parameters.Add("@pfOSMLOGOnSigner", SqlDbType.Bit).Value = 0
                .Parameters.Add("@pfOSMLOGPath", SqlDbType.NVarChar, 255).Value = 0
                .CommandText = " INSERT INTO [dbo].[PNRFinisherUsers] " &
                                "            ([pfPCC] " &
                                "            ,[pfUser] " &
                                "            ,[pfAgentQueue] " &
                                "            ,[pfAgentOPQueue] " &
                                "            ,[pfAgentName] " &
                                "            ,[pfAgentEmail] " &
                                "            ,[pfAirportName] " &
                                "            ,[pfAirlineLocator] " &
                                "            ,[pfClassOfService] " &
                                "            ,[pfBanElectricalEquipment] " &
                                "            ,[pfBrazilText] " &
                                "            ,[pfUSAText] " &
                                "            ,[pfTickets] " &
                                "            ,[pfPaxSegPerTkt] " &
                                "            ,[pfShowStopovers] " &
                                "            ,[pfShowTerminal] " &
                                "            ,[pfFlyingTime] " &
                                "            ,[pfCostCentre] " &
                                "            ,[pfSeating] " &
                                "            ,[pfVessel] " &
                                "            ,[pfPlainFormat] " &
                                "            ,[pfAdministrator] " &
                                "            ,[pfFormatStyle] " &
                                "            ,[pfOSMVesselGroup] " &
                                "            ,[pfOSMLOGPerPax] " &
                                "            ,[pfOSMLOGOnSigner] " &
                                "            ,[pfOSMLOGPath]) " &
                                "      VALUES " &
                                "            (@pfPCC " &
                                "            ,@pfUser " &
                                "            ,@pfAgentQueue " &
                                "            ,@pfAgentOPQueue " &
                                "            ,@pfAgentName " &
                                "            ,@pfAgentEmail " &
                                "            ,@pfAirportName " &
                                "            ,@pfAirlineLocator " &
                                "            ,@pfClassOfService " &
                                "            ,@pfBanElectricalEquipment " &
                                "            ,@pfBrazilText " &
                                "            ,@pfUSAText " &
                                "            ,@pfTickets " &
                                "            ,@pfPaxSegPerTkt " &
                                "            ,@pfShowStopovers " &
                                "            ,@pfShowTerminal " &
                                "            ,@pfFlyingTime " &
                                "            ,@pfCostCentre " &
                                "            ,@pfSeating " &
                                "            ,@pfVessel " &
                                "            ,@pfPlainFormat " &
                                "            ,@pfAdministrator " &
                                "            ,@pfFormatStyle " &
                                "            ,@pfOSMVesselGroup " &
                                "            ,@pfOSMLOGPerPax " &
                                "            ,@pfOSMLOGOnSigner " &
                                "            ,@pfOSMLOGPath )"

                pobjReader = .ExecuteScalar
            End With


        Else
            With mudtProps
                pErrorMessage = "Cannot update user " & .UserID & " for PCC " & .PCC & vbCrLf
                If Not .IsUserNameValid Then
                    pErrorMessage &= "User name required" & vbCrLf
                End If
                If Not .IsEmailValid Then
                    pErrorMessage &= "A valid email address is required" & vbCrLf
                End If
                If Not .IsQueueNumberValid Then
                    pErrorMessage &= "A valid queue number is required for ticketing" & vbCrLf
                End If
                If Not .IsOPQueueValid Then
                    pErrorMessage &= "A valid queue number is required for reminders"
                End If
            End With
            Throw New Exception(pErrorMessage)
        End If
    End Sub
End Class
