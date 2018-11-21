Option Strict On
Option Explicit On
Public Class DownsellCollection
    Inherits Collections.Generic.Dictionary(Of Integer, DownsellItem)
    Private mstrPCC As String = ""
    Private mstrUserID As String = ""
    Private mCustomers As New CustomerCollectionAll
    Public Function CountNonVerified(ByVal PCC As String, ByVal UserID As String) As Integer
        Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
        Dim pobjComm As New SqlClient.SqlCommand

        CountNonVerified = 0
        pobjConn.Open()
        pobjComm = pobjConn.CreateCommand
        mstrPCC = PCC
        mstrUserID = UserID

        With pobjComm
            .CommandType = CommandType.Text
            .Parameters.Add("@PCC", SqlDbType.NVarChar, 9).Value = PCC
            .Parameters.Add("@UserID", SqlDbType.NVarChar, 20).Value = UserID

            .CommandText = "SELECT COUNT(*)
            FROM AmadeusReports.dbo.DownsellPNRLog
            WHERE ISNULL(doVerifiedbyUser, 0) = 0 AND (doPCC + '|' + doUserGdsId in (
            SELECT  pfPCC + '|' + pfUser AS UserKey
            FROM AmadeusReports.dbo.PNRFinisherUserName
            LEFT JOIN AmadeusReports.dbo.PNRFinisherUsers
            ON pfnID = pfUsername_fk
            WHERE pfnUserTeamleaderID_fk = (SELECT PU.pfUserName_fk FROM AmadeusReports.dbo.PNRFinisherUsers PU WHERE PU.pfPCC = @PCC AND PU.pfUser = @UserID)
            OR pfUserName_fk IN (SELECT PU.pfUserName_fk FROM AmadeusReports.dbo.PNRFinisherUsers PU WHERE PU.pfPCC = @PCC AND PU.pfUser = @UserID))
            OR (SELECT pfnIsAdministrator
            FROM AmadeusReports.dbo.PNRFinisherUsers
            LEFT JOIN AmadeusReports.dbo.PNRFinisherUserName
            ON pfUserName_fk = pfnID
            WHERE pfPCC = @PCC AND pfUser = @UserID)=1)"
            CountNonVerified = CInt(.ExecuteScalar)
        End With
        pobjConn.Close()
    End Function
    Public Sub Load(ByVal PCC As String, ByVal UserID As String)
        Dim pCustomer As CustomerItem

        MyBase.Clear()
        If PCC <> "" And UserID <> "" Then
            mstrPCC = PCC
            mstrUserID = UserID

            If mCustomers Is Nothing OrElse mCustomers.Count = 0 Then
                mCustomers.Load()
            End If
            Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            Dim pobjReader As SqlClient.SqlDataReader
            Dim pobjClass As DownsellItem
            Dim pID As Integer

            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand
            MyBase.Clear()
            With pobjComm
                .CommandType = CommandType.Text
                .Parameters.Add("@PCC", SqlDbType.NVarChar, 9).Value = PCC
                .Parameters.Add("@UserID", SqlDbType.NVarChar, 20).Value = UserID
                .CommandText = "SELECT DISTINCT 1 AS OwnPNR
                          ,doPCC
                          ,doGDS
                          ,doPNR
                          ,doUserGdsId
                          ,doDateLogged
                          ,doDownsellDecision
                          ,doClientCode
                          ,doPaxName
                          ,doItinerary
                          ,doTotal
                          ,doDownsellTotal
                          ,doFareBasis
                          ,doDownsellFareBasis
                          ,doGDSCommand
                      FROM AmadeusReports.dbo.DownsellPNRLog
                      WHERE ISNULL(doVerifiedbyUser, 0) = 0 AND doPCC + '|' + doUserGdsId in (
                      SELECT pfPCC + '|' + pfUser AS UserKey
                    FROM AmadeusReports.dbo.PNRFinisherUsers
                    WHERE pfUserName_fk IN (SELECT PU.pfUserName_fk FROM AmadeusReports.dbo.PNRFinisherUsers PU WHERE PU.pfPCC = @PCC AND PU.pfUser = @UserID))
                    UNION
SELECT DISTINCT 2 AS OwnPNR
      ,doPCC
      ,doGDS
      ,doPNR
      ,doUserGdsId
      ,doDateLogged
      ,doDownsellDecision
      ,doClientCode
      ,doPaxName
      ,doItinerary
      ,doTotal
      ,doDownsellTotal
      ,doFareBasis
      ,doDownsellFareBasis
      ,doGDSCommand
  FROM AmadeusReports.dbo.DownsellPNRLog
  WHERE ISNULL(doVerifiedbyUser, 0) = 0 AND (doPCC + '|' + doUserGdsId in (
  SELECT  pfPCC + '|' + pfUser AS UserKey
  FROM AmadeusReports.dbo.PNRFinisherUserName
  LEFT JOIN AmadeusReports.dbo.PNRFinisherUsers
  ON pfnID = pfUsername_fk
  WHERE pfnUserTeamleaderID_fk = (SELECT PU.pfUserName_fk FROM AmadeusReports.dbo.PNRFinisherUsers PU WHERE PU.pfPCC = @PCC AND PU.pfUser = @UserID))
   OR (SELECT pfnIsAdministrator
  FROM AmadeusReports.dbo.PNRFinisherUsers
  LEFT JOIN AmadeusReports.dbo.PNRFinisherUserName
  ON pfUserName_fk = pfnID
  WHERE pfPCC = @PCC AND pfUser = @UserID)=1)
  ORDER BY OwnPNR,doPCC,doUserGdsId"
                pobjReader = .ExecuteReader
            End With

            With pobjReader
                Do While .Read
                    pID += 1
                    pobjClass = New DownsellItem
                    pCustomer = mCustomers.GetCustomerByCode(CStr(.Item("doClientCode")))
                    pobjClass.SetValues(CInt(.Item("OwnPNR")), CStr(.Item("doPCC")), CStr(.Item("doGDS")), CStr(.Item("doPNR")) _
                    , CStr(.Item("doUserGdsId")), CDate(.Item("doDateLogged")), CStr(.Item("doDownsellDecision")), CStr(.Item("doClientCode")) _
                          , pCustomer.Logo, pCustomer.AlertForDownsell, CStr(.Item("doPaxName")), CStr(.Item("doItinerary")), CDec(.Item("doTotal")), CDec(.Item("doDownsellTotal")) _
                          , CStr(.Item("doFareBasis")), CStr(.Item("doDownsellFareBasis")), CStr(.Item("doGDSCommand")))
                    MyBase.Add(pID, pobjClass)
                Loop
                .Close()
            End With
            pobjConn.Close()
        End If
    End Sub
    Public Sub IgnorePNR(ByVal PCC As String, ByVal PNR As String, ByVal VerificationReason As String)
        If mstrPCC <> "" And mstrUserID <> "" Then
            Dim pobjConn As New SqlClient.SqlConnection(UtilitiesDB.ConnectionStringPNR) ' ActiveConnection)
            Dim pobjComm As New SqlClient.SqlCommand
            pobjConn.Open()
            pobjComm = pobjConn.CreateCommand
            With pobjComm
                .CommandType = CommandType.Text
                .Parameters.Add("@PCC", SqlDbType.NVarChar, 20).Value = PCC
                .Parameters.Add("@PNR", SqlDbType.NVarChar, 6).Value = PNR
                .Parameters.Add("@UserId", SqlDbType.NVarChar, 30).Value = mstrPCC & "-" & mstrUserID
                .Parameters.Add("@Date", SqlDbType.DateTime).Value = Now
                .Parameters.Add("@VerificationReason", SqlDbType.NVarChar, 50).Value = VerificationReason
                .CommandText = "UPDATE AmadeusReports.dbo.DownsellPNRLog 
                            SET doVerifiedByUser = 1, doVerifiedUserId = @UserId, doVerifiedDate = @Date, doVerificationReason=@VerificationReason
                            WHERE doPCC = @PCC AND doPNR = @PNR"
                .ExecuteNonQuery()
            End With
            pobjConn.Close()
            Load(mstrPCC, mstrUserID)
        End If
    End Sub
End Class
