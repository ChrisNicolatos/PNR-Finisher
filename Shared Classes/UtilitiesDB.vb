Imports System.IO
Imports System.Text
Public NotInheritable Class UtilitiesDB
    '
    ' Prepares the SQL connection string for the Travel Force Cosmos database
    '
    ' The SQL connection string for the Travel Force Cosmos database
    ' Returns the connection string to be used for SQLConnection
    ' The options for the connection string are bound to the application and cannot be modified by the user
    ' 

    Private Const mstrDBConnectionsFileGT As String = "\\192.168.102.223\Common\Click-Once Applications\PNR Finisher ATH V2\Config\PNRFinisher.txt"
    Private Const mstrDBConnectionsFile As String = "\\ath2-svrdc1\PNR Finisher ATH Config\PNRFinisher.txt"
    Private Const msrtConfigFolder As String = "\\192.168.102.223\Common\Click-Once Applications\PNR Finisher ATH V2\Config"
    Private Shared mstrDBConnectionFileActual As String

    Private Shared mPnrDataSource As String = ""
    Private Shared mPnrDataCatalog As String = ""
    Private Shared mPnrUserName As String = ""
    Private Shared mPnrPassword As String = ""
    Private Sub New()
    End Sub
    Public Shared ReadOnly Property PNRDataSource As String
        Get
            PNRDataSource = mPnrDataSource
        End Get
    End Property
    Public Shared ReadOnly Property MyConfigPath As String
        Get
            MyConfigPath = msrtConfigFolder
        End Get
    End Property
    Public Shared ReadOnly Property PNRDataCatalog As String
        Get
            PNRDataCatalog = mPnrDataCatalog
        End Get
    End Property
    Public Shared ReadOnly Property PNRUserName As String
        Get
            PNRUserName = mPnrUserName
        End Get
    End Property
    Public Shared ReadOnly Property DBConnectionsFile As String
        Get
            DBConnectionsFile = mstrDBConnectionFileActual
        End Get
    End Property
    Public Shared ReadOnly Property ConnectionStringACC() As String
        Get
            If MySettings.PCCDBDataSource = "" Then
                ReadDBConnections()
            End If
            ConnectionStringACC = "Data Source=" & MySettings.PCCDBDataSource &
                                  ";Initial Catalog=" & MySettings.PCCDBInitialCatalog &
                                  ";User ID=" & MySettings.PCCDBUserId &
                                  ";Password=" & MySettings.PCCDBUserPassword

        End Get
    End Property
    Public Shared ReadOnly Property ConnectionStringPNR() As String
        Get
            If mPnrDataSource = "" Then
                ReadDBConnections()
            End If
            ConnectionStringPNR = "Data Source=" & mPnrDataSource &
                                  ";Initial Catalog=" & mPnrDataCatalog &
                                  ";User ID=" & mPnrUserName &
                                  ";Password=" & mPnrPassword
        End Get
    End Property
    Private Shared Sub ReadDBConnections()

        Dim pFileExists As Boolean = False

        If File.Exists(mstrDBConnectionsFile) Then
            mstrDBConnectionFileActual = mstrDBConnectionsFile
            pFileExists = True
        ElseIf File.Exists(mstrDBConnectionsFileGT) Then
            mstrDBConnectionFileActual = mstrDBConnectionsFileGT
            pFileExists = True
        End If

        If pFileExists Then
            Dim GDSData As StreamReader = File.OpenText(mstrDBConnectionFileActual)
            Dim xLine() As String = Split(GDSData.ReadToEnd, vbCrLf)
            GDSData.Close()

            If IsArray(xLine) Then
                For i As Integer = xLine.GetLowerBound(0) To xLine.GetUpperBound(0)
                    Dim pValues() As String = xLine(i).Split("=")
                    Select Case pValues(0).Trim.ToUpper
                        Case "DATASOURCEPNR"
                            mPnrDataSource = pValues(1).Trim
                        Case "DATACATALOGPNR"
                            mPnrDataCatalog = pValues(1).Trim
                        Case "DATAUSERNAMEPNR"
                            mPnrUserName = pValues(1).Trim
                        Case "DATAUSERPASSWORDPNR"
                            mPnrPassword = pValues(1).Trim
                    End Select
                Next
            Else
                Throw New Exception("Settings File Error" & vbCrLf & mstrDBConnectionFileActual)
            End If
        Else
            Throw New Exception("DB Connection file does not exist. Please contact you system administrator" & vbCrLf & mstrDBConnectionFileActual)
        End If

    End Sub
End Class
