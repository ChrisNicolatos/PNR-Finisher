Imports System.IO
Imports System.Text

Module modPNR

    Private mMySettings As Config
    Private mHomeSettings As Config
    Private mHomeSettingsExist As Boolean
    Private mstrRequestedPCC As String = ""
    Private mstrRequestedUser As String = ""
    Public Sub InitSettings()
        Try
            mstrRequestedPCC = ""
            mstrRequestedUser = ""

            mMySettings = New Config
            If Not mHomeSettingsExist Then
                mHomeSettings = mMySettings
                mHomeSettingsExist = True
            End If
        Catch ex As Exception
            If mHomeSettingsExist Then
                mMySettings = mHomeSettings
            Else
                Throw New Exception(ex.Message)
            End If
        End Try


    End Sub
    Public Sub InitSettings(ByVal mGDSUser As GDSUser)
        Try
            mstrRequestedPCC = mGDSUser.PCC
            mstrRequestedUser = mGDSUser.User

            mMySettings = New Config(mGDSUser)
            If Not mHomeSettingsExist Then
                mHomeSettings = mMySettings
                mHomeSettingsExist = True
            End If
        Catch ex As Exception
            If mHomeSettingsExist Then
                mMySettings = mHomeSettings
            Else
                Throw New Exception(ex.Message)
            End If
        End Try

    End Sub
    Public ReadOnly Property MySettings As Config
        Get
            MySettings = mMySettings
        End Get

    End Property
    Public ReadOnly Property MyHomeSettings As Config
        Get
            MyHomeSettings = mHomeSettings
        End Get
    End Property
    Public ReadOnly Property RequestedPCC As String
        Get
            RequestedPCC = mstrRequestedPCC
        End Get
    End Property
    Public ReadOnly Property RequestedUser As String
        Get
            RequestedUser = mstrRequestedUser
        End Get
    End Property

End Module
