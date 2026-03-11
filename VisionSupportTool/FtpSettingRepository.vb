Imports System.IO
Imports System.Xml.Serialization

Public Class FtpSettingsRepository
    Private ReadOnly _settingsDir As String
    Private ReadOnly _settingsFile As String

    Public Sub New(baseDir As String)
        _settingsDir = Path.Combine(baseDir, "Settings")
        _settingsFile = Path.Combine(_settingsDir, "ftphostsettings.xml")
    End Sub

    Private Sub EnsureSettingsDir()
        If Not Directory.Exists(_settingsDir) Then
            Directory.CreateDirectory(_settingsDir)
        End If
    End Sub

    Public Sub SaveSetting(cfg As FtpHostSetting)
        Try
            EnsureSettingsDir()

            Dim xs As New XmlSerializer(GetType(FtpHostSetting))
            Using fs As New FileStream(_settingsFile, FileMode.Create, FileAccess.Write, FileShare.None)
                xs.Serialize(fs, cfg)
            End Using

        Catch ex As Exception
            Throw New ApplicationException("FTP 설정 저장 실패: " & ex.Message, ex)
        End Try
    End Sub

    Public Function LoadSetting() As FtpHostSetting
        Try
            EnsureSettingsDir()

            If Not File.Exists(_settingsFile) Then
                Return New FtpHostSetting()
            End If

            Dim xs As New XmlSerializer(GetType(FtpHostSetting))
            Using fs As New FileStream(_settingsFile, FileMode.Open, FileAccess.Read, FileShare.Read)
                Dim cfg As FtpHostSetting = CType(xs.Deserialize(fs), FtpHostSetting)

                If cfg Is Nothing Then
                    Return New FtpHostSetting()
                End If

                Return cfg
            End Using

        Catch ex As Exception
            Throw New ApplicationException("FTP 설정 로드 실패: " & ex.Message, ex)
        End Try
    End Function
End Class