Imports System.IO
Imports System.Xml.Serialization
Imports System.ComponentModel

Public Class SettingsRepository
    Private ReadOnly _settingsDir As String
    Private ReadOnly _settingsFile As String

    Public Sub New(baseDir As String)
        _settingsDir = Path.Combine(baseDir, "Settings")
        _settingsFile = Path.Combine(_settingsDir, "foldersettings.xml")
    End Sub

    Private Sub EnsureSettingsDir()
        If Not Directory.Exists(_settingsDir) Then
            Directory.CreateDirectory(_settingsDir)
        End If
    End Sub

    Public Sub SaveSettings(list As BindingList(Of FolderSetting))
        Try
            EnsureSettingsDir()
            Dim xs As New XmlSerializer(GetType(List(Of FolderSetting)))
            Using fs As New FileStream(_settingsFile, FileMode.Create, FileAccess.Write, FileShare.None)
                xs.Serialize(fs, list.ToList())
            End Using
        Catch ex As Exception
            Throw New ApplicationException("설정 저장 실패: " & ex.Message, ex)
        End Try
    End Sub

    Public Function LoadSettings() As BindingList(Of FolderSetting)
        Dim result As New BindingList(Of FolderSetting)()

        Try
            EnsureSettingsDir()
            If Not File.Exists(_settingsFile) Then
                Return result
            End If

            Dim xs As New XmlSerializer(GetType(List(Of FolderSetting)))
            Using fs As New FileStream(_settingsFile, FileMode.Open, FileAccess.Read, FileShare.Read)
                Dim data As List(Of FolderSetting) = CType(xs.Deserialize(fs), List(Of FolderSetting))
                If data IsNot Nothing Then
                    For Each cfg In data
                        result.Add(cfg)
                    Next
                End If
            End Using
        Catch ex As Exception
            Throw New ApplicationException("설정 로드 실패: " & ex.Message, ex)
        End Try

        Return result
    End Function
End Class