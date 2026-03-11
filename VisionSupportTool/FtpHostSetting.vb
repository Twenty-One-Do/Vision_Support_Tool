Imports System

<Serializable()>
Public Class FtpHostSetting
    Public Property Enabled As Boolean = False
    Public Property RootFolder As String = "D:\FTP_ROOT"
    Public Property Port As Integer = 2121

    ' 테스트 끝나고 카메라 붙일 때는 0.0.0.0 또는 실제 서버 IP 사용
    Public Property BindAddress As String = "127.0.0.1"

    Public Property PassiveAddress As String = ""

    Public Property AllowAnonymous As Boolean = False
    Public Property UserName As String = "camera"
    Public Property Password As String = "1234"

    Public Property AutoCreateDirectories As Boolean = True
    Public Property AllowOverwrite As Boolean = True
    Public Property AutoStartOnProgramLaunch As Boolean = False
    Public Property LogToFile As Boolean = False
    Public Property CloneToTempOnCollision As Boolean = False
End Class