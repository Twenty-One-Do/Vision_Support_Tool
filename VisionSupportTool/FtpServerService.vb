Imports System
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks

Public Class FtpServerService

    Public Event OnLog(msg As String)

    Private _listener As TcpListener
    Private _cts As CancellationTokenSource
    Private _acceptTask As Task
    Private _setting As FtpHostSetting
    Private _isRunning As Boolean = False

    Public ReadOnly Property IsRunning As Boolean
        Get
            Return _isRunning
        End Get
    End Property

    Public ReadOnly Property CurrentSetting As FtpHostSetting
        Get
            Return _setting
        End Get
    End Property

    Public Sub Start(setting As FtpHostSetting)
        If setting Is Nothing Then
            Throw New ArgumentNullException("setting")
        End If

        If _isRunning Then
            RaiseEvent OnLog("[FTP] 이미 실행 중입니다.")
            Exit Sub
        End If

        If String.IsNullOrWhiteSpace(setting.RootFolder) Then
            Throw New ApplicationException("FTP 루트 폴더가 비어 있습니다.")
        End If

        If Not Directory.Exists(setting.RootFolder) Then
            Directory.CreateDirectory(setting.RootFolder)
        End If

        _setting = setting
        _cts = New CancellationTokenSource()

        Dim bindIp As IPAddress = IPAddress.Any

        If Not String.IsNullOrWhiteSpace(setting.BindAddress) AndAlso setting.BindAddress <> "0.0.0.0" Then
            bindIp = IPAddress.Parse(setting.BindAddress)
        End If

        _listener = New TcpListener(bindIp, setting.Port)
        _listener.Start()

        _acceptTask = Task.Run(Sub() AcceptLoop(_cts.Token))
        _isRunning = True

        RaiseEvent OnLog("[FTP] 서버 시작 완료 - Port=" & setting.Port & ", Root=" & setting.RootFolder)
    End Sub

    Public Sub [Stop]()
        If Not _isRunning Then Exit Sub

        Try
            If _cts IsNot Nothing Then
                _cts.Cancel()
            End If
        Catch
        End Try

        Try
            If _listener IsNot Nothing Then
                _listener.Stop()
            End If
        Catch
        End Try

        Try
            If _acceptTask IsNot Nothing Then
                _acceptTask.Wait(1000)
            End If
        Catch
        End Try

        _listener = Nothing
        _acceptTask = Nothing
        _cts = Nothing
        _isRunning = False

        RaiseEvent OnLog("[FTP] 서버 중지")
    End Sub

    Private Sub AcceptLoop(token As CancellationToken)
        RaiseEvent OnLog("[FTP] 연결 대기 시작")

        While Not token.IsCancellationRequested
            Try
                Dim client As TcpClient = _listener.AcceptTcpClient()
                Dim remote As String = ""
                Try
                    remote = CType(client.Client.RemoteEndPoint, IPEndPoint).ToString()
                Catch
                End Try

                RaiseEvent OnLog("[FTP] 클라이언트 접속: " & remote)

                Task.Run(Sub() HandleClient(client, token), token)

            Catch ex As SocketException
                If token.IsCancellationRequested Then Exit While
                RaiseEvent OnLog("[FTP] Accept 오류: " & ex.Message)

            Catch ex As ObjectDisposedException
                Exit While

            Catch ex As Exception
                If token.IsCancellationRequested Then Exit While
                RaiseEvent OnLog("[FTP] AcceptLoop 오류: " & ex.Message)
            End Try
        End While

        RaiseEvent OnLog("[FTP] 연결 대기 종료")
    End Sub

    Private Sub HandleClient(client As TcpClient, token As CancellationToken)
        Dim state As New ClientState()

        Try
            client.ReceiveTimeout = 30000
            client.SendTimeout = 30000

            Using ns As NetworkStream = client.GetStream()
                Using reader As New StreamReader(ns, Encoding.ASCII)
                    Using writer As New StreamWriter(ns, Encoding.ASCII)
                        writer.NewLine = vbCrLf
                        writer.AutoFlush = True

                        state.CurrentDir = "/"
                        state.Authenticated = False
                        state.ControlClient = client

                        SendReply(writer, 220, "VisionSupportTool FTP Server Ready")

                        While client.Connected AndAlso Not token.IsCancellationRequested
                            Dim line As String = reader.ReadLine()
                            If line Is Nothing Then Exit While

                            line = line.Trim()
                            If line = "" Then Continue While

                            RaiseEvent OnLog("[FTP][REQ] " & line)

                            Dim cmd As String = line
                            Dim arg As String = ""

                            Dim sp As Integer = line.IndexOf(" "c)
                            If sp >= 0 Then
                                cmd = line.Substring(0, sp).Trim().ToUpperInvariant()
                                arg = line.Substring(sp + 1).Trim()
                            Else
                                cmd = line.Trim().ToUpperInvariant()
                            End If

                            Select Case cmd
                                Case "USER"
                                    HandleUSER(writer, state, arg)

                                Case "PASS"
                                    HandlePASS(writer, state, arg)

                                Case "SYST"
                                    SendReply(writer, 215, "UNIX Type: L8")

                                Case "FEAT"
                                    writer.WriteLine("211-Features")
                                    writer.WriteLine(" PASV")
                                    writer.WriteLine(" PORT")
                                    writer.WriteLine(" SIZE")
                                    writer.WriteLine(" UTF8")
                                    writer.WriteLine("211 End")

                                Case "OPTS"
                                    SendReply(writer, 200, "OK")

                                Case "NOOP"
                                    SendReply(writer, 200, "OK")

                                Case "PWD", "XPWD"
                                    SendReply(writer, 257, """" & state.CurrentDir & """ is current directory")

                                Case "TYPE"
                                    SendReply(writer, 200, "Type set")

                                Case "CWD"
                                    HandleCWD(writer, state, arg)

                                Case "CDUP"
                                    HandleCDUP(writer, state)

                                Case "MKD", "XMKD"
                                    HandleMKD(writer, state, arg)

                                Case "PASV"
                                    HandlePASV(writer, state)

                                Case "PORT"
                                    HandlePORT(writer, state, arg)

                                Case "LIST", "NLST"
                                    HandleLIST(writer, state, arg)

                                Case "STOR"
                                    HandleSTOR(writer, state, arg)

                                Case "QUIT"
                                    SendReply(writer, 221, "Bye")
                                    Exit While

                                Case Else
                                    SendReply(writer, 502, "Command not implemented")
                            End Select
                        End While
                    End Using
                End Using
            End Using

        Catch ex As IOException
            RaiseEvent OnLog("[FTP] 클라이언트 연결 종료: " & ex.Message)

        Catch ex As Exception
            RaiseEvent OnLog("[FTP] HandleClient 오류: " & ex.Message)

        Finally
            Try
                CloseDataSocket(state)
            Catch
            End Try

            Try
                client.Close()
            Catch
            End Try
        End Try
    End Sub

    Private Sub HandleUSER(writer As StreamWriter, state As ClientState, user As String)
        state.UserName = user

        If _setting.AllowAnonymous Then
            If String.Equals(user, "anonymous", StringComparison.OrdinalIgnoreCase) Then
                SendReply(writer, 331, "Anonymous login ok, send password")
            Else
                SendReply(writer, 331, "User name okay, need password")
            End If
        Else
            If String.Equals(user, _setting.UserName, StringComparison.OrdinalIgnoreCase) Then
                SendReply(writer, 331, "User name okay, need password")
            Else
                SendReply(writer, 530, "Invalid user")
            End If
        End If
    End Sub

    Private Sub HandlePASS(writer As StreamWriter, state As ClientState, password As String)
        If _setting.AllowAnonymous AndAlso String.Equals(state.UserName, "anonymous", StringComparison.OrdinalIgnoreCase) Then
            state.Authenticated = True
            SendReply(writer, 230, "Login successful")
            RaiseEvent OnLog("[FTP] 로그인 성공(anonymous)")
            Exit Sub
        End If

        If String.Equals(state.UserName, _setting.UserName, StringComparison.OrdinalIgnoreCase) AndAlso
           String.Equals(password, _setting.Password, StringComparison.Ordinal) Then

            state.Authenticated = True
            SendReply(writer, 230, "Login successful")
            RaiseEvent OnLog("[FTP] 로그인 성공(" & state.UserName & ")")
        Else
            SendReply(writer, 530, "Login incorrect")
        End If
    End Sub

    Private Sub HandleCWD(writer As StreamWriter, state As ClientState, arg As String)
        If Not EnsureAuthenticated(writer, state) Then Exit Sub

        Try
            Dim newFtpDir As String = CombineFtpPath(state.CurrentDir, arg)
            Dim localDir As String = BuildSafeLocalDirectory(_setting.RootFolder, newFtpDir)

            If Directory.Exists(localDir) Then
                state.CurrentDir = NormalizeDirForFtp(newFtpDir)
                SendReply(writer, 250, "Directory changed to " & state.CurrentDir)
            Else
                SendReply(writer, 550, "Directory not found")
            End If

        Catch ex As Exception
            SendReply(writer, 550, ex.Message)
        End Try
    End Sub

    Private Sub HandleCDUP(writer As StreamWriter, state As ClientState)
        If Not EnsureAuthenticated(writer, state) Then Exit Sub

        Try
            Dim cur As String = NormalizeDirForFtp(state.CurrentDir)

            If cur = "/" Then
                SendReply(writer, 250, "Directory changed to /")
                Exit Sub
            End If

            Dim trimmed As String = cur.Trim("/"c)
            Dim parts() As String = trimmed.Split("/"c)

            If parts.Length <= 1 Then
                state.CurrentDir = "/"
            Else
                state.CurrentDir = "/" & String.Join("/", parts, 0, parts.Length - 1)
            End If

            SendReply(writer, 250, "Directory changed to " & state.CurrentDir)

        Catch ex As Exception
            SendReply(writer, 550, ex.Message)
        End Try
    End Sub

    Private Sub HandleMKD(writer As StreamWriter, state As ClientState, arg As String)
        If Not EnsureAuthenticated(writer, state) Then Exit Sub

        Try
            Dim ftpDir As String = CombineFtpPath(state.CurrentDir, arg)
            Dim localDir As String = BuildSafeLocalDirectory(_setting.RootFolder, ftpDir)

            If Not Directory.Exists(localDir) Then
                Directory.CreateDirectory(localDir)
            End If

            SendReply(writer, 257, """" & NormalizeDirForFtp(ftpDir) & """ created")

        Catch ex As Exception
            SendReply(writer, 550, ex.Message)
        End Try
    End Sub

    Private Sub HandlePASV(writer As StreamWriter, state As ClientState)
        If Not EnsureAuthenticated(writer, state) Then Exit Sub

        Try
            CloseDataSocket(state)

            state.PassiveListener = New TcpListener(IPAddress.Any, 0)
            state.PassiveListener.Start()

            Dim ep As IPEndPoint = CType(state.PassiveListener.LocalEndpoint, IPEndPoint)

            Dim addr As IPAddress = Nothing

            If Not String.IsNullOrWhiteSpace(_setting.PassiveAddress) Then
                addr = IPAddress.Parse(_setting.PassiveAddress)
            Else
                Dim localEp As IPEndPoint = CType(state.ControlClient.Client.LocalEndPoint, IPEndPoint)
                addr = localEp.Address
            End If

            Dim addrBytes() As Byte = addr.GetAddressBytes()
            Dim p1 As Integer = ep.Port \ 256
            Dim p2 As Integer = ep.Port Mod 256

            Dim msg As String = "Entering Passive Mode (" &
                                addrBytes(0) & "," &
                                addrBytes(1) & "," &
                                addrBytes(2) & "," &
                                addrBytes(3) & "," &
                                p1 & "," & p2 & ")"

            SendReply(writer, 227, msg)

        Catch ex As Exception
            SendReply(writer, 425, "Cannot open passive connection: " & ex.Message)
        End Try
    End Sub

    Private Sub HandlePORT(writer As StreamWriter, state As ClientState, arg As String)
        If Not EnsureAuthenticated(writer, state) Then Exit Sub

        Try
            CloseDataSocket(state)

            Dim parts() As String = arg.Split(","c)
            If parts.Length <> 6 Then
                SendReply(writer, 501, "Syntax error in PORT")
                Exit Sub
            End If

            Dim ip As String = parts(0) & "." & parts(1) & "." & parts(2) & "." & parts(3)
            Dim port As Integer = (CInt(parts(4)) * 256) + CInt(parts(5))

            state.ActiveEndPoint = New IPEndPoint(IPAddress.Parse(ip), port)
            SendReply(writer, 200, "PORT command successful")

        Catch ex As Exception
            SendReply(writer, 501, "Invalid PORT: " & ex.Message)
        End Try
    End Sub

    Private Sub HandleLIST(writer As StreamWriter, state As ClientState, arg As String)
        If Not EnsureAuthenticated(writer, state) Then Exit Sub

        Dim dataClient As TcpClient = Nothing

        Try
            Dim targetFtpDir As String

            If String.IsNullOrWhiteSpace(arg) Then
                targetFtpDir = state.CurrentDir
            Else
                targetFtpDir = CombineFtpPath(state.CurrentDir, arg)
            End If

            Dim localDir As String = BuildSafeLocalDirectory(_setting.RootFolder, targetFtpDir)

            If Not Directory.Exists(localDir) Then
                SendReply(writer, 550, "Directory not found")
                Exit Sub
            End If

            SendReply(writer, 150, "Opening data connection for directory list")

            dataClient = OpenDataConnection(state)

            Using dataStream As NetworkStream = dataClient.GetStream()
                Using sw As New StreamWriter(dataStream, Encoding.ASCII)
                    sw.NewLine = vbCrLf
                    sw.AutoFlush = True

                    Dim dirInfo As New DirectoryInfo(localDir)

                    Dim dirs() As DirectoryInfo = dirInfo.GetDirectories()
                    Dim files() As FileInfo = dirInfo.GetFiles()

                    Dim i As Integer
                    For i = 0 To dirs.Length - 1
                        sw.WriteLine("drwxr-xr-x 1 owner group 0 " &
                                     dirs(i).LastWriteTime.ToString("MMM dd yyyy") &
                                     " " & dirs(i).Name)
                    Next

                    For i = 0 To files.Length - 1
                        sw.WriteLine("-rw-r--r-- 1 owner group " &
                                     files(i).Length &
                                     " " & files(i).LastWriteTime.ToString("MMM dd yyyy") &
                                     " " & files(i).Name)
                    Next
                End Using
            End Using

            Try
                dataClient.Close()
            Catch
            End Try

            CloseDataSocket(state)
            SendReply(writer, 226, "Directory send OK")

        Catch ex As Exception
            Try
                If dataClient IsNot Nothing Then dataClient.Close()
            Catch
            End Try

            CloseDataSocket(state)
            SendReply(writer, 550, "LIST failed: " & ex.Message)
        End Try
    End Sub

    Private Sub HandleSTOR(writer As StreamWriter, state As ClientState, arg As String)
        If Not EnsureAuthenticated(writer, state) Then Exit Sub

        Dim dataClient As TcpClient = Nothing

        Try
            If String.IsNullOrWhiteSpace(arg) Then
                SendReply(writer, 501, "File name required")
                Exit Sub
            End If

            Dim ftpFullPath As String = CombineFtpPath(state.CurrentDir, arg)
            Dim localFile As String = BuildSafeLocalFilePath(_setting.RootFolder, ftpFullPath)
            Dim localDir As String = Path.GetDirectoryName(localFile)

            If Not Directory.Exists(localDir) Then
                If _setting.AutoCreateDirectories Then
                    Directory.CreateDirectory(localDir)
                Else
                    SendReply(writer, 550, "Target directory not found")
                    Exit Sub
                End If
            End If

            If File.Exists(localFile) Then
                If _setting.AllowOverwrite Then
                    ' 그대로 진행
                ElseIf _setting.CloneToTempOnCollision Then
                    Dim dir As String = Path.GetDirectoryName(localFile)
                    Dim name As String = Path.GetFileNameWithoutExtension(localFile)
                    Dim ext As String = Path.GetExtension(localFile)
                    localFile = Path.Combine(dir, name & "_" & DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") & ext)
                Else
                    SendReply(writer, 550, "File already exists")
                    Exit Sub
                End If
            End If

            SendReply(writer, 150, "Opening data connection for file upload")

            dataClient = OpenDataConnection(state)

            Using dataStream As NetworkStream = dataClient.GetStream()
                Using fs As New FileStream(localFile, FileMode.Create, FileAccess.Write, FileShare.None)
                    dataStream.CopyTo(fs)
                End Using
            End Using

            Try
                dataClient.Close()
            Catch
            End Try

            CloseDataSocket(state)

            RaiseEvent OnLog("[FTP][UPLOAD] 저장 완료: " & localFile)
            SendReply(writer, 226, "Transfer complete")

        Catch ex As Exception
            Try
                If dataClient IsNot Nothing Then dataClient.Close()
            Catch
            End Try

            CloseDataSocket(state)
            RaiseEvent OnLog("[FTP][ERROR] 업로드 실패: " & ex.Message)
            SendReply(writer, 550, "Upload failed: " & ex.Message)
        End Try
    End Sub

    Private Function OpenDataConnection(state As ClientState) As TcpClient
        If state.PassiveListener IsNot Nothing Then
            state.PassiveListener.Server.ReceiveTimeout = 30000
            state.PassiveListener.Server.SendTimeout = 30000
            Dim client As TcpClient = state.PassiveListener.AcceptTcpClient()
            Return client
        End If

        If state.ActiveEndPoint IsNot Nothing Then
            Dim client As New TcpClient()
            client.ReceiveTimeout = 30000
            client.SendTimeout = 30000
            client.Connect(state.ActiveEndPoint)
            Return client
        End If

        Throw New ApplicationException("데이터 연결(PASV/PORT)이 준비되지 않았습니다.")
    End Function

    Private Function EnsureAuthenticated(writer As StreamWriter, state As ClientState) As Boolean
        If state.Authenticated Then Return True

        SendReply(writer, 530, "Please login with USER and PASS")
        Return False
    End Function

    Private Sub SendReply(writer As StreamWriter, code As Integer, message As String)
        Dim line As String = code.ToString() & " " & message
        writer.WriteLine(line)
        RaiseEvent OnLog("[FTP][RES] " & line)
    End Sub

    Private Sub CloseDataSocket(state As ClientState)
        Try
            If state.PassiveListener IsNot Nothing Then
                state.PassiveListener.Stop()
            End If
        Catch
        End Try

        state.PassiveListener = Nothing
        state.ActiveEndPoint = Nothing
    End Sub

    Private Function NormalizeDirForFtp(pathValue As String) As String
        Dim p As String = NormalizeFtpPath(pathValue)

        If String.IsNullOrWhiteSpace(p) Then
            Return "/"
        End If

        If Not p.StartsWith("/") Then
            p = "/" & p
        End If

        While p.EndsWith("/") AndAlso p.Length > 1
            p = p.Substring(0, p.Length - 1)
        End While

        Return p
    End Function

    Private Class ClientState
        Public Property UserName As String
        Public Property Authenticated As Boolean
        Public Property CurrentDir As String
        Public Property ControlClient As TcpClient
        Public Property PassiveListener As TcpListener
        Public Property ActiveEndPoint As IPEndPoint
    End Class

End Class