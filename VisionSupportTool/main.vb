Imports System
Imports System.IO
Imports System.ComponentModel
Imports Microsoft.Win32

Public Class main
    ' ===== 필드 =====
    Private _repo As SettingsRepository
    Private _cleaner As CleanerService
    Private _scheduler As SchedulerService
    Private _folderSets As BindingList(Of FolderSetting)

    ' ===== 트레이 실행용 =====
    Private _tray As NotifyIcon
    Private _allowExit As Boolean = False

    ' ===== 자동 실행 등록 설정 =====
    Private Const STARTUP_RUN_KEY As String = "Software\Microsoft\Windows\CurrentVersion\Run"
    Private Const STARTUP_VALUE_NAME As String = "VisionSupportTool"

    ' 체크박스 초기화 중 이벤트 오작동 방지
    Private _isLoading As Boolean = False

    ' 시작 시 백그라운드 모드 여부
    Private _startBackground As Boolean = False

    ' ===== 폼 로드 =====
    Private Sub main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            _isLoading = True

            ' 시작 인자 확인
            _startBackground = HasBackgroundArgument()

            _repo = New SettingsRepository(Application.StartupPath)

            Try
                _folderSets = _repo.LoadSettings()
                LogLine("[LOAD] 설정 로드 완료: " & _folderSets.Count & "개")
            Catch ex As Exception
                _folderSets = New BindingList(Of FolderSetting)()
                LogLine("[WARN] 설정 로드 실패: " & ex.Message)
            End Try

            LB_FolderSetList.DataSource = _folderSets
            LB_FolderSetList.DisplayMember = "Display"

            _cleaner = New CleanerService()
            AddHandler _cleaner.OnLog, AddressOf HandleServiceLog
            AddHandler _cleaner.OnCapacityUpdate, AddressOf HandleCapacityUpdate
            AddHandler _cleaner.OnCountUpdate, AddressOf HandleCountUpdate

            _cleaner.DryRun = CB_DryRun.Checked
            _cleaner.UseRecycleBin = CB_RecycleBin.Checked

            _scheduler = New SchedulerService(_cleaner, _folderSets)
            AddHandler _scheduler.OnLog, AddressOf HandleServiceLog
            _scheduler.Start()

            InitTray()

            If CB_AutoStartup IsNot Nothing Then
                CB_AutoStartup.Checked = IsStartupRegistered()
            End If

            ' 검사 시작 기준: 드라이브 사용률 %
            NUD_ScanStartVol.Minimum = 0D
            NUD_ScanStartVol.Maximum = 90D

            LogLine("[INIT] 프로그램 초기화 완료.")

            ' 자동실행(/background)으로 시작된 경우 창 숨김
            If _startBackground Then
                BeginInvoke(New Action(AddressOf StartInBackground))
            End If

        Catch ex As Exception
            MessageBox.Show("초기화 중 오류: " & ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isLoading = False
        End Try
    End Sub

    Private Function HasBackgroundArgument() As Boolean
        Try
            Dim args() As String = Environment.GetCommandLineArgs()
            If args Is Nothing Then Return False

            Dim i As Integer
            For i = 0 To args.Length - 1
                Dim s As String = args(i)
                If s IsNot Nothing Then
                    s = s.Trim().ToLower()
                    If s = "/background" OrElse s = "-background" OrElse s = "/tray" OrElse s = "-tray" Then
                        Return True
                    End If
                End If
            Next
        Catch
        End Try

        Return False
    End Function

    Private Sub StartInBackground()
        Try
            Me.Hide()
            Me.WindowState = FormWindowState.Minimized
            Me.ShowInTaskbar = False

            If _tray IsNot Nothing Then
                _tray.Visible = True
            End If

            LogLine("[AUTO] 백그라운드 모드로 시작됨")
        Catch
        End Try
    End Sub

    ' ===== 자동 실행 등록 상태 확인 =====
    Private Function IsStartupRegistered() As Boolean
        Try
            Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey(STARTUP_RUN_KEY, False)
            If key Is Nothing Then Return False

            Dim val As Object = key.GetValue(STARTUP_VALUE_NAME, Nothing)
            key.Close()

            If val Is Nothing Then Return False

            Dim currentExe As String = """" & Application.ExecutablePath & """ /background"
            Return String.Equals(val.ToString(), currentExe, StringComparison.OrdinalIgnoreCase)
        Catch ex As Exception
            LogLine("[WARN] 시작프로그램 등록 상태 확인 실패: " & ex.Message)
            Return False
        End Try
    End Function

    ' ===== 자동 실행 등록 =====
    Private Sub RegisterStartup()
        Try
            Dim exePath As String = """" & Application.ExecutablePath & """ /background"
            Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey(STARTUP_RUN_KEY, True)

            If key Is Nothing Then
                LogLine("[WARN] 시작프로그램 레지스트리 키를 열 수 없음")
                Exit Sub
            End If

            key.SetValue(STARTUP_VALUE_NAME, exePath, RegistryValueKind.String)
            key.Close()

            LogLine("[AUTO] 시작프로그램 등록 완료: " & STARTUP_VALUE_NAME)
        Catch ex As Exception
            LogLine("[WARN] 시작프로그램 등록 실패: " & ex.Message)
        End Try
    End Sub

    ' ===== 자동 실행 해제 =====
    Private Sub UnregisterStartup()
        Try
            Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey(STARTUP_RUN_KEY, True)
            If key Is Nothing Then Exit Sub

            key.DeleteValue(STARTUP_VALUE_NAME, False)
            key.Close()

            LogLine("[AUTO] 시작프로그램 해제 완료: " & STARTUP_VALUE_NAME)
        Catch ex As Exception
            LogLine("[WARN] 시작프로그램 해제 실패: " & ex.Message)
        End Try
    End Sub

    ' ===== 자동 실행 체크박스 변경 =====
    Private Sub CB_AutoStartup_CheckedChanged(sender As Object, e As EventArgs) Handles CB_AutoStartup.CheckedChanged
        If _isLoading Then Return

        Try
            If CB_AutoStartup.Checked Then
                RegisterStartup()
            Else
                UnregisterStartup()
            End If
        Catch ex As Exception
            LogLine("[WARN] 자동 실행 설정 변경 실패: " & ex.Message)
        End Try
    End Sub

    ' ===== 트레이 =====
    Private Sub InitTray()
        Try
            _tray = New NotifyIcon()
            _tray.Icon = Me.Icon
            _tray.Visible = True
            _tray.Text = "Folder Cleaner (백그라운드 실행중)"

            Dim menu As New ContextMenuStrip()
            menu.Items.Add("열기", Nothing, AddressOf Tray_Open_Click)
            menu.Items.Add("지금 전체 실행", Nothing, AddressOf Tray_RunAll_Click)
            menu.Items.Add(New ToolStripSeparator())
            menu.Items.Add("종료", Nothing, AddressOf Tray_Exit_Click)

            _tray.ContextMenuStrip = menu
            AddHandler _tray.DoubleClick, AddressOf Tray_Open_Click
        Catch ex As Exception
            LogLine("[WARN] 트레이 초기화 실패: " & ex.Message)
        End Try
    End Sub

    ' ===== 폼 종료(닫기) 처리 =====
    Private Sub main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Not _allowExit Then
            e.Cancel = True
            Me.Hide()
            Me.ShowInTaskbar = False

            Try
                If _tray IsNot Nothing Then
                    _tray.BalloonTipTitle = "백그라운드 실행"
                    _tray.BalloonTipText = "창은 닫혔지만 백그라운드에서 계속 실행 중입니다. (트레이 아이콘 우클릭)"
                    _tray.ShowBalloonTip(1500)
                End If
            Catch
            End Try

            Return
        End If

        Try
            _repo.SaveSettings(_folderSets)
            LogLine("[SAVE] 설정 저장 완료.")
        Catch ex As Exception
            LogLine("[ERROR] 설정 저장 실패: " & ex.Message)
        End Try

        Try
            If _scheduler IsNot Nothing Then _scheduler.StopRunner()
        Catch
        End Try

        Try
            If _tray IsNot Nothing Then
                _tray.Visible = False
                _tray.Dispose()
                _tray = Nothing
            End If
        Catch
        End Try
    End Sub

    ' ===== 트레이 메뉴 핸들러 =====
    Private Sub Tray_Open_Click(sender As Object, e As EventArgs)
        Try
            Me.Show()
            Me.WindowState = FormWindowState.Normal
            Me.ShowInTaskbar = True
            Me.Activate()
        Catch
        End Try
    End Sub

    Private Sub Tray_RunAll_Click(sender As Object, e As EventArgs)
        Try
            If _scheduler IsNot Nothing Then
                _scheduler.RunAllNow()
            End If
        Catch ex As Exception
            LogLine("[ERROR] 트레이 전체 실행 실패: " & ex.Message)
        End Try
    End Sub

    Private Sub Tray_Exit_Click(sender As Object, e As EventArgs)
        _allowExit = True
        Try
            Me.ShowInTaskbar = True
            Me.Close()
        Catch
        End Try
    End Sub

    ' ===== 로그 헬퍼 =====
    Private Sub LogLine(msg As String)
        If TB_Logs.InvokeRequired Then
            TB_Logs.Invoke(New Action(Of String)(AddressOf LogLine), msg)
        Else
            TB_Logs.AppendText("[" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "] " & msg & Environment.NewLine)
        End If
    End Sub

    ' ===== CleanerService / SchedulerService 이벤트 핸들러 =====
    Private Sub HandleServiceLog(msg As String)
        LogLine(msg)
    End Sub

    Private Sub HandleCapacityUpdate(cfgName As String, totalPercent As Decimal, startPercent As Decimal)
        Try
            If PB_VolPerScanVol.InvokeRequired Then
                PB_VolPerScanVol.Invoke(New Action(Of String, Decimal, Decimal)(AddressOf HandleCapacityUpdate), cfgName, totalPercent, startPercent)
            Else
                Dim percent As Integer = 0
                If startPercent > 0D Then
                    percent = CInt(Math.Min(100, (totalPercent / startPercent) * 100))
                End If
                PB_VolPerScanVol.Value = percent
            End If
        Catch
        End Try
    End Sub

    Private Sub HandleCountUpdate(cfgName As String, count As Integer)
        LogLine("[" & cfgName & "] 대상 수: " & count)
    End Sub

    ' ===== 체크박스 상태 변경 =====
    Private Sub CB_DryRun_CheckedChanged(sender As Object, e As EventArgs) Handles CB_DryRun.CheckedChanged
        If _cleaner Is Nothing Then Return
        _cleaner.DryRun = CB_DryRun.Checked
        LogLine("드라이런 모드: " & If(_cleaner.DryRun, "ON", "OFF"))
    End Sub

    Private Sub CB_RecycleBin_CheckedChanged(sender As Object, e As EventArgs) Handles CB_RecycleBin.CheckedChanged
        If _cleaner Is Nothing Then Return
        _cleaner.UseRecycleBin = CB_RecycleBin.Checked
        LogLine("휴지통 삭제 모드: " & If(_cleaner.UseRecycleBin, "ON", "OFF"))
    End Sub

    ' ===== 폴더 경로 선택 =====
    Private Sub Btn_SelectFolderPath_Click(sender As Object, e As EventArgs) Handles Btn_SelectFolderPath.Click
        Using dlg As New FolderBrowserDialog()
            dlg.Description = "폴더 경로를 선택하세요."
            If Directory.Exists(TB_FolderPath.Text) Then
                dlg.SelectedPath = TB_FolderPath.Text
            End If

            If dlg.ShowDialog() = DialogResult.OK Then
                TB_FolderPath.Text = dlg.SelectedPath
                LogLine("폴더 선택: " & dlg.SelectedPath)
            End If
        End Using
    End Sub

    ' ===== 리스트 선택 시 설정 불러오기 =====
    Private Sub LB_FolderSetList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LB_FolderSetList.SelectedIndexChanged
        Dim idx As Integer = LB_FolderSetList.SelectedIndex
        If idx < 0 OrElse idx >= _folderSets.Count Then Return
        Dim cfg As FolderSetting = _folderSets(idx)

        TB_FolderSetName.Text = cfg.FolderName
        TB_Memo.Text = cfg.Memo
        TB_FolderPath.Text = cfg.FolderPath
        NUD_ScanStartVol.Value = cfg.StartThresholdGB
        NUD_ImageLife.Value = cfg.ImageTTL_Days
        NUD_ScanPeriod.Value = cfg.ScanInterval_Min
    End Sub

    ' ===== 설정 추가 =====
    Private Sub Btn_FolderSetList_Add_Click(sender As Object, e As EventArgs) Handles Btn_FolderSetList_Add.Click
        Try
            Dim cfg As New FolderSetting() With {
                .FolderName = TB_FolderSetName.Text.Trim(),
                .Memo = TB_Memo.Text.Trim(),
                .FolderPath = TB_FolderPath.Text.Trim(),
                .StartThresholdGB = NUD_ScanStartVol.Value,
                .ImageTTL_Days = CInt(NUD_ImageLife.Value),
                .ScanInterval_Min = CInt(NUD_ScanPeriod.Value)
            }

            _folderSets.Add(cfg)
            _repo.SaveSettings(_folderSets)
            _scheduler.UpdateFolderSets(_folderSets)
            LogLine("[ADD] " & cfg.Display)
        Catch ex As Exception
            MessageBox.Show("추가 실패: " & ex.Message)
        End Try
    End Sub

    ' ===== 설정 수정 =====
    Private Sub Btn_FolderSetList_Edit_Click(sender As Object, e As EventArgs) Handles Btn_FolderSetList_Edit.Click
        Dim idx As Integer = LB_FolderSetList.SelectedIndex
        If idx < 0 OrElse idx >= _folderSets.Count Then Return

        Try
            Dim cfg As FolderSetting = _folderSets(idx)
            cfg.FolderName = TB_FolderSetName.Text.Trim()
            cfg.Memo = TB_Memo.Text.Trim()
            cfg.FolderPath = TB_FolderPath.Text.Trim()
            cfg.StartThresholdGB = NUD_ScanStartVol.Value
            cfg.ImageTTL_Days = CInt(NUD_ImageLife.Value)
            cfg.ScanInterval_Min = CInt(NUD_ScanPeriod.Value)

            LB_FolderSetList.DataSource = Nothing
            LB_FolderSetList.DisplayMember = Nothing
            LB_FolderSetList.DataSource = _folderSets
            LB_FolderSetList.DisplayMember = "Display"

            _repo.SaveSettings(_folderSets)
            LogLine("[EDIT] " & cfg.Display)
        Catch ex As Exception
            MessageBox.Show("수정 실패: " & ex.Message)
        End Try
    End Sub

    ' ===== 설정 삭제 =====
    Private Sub Btn_FolderSetList_Del_Click(sender As Object, e As EventArgs) Handles Btn_FolderSetList_Del.Click
        Dim idx As Integer = LB_FolderSetList.SelectedIndex
        If idx < 0 OrElse idx >= _folderSets.Count Then Return
        Dim cfg As FolderSetting = _folderSets(idx)

        Dim result As DialogResult = MessageBox.Show( _
            String.Format("'{0}' 폴더 설정을 정말 삭제하시겠습니까?" & vbCrLf & "이 작업은 되돌릴 수 없습니다.", cfg.FolderName), _
            "삭제 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)

        If result <> DialogResult.Yes Then
            LogLine("[CANCEL] '" & cfg.FolderName & "' 삭제 취소됨.")
            Return
        End If

        _folderSets.RemoveAt(idx)
        _repo.SaveSettings(_folderSets)
        _scheduler.UpdateFolderSets(_folderSets)
        LogLine("[DEL] " & cfg.Display)
    End Sub

    ' ===== 즉시 실행 (선택 항목) =====
    Private Sub Btn_RunSelectedNow_Click(sender As Object, e As EventArgs) Handles Btn_RunSelectedNow.Click
        Dim idx As Integer = LB_FolderSetList.SelectedIndex
        If idx < 0 OrElse idx >= _folderSets.Count Then
            MessageBox.Show("즉시 검사할 항목을 선택하세요.")
            Return
        End If
        _scheduler.RunOneNow(_folderSets(idx))
    End Sub

    ' ===== 즉시 실행 (전체) =====
    Private Sub Btn_RunAllNow_Click(sender As Object, e As EventArgs) Handles Btn_RunAllNow.Click
        _scheduler.RunAllNow()
    End Sub
End Class