Imports System.ComponentModel
Imports System.Threading.Tasks

Public Class SchedulerService
    Private ReadOnly _timer As Timer
    Private ReadOnly _lastRun As New Dictionary(Of FolderSetting, DateTime)
    Private ReadOnly _cleaner As CleanerService
    Private _folderSets As BindingList(Of FolderSetting)

    ' 동시 실행 방지용
    Private ReadOnly _sync As New Object()
    Private ReadOnly _running As New HashSet(Of FolderSetting)()

    Public Event OnLog(msg As String)

    Public Sub New(cleaner As CleanerService, folderSets As BindingList(Of FolderSetting))
        _cleaner = cleaner
        _folderSets = folderSets

        _timer = New Timer()
        _timer.Interval = 60 * 1000 ' 1분
        AddHandler _timer.Tick, AddressOf TickHandler
    End Sub

    Public Sub Start()
        _timer.Start()
        RaiseEvent OnLog("스케줄러 시작")
    End Sub

    Public Sub StopRunner()
        _timer.Stop()
        RaiseEvent OnLog("스케줄러 정지")
    End Sub

    ' 설정 목록 바뀌었을 때 교체 가능하도록
    Public Sub UpdateFolderSets(newList As BindingList(Of FolderSetting))
        _folderSets = newList
    End Sub

    ' 강제 실행 (선택 항목만) - 백그라운드 실행
    Public Sub RunOneNow(cfg As FolderSetting)
        If cfg Is Nothing Then Return

        SyncLock _sync
            _lastRun(cfg) = DateTime.Now
        End SyncLock

        RaiseEvent OnLog("[MANUAL] '" & cfg.FolderName & "' 즉시 검사 요청")
        RunPolicyBackground(cfg, "MANUAL")
    End Sub

    ' 강제 실행 (전체) - 백그라운드 실행
    Public Sub RunAllNow()
        For Each cfg In _folderSets.ToList()
            RunOneNow(cfg)
        Next
    End Sub

    ' 타이머 Tick 시 호출 (UI 스레드) → 실제 작업은 Task로 분리
    Private Sub TickHandler(sender As Object, e As EventArgs)
        Dim now As DateTime = DateTime.Now
        Dim minuteNow As Integer = now.Minute

        For Each cfg In _folderSets.ToList()
            If cfg Is Nothing Then Continue For

            ' 매 시간 cfg.ScanInterval_Min 분에 실행 (0~59)
            If minuteNow <> cfg.ScanInterval_Min Then
                Continue For
            End If

            Dim lr As DateTime = DateTime.MinValue
            SyncLock _sync
                If _lastRun.ContainsKey(cfg) Then lr = _lastRun(cfg)
            End SyncLock

            ' 같은 날짜, 같은 시(hour)에는 1번만
            If lr <> DateTime.MinValue AndAlso lr.Hour = now.Hour AndAlso lr.Date = now.Date Then
                RaiseEvent OnLog("[SCHEDULE] '" & cfg.FolderName & "' " &
                                  now.ToString("HH:mm") &
                                  "분 도달했지만 이미 실행됨 → 스킵")
                Continue For
            End If

            SyncLock _sync
                _lastRun(cfg) = now
            End SyncLock

            RaiseEvent OnLog("[SCHEDULE] '" & cfg.FolderName & "' " &
                              now.ToString("HH:mm") &
                              "분 스케줄 도달 → 검사 시작")

            RunPolicyBackground(cfg, "SCHEDULE")
        Next
    End Sub

    ' ===== 실제 백그라운드 실행 =====
    Private Sub RunPolicyBackground(cfg As FolderSetting, reason As String)
        ' 중복 실행 방지
        SyncLock _sync
            If _running.Contains(cfg) Then
                RaiseEvent OnLog("[" & reason & "] '" & cfg.FolderName & "' 이미 실행 중 → 스킵")
                Return
            End If
            _running.Add(cfg)
        End SyncLock

        ' 백그라운드 작업 시작
        Task.Run(Sub()
                     Try
                         _cleaner.RunPolicy(cfg)
                     Catch ex As Exception
                         RaiseEvent OnLog("[" & cfg.FolderName & "] 백그라운드 실행 오류: " & ex.Message)
                     Finally
                         SyncLock _sync
                             _running.Remove(cfg)
                         End SyncLock
                         RaiseEvent OnLog("[" & reason & "] '" & cfg.FolderName & "' 실행 종료")
                     End Try
                 End Sub)
    End Sub
End Class