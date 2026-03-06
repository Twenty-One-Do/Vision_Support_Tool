Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO

Public Class CleanerService
    ' 실행 옵션 (전역 스위치들)
    Public Property DryRun As Boolean = True
    Public Property UseRecycleBin As Boolean = False

    ' UI 업데이트를 main에 전달하기 위한 이벤트
    Public Event OnLog(msg As String)
    Public Event OnCapacityUpdate(cfgName As String, totalPercent As Decimal, startPercent As Decimal)
    Public Event OnCountUpdate(cfgName As String, count As Integer)

    ' ====== "폴더명 전체가 날짜"만 허용 ======
    ' 허용: yyyyMMdd, yyyy.MM.dd, yyyy_MM_dd, yyyy-MM-dd
    Private Shared ReadOnly DateOnlyRegex As New Regex( _
        "^(?<y>\d{4})(?:[._-]?(?<m>\d{1,2}))(?:[._-]?(?<d>\d{1,2}))$", _
        RegexOptions.Compiled Or RegexOptions.CultureInvariant _
    )

    Public Sub RunPolicy(cfg As FolderSetting)
        Try
            If cfg Is Nothing Then Exit Sub

            Dim root As String = cfg.FolderPath
            If String.IsNullOrWhiteSpace(root) Then
                RaiseEvent OnLog("[WARN] FolderPath가 비어있음")
                Exit Sub
            End If
            If Not Directory.Exists(root) Then
                RaiseEvent OnLog("[" & cfg.FolderName & "] 경로가 존재하지 않음: " & root)
                Exit Sub
            End If

            ' =========================================================
            ' 1) 드라이브 사용률 기준 검사 시작 여부 판단
            '    - 기본 D: 드라이브 사용
            '    - 없으면 C: 드라이브 사용
            '    - 시스템 보고값만 사용, 파일 전수조사 금지
            ' =========================================================
            Dim driveLetter As String = "D:\"
            Dim driveInfo As DriveInfo = Nothing

            Try
                If DriveExists("D") Then
                    driveInfo = New DriveInfo("D")
                Else
                    RaiseEvent OnLog("[" & cfg.FolderName & "] D드라이브 감지 실패, C드라이브로 검사합니다.")
                    driveInfo = New DriveInfo("C")
                    driveLetter = "C:\"
                End If
            Catch ex As Exception
                RaiseEvent OnLog("[" & cfg.FolderName & "] 드라이브 정보 조회 실패: " & ex.Message)
                Exit Sub
            End Try

            If driveInfo Is Nothing OrElse Not driveInfo.IsReady Then
                RaiseEvent OnLog("[" & cfg.FolderName & "] 대상 드라이브를 사용할 수 없습니다: " & driveLetter)
                Exit Sub
            End If

            Dim totalPercent As Decimal = 0D

            If cfg.StartThresholdGB > 0D Then
                Dim totalSize As Decimal = CDec(driveInfo.TotalSize)
                Dim freeSize As Decimal = CDec(driveInfo.AvailableFreeSpace)
                Dim usedSize As Decimal = totalSize - freeSize

                If totalSize > 0D Then
                    totalPercent = Math.Round((usedSize / totalSize) * 100D, 2)
                End If

                RaiseEvent OnCapacityUpdate(cfg.FolderName, totalPercent, cfg.StartThresholdGB)
                RaiseEvent OnCountUpdate(cfg.FolderName, 0)

                If totalPercent < cfg.StartThresholdGB Then
                    RaiseEvent OnLog("[" & cfg.FolderName & "] " & driveLetter & " 사용률 " & totalPercent & "% < 기준 " & cfg.StartThresholdGB & "% → 삭제 안 함.")
                    Exit Sub
                Else
                    RaiseEvent OnLog("[" & cfg.FolderName & "] " & driveLetter & " 사용률 " & totalPercent & "% >= 기준 " & cfg.StartThresholdGB & "% → 삭제 진행")
                End If
            Else
                RaiseEvent OnCapacityUpdate(cfg.FolderName, 0D, 0D)
                RaiseEvent OnCountUpdate(cfg.FolderName, 0)
                RaiseEvent OnLog("[" & cfg.FolderName & "] 시작 기준 퍼센트가 0이므로 드라이브 사용률 검사 없이 삭제 로직으로 진행")
            End If

            ' =========================================================
            ' 2) TTL → 기준 날짜 자동 계산 (로컬 날짜 기준)
            ' =========================================================
            Dim cutoffDate As Date = Date.Today.AddDays(-cfg.ImageTTL_Days)

            RaiseEvent OnLog("[" & cfg.FolderName & "] 폴더명 날짜 기준 삭제 시작. cutoff <= " &
                             cutoffDate.ToString("yyyy-MM-dd") &
                             " (TTL " & cfg.ImageTTL_Days & "일)")

            ' =========================================================
            ' 3) 하위 모든 폴더 전수조사 → "폴더명 전체가 날짜"인 폴더만 후보
            '    (파일 전수조사는 아니므로 허용)
            ' =========================================================
            Dim allDirs As IEnumerable(Of String)
            Try
                allDirs = Directory.EnumerateDirectories(root, "*", System.IO.SearchOption.AllDirectories)
            Catch ex As Exception
                RaiseEvent OnLog("[" & cfg.FolderName & "] 폴더 나열 실패(" & root & "): " & ex.Message)
                Exit Sub
            End Try

            Dim rawCandidates As New List(Of CandidateDir)()

            For Each dirPath As String In allDirs
                Dim name As String = Path.GetFileName(dirPath)
                Dim d As Date
                If TryParseDateOnlyFolderName(name, d) Then
                    If d <= cutoffDate Then
                        Dim c As New CandidateDir()
                        c.DirPath = dirPath
                        c.FolderDate = d
                        rawCandidates.Add(c)
                    End If
                End If
            Next

            If rawCandidates.Count = 0 Then
                RaiseEvent OnLog("[" & cfg.FolderName & "] 삭제 대상 폴더 없음.")
                Exit Sub
            End If

            ' =========================================================
            ' 3-1) 부모/자식이 둘 다 후보면 "부모만" 남기기
            ' =========================================================
            Dim candidates As New List(Of CandidateDir)()
            Dim selected As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            rawCandidates.Sort(AddressOf CompareByPathLengthAsc)

            For Each c As CandidateDir In rawCandidates
                If Not IsUnderAnySelectedParent(c.DirPath, selected) Then
                    candidates.Add(c)
                    selected.Add(NormalizeDirPath(c.DirPath))
                End If
            Next

            RaiseEvent OnLog("[" & cfg.FolderName & "] 삭제 대상 폴더(부모 기준) " & candidates.Count & "개 탐지")

            ' =========================================================
            ' 4) 삭제 루프
            ' =========================================================
            candidates.Sort(AddressOf CompareByPathLengthDesc)

            Dim delCount As Integer = 0

            For Each c As CandidateDir In candidates
                Try
                    If DryRun Then
                        RaiseEvent OnLog("[" & cfg.FolderName & "][DRY-RUN] 삭제 대상 폴더: " &
                                         c.DirPath & " (date=" & c.FolderDate.ToString("yyyy-MM-dd") & ")")
                    Else
                        DeleteDirectory(c.DirPath)
                        delCount += 1
                        RaiseEvent OnLog("[" & cfg.FolderName & "] 삭제 완료: " & c.DirPath)
                    End If
                Catch ex As Exception
                    RaiseEvent OnLog("[" & cfg.FolderName & "] 삭제 실패: " & c.DirPath & " - " & ex.Message)
                End Try
            Next

            ' =========================================================
            ' 5) 요약 로그
            ' =========================================================
            If DryRun Then
                RaiseEvent OnLog("[" & cfg.FolderName & "][DRY-RUN] 실제 삭제 없음.")
            Else
                RaiseEvent OnLog("[" & cfg.FolderName & "] 폴더 삭제 " & delCount & "개 완료.")
            End If

        Catch ex As Exception
            RaiseEvent OnLog("[" & cfg.FolderName & "] 오류: " & ex.Message)
        End Try
    End Sub

    Private Function DriveExists(driveName As String) As Boolean
        Try
            For Each d As DriveInfo In DriveInfo.GetDrives()
                If String.Equals(d.Name.TrimEnd("\"c), driveName & ":", StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
        Catch
        End Try

        Return False
    End Function

    ' ===== 후보 타입 =====
    Private Class CandidateDir
        Public DirPath As String
        Public FolderDate As Date
    End Class

    Private Shared Function CompareByPathLengthAsc(a As CandidateDir, b As CandidateDir) As Integer
        Return a.DirPath.Length.CompareTo(b.DirPath.Length)
    End Function

    Private Shared Function CompareByPathLengthDesc(a As CandidateDir, b As CandidateDir) As Integer
        Return b.DirPath.Length.CompareTo(a.DirPath.Length)
    End Function

    ' ----- 날짜 폴더명 파싱 -----
    Private Function TryParseDateOnlyFolderName(folderName As String, ByRef result As Date) As Boolean
        result = Date.MinValue
        If String.IsNullOrWhiteSpace(folderName) Then Return False

        Dim m As Match = DateOnlyRegex.Match(folderName.Trim())
        If Not m.Success Then Return False

        Dim y As Integer, mo As Integer, d As Integer
        If Not Integer.TryParse(m.Groups("y").Value, y) Then Return False
        If Not Integer.TryParse(m.Groups("m").Value, mo) Then Return False
        If Not Integer.TryParse(m.Groups("d").Value, d) Then Return False

        If y < 1900 OrElse y > 3000 Then Return False
        If mo < 1 OrElse mo > 12 Then Return False

        Dim dimm As Integer = DateTime.DaysInMonth(y, mo)
        If d < 1 OrElse d > dimm Then Return False

        result = New Date(y, mo, d)
        Return True
    End Function

    ' ----- 삭제 -----
    Private Sub DeleteDirectory(dirPath As String)
        If Not Directory.Exists(dirPath) Then Exit Sub

        If UseRecycleBin Then
            FileSystem.DeleteDirectory(dirPath, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin)
        Else
            Directory.Delete(dirPath, recursive:=True)
        End If
    End Sub

    ' ----- 경로 정규화/부모 포함 판정 -----
    Private Function NormalizeDirPath(p As String) As String
        If String.IsNullOrWhiteSpace(p) Then Return ""
        Dim x As String = p.Trim()
        Return x.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
    End Function

    Private Function IsUnderAnySelectedParent(path As String, selectedParents As HashSet(Of String)) As Boolean
        Dim cur As DirectoryInfo = Nothing
        Try
            cur = New DirectoryInfo(path)
        Catch
            Return False
        End Try

        While cur IsNot Nothing
            Dim parentPath As String = NormalizeDirPath(cur.FullName)
            If selectedParents.Contains(parentPath) Then Return True
            cur = cur.Parent
        End While

        Return False
    End Function
End Class