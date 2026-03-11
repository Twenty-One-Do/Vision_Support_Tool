Imports System.IO
Imports System.Text

Public Module FtpPathHelper

    Public Function NormalizeFtpPath(pathValue As String) As String
        If String.IsNullOrWhiteSpace(pathValue) Then
            Return ""
        End If

        Dim p As String = pathValue.Trim()
        p = p.Replace("\"c, "/"c)

        While p.Contains("//")
            p = p.Replace("//", "/")
        End While

        Return p
    End Function

    Public Function NormalizeRelativeDirectory(pathValue As String) As String
        Dim p As String = NormalizeFtpPath(pathValue)

        If p = "" Then Return ""

        While p.StartsWith("/")
            p = p.Substring(1)
        End While

        While p.EndsWith("/")
            p = p.Substring(0, p.Length - 1)
        End While

        Return p
    End Function

    Public Function CombineFtpPath(currentDir As String, inputPath As String) As String
        Dim arg As String = NormalizeFtpPath(inputPath)
        Dim cur As String = NormalizeFtpPath(currentDir)

        If String.IsNullOrWhiteSpace(arg) Then
            Return NormalizeFtpPath(cur)
        End If

        If arg.StartsWith("/") Then
            Return NormalizeFtpPath(arg)
        End If

        If String.IsNullOrWhiteSpace(cur) OrElse cur = "/" Then
            Return "/" & arg.Trim("/"c)
        End If

        Return NormalizeFtpPath(cur.TrimEnd("/"c) & "/" & arg.Trim("/"c))
    End Function

    Public Function SanitizeFileName(name As String) As String
        If name Is Nothing Then Return ""

        Dim invalid() As Char = Path.GetInvalidFileNameChars()
        Dim sb As New StringBuilder()

        Dim i As Integer
        For i = 0 To name.Length - 1
            Dim ch As Char = name(i)
            If Array.IndexOf(invalid, ch) >= 0 Then
                sb.Append("_")
            Else
                sb.Append(ch)
            End If
        Next

        Return sb.ToString().Trim()
    End Function

    Public Function BuildSafeLocalDirectory(rootFolder As String, ftpDirectory As String) As String
        If String.IsNullOrWhiteSpace(rootFolder) Then
            Throw New ArgumentException("루트 폴더가 비어 있습니다.")
        End If

        Dim rootFull As String = Path.GetFullPath(rootFolder)
        Dim relativeDir As String = NormalizeRelativeDirectory(ftpDirectory)

        If relativeDir = "" Then
            Return rootFull
        End If

        Dim parts() As String = relativeDir.Split("/"c)
        Dim localPath As String = rootFull

        Dim i As Integer
        For i = 0 To parts.Length - 1
            Dim seg As String = parts(i).Trim()

            If seg = "" Then Continue For
            If seg = "." OrElse seg = ".." Then
                Throw New UnauthorizedAccessException("상위 폴더 경로(..)는 허용되지 않습니다.")
            End If

            seg = SanitizeFileName(seg)
            If seg = "" Then
                Throw New UnauthorizedAccessException("유효하지 않은 폴더명이 포함되어 있습니다.")
            End If

            localPath = Path.Combine(localPath, seg)
        Next

        localPath = Path.GetFullPath(localPath)

        If Not localPath.StartsWith(rootFull, StringComparison.OrdinalIgnoreCase) Then
            Throw New UnauthorizedAccessException("루트 폴더 밖 경로는 허용되지 않습니다.")
        End If

        Return localPath
    End Function

    Public Function BuildSafeLocalFilePath(rootFolder As String, ftpFullPath As String) As String
        If String.IsNullOrWhiteSpace(ftpFullPath) Then
            Throw New ArgumentException("업로드 파일 경로가 비어 있습니다.")
        End If

        Dim normalized As String = NormalizeFtpPath(ftpFullPath)

        Dim fileName As String = Path.GetFileName(normalized)
        fileName = SanitizeFileName(fileName)

        If String.IsNullOrWhiteSpace(fileName) Then
            Throw New ArgumentException("유효한 파일명이 아닙니다.")
        End If

        Dim ftpDir As String = Path.GetDirectoryName(normalized)
        If ftpDir Is Nothing Then ftpDir = ""
        ftpDir = ftpDir.Replace("\"c, "/"c)

        Dim localDir As String = BuildSafeLocalDirectory(rootFolder, ftpDir)
        Dim fullFile As String = Path.Combine(localDir, fileName)
        Dim rootFull As String = Path.GetFullPath(rootFolder)
        fullFile = Path.GetFullPath(fullFile)

        If Not fullFile.StartsWith(rootFull, StringComparison.OrdinalIgnoreCase) Then
            Throw New UnauthorizedAccessException("루트 폴더 밖 파일 경로는 허용되지 않습니다.")
        End If

        Return fullFile
    End Function

End Module