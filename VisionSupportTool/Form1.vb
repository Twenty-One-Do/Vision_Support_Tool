Imports System.ComponentModel

Public Class FolderSetting
    Public Property FolderName As String        ' TB_FolderSetName
    Public Property Memo As String              ' TB_Memo
    Public Property FolderPath As String        ' TB_FolderPath
    Public Property StartThresholdGB As Decimal ' NUD_ScanStartVol (0~90 %) - 드라이브 사용률 기준

    Public Property ImageTTL_Days As Integer    ' NUD_ImageLife (일)
    Public Property ScanInterval_Min As Integer ' NUD_ScanPeriod (0~59 분) — 매 시간 이 분에 실행

    Public ReadOnly Property Display As String
        Get
            Return FolderName
        End Get
    End Property
End Class