Attribute VB_Name = "modPerfLog"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

Private Const PERF_LOG_FILENAME As String = "invSys.Inventory.Sync.log"

Private mPerfRunId As String
Private mPerfStartMs As Double
Private mPerfLastMs As Double
Private mPerfActive As Boolean

Public Sub BeginTransaction(ByVal label As String)
    Dim tickNow As Double

    tickNow = GetTickMsPerf()
    mPerfRunId = Trim$(label) & "-" & Format$(Now, "hhmmss")
    mPerfStartMs = tickNow
    mPerfLastMs = tickNow
    mPerfActive = True
    AppendPerfLine "[PERF-BEGIN] " & mPerfRunId & " | wall=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Public Sub MarkSegment(ByVal segmentName As String)
    Dim tickNow As Double

    If Not mPerfActive Then Exit Sub

    tickNow = GetTickMsPerf()
    AppendPerfLine "[PERF] " & mPerfRunId & " | " & Trim$(segmentName) & _
                   " | seg=" & Format$(tickNow - mPerfLastMs, "0") & "ms" & _
                   " | total=" & Format$(tickNow - mPerfStartMs, "0") & "ms"
    mPerfLastMs = tickNow
End Sub

Public Sub EndTransaction(ByVal resultText As String)
    Dim tickNow As Double

    If Not mPerfActive Then Exit Sub

    tickNow = GetTickMsPerf()
    AppendPerfLine "[PERF-END] " & mPerfRunId & " | " & Trim$(resultText) & _
                   " | total=" & Format$(tickNow - mPerfStartMs, "0") & "ms"
    mPerfRunId = vbNullString
    mPerfStartMs = 0
    mPerfLastMs = 0
    mPerfActive = False
End Sub

Public Function IsTransactionActive() As Boolean
    IsTransactionActive = mPerfActive
End Function

Private Function GetTickMsPerf() As Double
    GetTickMsPerf = CDbl(timeGetTime())
End Function

Private Sub AppendPerfLine(ByVal lineText As String)
    Dim fileNum As Integer
    Dim logPath As String

    On Error Resume Next
    logPath = ResolvePerfLogPath()
    fileNum = FreeFile
    Open logPath For Append As #fileNum
    Print #fileNum, lineText
    Close #fileNum
    On Error GoTo 0
End Sub

Private Function ResolvePerfLogPath() As String
    Dim rootPath As String

    rootPath = Trim$(Environ$("TEMP"))
    If rootPath = "" Then rootPath = CurDir$
    If Right$(rootPath, 1) <> "\" Then rootPath = rootPath & "\"
    ResolvePerfLogPath = rootPath & PERF_LOG_FILENAME
End Function
