Attribute VB_Name = "modInventoryInit"
Option Explicit

Private gNextSourceSync As Date
Private gSourceSyncScheduled As Boolean
Private Const SOURCE_SYNC_INTERVAL_SECONDS As Long = 2
Private Const SOURCE_SYNC_IDLE_INTERVAL_SECONDS As Long = 2
Private Const SOURCE_SYNC_LOG_FILENAME As String = "invSys.Inventory.Sync.log"
Private Const INVENTORY_DOMAIN_CONTRACT_VERSION As String = "R1-INVENTORY-1"

Public Sub InitInventoryDomainAddin()
    ' D3 boundary: loading the Inventory Domain must be inert. Snapshot
    ' publication and catalog commands are invoked explicitly by Core against
    ' an identified authoritative workbook; the Domain never scans open
    ' operator workbooks or subscribes to Application.WorkbookOpen.
End Sub

Public Sub Auto_Open()
    InitInventoryDomainAddin
End Sub

Public Function GetInventoryDomainContractVersion() As String
    GetInventoryDomainContractVersion = INVENTORY_DOMAIN_CONTRACT_VERSION
End Function

Public Function DiagnoseInventoryDomain() As String
    DiagnoseInventoryDomain = _
        "ContractVersion=" & INVENTORY_DOMAIN_CONTRACT_VERSION & _
        "|Workbook=" & ThisWorkbook.Name & _
        "|IsAddin=" & CStr(ThisWorkbook.IsAddin) & _
        "|StartupOperatorMutation=False" & _
        "|LegacyDirectWrites=False" & _
        "|UndoModel=CompensatingEvent" & _
        "|Authority=WHx.invSys.Data.Inventory.xlsb"
End Function

Public Sub ScheduleSourceWorkbookSync(Optional ByVal delaySeconds As Long = 3)
    ' Compatibility entry point for older Core builds. Cross-workbook
    ' canonical pulls bypassed the snapshot contract and broke add-in
    ' isolation, so cancel any old timer and schedule nothing.
    On Error Resume Next
    If gSourceSyncScheduled Then
        Application.OnTime EarliestTime:=gNextSourceSync, _
                           Procedure:=BuildSourceSyncProcedureInit(), _
                           Schedule:=False
    End If
    On Error GoTo 0

    gSourceSyncScheduled = False
    AppendSyncLogEntry "SCHEDULE_DISABLED", _
        "Workbook=" & ThisWorkbook.Name & "|Reason=OperatorReadModelsAreCoreSnapshotOwned"
End Sub

Public Sub SyncSourceWorkbookFromCanonicalRuntime()
    ' Compatibility target for a timer queued by an older loaded build.
    ' Never inspect or mutate open operator workbooks and never reschedule.
    gSourceSyncScheduled = False
    AppendSyncLogEntry "SYNC_DISABLED", _
        "Workbook=" & ThisWorkbook.Name & "|Reason=OperatorReadModelsAreCoreSnapshotOwned"
End Sub

Public Function GetSyncLogPath() As String
    GetSyncLogPath = ResolveSyncLogPathInit()
End Function

Public Sub ResetSyncLog()
    Dim logPath As String

    On Error Resume Next
    logPath = ResolveSyncLogPathInit()
    If Len(Dir$(logPath)) > 0 Then Kill logPath
    On Error GoTo 0
End Sub

Public Sub AppendSyncLogEntry(ByVal tag As String, ByVal valueText As String)
    Dim fileNum As Integer
    Dim logPath As String

    On Error Resume Next
    logPath = ResolveSyncLogPathInit()
    fileNum = FreeFile
    Open logPath For Append As #fileNum
    Print #fileNum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & tag & " | " & valueText
    Close #fileNum
    On Error GoTo 0
End Sub

Private Function ShouldSyncSourceWorkbookInit(ByVal wb As Workbook) As Boolean
    Dim wbName As String

    If wb Is Nothing Then Exit Function
    If wb.IsAddin Then Exit Function

    wbName = LCase$(Trim$(wb.Name))
    If wbName = "" Then Exit Function
    If Left$(wbName, 2) = "~$" Then Exit Function
    If wbName Like "*.xla" Or wbName Like "*.xlam" Then Exit Function
    If wbName Like "*.invsys.*.xls*" Then Exit Function
    If wbName Like "invsys.inbox.*.xls*" Then Exit Function
    If wbName Like "*.outbox.events.xls*" Then Exit Function
    If wbName Like "*.snapshot.inventory.xls*" Then Exit Function

    If wbName Like "*inventory_management*.xls*" Then
        ShouldSyncSourceWorkbookInit = True
        Exit Function
    End If

    ShouldSyncSourceWorkbookInit = WorkbookHasSyncTableInit(wb, "invSys") _
        And (WorkbookHasSyncTableInit(wb, "ReceivedTally") _
             Or WorkbookHasSyncTableInit(wb, "ShipmentsTally") _
             Or WorkbookHasSyncTableInit(wb, "ProductionOutput") _
             Or WorkbookHasSyncTableInit(wb, "Recipes"))
End Function

Private Function IsSourceSyncSchedulerHostInit() As Boolean
    Dim wbName As String

    On Error Resume Next
    wbName = LCase$(Trim$(ThisWorkbook.Name))
    If ThisWorkbook.IsAddin Then
        IsSourceSyncSchedulerHostInit = True
    ElseIf wbName Like "*.xla" Or wbName Like "*.xlam" Then
        IsSourceSyncSchedulerHostInit = True
    End If
    On Error GoTo 0
End Function

Private Function BuildSourceSyncProcedureInit() As String
    BuildSourceSyncProcedureInit = "'" & Replace$(ThisWorkbook.Name, "'", "''") & "'!modInventoryInit.SyncSourceWorkbookFromCanonicalRuntime"
End Function

Private Function WorkbookHasSyncTableInit(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject

    If wb Is Nothing Then Exit Function

    On Error Resume Next
    For Each ws In wb.Worksheets
        Set lo = ws.ListObjects(tableName)
        If Not lo Is Nothing Then
            WorkbookHasSyncTableInit = True
            Exit Function
        End If
        Set lo = Nothing
    Next ws
    On Error GoTo 0
End Function

Private Function ResolveSyncLogPathInit() As String
    Dim rootPath As String

    rootPath = Trim$(Environ$("TEMP"))
    If rootPath = "" Then rootPath = ThisWorkbook.Path
    If rootPath = "" Then rootPath = CurDir$
    If Right$(rootPath, 1) <> "\" Then rootPath = rootPath & "\"

    ResolveSyncLogPathInit = rootPath & SOURCE_SYNC_LOG_FILENAME
End Function
