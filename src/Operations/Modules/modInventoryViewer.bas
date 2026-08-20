Attribute VB_Name = "modInventoryViewer"
Option Explicit

Private mInventoryViewer As frmInventoryViewer
Private mViewerGeneration As Long

Public Sub OpenInventoryViewer()
    Dim warehouseId As String

    If Not modAuth.IsSignedIn() Then
        MsgBox "Sign in to invSys before opening Inventory Viewer.", _
            vbInformation, "invSys Inventory Viewer"
        Exit Sub
    End If
    warehouseId = Trim$(modNasConnection.GetCurrentTargetWarehouseId())
    If warehouseId = "" Then
        MsgBox "Select a warehouse before opening Inventory Viewer.", _
            vbInformation, "invSys Inventory Viewer"
        Exit Sub
    End If

    If mInventoryViewer Is Nothing Then
        Set mInventoryViewer = New frmInventoryViewer
        mViewerGeneration = mViewerGeneration + 1
        mInventoryViewer.SetGeneration mViewerGeneration
    End If
    mInventoryViewer.SetWarehouse warehouseId
    mInventoryViewer.RefreshInventory
    If Not mInventoryViewer.Visible Then mInventoryViewer.Show vbModeless
End Sub

Public Function LoadInventoryViewerEvents(Optional ByVal publishedPayload As String = "") As String
    Dim corePayload As String
    Dim shippingPayload As String
    Dim coreLines As Variant
    Dim shippingLines As Variant
    Dim coreHeader As Variant
    Dim shippingHeader As Variant
    Dim resultText As String
    Dim lineIndex As Long
    Dim rowCount As Long

    corePayload = publishedPayload
    If Trim$(corePayload) = "" Then corePayload = modInventoryViewerData.LoadCurrentInventoryEventViewerData()
    shippingPayload = modTS_Shipments.LoadShippingViewerSupplementEvents()
    coreLines = Split(corePayload, vbCrLf)
    shippingLines = Split(shippingPayload, vbCrLf)
    coreHeader = Split(CStr(coreLines(0)), vbTab)
    shippingHeader = Split(CStr(shippingLines(0)), vbTab)

    If UBound(coreHeader) >= 1 And StrComp(CStr(coreHeader(0)), "OK", vbTextCompare) = 0 Then
        resultText = CStr(coreLines(0))
        If UBound(coreHeader) >= 3 Then rowCount = CLng(Val(CStr(coreHeader(3))))
        For lineIndex = 1 To UBound(coreLines)
            If Trim$(CStr(coreLines(lineIndex))) <> "" Then resultText = resultText & vbCrLf & CStr(coreLines(lineIndex))
        Next lineIndex
    ElseIf UBound(shippingHeader) >= 1 And StrComp(CStr(shippingHeader(0)), "OK", vbTextCompare) = 0 Then
        resultText = "OK" & vbTab & modNasConnection.GetCurrentTargetWarehouseId() & vbTab & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & "0"
    Else
        LoadInventoryViewerEvents = corePayload
        Exit Function
    End If

    If UBound(shippingHeader) >= 1 And StrComp(CStr(shippingHeader(0)), "OK", vbTextCompare) = 0 Then
        If UBound(shippingHeader) >= 3 Then rowCount = rowCount + CLng(Val(CStr(shippingHeader(3))))
        For lineIndex = 1 To UBound(shippingLines)
            If Trim$(CStr(shippingLines(lineIndex))) <> "" Then resultText = resultText & vbCrLf & CStr(shippingLines(lineIndex))
        Next lineIndex
    End If
    coreLines = Split(resultText, vbCrLf)
    coreHeader = Split(CStr(coreLines(0)), vbTab)
    If UBound(coreHeader) >= 3 Then
        coreHeader(3) = CStr(rowCount)
        coreLines(0) = Join(coreHeader, vbTab)
        resultText = Join(coreLines, vbCrLf)
    End If
    LoadInventoryViewerEvents = resultText
End Function

Public Sub UnregisterInventoryViewer(ByVal formInstance As Object)
    On Error Resume Next
    If Not mInventoryViewer Is Nothing Then
        If formInstance Is Nothing Or mInventoryViewer Is formInstance Then Set mInventoryViewer = Nothing
    End If
    On Error GoTo 0
End Sub

Public Function RunInventoryViewerActionForTest() As String
    OpenInventoryViewer
    If mInventoryViewer Is Nothing Then
        RunInventoryViewerActionForTest = "FAIL|FormNotOpen"
    Else
        RunInventoryViewerActionForTest = mInventoryViewer.TestReport()
    End If
End Function

Public Function RunInventoryViewerFilterForTest(ByVal filterText As String) As String
    If mInventoryViewer Is Nothing Then OpenInventoryViewer
    If mInventoryViewer Is Nothing Then
        RunInventoryViewerFilterForTest = "FAIL|FormNotOpen"
    Else
        RunInventoryViewerFilterForTest = mInventoryViewer.TestApplySearch(filterText)
    End If
End Function

Public Function RunInventoryViewerEventsForTest() As String
    If mInventoryViewer Is Nothing Then OpenInventoryViewer
    If mInventoryViewer Is Nothing Then
        RunInventoryViewerEventsForTest = "FAIL|FormNotOpen"
    Else
        RunInventoryViewerEventsForTest = mInventoryViewer.TestEventsReport()
    End If
End Function

Public Sub CloseInventoryViewerForTest()
    If Not mInventoryViewer Is Nothing Then Unload mInventoryViewer
End Sub
