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

Public Sub CloseInventoryViewerForTest()
    If Not mInventoryViewer Is Nothing Then Unload mInventoryViewer
End Sub
