Attribute VB_Name = "TestCoreRoleUiAccess"
Option Explicit

Public Sub RunCoreRoleUiAccessTests()
    Dim passed As Long
    Dim failed As Long

    Tally TestCanCurrentUserPerformCapability_Allow(), passed, failed
    Tally TestCanCurrentUserPerformCapability_Deny(), passed, failed
    Tally TestApplyShapeCapability_TogglesVisibility(), passed, failed

    Debug.Print "Core.RoleUiAccess tests - Passed: " & passed & " Failed: " & failed
End Sub

Public Function TestCanCurrentUserPerformCapability_Allow() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim errorMessage As String

    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WHUI1", "UI1", "RECEIVE")
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WHUI1")
    TestPhase2Helpers.AddCapability wbAuth, "user1", "RECEIVE_POST", "WHUI1", "UI1", "ACTIVE"

    On Error GoTo CleanFail
    If Not modRoleUiAccess.CanCurrentUserPerformCapability("RECEIVE_POST", "user1", "WHUI1", "UI1", errorMessage) Then GoTo CleanExit

    TestCanCurrentUserPerformCapability_Allow = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestCanCurrentUserPerformCapability_Deny() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim errorMessage As String

    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WHUI2", "UI2", "SHIP")
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WHUI2")

    On Error GoTo CleanFail
    If modRoleUiAccess.CanCurrentUserPerformCapability("SHIP_POST", "user1", "WHUI2", "UI2", errorMessage) Then GoTo CleanExit
    If InStr(1, errorMessage, "SHIP_POST", vbTextCompare) = 0 Then GoTo CleanExit

    TestCanCurrentUserPerformCapability_Deny = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestApplyShapeCapability_TogglesVisibility() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbUi As Workbook
    Dim ws As Worksheet
    Dim shp As Shape

    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WHUI3", "UI3", "PROD")
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WHUI3")
    Set wbUi = Application.Workbooks.Add
    Set ws = wbUi.Worksheets(1)
    Set shp = ws.Shapes.AddFormControl(xlButtonControl, 10, 10, 120, 18)
    shp.Name = "btnProdPost"
    shp.TextFrame.Characters.Text = "Post"

    On Error GoTo CleanFail
    modRoleUiAccess.ApplyShapeCapability ws, "btnProdPost", "PROD_POST", "user1", "WHUI3", "UI3"
    If shp.Visible <> msoFalse Then GoTo CleanExit

    TestPhase2Helpers.AddCapability wbAuth, "user1", "PROD_POST", "WHUI3", "UI3", "ACTIVE"
    modRoleUiAccess.ApplyShapeCapability ws, "btnProdPost", "PROD_POST", "user1", "WHUI3", "UI3"
    If shp.Visible <> msoTrue Then GoTo CleanExit

    TestApplyShapeCapability_TogglesVisibility = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbUi
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Sub Tally(ByVal resultIn As Long, ByRef passed As Long, ByRef failed As Long)
    If resultIn = 1 Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
End Sub
