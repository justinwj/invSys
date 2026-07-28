Attribute VB_Name = "modBoxingService"
Option Explicit

Public Function LoadBoxDesigns(ByVal operatorWb As Workbook, _
                               Optional ByVal includeActive As Boolean = True, _
                               Optional ByVal includeArchived As Boolean = False) As Variant
    If operatorWb Is Nothing Then Exit Function
    LoadBoxDesigns = modTS_Shipments.BoxBuilderFormLoadSavedBoxes( _
        includeActive, includeArchived, False, operatorWb)
End Function

Public Function LoadBoxMakerChoices(ByVal operatorWb As Workbook, _
                                    Optional ByRef report As String = "") As Variant
    If operatorWb Is Nothing Then
        report = "The captured Shipping operator workbook was not provided."
        Exit Function
    End If
    LoadBoxMakerChoices = modTS_Shipments.BoxMakerFormLoadSavedBoxes(operatorWb, False)
    report = "OK"
End Function

Public Function LoadBoxDesignVersions(ByVal operatorWb As Workbook, _
                                      ByVal packageRow As Long) As Variant
    If operatorWb Is Nothing Or packageRow <= 0 Then Exit Function
    LoadBoxDesignVersions = modTS_Shipments.BoxBuilderFormLoadVersions( _
        packageRow, operatorWb)
End Function

Public Function LoadBoxDesignComponents(ByVal operatorWb As Workbook, _
                                        ByVal packageRow As Long, _
                                        ByVal versionLabel As String) As Variant
    If operatorWb Is Nothing Or packageRow <= 0 Then Exit Function
    LoadBoxDesignComponents = modTS_Shipments.BoxBuilderFormLoadVersionComponents( _
        packageRow, versionLabel, operatorWb)
End Function

Public Function LoadBoxMakerVersions(ByVal operatorWb As Workbook, _
                                     ByVal packageRow As Long) As Variant
    If operatorWb Is Nothing Or packageRow <= 0 Then Exit Function
    LoadBoxMakerVersions = modTS_Shipments.BoxMakerFormLoadVersions( _
        packageRow, operatorWb)
End Function

Public Function LoadBoxMakerComponents(ByVal operatorWb As Workbook, _
                                       ByVal packageRow As Long, _
                                       ByVal versionLabel As String) As Variant
    If operatorWb Is Nothing Or packageRow <= 0 Then Exit Function
    LoadBoxMakerComponents = modTS_Shipments.BoxMakerFormLoadVersionComponents( _
        packageRow, versionLabel, operatorWb)
End Function

Public Function LoadComponentChoices(ByVal operatorWb As Workbook) As Variant
    If operatorWb Is Nothing Then Exit Function
    LoadComponentChoices = modTS_Shipments.LoadShippingComponentPickerItems(operatorWb)
End Function

Public Function SaveBoxDesign(ByVal operatorWb As Workbook, _
                              ByVal boxName As String, _
                              ByVal boxUom As String, _
                              ByVal boxLocation As String, _
                              ByVal boxDescription As String, _
                              ByVal componentRows As Variant, _
                              ByVal saveAction As String, _
                              ByVal versionLabel As String, _
                              ByVal statusText As String, _
                              ByRef report As String) As Boolean
    If operatorWb Is Nothing Then
        report = "The captured Shipping operator workbook was not provided."
        Exit Function
    End If
    If Not modRoleUiAccess.RequireCurrentUserCapability("SHIP_POST") Then
        report = "SHIP_POST capability is required."
        Exit Function
    End If
    If Trim$(boxName) = "" Or Trim$(boxUom) = "" Or IsEmpty(componentRows) Then
        report = "Box name, UOM, and at least one component are required."
        Exit Function
    End If

    modTS_Shipments.CommitBoxBuilderFormState _
        boxName, boxUom, boxLocation, boxDescription, componentRows, _
        saveAction, versionLabel, statusText, operatorWb
    report = "Box design action completed."
    SaveBoxDesign = True
End Function

Public Function DeleteBoxDesignVersion(ByVal operatorWb As Workbook, _
                                       ByVal packageRow As Long, _
                                       ByVal versionLabel As String, _
                                       ByRef report As String) As Boolean
    If Not ValidateBoxMaintenanceRequest(operatorWb, packageRow, report) Then Exit Function
    If Trim$(versionLabel) = "" Then
        report = "Select a box version before deleting."
        Exit Function
    End If
    DeleteBoxDesignVersion = modTS_Shipments.DeleteBoxDesignVersionForWorkbook( _
        operatorWb, packageRow, versionLabel, report)
End Function

Public Function ArchiveBoxDesign(ByVal operatorWb As Workbook, _
                                 ByVal packageRow As Long, _
                                 ByRef report As String) As Boolean
    If Not ValidateBoxMaintenanceRequest(operatorWb, packageRow, report) Then Exit Function
    ArchiveBoxDesign = modTS_Shipments.ArchiveBoxDesignForWorkbook( _
        operatorWb, packageRow, report)
End Function

Public Function DeleteBoxDesign(ByVal operatorWb As Workbook, _
                                ByVal packageRow As Long, _
                                ByRef report As String) As Boolean
    If Not ValidateBoxMaintenanceRequest(operatorWb, packageRow, report) Then Exit Function
    DeleteBoxDesign = modTS_Shipments.DeleteBoxDesignForWorkbook( _
        operatorWb, packageRow, report)
End Function

Private Function ValidateBoxMaintenanceRequest(ByVal operatorWb As Workbook, _
                                               ByVal packageRow As Long, _
                                               ByRef report As String) As Boolean
    If operatorWb Is Nothing Then
        report = "The captured Shipping operator workbook was not provided."
        Exit Function
    End If
    If packageRow <= 0 Then
        report = "Select a saved box design."
        Exit Function
    End If
    If Not modRoleUiAccess.RequireCurrentUserCapability("ADMIN_MAINT") Then
        report = "ADMIN_MAINT capability is required."
        Exit Function
    End If
    ValidateBoxMaintenanceRequest = True
End Function

Public Function PostBoxMakerAction(ByVal operatorWb As Workbook, _
                                   ByVal packageRow As Long, _
                                   ByVal boxName As String, _
                                   ByVal boxUom As String, _
                                   ByVal boxLocation As String, _
                                   ByVal boxDescription As String, _
                                   ByVal versionLabel As String, _
                                   ByVal boxQty As Double, _
                                   ByVal componentRows As Variant, _
                                   ByVal actionText As String, _
                                   ByRef report As String) As Boolean
    Dim syncCompleted As Boolean

    If operatorWb Is Nothing Then
        report = "The captured Shipping operator workbook was not provided."
        Exit Function
    End If
    If Not modRoleUiAccess.RequireCurrentUserCapability("SHIP_POST") Then
        report = "SHIP_POST capability is required."
        Exit Function
    End If
    PostBoxMakerAction = modTS_Shipments.CommitBoxMakerFormAction( _
        packageRow, boxName, boxUom, boxLocation, boxDescription, _
        versionLabel, boxQty, componentRows, report, actionText, _
        syncCompleted, Empty, operatorWb)
End Function

Public Function NasInventoryIsReadOnly() As Boolean
    ' Boxing services consume NAS Inv as a read-only canonical input.
    NasInventoryIsReadOnly = True
End Function

Public Function ProjectedComponentInventory(ByVal nasInventory As Double, _
                                            ByVal pendingBuildQty As Double, _
                                            ByVal pendingUnboxQty As Double) As Double
    ProjectedComponentInventory = nasInventory - pendingBuildQty + pendingUnboxQty
End Function
