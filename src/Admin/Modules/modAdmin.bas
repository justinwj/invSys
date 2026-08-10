Attribute VB_Name = "modAdmin"
Option Explicit

Private Const ADMIN_DEMO_INVENTORY_QTY As Double = 1000#
Private mSeedCallbackAutomationEnabled As Boolean
Private mSeedCallbackAutomationWarehouseId As String
Private mSeedCallbackAutomationStationId As String
Private mSeedCallbackAutomationUserId As String

Sub Admin_Click()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    Call modAdminConsole.OpenAdminConsole(targetWb, report)
End Sub

Sub Set_CurrentUser()
    modRoleEventWriter.PromptSetCurrentUser
End Sub

Sub Open_CreateDeleteUser()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If Not modLocalAddinsRegistration.EnsureLocalInvSysAddinsRegistered("", report) Then
        MsgBox "Current invSys add-ins are not registered cleanly for this Excel session." & vbCrLf & vbCrLf & _
               report, vbExclamation, "invSys Admin"
        Exit Sub
    End If
    frmCreateDeleteUser.Show
End Sub

Sub Open_CreateWarehouse()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If Not modLocalAddinsRegistration.EnsureLocalInvSysAddinsRegistered("", report) Then
        MsgBox "Current invSys add-ins are not registered cleanly for this Excel session." & vbCrLf & vbCrLf & _
               report, vbExclamation, "invSys Admin"
        Exit Sub
    End If
    frmCreateWarehouse.Show
End Sub

Sub Admin_SetupTesterStation_Click()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If Not modLocalAddinsRegistration.EnsureLocalInvSysAddinsRegistered("", report) Then
        MsgBox "Current invSys add-ins are not registered cleanly for this Excel session." & vbCrLf & vbCrLf & _
               report, vbExclamation, "invSys Admin"
        Exit Sub
    End If
    frmSetupTesterStation.Show
End Sub

Sub Open_SetupTesterStation()
    Admin_SetupTesterStation_Click
End Sub

Sub Open_LastTesterWorkbook()
    If modTesterSetup.OpenTesterReceivingWorkbook("") Then
        MsgBox "Tester receiving workbook opened. Use Refresh Inventory, then run Confirm Writes.", vbInformation, "invSys Admin"
    Else
        MsgBox "No tester receiving workbook is available in this Excel session. Run Generate Test Warehouse first.", vbExclamation, "invSys Admin"
    End If
End Sub

Sub Open_WarehouseDirectory()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    PromptForWarehouseDirectoryRootIfNeeded
    If modAdminConsole.OpenWarehouseDirectory(targetWb, report) Then
        MsgBox "Warehouse directory refreshed.", vbInformation, "invSys Admin"
    Else
        If Len(Trim$(report)) = 0 Then report = "Warehouse directory could not be opened."
        MsgBox report, vbExclamation, "invSys Admin"
    End If
End Sub

Sub Open_Settings()
    Dim report As String
    Dim targetWb As Workbook

    If Not modRoleUiAccess.RequireCurrentUserCapabilityCached("ADMIN_MAINT", "", report) Then Exit Sub
    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If Not modConfig.IsLoaded() Then
        If Not modConfig.LoadConfig("", "") Then
            MsgBox "Canonical config could not be loaded." & vbCrLf & vbCrLf & _
                   modConfig.Validate(), vbExclamation, "invSys Settings"
            Exit Sub
        End If
    End If
    frmAdminSettings.Show
End Sub

Public Function AdminSettingsFormInitializeSmokeForWorkbook(ByVal operatorWb As Workbook) As String
    On Error GoTo FailSmoke

    Dim frm As frmAdminSettings
    Dim detail As String

    Set frm = New frmAdminSettings
    detail = frm.TestInitializeConfigEditor()
    If InStr(1, detail, "Rows=28", vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 7320, "AdminSettingsFormInitializeSmokeForWorkbook", _
                  "Admin config editor did not load all 28 config keys. " & detail
    End If
    AdminSettingsFormInitializeSmokeForWorkbook = "OK|" & detail

CleanExit:
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    Set frm = Nothing
    On Error GoTo 0
    Exit Function

FailSmoke:
    AdminSettingsFormInitializeSmokeForWorkbook = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
    Resume CleanExit
End Function

Sub Add_WarehouseDirectoryRoot()
    Dim report As String
    Dim targetWb As Workbook
    Dim rootPath As String

    rootPath = InputBox("Enter a NAS/server warehouse hub folder or a specific warehouse runtime folder to include in Admin warehouse scans.", _
                        "invSys Admin - Warehouse Root", _
                        "\\100.84.136.19\invSysWH1")
    rootPath = Trim$(rootPath)
    If rootPath = "" Then Exit Sub

    modAdminConsole.RememberWarehouseScanRoot rootPath
    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If modAdminConsole.OpenWarehouseDirectory(targetWb, report) Then
        MsgBox "Warehouse root remembered and directory refreshed.", vbInformation, "invSys Admin"
    Else
        If Len(Trim$(report)) = 0 Then report = "Warehouse directory could not be opened."
        MsgBox report, vbExclamation, "invSys Admin"
    End If
End Sub

Sub Seed_DemoInventory()
    Dim warehouseId As String
    Dim stationId As String
    Dim userId As String
    Dim report As String
    Dim stage As String

    On Error GoTo FailSeedCallback

    stage = "context resolution"
    If Not ResolveSeedInventoryContext(warehouseId, stationId, userId, report) Then
        If Not mSeedCallbackAutomationEnabled Then MsgBox report, vbExclamation, "invSys Admin"
        GoTo CleanExit
    End If

    stage = "queue and processor application"
    If modAdminInventorySeed.SeedDemoInventoryForWarehouse(warehouseId, stationId, userId, report) Then
        If Not mSeedCallbackAutomationEnabled Then MsgBox report, vbInformation, "invSys Admin"
    Else
        If Not mSeedCallbackAutomationEnabled Then MsgBox report, vbExclamation, "invSys Admin"
    End If

CleanExit:
    mSeedCallbackAutomationEnabled = False
    mSeedCallbackAutomationWarehouseId = ""
    mSeedCallbackAutomationStationId = ""
    mSeedCallbackAutomationUserId = ""
    Exit Sub

FailSeedCallback:
    report = "Seed Demo Inventory failed at " & stage & "." & vbCrLf & _
             "Error " & CStr(Err.Number) & vbCrLf & _
             "Source: " & SanitizeSeedCallbackErrorText(Err.Source) & vbCrLf & _
             SanitizeSeedCallbackErrorText(Err.Description)
    If Not mSeedCallbackAutomationEnabled Then MsgBox report, vbExclamation, "invSys Admin"
    Resume CleanExit
End Sub

Public Sub SetSeedInventorySelectionForAutomation(ByVal warehouseId As String, _
                                                  ByVal stationId As String, _
                                                  ByVal userId As String)
    mSeedCallbackAutomationWarehouseId = Trim$(warehouseId)
    mSeedCallbackAutomationStationId = Trim$(stationId)
    mSeedCallbackAutomationUserId = Trim$(userId)
    mSeedCallbackAutomationEnabled = True
End Sub

Sub Add_InventoryItem()
    Dim warehouseId As String
    Dim stationId As String
    Dim userId As String
    Dim report As String
    Dim sku As String
    Dim rowVal As Long
    Dim defaultLocation As String
    Dim catalogItems As Object
    Dim addForm As frmAddInventoryItem
    Dim accepted As Boolean
    Dim isEdit As Boolean
    Dim formSku As String
    Dim formRow As Long
    Dim formItemName As String
    Dim formUom As String
    Dim formLocation As String
    Dim formQty As Double
    Dim formDescription As String
    Dim formVendorName As String
    Dim formVendorCode As String
    Dim formCategory As String
    Dim formExternalCode As String
    Dim formImagePath As String
    Dim formEditReason As String
    Dim formCustomFields As Object
    Dim editStamp As String
    Dim actionSucceeded As Boolean

    If Not ResolveAdminCurrentTargetContext(warehouseId, stationId, userId, report) Then
        MsgBox report, vbExclamation, "invSys Admin"
        Exit Sub
    End If

    rowVal = NextInventoryRowSuggestionAdmin(warehouseId)
    sku = GenerateInventorySkuAdmin(warehouseId, rowVal)
    defaultLocation = Trim$(modConfig.GetString("DefaultLocation", ""))
    Set catalogItems = LoadInventoryCatalogItemsAdmin(warehouseId)

    Set addForm = New frmAddInventoryItem
    addForm.Configure warehouseId, stationId, userId, sku, rowVal, defaultLocation
    LoadCatalogItemsIntoAddInventoryForm addForm, catalogItems

    Do
        addForm.Show vbModal

        On Error GoTo FormUnavailable
        accepted = addForm.Accepted
        On Error GoTo 0
        If Not accepted Then Exit Do

        On Error GoTo FormUnavailable
        isEdit = addForm.EditMode
        formSku = addForm.GeneratedSku
        formRow = addForm.GeneratedRow
        formItemName = addForm.ItemName
        formUom = addForm.Uom
        formLocation = addForm.LocationValue
        formQty = addForm.StartingQty
        formDescription = addForm.DescriptionValue
        formVendorName = addForm.VendorName
        formVendorCode = addForm.VendorCode
        formCategory = addForm.Category
        formExternalCode = addForm.ExternalCode
        formImagePath = addForm.ImagePath
        formEditReason = addForm.EditReason
        Set formCustomFields = addForm.CustomFields
        On Error GoTo 0

        report = vbNullString
        actionSucceeded = False
        If isEdit Then
            editStamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")
            formCustomFields("LAST_EDIT_REASON") = formEditReason
            formCustomFields("LAST_EDIT_AT") = editStamp
            formCustomFields("LAST_EDIT_USER") = userId
            formCustomFields("EDIT_HISTORY_APPEND") = editStamp & " | User=" & userId & " | SKU=" & formSku & " | Reason=" & formEditReason
            actionSucceeded = UpdateInventoryItemCatalogForWarehouse(warehouseId, formSku, formItemName, _
                                                                     formUom, formLocation, _
                                                                     formDescription, formVendorName, _
                                                                     formVendorCode, formCategory, _
                                                                     formExternalCode, formImagePath, _
                                                                     formCustomFields, report)
            If actionSucceeded And Not IsNonCountedCustomFieldsAdmin(formCustomFields) Then
                Dim qtyReport As String
                If SetInventoryQuantityForWarehouse(warehouseId, stationId, userId, formRow, formSku, _
                                                    formItemName, formUom, formLocation, formQty, formEditReason, qtyReport) Then
                    report = report & vbCrLf & vbCrLf & qtyReport
                Else
                    report = "Inventory item catalog fields were updated, but set qty failed." & vbCrLf & qtyReport
                    actionSucceeded = False
                End If
            End If
        Else
            actionSucceeded = AddInventoryItemForWarehouse(warehouseId, stationId, userId, formRow, _
                                                           formSku, formItemName, _
                                                           formUom, formLocation, _
                                                           formQty, formDescription, _
                                                           formVendorName, formVendorCode, _
                                                           formCategory, formExternalCode, _
                                                           formImagePath, formCustomFields, report)
        End If

        If actionSucceeded Then
            MsgBox report, vbInformation, "invSys Admin"
            rowVal = NextInventoryRowSuggestionAdmin(warehouseId)
            sku = GenerateInventorySkuAdmin(warehouseId, rowVal)
            Set catalogItems = LoadInventoryCatalogItemsAdmin(warehouseId)
            addForm.Configure warehouseId, stationId, userId, sku, rowVal, defaultLocation
            LoadCatalogItemsIntoAddInventoryForm addForm, catalogItems
        Else
            MsgBox report, vbExclamation, "invSys Admin"
        End If
    Loop

CleanExit:
    On Error Resume Next
    Unload addForm
    On Error GoTo 0
    Exit Sub

FormUnavailable:
    report = Err.Description
    On Error GoTo 0
    If Trim$(report) <> "" Then
        If InStr(1, report, "callee", vbTextCompare) = 0 _
           And InStr(1, report, "connections are invalid", vbTextCompare) = 0 Then
            MsgBox report, vbExclamation, "invSys Admin"
        End If
    End If
    Resume CleanExit
End Sub

Public Function AddInventoryItemForWarehouse(ByVal warehouseId As String, _
                                             ByVal stationId As String, _
                                             ByVal userId As String, _
                                             ByVal rowVal As Long, _
                                             ByVal sku As String, _
                                             ByVal itemName As String, _
                                             ByVal uom As String, _
                                             ByVal locationVal As String, _
                                             ByVal qty As Double, _
                                             Optional ByVal description As String = "", _
                                             Optional ByVal vendorName As String = "", _
                                             Optional ByVal vendorCode As String = "", _
                                             Optional ByVal category As String = "", _
                                             Optional ByVal externalCode As String = "", _
                                             Optional ByVal imagePath As String = "", _
                                             Optional ByVal customFields As Object = Nothing, _
                                             Optional ByRef report As String = "") As Boolean
    Dim payloadItems As Collection
    Dim item As Object
    Dim customKey As Variant
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim queueError As String
    Dim batchReport As String
    Dim processedCount As Long
    Dim inboxReport As String

    On Error GoTo FailAdd

    warehouseId = Trim$(warehouseId)
    stationId = Trim$(stationId)
    userId = Trim$(userId)
    sku = Trim$(sku)
    itemName = Trim$(itemName)
    uom = Trim$(uom)
    locationVal = Trim$(locationVal)

    If warehouseId = "" Then report = "WarehouseId is required.": Exit Function
    If stationId = "" Then stationId = "S1"
    If userId = "" Then report = "Admin user is required.": Exit Function
    If sku = "" Then report = "SKU is required.": Exit Function
    If itemName = "" Then itemName = sku
    If uom = "" Then report = "UOM is required.": Exit Function
    If rowVal <= 0 Then report = "Inventory ROW id must be positive.": Exit Function
    If IsNonCountedCustomFieldsAdmin(customFields) Then
        qty = 0#
    ElseIf qty <= 0 Then
        report = "Starting quantity must be greater than zero."
        Exit Function
    End If

    If Not EnsureDemoStationInboxes(warehouseId, stationId, inboxReport) Then
        report = inboxReport
        Exit Function
    End If

    Set payloadItems = New Collection
    Set item = modRoleEventWriter.CreatePayloadItem(rowVal, sku, qty, locationVal, "Admin add inventory item", "IMPORT")
    item("ITEM_CODE") = sku
    item("ITEM") = itemName
    item("UOM") = uom
    item("LOCATION") = locationVal
    item("TOTAL INV") = qty
    item("QtyAvailable") = qty
    item("DESCRIPTION") = description
    item("VENDOR(s)") = vendorName
    item("VENDOR_CODE") = vendorCode
    item("CATEGORY") = category
    item("EXTERNAL_CODE") = externalCode
    AddPictureReferencesToPayloadAdmin item, imagePath
    If IsNonCountedCustomFieldsAdmin(customFields) Then
        item("TRACK_QTY") = "FALSE"
        item("ITEM_KIND") = NonCountedItemKindAdmin(customFields)
    End If
    If Not customFields Is Nothing Then
        For Each customKey In customFields.Keys
            If Trim$(CStr(customKey)) <> "" Then item(Trim$(CStr(customKey))) = customFields(customKey)
        Next customKey
    End If
    payloadItems.Add item

    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)
    If payloadJson = "" Or payloadJson = "[]" Then
        report = "Inventory item payload was empty."
        Exit Function
    End If

    If Not modRoleEventWriter.QueueMigrationSeedEvent(warehouseId, stationId, userId, payloadJson, _
                                                      "ADMIN_ADD_INVENTORY_ITEM", "Admin add inventory item " & sku, _
                                                      0, Nothing, eventIdOut, queueError, "") Then
        report = "Inventory item event could not be queued: " & queueError & vbCrLf & _
                 "Use Users & Roles to grant ADMIN_MAINT to '" & userId & "' for " & warehouseId & " / " & stationId & "."
        Exit Function
    End If

    processedCount = modProcessor.RunBatch(warehouseId, 0, batchReport)
    If processedCount < 1 Then
        report = "Inventory item event was queued but not applied. " & batchReport
        Exit Function
    End If

    report = "Inventory item added." & vbCrLf & _
             "Warehouse: " & warehouseId & vbCrLf & _
             "SKU: " & sku & vbCrLf & _
             "Item: " & itemName & vbCrLf & _
             "Starting quantity: " & IIf(IsNonCountedCustomFieldsAdmin(customFields), "not counted", CStr(qty)) & vbCrLf & _
             "Processor: " & batchReport & vbCrLf & _
             "Refresh inventory in any open role workbook to see the new item."
    AddInventoryItemForWarehouse = True
    Exit Function

FailAdd:
    report = "AddInventoryItem failed: " & Err.Description
End Function

Private Function AddInventoryQuantityForWarehouse(ByVal warehouseId As String, _
                                                  ByVal stationId As String, _
                                                  ByVal userId As String, _
                                                  ByVal rowVal As Long, _
                                                  ByVal sku As String, _
                                                  ByVal itemName As String, _
                                                  ByVal uom As String, _
                                                  ByVal locationVal As String, _
                                                  ByVal qty As Double, _
                                                  ByRef report As String) As Boolean
    Dim payloadItems As Collection
    Dim item As Object
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim queueError As String
    Dim batchReport As String
    Dim processedCount As Long
    Dim inboxReport As String

    On Error GoTo FailAdjust

    warehouseId = Trim$(warehouseId)
    stationId = Trim$(stationId)
    userId = Trim$(userId)
    sku = Trim$(sku)
    itemName = Trim$(itemName)
    uom = Trim$(uom)
    locationVal = Trim$(locationVal)

    If warehouseId = "" Then report = "WarehouseId is required.": Exit Function
    If stationId = "" Then stationId = "S1"
    If userId = "" Then report = "Admin user is required.": Exit Function
    If sku = "" Then report = "SKU is required.": Exit Function
    If rowVal <= 0 Then report = "Inventory ROW id must be positive.": Exit Function
    If qty <= 0 Then report = "Add qty must be greater than zero.": Exit Function

    If Not EnsureDemoStationInboxes(warehouseId, stationId, inboxReport) Then
        report = inboxReport
        Exit Function
    End If

    Set payloadItems = New Collection
    Set item = modRoleEventWriter.CreatePayloadItem(rowVal, sku, qty, locationVal, "Admin add inventory quantity", "IMPORT")
    item("ITEM_CODE") = sku
    item("ITEM") = itemName
    item("UOM") = uom
    item("LOCATION") = locationVal
    payloadItems.Add item

    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)
    If payloadJson = "" Or payloadJson = "[]" Then
        report = "Inventory quantity payload was empty."
        Exit Function
    End If

    If Not modRoleEventWriter.QueueMigrationSeedEvent(warehouseId, stationId, userId, payloadJson, _
                                                      "ADMIN_ADD_INVENTORY_QTY", "Admin add inventory quantity " & sku, _
                                                      0, Nothing, eventIdOut, queueError, "") Then
        report = "Inventory quantity event could not be queued: " & queueError & vbCrLf & _
                 "Use Users & Roles to grant ADMIN_MAINT to '" & userId & "' for " & warehouseId & " / " & stationId & "."
        Exit Function
    End If

    processedCount = modProcessor.RunBatch(warehouseId, 0, batchReport)
    If processedCount < 1 Then
        report = "Inventory quantity event was queued but not applied. " & batchReport
        Exit Function
    End If

    report = "Inventory quantity added." & vbCrLf & _
             "Warehouse: " & warehouseId & vbCrLf & _
             "SKU: " & sku & vbCrLf & _
             "Item: " & itemName & vbCrLf & _
             "Added quantity: " & CStr(qty) & vbCrLf & _
             "Processor: " & batchReport & vbCrLf & _
             "Refresh inventory in any open role workbook to see the updated quantity."
    AddInventoryQuantityForWarehouse = True
    Exit Function

FailAdjust:
    report = "AddInventoryQuantity failed: " & Err.Description
End Function

Private Function SetInventoryQuantityForWarehouse(ByVal warehouseId As String, _
                                                  ByVal stationId As String, _
                                                  ByVal userId As String, _
                                                  ByVal rowVal As Long, _
                                                  ByVal sku As String, _
                                                  ByVal itemName As String, _
                                                  ByVal uom As String, _
                                                  ByVal locationVal As String, _
                                                  ByVal targetQty As Double, _
                                                  ByVal editReason As String, _
                                                  ByRef report As String) As Boolean
    Dim currentQty As Double
    Dim deltaQty As Double
    Dim payloadItems As Collection
    Dim item As Object
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim queueError As String
    Dim batchReport As String
    Dim processedCount As Long
    Dim inboxReport As String
    Dim noteText As String

    On Error GoTo FailSet

    warehouseId = Trim$(warehouseId)
    stationId = Trim$(stationId)
    userId = Trim$(userId)
    sku = Trim$(sku)
    itemName = Trim$(itemName)
    uom = Trim$(uom)
    locationVal = Trim$(locationVal)
    editReason = Trim$(editReason)

    If warehouseId = "" Then report = "WarehouseId is required.": Exit Function
    If stationId = "" Then stationId = "S1"
    If userId = "" Then report = "Admin user is required.": Exit Function
    If sku = "" Then report = "SKU is required.": Exit Function
    If rowVal <= 0 Then report = "Inventory ROW id must be positive.": Exit Function
    If targetQty < 0 Then report = "Set qty cannot be negative.": Exit Function
    If editReason = "" Then report = "Why the edit is required before changing quantity.": Exit Function

    currentQty = ResolveInventoryQtyOnHandAdmin(warehouseId, sku)
    deltaQty = targetQty - currentQty
    If Abs(deltaQty) < 0.0000001 Then
        report = "Inventory quantity already matched target." & vbCrLf & _
                 "Warehouse: " & warehouseId & vbCrLf & _
                 "SKU: " & sku & vbCrLf & _
                 "Target quantity: " & CStr(targetQty)
        SetInventoryQuantityForWarehouse = True
        Exit Function
    End If

    If Not EnsureDemoStationInboxes(warehouseId, stationId, inboxReport) Then
        report = inboxReport
        Exit Function
    End If

    noteText = "Admin set inventory quantity. Reason: " & editReason & _
               "; PreviousQty=" & CStr(currentQty) & "; TargetQty=" & CStr(targetQty)

    Set payloadItems = New Collection
    Set item = modRoleEventWriter.CreatePayloadItem(rowVal, sku, deltaQty, locationVal, noteText, "ADJUST")
    item("ITEM_CODE") = sku
    item("ITEM") = itemName
    item("UOM") = uom
    item("LOCATION") = locationVal
    item("PreviousQty") = currentQty
    item("TargetQty") = targetQty
    item("Reason") = editReason
    payloadItems.Add item

    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)
    If payloadJson = "" Or payloadJson = "[]" Then
        report = "Inventory adjustment payload was empty."
        Exit Function
    End If

    If Not modRoleEventWriter.QueueAdminInventoryAdjustEvent(warehouseId, stationId, userId, payloadJson, _
                                                             noteText, 0, eventIdOut, queueError, "") Then
        report = "Inventory adjustment event could not be queued: " & queueError & vbCrLf & _
                 "Use Users & Roles to grant ADMIN_MAINT to '" & userId & "' for " & warehouseId & " / " & stationId & "."
        Exit Function
    End If

    processedCount = modProcessor.RunBatch(warehouseId, 0, batchReport)
    If processedCount < 1 Then
        report = "Inventory adjustment event was queued but not applied. " & batchReport
        Exit Function
    End If

    report = "Inventory quantity set." & vbCrLf & _
             "Warehouse: " & warehouseId & vbCrLf & _
             "SKU: " & sku & vbCrLf & _
             "Item: " & itemName & vbCrLf & _
             "Previous quantity: " & CStr(currentQty) & vbCrLf & _
             "Target quantity: " & CStr(targetQty) & vbCrLf & _
             "Adjustment delta: " & CStr(deltaQty) & vbCrLf & _
             "Reason: " & editReason & vbCrLf & _
             "Processor: " & batchReport & vbCrLf & _
             "Refresh inventory in any open role workbook to see the updated quantity."
    SetInventoryQuantityForWarehouse = True
    Exit Function

FailSet:
    report = "SetInventoryQuantity failed: " & Err.Description
End Function

Public Function UpdateInventoryItemCatalogForWarehouse(ByVal warehouseId As String, _
                                                       ByVal sku As String, _
                                                       ByVal itemName As String, _
                                                       ByVal uom As String, _
                                                       ByVal locationVal As String, _
                                                       Optional ByVal description As String = "", _
                                                       Optional ByVal vendorName As String = "", _
                                                       Optional ByVal vendorCode As String = "", _
                                                       Optional ByVal category As String = "", _
                                                       Optional ByVal externalCode As String = "", _
                                                       Optional ByVal imagePath As String = "", _
                                                       Optional ByVal customFields As Object = Nothing, _
                                                       Optional ByRef report As String = "") As Boolean
    Dim path As String
    Dim wb As Workbook
    Dim openedHere As Boolean
    Dim lo As ListObject
    Dim loBalance As ListObject
    Dim rowIndex As Long
    Dim customKey As Variant

    On Error GoTo FailUpdate

    warehouseId = Trim$(warehouseId)
    sku = Trim$(sku)
    itemName = Trim$(itemName)
    uom = Trim$(uom)

    If warehouseId = "" Then report = "WarehouseId is required.": Exit Function
    If sku = "" Then report = "SKU is required.": Exit Function
    If itemName = "" Then report = "Item name is required.": Exit Function
    If uom = "" Then report = "UOM is required.": Exit Function

    path = modProcessor.ResolveInventoryWorkbookPathForAutomation(warehouseId)
    If path = "" Then report = "Inventory workbook path could not be resolved for " & warehouseId & ".": Exit Function

    Set wb = FindOpenWorkbookByFullNameAdmin(path)
    If wb Is Nothing Then
        If Len(Dir$(path, vbNormal)) = 0 Then
            report = "Inventory workbook was not found: " & path
            Exit Function
        End If
        Set wb = Application.Workbooks.Open(path, UpdateLinks:=False, ReadOnly:=False, AddToMru:=False)
        openedHere = True
    End If
    If wb.ReadOnly Then
        report = "Inventory workbook is open read-only. Close other copies and try again."
        GoTo CleanExit
    End If

    Set lo = FindListObjectByNameAdminLocal(wb, "tblSkuCatalog")
    If lo Is Nothing Then
        report = "tblSkuCatalog was not found in " & wb.Name & "."
        GoTo CleanExit
    End If
    rowIndex = FindCatalogRowBySkuAdmin(lo, sku)
    If rowIndex <= 0 Then
        report = "Inventory item was not found in catalog: " & sku
        GoTo CleanExit
    End If

    SetSheetProtectionAdminLocal lo.Parent, False
    SetCatalogValueAdmin lo, rowIndex, "ITEM_CODE", sku
    SetCatalogValueAdmin lo, rowIndex, "ITEM", itemName
    SetCatalogValueAdmin lo, rowIndex, "UOM", uom
    SetCatalogValueAdmin lo, rowIndex, "LOCATION", locationVal
    SetCatalogValueAdmin lo, rowIndex, "DESCRIPTION", description
    SetCatalogValueAdmin lo, rowIndex, "VENDOR(s)", vendorName
    SetCatalogValueAdmin lo, rowIndex, "VENDOR_CODE", vendorCode
    SetCatalogValueAdmin lo, rowIndex, "CATEGORY", category
    SetCatalogValueAdmin lo, rowIndex, "EXTERNAL_CODE", externalCode
    SetPictureReferencesInCatalogAdmin lo, rowIndex, imagePath
    If Not customFields Is Nothing Then
        For Each customKey In customFields.Keys
            If Trim$(CStr(customKey)) <> "" Then
                If StrComp(Trim$(CStr(customKey)), "EDIT_HISTORY_APPEND", vbTextCompare) = 0 Then
                    AppendCatalogEditHistoryAdmin lo, rowIndex, CStr(customFields(customKey))
                Else
                    SetCatalogValueAdmin lo, rowIndex, Trim$(CStr(customKey)), customFields(customKey)
                End If
            End If
        Next customKey
    End If
    SetSheetProtectionAdminLocal lo.Parent, True
    wb.Save

    report = "Inventory item updated." & vbCrLf & _
             "Warehouse: " & warehouseId & vbCrLf & _
             "SKU: " & sku & vbCrLf & _
             "Item: " & itemName & vbCrLf & _
             "Refresh inventory in any open role workbook to see the updated catalog fields."
    UpdateInventoryItemCatalogForWarehouse = True

CleanExit:
    On Error Resume Next
    If Not lo Is Nothing Then SetSheetProtectionAdminLocal lo.Parent, True
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function

FailUpdate:
    report = "UpdateInventoryItemCatalog failed: " & Err.Description
    Resume CleanExit
End Function

Private Sub AppendCatalogEditHistoryAdmin(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal entryText As String)
    Dim existingText As String

    entryText = Trim$(entryText)
    If entryText = "" Then Exit Sub

    existingText = CatalogCellAdmin(lo, rowIndex, "EDIT_HISTORY")
    If existingText <> "" Then
        SetCatalogValueAdmin lo, rowIndex, "EDIT_HISTORY", existingText & vbLf & entryText
    Else
        SetCatalogValueAdmin lo, rowIndex, "EDIT_HISTORY", entryText
    End If
End Sub

Private Function GenerateInventorySkuAdmin(ByVal warehouseId As String, ByVal rowVal As Long) As String
    Dim rowPart As String
    Dim whPart As String

    whPart = UCase$(Trim$(warehouseId))
    whPart = Replace(whPart, " ", "")
    whPart = Replace(whPart, "-", "")
    If Len(whPart) > 4 Then whPart = Left$(whPart, 4)
    If whPart = "" Then whPart = "INV"

    rowPart = Base36Admin(rowVal)
    Do While Len(rowPart) < 6
        rowPart = "0" & rowPart
    Loop
    GenerateInventorySkuAdmin = "ITM-" & whPart & "-" & rowPart
End Function

Private Function LoadInventoryCatalogItemsAdmin(ByVal warehouseId As String) As Collection
    Dim path As String
    Dim wb As Workbook
    Dim openedHere As Boolean
    Dim lo As ListObject
    Dim loBalance As ListObject
    Dim rowIndex As Long
    Dim item As Object
    Dim result As Collection

    On Error GoTo CleanExit
    Set result = New Collection
    path = modProcessor.ResolveInventoryWorkbookPathForAutomation(warehouseId)
    If path <> "" Then
        Set wb = FindOpenWorkbookByFullNameAdmin(path)
        If wb Is Nothing Then
            If Len(Dir$(path, vbNormal)) > 0 Then
                Set wb = Application.Workbooks.Open(path, UpdateLinks:=False, ReadOnly:=True, AddToMru:=False)
                openedHere = True
            End If
        End If
    End If

    If Not wb Is Nothing Then
        Set lo = FindListObjectByNameAdminLocal(wb, "tblSkuCatalog")
        Set loBalance = FindListObjectByNameAdminLocal(wb, "tblSkuBalance")
        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                For rowIndex = 1 To lo.ListRows.Count
                    If CatalogCellAdmin(lo, rowIndex, "SKU") <> "" Then
                        Set item = CreateObject("Scripting.Dictionary")
                        item.CompareMode = vbTextCompare
                        item("SKU") = CatalogCellAdmin(lo, rowIndex, "SKU")
                        item("ROW") = CatalogCellAdmin(lo, rowIndex, "ROW")
                        item("ITEM") = CatalogCellAdmin(lo, rowIndex, "ITEM")
                        item("UOM") = CatalogCellAdmin(lo, rowIndex, "UOM")
                        item("LOCATION") = CatalogCellAdmin(lo, rowIndex, "LOCATION")
                        item("DESCRIPTION") = CatalogCellAdmin(lo, rowIndex, "DESCRIPTION")
                        item("VENDOR(s)") = CatalogCellAdmin(lo, rowIndex, "VENDOR(s)")
                        item("VENDOR_CODE") = CatalogCellAdmin(lo, rowIndex, "VENDOR_CODE")
                        item("CATEGORY") = CatalogCellAdmin(lo, rowIndex, "CATEGORY")
                        item("EXTERNAL_CODE") = CatalogCellAdmin(lo, rowIndex, "EXTERNAL_CODE")
                        item("IMAGE_PATH") = CombinedPictureReferencesAdmin(lo, rowIndex)
                        item("TRACK_QTY") = CatalogCellAdmin(lo, rowIndex, "TRACK_QTY")
                        item("ITEM_KIND") = CatalogCellAdmin(lo, rowIndex, "ITEM_KIND")
                        item("QTY_ON_HAND") = ResolveQtyOnHandFromBalanceAdmin(loBalance, item("SKU"))
                        result.Add item
                    End If
                Next rowIndex
            End If
        End If
    End If

CleanExit:
    On Error Resume Next
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Set LoadInventoryCatalogItemsAdmin = result
End Function

Private Sub LoadCatalogItemsIntoAddInventoryForm(ByVal targetForm As frmAddInventoryItem, ByVal catalogItems As Object)
    Dim item As Variant

    If targetForm Is Nothing Then Exit Sub
    If catalogItems Is Nothing Then Exit Sub
    For Each item In catalogItems
        targetForm.AddCatalogItem CatalogItemTextAdmin(item, "SKU"), _
                                  CatalogItemTextAdmin(item, "ROW"), _
                                  CatalogItemTextAdmin(item, "ITEM"), _
                                  CatalogItemTextAdmin(item, "UOM"), _
                                  CatalogItemTextAdmin(item, "LOCATION"), _
                                  CatalogItemTextAdmin(item, "DESCRIPTION"), _
                                  CatalogItemTextAdmin(item, "VENDOR(s)"), _
                                  CatalogItemTextAdmin(item, "VENDOR_CODE"), _
                                  CatalogItemTextAdmin(item, "CATEGORY"), _
                                  CatalogItemTextAdmin(item, "EXTERNAL_CODE"), _
                                  CatalogItemTextAdmin(item, "IMAGE_PATH"), _
                                  CatalogItemTextAdmin(item, "TRACK_QTY"), _
                                  CatalogItemTextAdmin(item, "ITEM_KIND"), _
                                  CatalogItemTextAdmin(item, "QTY_ON_HAND")
    Next item
End Sub

Private Function IsNonCountedCustomFieldsAdmin(ByVal customFields As Object) As Boolean
    Dim trackQty As String
    Dim itemKind As String

    On Error Resume Next
    If Not customFields Is Nothing Then
        trackQty = UCase$(Trim$(CStr(customFields("TRACK_QTY"))))
        itemKind = UCase$(Trim$(CStr(customFields("ITEM_KIND"))))
    End If
    On Error GoTo 0

    IsNonCountedCustomFieldsAdmin = (trackQty = "FALSE" Or trackQty = "NO" Or trackQty = "0" _
                                     Or itemKind = "UTILITY" Or itemKind = "SERVICE" Or itemKind = "NON_COUNTED")
End Function

Private Function NonCountedItemKindAdmin(ByVal customFields As Object) As String
    Dim itemKind As String

    On Error Resume Next
    If Not customFields Is Nothing Then itemKind = UCase$(Trim$(CStr(customFields("ITEM_KIND"))))
    On Error GoTo 0

    Select Case itemKind
        Case "UTILITY", "SERVICE", "NON_COUNTED"
            NonCountedItemKindAdmin = itemKind
        Case Else
            NonCountedItemKindAdmin = "NON_COUNTED"
    End Select
End Function

Private Function CatalogItemTextAdmin(ByVal item As Variant, ByVal fieldName As String) As String
    On Error Resume Next
    CatalogItemTextAdmin = Trim$(CStr(item(fieldName)))
    On Error GoTo 0
End Function

Private Function ResolveInventoryQtyOnHandAdmin(ByVal warehouseId As String, ByVal sku As String) As Double
    Dim path As String
    Dim wb As Workbook
    Dim openedHere As Boolean
    Dim loBalance As ListObject

    On Error GoTo CleanExit
    warehouseId = Trim$(warehouseId)
    sku = Trim$(sku)
    If warehouseId = "" Or sku = "" Then Exit Function

    path = modProcessor.ResolveInventoryWorkbookPathForAutomation(warehouseId)
    If path = "" Then Exit Function
    Set wb = FindOpenWorkbookByFullNameAdmin(path)
    If wb Is Nothing Then
        If Len(Dir$(path, vbNormal)) = 0 Then Exit Function
        Set wb = Application.Workbooks.Open(path, UpdateLinks:=False, ReadOnly:=True, AddToMru:=False)
        openedHere = True
    End If
    Set loBalance = FindListObjectByNameAdminLocal(wb, "tblSkuBalance")
    ResolveInventoryQtyOnHandAdmin = ResolveQtyOnHandFromBalanceAdmin(loBalance, sku)

CleanExit:
    On Error Resume Next
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Function

Private Function ResolveQtyOnHandFromBalanceAdmin(ByVal loBalance As ListObject, ByVal sku As String) As Double
    Dim cSku As Long
    Dim cQty As Long
    Dim rowIndex As Long

    If loBalance Is Nothing Then Exit Function
    If loBalance.DataBodyRange Is Nothing Then Exit Function
    sku = Trim$(sku)
    If sku = "" Then Exit Function
    cSku = ColumnIndexAdminLocal(loBalance, "SKU")
    cQty = ColumnIndexAdminLocal(loBalance, "QtyOnHand")
    If cSku = 0 Or cQty = 0 Then Exit Function
    For rowIndex = 1 To loBalance.ListRows.Count
        If StrComp(Trim$(CStr(loBalance.DataBodyRange.Cells(rowIndex, cSku).Value)), sku, vbTextCompare) = 0 Then
            ResolveQtyOnHandFromBalanceAdmin = CDbl(Val(CStr(loBalance.DataBodyRange.Cells(rowIndex, cQty).Value)))
            Exit Function
        End If
    Next rowIndex
End Function

Private Function Base36Admin(ByVal valueIn As Long) As String
    Const DIGITS As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim n As Long
    Dim result As String

    If valueIn <= 0 Then
        Base36Admin = "0"
        Exit Function
    End If
    n = valueIn
    Do While n > 0
        result = Mid$(DIGITS, (n Mod 36) + 1, 1) & result
        n = n \ 36
    Loop
    Base36Admin = result
End Function

Private Function ResolveSeedInventoryContext(ByRef warehouseId As String, _
                                             ByRef stationId As String, _
                                             ByRef userId As String, _
                                             ByRef report As String) As Boolean
    Dim warehouseOptions As Collection
    Dim runtimeRoot As String
    Dim formReport As String
    Dim item As Variant
    Dim selectionFound As Boolean

    warehouseId = Trim$(modConfig.GetWarehouseId())
    stationId = Trim$(modConfig.GetStationId())
    If warehouseId = "" Then warehouseId = Trim$(modConfig.GetString("WarehouseId", ""))
    If stationId = "" Then stationId = Trim$(modConfig.GetString("StationId", "S1"))

    userId = Trim$(modRoleEventWriter.ResolveCurrentUserId())
    If userId = "" Then userId = Trim$(Application.UserName)

    Set warehouseOptions = modAdminConsole.GetWarehouseDirectoryOptions(Nothing, formReport, True)
    If warehouseOptions Is Nothing Then
        report = "No warehouse configs were found. Use Add Warehouse Root or View Warehouses first."
        Exit Function
    End If
    If warehouseOptions.Count = 0 Then
        If Trim$(formReport) = "" Or StrComp(formReport, "OK", vbTextCompare) = 0 Then
            formReport = "No warehouse configs were found. Use Add Warehouse Root or View Warehouses first."
        End If
        report = formReport
        Exit Function
    End If

    If mSeedCallbackAutomationEnabled Then
        warehouseId = mSeedCallbackAutomationWarehouseId
        stationId = mSeedCallbackAutomationStationId
        userId = mSeedCallbackAutomationUserId
        If stationId = "" Then stationId = "S1"
        For Each item In warehouseOptions
            If StrComp(Trim$(CStr(item(1))), warehouseId, vbTextCompare) = 0 _
               And StrComp(Trim$(CStr(item(2))), stationId, vbTextCompare) = 0 Then
                runtimeRoot = Trim$(CStr(item(3)))
                selectionFound = True
                Exit For
            End If
        Next item
        If Not selectionFound Then
            report = "The selected warehouse/station is not available in the current target."
            Exit Function
        End If
    Else
        frmSeedInventory.Configure warehouseOptions, warehouseId, stationId, userId
        frmSeedInventory.Show
        If Not frmSeedInventory.Accepted Then
            report = "Seed inventory cancelled."
            Unload frmSeedInventory
            Exit Function
        End If

        warehouseId = Trim$(frmSeedInventory.SelectedWarehouseId)
        stationId = Trim$(frmSeedInventory.SelectedStationId)
        runtimeRoot = Trim$(frmSeedInventory.SelectedRuntimeRoot)
        userId = Trim$(frmSeedInventory.SelectedUserId)
        Unload frmSeedInventory
    End If

    If warehouseId = "" Then
        report = "WarehouseId is required."
        Exit Function
    End If
    If stationId = "" Then stationId = "S1"
    If userId = "" Then
        report = "Admin user is required."
        Exit Function
    End If
    If runtimeRoot <> "" Then modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot

    If Not modConfig.LoadConfig(warehouseId, stationId) Then
        report = "Config load failed: " & modConfig.Validate()
        Exit Function
    End If

    ResolveSeedInventoryContext = True
End Function

Private Function SanitizeSeedCallbackErrorText(ByVal valueText As String) As String
    valueText = Replace$(Replace$(Trim$(valueText), vbCr, " "), vbLf, " ")
    Do While InStr(1, valueText, "  ", vbBinaryCompare) > 0
        valueText = Replace$(valueText, "  ", " ")
    Loop
    If InStr(1, valueText, "\", vbBinaryCompare) > 0 Then
        valueText = "<redacted-path>"
    ElseIf Len(valueText) > 240 Then
        valueText = Left$(valueText, 240)
    End If
    If valueText = "" Then valueText = "<none>"
    SanitizeSeedCallbackErrorText = valueText
End Function

Private Function ResolveAdminCurrentTargetContext(ByRef warehouseId As String, _
                                                  ByRef stationId As String, _
                                                  ByRef userId As String, _
                                                  ByRef report As String) As Boolean
    Dim target As WarehouseTarget
    Dim accessReport As String

    If Not modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", accessReport) Then
        report = accessReport
        If Trim$(report) = "" Then report = "Sign in as an admin user and connect a warehouse target first."
        Exit Function
    End If

    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then
        report = "Connect a warehouse target first."
        Exit Function
    End If

    warehouseId = Trim$(target.WarehouseId)
    stationId = Trim$(target.StationId)
    userId = Trim$(modAuth.GetCurrentUserId())
    If stationId = "" Then stationId = "S1"
    If warehouseId = "" Then
        report = "Current warehouse target is missing WarehouseId."
        Exit Function
    End If
    If userId = "" Then
        report = "Current invSys user is not signed in."
        Exit Function
    End If
    If Not modConfig.LoadConfig(warehouseId, stationId) Then
        report = "Config load failed: " & modConfig.Validate()
        Exit Function
    End If

    ResolveAdminCurrentTargetContext = True
End Function

Private Function NextInventoryRowSuggestionAdmin(ByVal warehouseId As String) As Long
    Dim path As String
    Dim wb As Workbook
    Dim openedHere As Boolean
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim cRow As Long
    Dim maxRow As Long
    Dim rawValue As Variant

    On Error GoTo Fallback
    path = modProcessor.ResolveInventoryWorkbookPathForAutomation(warehouseId)
    If path <> "" Then
        Set wb = FindOpenWorkbookByFullNameAdmin(path)
        If wb Is Nothing Then
            If Len(Dir$(path, vbNormal)) > 0 Then
                Set wb = Application.Workbooks.Open(path, UpdateLinks:=False, ReadOnly:=True, AddToMru:=False)
                openedHere = True
            End If
        End If
    End If

    If Not wb Is Nothing Then
        Set lo = FindListObjectByNameAdminLocal(wb, "tblSkuCatalog")
        If Not lo Is Nothing Then
            cRow = ColumnIndexAdminLocal(lo, "ROW")
            If cRow > 0 And Not lo.DataBodyRange Is Nothing Then
                For rowIndex = 1 To lo.ListRows.Count
                    rawValue = lo.DataBodyRange.Cells(rowIndex, cRow).Value
                    If IsNumeric(rawValue) Then
                        If CLng(Val(CStr(rawValue))) > maxRow Then maxRow = CLng(Val(CStr(rawValue)))
                    End If
                Next rowIndex
            End If
        End If
    End If

    If maxRow > 0 Then
        NextInventoryRowSuggestionAdmin = maxRow + 1
    Else
        NextInventoryRowSuggestionAdmin = CLng((DateDiff("s", DateSerial(2026, 1, 1), Now) Mod 900000) + 10000)
    End If

CleanExit:
    On Error Resume Next
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function

Fallback:
    NextInventoryRowSuggestionAdmin = CLng((DateDiff("s", DateSerial(2026, 1, 1), Now) Mod 900000) + 10000)
    Resume CleanExit
End Function

Private Function FindOpenWorkbookByFullNameAdmin(ByVal fullName As String) As Workbook
    Dim wb As Workbook

    fullName = LCase$(Trim$(fullName))
    If fullName = "" Then Exit Function
    For Each wb In Application.Workbooks
        If LCase$(Trim$(wb.FullName)) = fullName Then
            Set FindOpenWorkbookByFullNameAdmin = wb
            Exit Function
        End If
    Next wb
End Function

Private Function FindListObjectByNameAdminLocal(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindListObjectByNameAdminLocal = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindListObjectByNameAdminLocal Is Nothing Then Exit Function
    Next ws
End Function

Private Function ColumnIndexAdminLocal(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            ColumnIndexAdminLocal = i
            Exit Function
        End If
    Next i
End Function

Private Function FindCatalogRowBySkuAdmin(ByVal lo As ListObject, ByVal sku As String) As Long
    Dim cSku As Long
    Dim rowIndex As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    sku = Trim$(sku)
    If sku = "" Then Exit Function

    cSku = ColumnIndexAdminLocal(lo, "SKU")
    If cSku = 0 Then cSku = ColumnIndexAdminLocal(lo, "ITEM_CODE")
    If cSku = 0 Then Exit Function

    For rowIndex = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, cSku).Value)), sku, vbTextCompare) = 0 Then
            FindCatalogRowBySkuAdmin = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function CatalogCellAdmin(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As String
    Dim idx As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    idx = ColumnIndexAdminLocal(lo, columnName)
    If idx = 0 Then Exit Function
    CatalogCellAdmin = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, idx).Value))
End Function

Private Sub SetCatalogValueAdmin(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueIn As Variant)
    Dim idx As Long

    If lo Is Nothing Then Exit Sub
    If rowIndex <= 0 Then Exit Sub
    columnName = NormalizeCatalogColumnNameAdmin(columnName)
    If columnName = "" Then Exit Sub
    idx = EnsureCatalogColumnAdmin(lo, columnName)
    If idx = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueIn
End Sub

Private Function EnsureCatalogColumnAdmin(ByVal lo As ListObject, ByVal columnName As String) As Long
    If lo Is Nothing Then Exit Function
    columnName = NormalizeCatalogColumnNameAdmin(columnName)
    If columnName = "" Then Exit Function
    EnsureCatalogColumnAdmin = ColumnIndexAdminLocal(lo, columnName)
    If EnsureCatalogColumnAdmin > 0 Then Exit Function
    EnsureCatalogColumnAdmin = lo.ListColumns.Add.Index
    lo.ListColumns(EnsureCatalogColumnAdmin).Name = columnName
End Function

Private Function NormalizeCatalogColumnNameAdmin(ByVal columnName As String) As String
    columnName = Replace(columnName, vbCr, " ")
    columnName = Replace(columnName, vbLf, " ")
    columnName = Replace(columnName, vbTab, " ")
    columnName = Replace(columnName, "[", "(")
    columnName = Replace(columnName, "]", ")")
    Do While InStr(1, columnName, "  ", vbBinaryCompare) > 0
        columnName = Replace(columnName, "  ", " ")
    Loop
    columnName = Trim$(columnName)
    If Len(columnName) > 48 Then columnName = Left$(columnName, 48)
    NormalizeCatalogColumnNameAdmin = columnName
End Function

Private Sub SetSheetProtectionAdminLocal(ByVal ws As Worksheet, ByVal protectAfter As Boolean)
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    If protectAfter Then
        ws.Protect UserInterfaceOnly:=True
    ElseIf ws.ProtectContents Then
        ws.Unprotect
    End If
    On Error GoTo 0
End Sub

Private Sub AddPictureReferencesToPayloadAdmin(ByVal item As Object, ByVal imagePath As String)
    Dim refs As Collection
    Dim i As Long

    If item Is Nothing Then Exit Sub
    Set refs = ParsePictureReferencesAdmin(imagePath)
    If refs.Count = 0 Then
        item("IMAGE_PATH") = ""
        Exit Sub
    End If
    item("IMAGE_PATH") = refs(1)
    For i = 2 To refs.Count
        item("IMAGE_PATH_" & CStr(i)) = refs(i)
    Next i
End Sub

Private Sub SetPictureReferencesInCatalogAdmin(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal imagePath As String)
    Dim refs As Collection
    Dim i As Long
    Dim columnName As String

    Set refs = ParsePictureReferencesAdmin(imagePath)
    If refs.Count > 0 Then
        SetCatalogValueAdmin lo, rowIndex, "IMAGE_PATH", refs(1)
    Else
        SetCatalogValueAdmin lo, rowIndex, "IMAGE_PATH", ""
    End If
    For i = 2 To 12
        columnName = "IMAGE_PATH_" & CStr(i)
        If i <= refs.Count Then
            SetCatalogValueAdmin lo, rowIndex, columnName, refs(i)
        ElseIf ColumnIndexAdminLocal(lo, columnName) > 0 Then
            SetCatalogValueAdmin lo, rowIndex, columnName, ""
        End If
    Next i
End Sub

Private Function CombinedPictureReferencesAdmin(ByVal lo As ListObject, ByVal rowIndex As Long) As String
    Dim parts As Collection
    Dim i As Long
    Dim valueText As String
    Dim result As String

    Set parts = New Collection
    valueText = CatalogCellAdmin(lo, rowIndex, "IMAGE_PATH")
    If valueText <> "" Then parts.Add valueText
    For i = 2 To 12
        valueText = CatalogCellAdmin(lo, rowIndex, "IMAGE_PATH_" & CStr(i))
        If valueText <> "" Then parts.Add valueText
    Next i
    For i = 1 To parts.Count
        If result <> "" Then result = result & "; "
        result = result & CStr(parts(i))
    Next i
    CombinedPictureReferencesAdmin = result
End Function

Private Function ParsePictureReferencesAdmin(ByVal imagePath As String) As Collection
    Dim result As Collection
    Dim rawParts As Variant
    Dim part As Variant
    Dim valueText As String

    Set result = New Collection
    imagePath = Replace(imagePath, "|", ";")
    imagePath = Replace(imagePath, vbCr, ";")
    imagePath = Replace(imagePath, vbLf, ";")
    rawParts = Split(imagePath, ";")
    For Each part In rawParts
        valueText = Trim$(CStr(part))
        If valueText <> "" Then result.Add valueText
    Next part
    Set ParsePictureReferencesAdmin = result
End Function

Private Function SeedDemoInventoryForWarehouse(ByVal warehouseId As String, _
                                               ByVal stationId As String, _
                                               ByVal userId As String, _
                                               ByRef report As String) As Boolean
    Dim payloadItems As Collection
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim queueError As String
    Dim batchReport As String
    Dim processedCount As Long
    Dim inboxReport As String

    On Error GoTo FailSeed

    If Not EnsureDemoStationInboxes(warehouseId, stationId, inboxReport) Then
        report = inboxReport
        Exit Function
    End If

    Set payloadItems = BuildAdminDemoInventoryPayload()
    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)
    If payloadJson = "" Or payloadJson = "[]" Then
        report = "Demo inventory payload was empty."
        Exit Function
    End If

    If Not modRoleEventWriter.QueueInventoryCreateEvent(warehouseId, stationId, userId, payloadJson, _
                                                        "Admin demo inventory seed", _
                                                        0, Nothing, eventIdOut, queueError, "") Then
        report = "Seed event could not be queued: " & queueError & vbCrLf & _
                 "Use Users & Roles to grant ADMIN_MAINT to '" & userId & "' for " & warehouseId & " / " & stationId & "."
        Exit Function
    End If

    processedCount = modProcessor.RunBatch(warehouseId, 0, batchReport)
    If processedCount < 1 Then
        report = "Seed event was queued but not applied. " & batchReport
        Exit Function
    End If

    report = "Demo inventory seeded." & vbCrLf & _
             "Warehouse: " & warehouseId & vbCrLf & _
             "Applied events: " & CStr(processedCount) & vbCrLf & _
             "Processor: " & batchReport & vbCrLf & _
             "Now refresh inventory in any open role workbook, or reopen the picker to auto-refresh."
    SeedDemoInventoryForWarehouse = True
    Exit Function

FailSeed:
    report = "SeedDemoInventory failed: " & Err.Description
End Function

Private Function EnsureDemoStationInboxes(ByVal warehouseId As String, _
                                          ByVal stationId As String, _
                                          ByRef report As String) As Boolean
    Dim inboxPath As String
    Dim stepReport As String

    If Not modConfig.EnsureStationInbox(warehouseId, stationId, "RECEIVE", "", inboxPath, stepReport) Then
        report = "Receiving inbox could not be created or repaired: " & stepReport
        Exit Function
    End If

    inboxPath = ""
    stepReport = ""
    If Not modConfig.EnsureStationInbox(warehouseId, stationId, "SHIP", "", inboxPath, stepReport) Then
        report = "Shipping inbox could not be created or repaired: " & stepReport
        Exit Function
    End If

    inboxPath = ""
    stepReport = ""
    If Not modConfig.EnsureStationInbox(warehouseId, stationId, "PRODUCTION", "", inboxPath, stepReport) Then
        report = "Production inbox could not be created or repaired: " & stepReport
        Exit Function
    End If

    report = "OK"
    EnsureDemoStationInboxes = True
End Function

Private Function BuildAdminDemoInventoryPayload() As Collection
    Dim csvPath As String
    Dim payload As Collection

    csvPath = ResolveAdminDemoInventoryCsvPath()
    If csvPath <> "" Then
        Set payload = BuildAdminDemoInventoryPayloadFromCsv(csvPath)
        If Not payload Is Nothing Then
            If payload.Count > 0 Then
                Set BuildAdminDemoInventoryPayload = payload
                Exit Function
            End If
        End If
    End If

    Set BuildAdminDemoInventoryPayload = BuildAdminDemoInventoryFallbackPayload()
End Function

Private Function BuildAdminDemoInventoryFallbackPayload() As Collection
    Dim rows As Collection
    Dim item As Object

    Set rows = New Collection

    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem( _
        modRoleEventWriter.CreateSystemKey(), "DEMO-RAW-BLACK-TEA", ADMIN_DEMO_INVENTORY_QTY, _
        "NAS-A1", "GOOD", "", "Admin demo seed")
    item("ITEM_CODE") = "DEMO-RAW-BLACK-TEA"
    item("ITEM") = "Black Tea Base"
    item("UOM") = "LB"
    item("TOTAL INV") = ADMIN_DEMO_INVENTORY_QTY
    item("QtyAvailable") = ADMIN_DEMO_INVENTORY_QTY
    item("DESCRIPTION") = "Demo raw black tea for receiving tests"
    item("VENDOR(s)") = "Demo Vendor"
    item("CATEGORY") = "Raw Material"
    rows.Add item

    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem( _
        modRoleEventWriter.CreateSystemKey(), "DEMO-SPICE-CARDAMOM", ADMIN_DEMO_INVENTORY_QTY, _
        "NAS-A2", "GOOD", "", "Admin demo seed")
    item("ITEM_CODE") = "DEMO-SPICE-CARDAMOM"
    item("ITEM") = "Cardamom Pods"
    item("UOM") = "LB"
    item("TOTAL INV") = ADMIN_DEMO_INVENTORY_QTY
    item("QtyAvailable") = ADMIN_DEMO_INVENTORY_QTY
    item("DESCRIPTION") = "Demo spice inventory for receiving tests"
    item("VENDOR(s)") = "Demo Vendor"
    item("CATEGORY") = "Spice"
    rows.Add item

    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem( _
        modRoleEventWriter.CreateSystemKey(), "DEMO-PKG-TIN", ADMIN_DEMO_INVENTORY_QTY, _
        "NAS-P1", "GOOD", "", "Admin demo seed")
    item("ITEM_CODE") = "DEMO-PKG-TIN"
    item("ITEM") = "Retail Tea Tin"
    item("UOM") = "EA"
    item("TOTAL INV") = ADMIN_DEMO_INVENTORY_QTY
    item("QtyAvailable") = ADMIN_DEMO_INVENTORY_QTY
    item("DESCRIPTION") = "Demo packaging item for picker tests"
    item("VENDOR(s)") = "Demo Vendor"
    item("CATEGORY") = "Packaging"
    rows.Add item

    Set BuildAdminDemoInventoryFallbackPayload = rows
End Function

Private Function BuildAdminDemoInventoryPayloadFromCsv(ByVal csvPath As String) As Collection
    On Error GoTo FailCsv

    Dim fso As Object
    Dim textStream As Object
    Dim headerLine As String
    Dim fields As Collection
    Dim headers As Object
    Dim rows As Collection
    Dim lineText As String
    Dim item As Object
    Dim sku As String
    Dim itemName As String
    Dim uom As String
    Dim location As String
    Dim category As String
    Dim qty As Double

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(csvPath) Then Exit Function

    Set textStream = fso.OpenTextFile(csvPath, 1, False)
    If textStream.AtEndOfStream Then GoTo CleanExit

    headerLine = textStream.ReadLine
    Set headers = CsvHeaderMapAdmin(ParseCsvLineAdmin(headerLine))
    Set rows = New Collection

    Do While Not textStream.AtEndOfStream
        lineText = textStream.ReadLine
        If Trim$(lineText) = "" Then GoTo NextLine

        Set fields = ParseCsvLineAdmin(lineText)
        sku = CsvFieldAdmin(fields, headers, "ITEM_CODE")
        itemName = CsvFieldAdmin(fields, headers, "ITEM")
        If sku = "" And itemName = "" Then GoTo NextLine
        If sku = "" Then sku = itemName

        uom = CsvFieldAdmin(fields, headers, "UOM")
        location = CsvFieldAdmin(fields, headers, "LOCATION")
        category = CsvFieldAdmin(fields, headers, "CATEGORY")
        qty = Val(CsvFieldAdmin(fields, headers, "QUANTITY"))
        If qty <= 0 Then qty = ResolveDemoSeedQuantityAdmin(category, CsvFieldAdmin(fields, headers, "PHASE"), uom)

        Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem( _
            modRoleEventWriter.CreateSystemKey(), sku, qty, location, "GOOD", "", _
            "Admin CSV demo inventory seed")
        item("ITEM_CODE") = sku
        item("ITEM") = itemName
        item("UOM") = uom
        item("LOCATION") = location
        item("TOTAL INV") = qty
        item("QtyAvailable") = qty
        item("DESCRIPTION") = CsvFieldAdmin(fields, headers, "DESCRIPTION")
        item("VENDOR(s)") = CsvFieldAdmin(fields, headers, "VENDOR(s)")
        item("VENDOR_CODE") = CsvFieldAdmin(fields, headers, "VENDOR_CODE")
        item("CATEGORY") = category
        If CsvFieldAdmin(fields, headers, "SUBSTITUTION") <> "" Then item("SUBSTITUTION") = CsvFieldAdmin(fields, headers, "SUBSTITUTION")
        If CsvFieldAdmin(fields, headers, "PHASE") <> "" Then item("PHASE") = CsvFieldAdmin(fields, headers, "PHASE")
        If CsvFieldAdmin(fields, headers, "ASSIGNEE") <> "" Then item("ASSIGNEE") = CsvFieldAdmin(fields, headers, "ASSIGNEE")
        rows.Add item
NextLine:
    Loop

    Set BuildAdminDemoInventoryPayloadFromCsv = rows

CleanExit:
    On Error Resume Next
    If Not textStream Is Nothing Then textStream.Close
    On Error GoTo 0
    Exit Function

FailCsv:
    Resume CleanExit
End Function

Private Function NormalizeAdminDemoInventoryRow(ByVal rowVal As Long, ByVal sku As String, ByVal fallbackRow As Long) As Long
    Dim demoRow As Long

    demoRow = AdminDemoInventoryRowForSku(sku)
    If demoRow > 0 Then
        NormalizeAdminDemoInventoryRow = demoRow
    ElseIf rowVal > 0 Then
        NormalizeAdminDemoInventoryRow = rowVal
    Else
        NormalizeAdminDemoInventoryRow = fallbackRow
    End If
End Function

Private Function AdminDemoInventoryRowForSku(ByVal sku As String) As Long
    Select Case UCase$(Trim$(sku))
        Case "DEMO-RAW-BLACK-TEA"
            AdminDemoInventoryRowForSku = 9001
        Case "DEMO-RAW-FILTERED-WATER"
            AdminDemoInventoryRowForSku = 9002
        Case "DEMO-RAW-CARDAMOM", "DEMO-SPICE-CARDAMOM"
            AdminDemoInventoryRowForSku = 9003
        Case "DEMO-FG-CLASSIC-CHAI"
            AdminDemoInventoryRowForSku = 9016
        Case "DEMO-PKG-TIN"
            AdminDemoInventoryRowForSku = 9021
    End Select
End Function

Private Function ResolveAdminDemoInventoryCsvPath() As String
    Dim fso As Object
    Dim candidates As Collection
    Dim basePath As String
    Dim parentPath As String
    Dim candidate As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set candidates = New Collection

    basePath = ThisWorkbook.Path
    If basePath <> "" Then
        candidates.Add basePath & "\assets\inv.sample.data.csv"
        parentPath = fso.GetParentFolderName(basePath)
        If parentPath <> "" Then
            candidates.Add parentPath & "\assets\inv.sample.data.csv"
            parentPath = fso.GetParentFolderName(parentPath)
            If parentPath <> "" Then candidates.Add parentPath & "\assets\inv.sample.data.csv"
        End If
    End If

    On Error Resume Next
    candidates.Add CurDir$ & "\assets\inv.sample.data.csv"
    candidates.Add CurDir$ & "\..\assets\inv.sample.data.csv"
    On Error GoTo 0

    For Each candidate In candidates
        If fso.FileExists(CStr(candidate)) Then
            ResolveAdminDemoInventoryCsvPath = CStr(candidate)
            Exit Function
        End If
    Next candidate
End Function

Private Function CsvHeaderMapAdmin(ByVal headers As Collection) As Object
    Dim result As Object
    Dim i As Long
    Dim headerText As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    For i = 1 To headers.Count
        headerText = Trim$(CStr(headers(i)))
        If i = 1 Then headerText = Replace$(headerText, ChrW$(&HFEFF), "")
        If headerText <> "" Then result(headerText) = i
    Next i
    Set CsvHeaderMapAdmin = result
End Function

Private Function CsvFieldAdmin(ByVal fields As Collection, ByVal headers As Object, ByVal headerName As String) As String
    Dim idx As Long

    If fields Is Nothing Then Exit Function
    If headers Is Nothing Then Exit Function
    If Not headers.Exists(headerName) Then Exit Function
    idx = CLng(headers(headerName))
    If idx <= 0 Or idx > fields.Count Then Exit Function
    CsvFieldAdmin = Trim$(CStr(fields(idx)))
End Function

Private Function ParseCsvLineAdmin(ByVal lineText As String) As Collection
    Dim result As Collection
    Dim i As Long
    Dim ch As String
    Dim current As String
    Dim inQuotes As Boolean

    Set result = New Collection
    For i = 1 To Len(lineText)
        ch = Mid$(lineText, i, 1)
        If ch = """" Then
            If inQuotes And i < Len(lineText) And Mid$(lineText, i + 1, 1) = """" Then
                current = current & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf ch = "," And Not inQuotes Then
            result.Add current
            current = ""
        Else
            current = current & ch
        End If
    Next i
    result.Add current
    Set ParseCsvLineAdmin = result
End Function

Private Function ResolveDemoSeedQuantityAdmin(ByVal category As String, ByVal phase As String, ByVal uom As String) As Double
    ResolveDemoSeedQuantityAdmin = ADMIN_DEMO_INVENTORY_QTY
End Function

Private Sub PromptForWarehouseDirectoryRootIfNeeded()
    Dim rootPath As String

    If modAdminConsole.HasRememberedWarehouseScanRoots() Then Exit Sub
    rootPath = InputBox("Optional: enter a NAS/server warehouse root to include in this warehouse scan. Leave blank to scan only local/open warehouse configs.", _
                        "invSys Admin - Warehouse Root", _
                        "\\100.84.136.19\invSysWH1")
    rootPath = Trim$(rootPath)
    If rootPath <> "" Then modAdminConsole.RememberWarehouseScanRoot rootPath
End Sub

Sub Verify_AddinsPublished()
    Dim report As String
    Dim detail As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If modAddinsPublish.VerifyAddinsPublished() Then
        MsgBox "All required add-ins are published." & vbCrLf & modAddinsPublish.GetLastAddinsPublishReport(), vbInformation, "invSys Admin"
    Else
        detail = modAddinsPublish.GetLastAddinsPublishReport()
        If Len(detail) = 0 Then detail = "One or more required add-ins are missing or zero-byte."
        If InStr(1, detail, "PathSharePointRoot is not configured", vbTextCompare) > 0 Then
            detail = detail & vbCrLf & _
                     "Use Create New Warehouse or Setup Tester Station to choose the locally synced invSys SharePoint root first."
        End If
        MsgBox "Add-ins publish verification failed." & vbCrLf & detail, vbExclamation, "invSys Admin"
    End If
End Sub

Sub Export_LoadedPackageReport()
    Dim report As String
    Dim pathOut As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If modPackageDiagnostics.ExportLoadedPackageReport("", "", "", pathOut, report) Then
        MsgBox "Loaded package report written to:" & vbCrLf & pathOut, vbInformation, "invSys Admin"
    Else
        If Len(Trim$(report)) = 0 Then report = "Loaded package report export failed."
        MsgBox report, vbExclamation, "invSys Admin"
    End If
End Sub

Sub Admin_RetireMigrateWarehouse_Click()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    frmRetireMigrateWarehouse.Show
End Sub

Sub Open_RetireMigrateWarehouse()
    Admin_RetireMigrateWarehouse_Click
End Sub

Public Sub Scheduler_RunWarehouseBatch()
    PublishSchedulerResult modAdminConsole.RunScheduledWarehouseBatchForAutomation("", 0)
End Sub

Public Sub Scheduler_RunWarehousePublish()
    PublishSchedulerResult modAdminConsole.RunScheduledWarehousePublishForAutomation("", "")
End Sub

Public Sub Scheduler_RunHQAggregation()
    PublishSchedulerResult modAdminConsole.RunScheduledHQAggregationForAutomation("", "")
End Sub

Private Sub PublishSchedulerResult(ByVal resultText As String)
    Debug.Print resultText
    On Error Resume Next
    Application.StatusBar = resultText
    On Error GoTo 0
End Sub

Public Function ResolveInteractiveAdminWorkbook(Optional ByVal allowAddinFallback As Boolean = True) As Workbook
    Set ResolveInteractiveAdminWorkbook = modAdminWorkbookTarget.ResolveAdminTargetWorkbook(Nothing, ThisWorkbook, allowAddinFallback)
End Function

''''''''''''''''''''''''''''''''''''
' This module contains administrative functions for the application.
' It includes functions to manage user accounts, roles, and permissions. yada yada
' It also includes functions to manage application settings and configurations.
' Shared administrative helpers used by the active Admin ribbon actions and forms.
''''''''''''''''''''''''''''''''''''
