VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cOperationsAnchorManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
#End If

Private mForm As Object
Private mItems As Collection
Private mMinWidth As Double
Private mMinHeight As Double
Private mApplying As Boolean

Public Sub ConfigureForForm(ByVal targetForm As Object, _
                            Optional ByVal minimumWidth As Double = 0, _
                            Optional ByVal minimumHeight As Double = 0)
    If targetForm Is Nothing Then
        Err.Raise vbObjectError + 7522, "cOperationsAnchorManager.ConfigureForForm", _
                  "A target form is required."
    End If

    Set mForm = targetForm
    Set mItems = New Collection
    mMinWidth = minimumWidth
    mMinHeight = minimumHeight
End Sub

Public Sub RegisterControl(ByVal targetControl As Object, ByVal anchorMask As Long)
    Dim item As cOperationsAnchorItem

    If mItems Is Nothing Then
        Err.Raise vbObjectError + 7523, "cOperationsAnchorManager.RegisterControl", _
                  "Configure the anchor manager before registering controls."
    End If

    Set item = New cOperationsAnchorItem
    item.CaptureAnchor targetControl, anchorMask
    mItems.Add item
End Sub

Public Sub RegisterProportionalVerticalControl(ByVal targetControl As Object, _
                                               ByVal horizontalAnchorMask As Long, _
                                               ByVal verticalTopFactor As Double, _
                                               ByVal verticalHeightFactor As Double)
    Dim item As cOperationsAnchorItem

    If mItems Is Nothing Then
        Err.Raise vbObjectError + 7525, "cOperationsAnchorManager.RegisterProportionalVerticalControl", _
                  "Configure the anchor manager before registering controls."
    End If

    Set item = New cOperationsAnchorItem
    item.CaptureProportionalVerticalAnchor targetControl, horizontalAnchorMask, _
                                           verticalTopFactor, verticalHeightFactor
    mItems.Add item
End Sub

Public Sub ApplyAnchoredLayout()
    Dim item As cOperationsAnchorItem
#If VBA7 Then
    Dim formHwnd As LongPtr
#Else
    Dim formHwnd As Long
#End If

    If mApplying Then Exit Sub
    If mForm Is Nothing Then Exit Sub
    If mItems Is Nothing Then Exit Sub

    On Error GoTo CleanFail
    mApplying = True
    formHwnd = modUserFormResizeWin.GetUserFormWindowHandle(mForm)
    If formHwnd = 0 Or (IsIconic(formHwnd) = 0 And IsZoomed(formHwnd) = 0) Then
        ApplyMinimumFormSize
    End If
    For Each item In mItems
        item.ApplyAnchoredControl
    Next item

CleanExit:
    mApplying = False
    Exit Sub

CleanFail:
    If Err.Number = 384 Then Resume CleanExit
    mApplying = False
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Property Get RegisteredControlCount() As Long
    If mItems Is Nothing Then Exit Property
    RegisteredControlCount = mItems.Count
End Property

Private Sub ApplyMinimumFormSize()
    If mMinWidth > 0 Then
        If CDbl(mForm.Width) < mMinWidth Then mForm.Width = mMinWidth
    End If
    If mMinHeight > 0 Then
        If CDbl(mForm.Height) < mMinHeight Then mForm.Height = mMinHeight
    End If
End Sub
