VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cOperationsAnchorItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mControl As Object
Private mParent As Object
Private mAnchorMask As Long
Private mStartingTop As Double
Private mStartingLeft As Double
Private mStartingHeight As Double
Private mStartingWidth As Double
Private mParentStartingHeight As Double
Private mParentStartingWidth As Double
Private mUseProportionalVertical As Boolean
Private mVerticalTopFactor As Double
Private mVerticalHeightFactor As Double

Public Sub CaptureAnchor(ByVal targetControl As Object, ByVal anchorMask As Long)
    If targetControl Is Nothing Then
        Err.Raise vbObjectError + 7520, "cOperationsAnchorItem.CaptureAnchor", _
                  "An anchor target control is required."
    End If

    Set mControl = targetControl
    Set mParent = ResolveParent(targetControl)
    If mParent Is Nothing Then
        Err.Raise vbObjectError + 7521, "cOperationsAnchorItem.CaptureAnchor", _
                  "The anchor target has no layout parent."
    End If

    mAnchorMask = anchorMask
    mStartingTop = CDbl(targetControl.Top)
    mStartingLeft = CDbl(targetControl.Left)
    mStartingHeight = CDbl(targetControl.Height)
    mStartingWidth = CDbl(targetControl.Width)
    mParentStartingHeight = ParentInsideExtent(mParent, False)
    mParentStartingWidth = ParentInsideExtent(mParent, True)
End Sub

Public Sub CaptureProportionalVerticalAnchor(ByVal targetControl As Object, _
                                             ByVal horizontalAnchorMask As Long, _
                                             ByVal verticalTopFactor As Double, _
                                             ByVal verticalHeightFactor As Double)
    If verticalTopFactor < 0 Or verticalHeightFactor < 0 Then
        Err.Raise vbObjectError + 7524, "cOperationsAnchorItem.CaptureProportionalVerticalAnchor", _
                  "Vertical layout factors must be nonnegative."
    End If

    CaptureAnchor targetControl, horizontalAnchorMask
    mUseProportionalVertical = True
    mVerticalTopFactor = verticalTopFactor
    mVerticalHeightFactor = verticalHeightFactor
End Sub

Public Sub ApplyAnchoredControl()
    Dim widthDelta As Double
    Dim heightDelta As Double
    Dim newLeft As Double
    Dim newTop As Double
    Dim newWidth As Double
    Dim newHeight As Double

    If mControl Is Nothing Or mParent Is Nothing Then Exit Sub

    widthDelta = ParentInsideExtent(mParent, True) - mParentStartingWidth
    heightDelta = ParentInsideExtent(mParent, False) - mParentStartingHeight
    newLeft = mStartingLeft
    newTop = mStartingTop
    newWidth = mStartingWidth
    newHeight = mStartingHeight

    If HasAnchor(OPERATIONS_ANCHOR_LEFT) And HasAnchor(OPERATIONS_ANCHOR_RIGHT) Then
        newWidth = Nonnegative(mStartingWidth + widthDelta)
    ElseIf (Not HasAnchor(OPERATIONS_ANCHOR_LEFT)) And HasAnchor(OPERATIONS_ANCHOR_RIGHT) Then
        newLeft = mStartingLeft + widthDelta
    End If

    If mUseProportionalVertical Then
        newTop = mStartingTop + heightDelta * mVerticalTopFactor
        newHeight = Nonnegative(mStartingHeight + heightDelta * mVerticalHeightFactor)
    ElseIf HasAnchor(OPERATIONS_ANCHOR_TOP) And HasAnchor(OPERATIONS_ANCHOR_BOTTOM) Then
        newHeight = Nonnegative(mStartingHeight + heightDelta)
    ElseIf (Not HasAnchor(OPERATIONS_ANCHOR_TOP)) And HasAnchor(OPERATIONS_ANCHOR_BOTTOM) Then
        newTop = mStartingTop + heightDelta
    End If

    mControl.Move newLeft, newTop, newWidth, newHeight
End Sub

Private Function HasAnchor(ByVal anchorFlag As Long) As Boolean
    HasAnchor = ((mAnchorMask And anchorFlag) = anchorFlag)
End Function

Private Function ResolveParent(ByVal targetControl As Object) As Object
    On Error Resume Next
    Set ResolveParent = CallByName(targetControl, "Container", VbGet)
    If ResolveParent Is Nothing Then
        Err.Clear
        Set ResolveParent = targetControl.Parent
    End If
    On Error GoTo 0
End Function

Private Function ParentInsideExtent(ByVal parentObject As Object, _
                                    ByVal horizontal As Boolean) As Double
    Dim pageHost As Object
    Dim primaryProperty As String
    Dim fallbackProperty As String
    Dim pageInset As Double

    On Error Resume Next
    If horizontal Then
        primaryProperty = "InsideWidth"
        fallbackProperty = "Width"
        pageInset = 20
    Else
        primaryProperty = "InsideHeight"
        fallbackProperty = "Height"
        pageInset = 20
    End If
    If TypeName(parentObject) = "Page" Then
        Set pageHost = parentObject.Parent
        ParentInsideExtent = CDbl(CallByName(pageHost, fallbackProperty, VbGet)) - pageInset
        On Error GoTo 0
        Exit Function
    End If
    ParentInsideExtent = CDbl(CallByName(parentObject, primaryProperty, VbGet))
    If Err.Number <> 0 Then
        Err.Clear
        ParentInsideExtent = CDbl(CallByName(parentObject, fallbackProperty, VbGet))
    End If
    On Error GoTo 0
End Function

Private Function Nonnegative(ByVal value As Double) As Double
    If value < 0 Then
        Nonnegative = 0
    Else
        Nonnegative = value
    End If
End Function
