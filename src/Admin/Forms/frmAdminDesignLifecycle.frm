VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdminDesignLifecycle
   Caption         =   "Designs Lifecycle"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAdminDesignLifecycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mLstDesigns As MSForms.ListBox
Private WithEvents mBtnRefresh As MSForms.CommandButton
Private WithEvents mBtnRelease As MSForms.CommandButton
Private WithEvents mBtnObsolete As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton
Private mTxtNote As MSForms.TextBox
Private mLblStatus As MSForms.Label
Private mBuilt As Boolean

Private Sub UserForm_Initialize()
    BuildLayout
    RefreshDesigns
End Sub

Private Sub BuildLayout()
    If mBuilt Then Exit Sub

    Me.Caption = "invSys Admin - Designs Lifecycle"
    Me.Width = 760
    Me.Height = 430

    AddLabel "Design versions", 12, 12, 180, 18, True
    Set mBtnRefresh = AddButton("btnRefresh", "Refresh", 642, 8, 88, 24)

    Set mLstDesigns = Me.Controls.Add("Forms.ListBox.1", "lstDesigns", True)
    With mLstDesigns
        .Left = 12
        .Top = 36
        .Width = 718
        .Height = 252
        .ColumnCount = 6
        .ColumnHeads = False
        .ColumnWidths = "100 pt;55 pt;70 pt;155 pt;220 pt;75 pt"
        .MultiSelect = fmMultiSelectSingle
    End With

    AddLabel "Audit note", 12, 300, 90, 18, True
    Set mTxtNote = Me.Controls.Add("Forms.TextBox.1", "txtNote", True)
    With mTxtNote
        .Left = 12
        .Top = 320
        .Width = 718
        .Height = 42
        .MultiLine = True
        .EnterKeyBehavior = True
    End With

    Set mBtnRelease = AddButton("btnRelease", "Release selected", 12, 370, 118, 28)
    Set mBtnObsolete = AddButton("btnObsolete", "Obsolete selected", 138, 370, 125, 28)
    Set mBtnClose = AddButton("btnClose", "Close", 642, 370, 88, 28)

    Set mLblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    With mLblStatus
        .Left = 276
        .Top = 374
        .Width = 354
        .Height = 32
        .WordWrap = True
        .Caption = ""
    End With
    mBuilt = True
End Sub

Private Sub RefreshDesigns()
    On Error GoTo FailRefresh

    Dim designs As Variant
    Dim r As Long
    Dim c As Long

    If Not mBuilt Then BuildLayout
    mLstDesigns.Clear
    designs = modDesignsDomainBridge.ListDesignsBridge(Nothing, "")
    If IsEmpty(designs) Or Not IsArray(designs) Then
        ShowStatus "No Designs Domain versions are available."
        Exit Sub
    End If

    For r = LBound(designs, 1) To UBound(designs, 1)
        mLstDesigns.AddItem CStr(designs(r, 1))
        For c = 2 To 6
            mLstDesigns.List(mLstDesigns.ListCount - 1, c - 1) = CStr(designs(r, c))
        Next c
    Next r
    ShowStatus CStr(mLstDesigns.ListCount) & " design version(s)."
    Exit Sub

FailRefresh:
    ShowStatus "Refresh failed: " & Err.Description
End Sub

Private Sub ExecuteSelected(ByVal eventType As String)
    Dim designId As String
    Dim designVersion As String
    Dim report As String

    If mLstDesigns.ListIndex < 0 Then
        ShowStatus "Select a design version first."
        Exit Sub
    End If
    designId = CStr(mLstDesigns.List(mLstDesigns.ListIndex, 0))
    designVersion = CStr(mLstDesigns.List(mLstDesigns.ListIndex, 1))
    If eventType = "DESIGN_OBSOLETE" Then
        If MsgBox("Obsolete " & designId & " version " & designVersion & "?", _
                  vbQuestion Or vbYesNo Or vbDefaultButton2, _
                  "invSys Admin") <> vbYes Then Exit Sub
    End If

    If modAdminDesignLifecycle.ExecuteDesignLifecycleCommand( _
        eventType, designId, designVersion, mTxtNote.Text, report) Then
        mTxtNote.Text = ""
        RefreshDesigns
        ShowStatus report
    Else
        ShowStatus report
    End If
End Sub

Private Sub mBtnRefresh_Click()
    RefreshDesigns
End Sub

Private Sub mBtnRelease_Click()
    ExecuteSelected "DESIGN_RELEASE"
End Sub

Private Sub mBtnObsolete_Click()
    ExecuteSelected "DESIGN_OBSOLETE"
End Sub

Private Sub mBtnClose_Click()
    Unload Me
End Sub

Private Sub ShowStatus(ByVal textValue As String)
    If Not mLblStatus Is Nothing Then mLblStatus.Caption = textValue
End Sub

Private Sub AddLabel(ByVal captionValue As String, ByVal leftValue As Single, _
                     ByVal topValue As Single, ByVal widthValue As Single, _
                     ByVal heightValue As Single, ByVal boldValue As Boolean)
    Dim labelControl As MSForms.Label
    Set labelControl = Me.Controls.Add("Forms.Label.1", _
        "lbl" & CStr(Me.Controls.Count + 1), True)
    With labelControl
        .Caption = captionValue
        .Left = leftValue
        .Top = topValue
        .Width = widthValue
        .Height = heightValue
        .Font.Bold = boldValue
    End With
End Sub

Private Function AddButton(ByVal controlName As String, ByVal captionValue As String, _
                           ByVal leftValue As Single, ByVal topValue As Single, _
                           ByVal widthValue As Single, ByVal heightValue As Single) As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    With AddButton
        .Caption = captionValue
        .Left = leftValue
        .Top = topValue
        .Width = widthValue
        .Height = heightValue
    End With
End Function

Public Function TestLayoutReady() As Long
    If Not mBuilt Then BuildLayout
    If Not mLstDesigns Is Nothing _
       And Not mBtnRelease Is Nothing _
       And Not mBtnObsolete Is Nothing _
       And mLstDesigns.ColumnCount = 6 Then TestLayoutReady = 1
End Function
