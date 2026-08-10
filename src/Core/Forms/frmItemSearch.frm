VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmItemSearch
   Caption         =   "Item Search"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6480
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmItemSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

' cDynItemSearch owns all controls, role profiles, event wiring, and behavior.
' This form is intentionally a clean Core-owned runtime canvas.
Private mAnchors As cFormAnchorManager
Private mResizeEnabled As Boolean

Private Sub UserForm_Activate()
    If Not mResizeEnabled Then
        modUserFormResizeWin.EnableResizableUserForm Me, True, True
        mResizeEnabled = True
    End If
    If Not mAnchors Is Nothing Then mAnchors.ResizeControls
End Sub

Private Sub UserForm_Layout()
    If Not mAnchors Is Nothing Then mAnchors.ResizeControls
End Sub

Private Sub UserForm_Terminate()
    Set mAnchors = Nothing
End Sub

Public Sub ConfigureRuntimeLayout(ByVal searchBox As Object, _
                                  ByVal shippingFilter As Object, _
                                  ByVal resultsList As Object, _
                                  ByVal descriptionBox As Object)
    Set mAnchors = modDynamicForms.CreateFormAnchorManager()
    mAnchors.Initialize Me, 480, 420
    mAnchors.Add searchBox, anchorLeft Or anchorTop Or anchorRight
    mAnchors.Add shippingFilter, anchorLeft Or anchorTop Or anchorRight
    mAnchors.Add resultsList, anchorLeft Or anchorTop Or anchorRight Or anchorBottom
    mAnchors.Add descriptionBox, anchorLeft Or anchorRight Or anchorBottom
End Sub
