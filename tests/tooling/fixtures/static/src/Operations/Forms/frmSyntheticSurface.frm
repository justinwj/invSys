VERSION 5.00
Begin VB.UserForm frmSyntheticSurface
   Caption         =   "Synthetic Operations"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   StartUpPosition =   1
   Begin MSForms.CommandButton cmdApply
      Caption         =   "Apply"
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   960
   End
   Begin MSForms.TextBox txtSystemKey
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3000
   End
End
Attribute VB_Name = "frmSyntheticSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.txtSystemKey.Value = "SYNTHETIC-SYSTEM-KEY"
End Sub

Private Sub cmdApply_Click()
    modSyntheticSurface.SyntheticProcessorHandler CStr(Me.txtSystemKey.Value)
End Sub
