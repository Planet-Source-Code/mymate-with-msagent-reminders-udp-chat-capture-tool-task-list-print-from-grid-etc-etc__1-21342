VERSION 5.00
Begin VB.Form frmTag 
   ClientHeight    =   600
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   2430
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleWidth      =   2430
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_DblClick()
Unload Me
End Sub

Public Function ShowThis(sString As String)
Label1 = sString
Me.Top = 0
Me.Left = 0
Me.Show

End Function

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label1_DblClick()
Unload Me
End Sub
