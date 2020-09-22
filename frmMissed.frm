VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMissed 
   Caption         =   "Missed Reminders"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   615
      Left            =   7800
      Picture         =   "frmMissed.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   615
      Left            =   6360
      Picture         =   "frmMissed.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid grdShow 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      FixedCols       =   3
      BackColor       =   -2147483639
      BackColorBkg    =   12632256
      HighLight       =   2
      GridLines       =   2
      BorderStyle     =   0
   End
   Begin VB.Label lblType 
      Caption         =   "Missed Reminders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   6015
   End
End
Attribute VB_Name = "frmMissed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function CheckforOldReminders(grdX As MSFlexGrid)
Dim ii As Integer
Dim bb As Integer
Dim bFoundMissed As Boolean

gbCheck = StartGrid2(grdShow)

DoEvents

For ii = 0 To grdX.Rows - 1
  If grdX.TextMatrix(ii, 4) = "Pending" Then
    If Format(grdX.TextMatrix(ii, 0), "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
        bFoundMissed = True
        grdX.TextMatrix(ii, 4) = "Missed"
        grdShow.AddItem Format(grdX.TextMatrix(ii, 0), "dd/mm/yyyy") & Chr(9) & Format(grdX.TextMatrix(ii, 1), "hh:mm") & Chr(9) & grdX.TextMatrix(ii, 2) & Chr(9) & grdX.TextMatrix(ii, 3) & Chr(9) & grdX.TextMatrix(ii, 4)
    End If
  End If
  
  DoEvents
  
Next ii

If bFoundMissed = True Then
    Me.Show vbModal
  Else
    Unload Me
  End If
End Function

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
gbCheck = frmPrint.PrintGrid("Missed Reminders", grdShow)
End Sub

Private Sub grdShow_DblClick()
MsgBox grdShow.TextMatrix(grdShow.MouseRow, 3), vbInformation, "Reminder Details"
End Sub
