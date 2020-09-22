VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTasks 
   Caption         =   "Tasks"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   Icon            =   "frmTasks.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   615
      Left            =   7800
      Picture         =   "frmTasks.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Task"
      Height          =   615
      Left            =   0
      Picture         =   "frmTasks.frx":2AAC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change Task"
      Height          =   615
      Left            =   1440
      Picture         =   "frmTasks.frx":524E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Task"
      Height          =   615
      Left            =   2880
      Picture         =   "frmTasks.frx":79F0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   615
      Left            =   6240
      Picture         =   "frmTasks.frx":A192
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid grdMain 
      Height          =   5895
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10398
      _Version        =   393216
      Cols            =   4
      BackColor       =   -2147483639
      BackColorBkg    =   12632256
      HighLight       =   2
      GridLines       =   2
      BorderStyle     =   0
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "Tasks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   7
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label lblType 
      Caption         =   "Tasks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Index           =   0
      Left            =   4500
      TabIndex        =   6
      Top             =   6150
      Width           =   1695
   End
End
Attribute VB_Name = "frmTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colRems As New Collection

Private Sub cmdAdd_Click()

merlin.StopAll

giType2 = 1 ' New Task
frmMymate.tmrIrri.Enabled = False

merlin.Play "Read"
merlin.Speak "Well, it is about time you did something around here !"


DoEvents
frmTaskMaint.Show
frmTaskMaint.DTPicker1.Value = Now()
DoEvents

End Sub

Private Sub cmdChange_Click()

If grdMain.Row < 1 Or Len(grdMain.TextMatrix(grdMain.Row, 0)) < 4 Then
  MsgBox "Please select a task to change from the grid", vbInformation
  Exit Sub
End If

merlin.StopAll

giType2 = 2 ' Change Reminder
glRemKey2 = grdMain.Row
frmMymate.tmrIrri.Enabled = False

merlin.Play "Read"
merlin.Speak "Oh. If you did it right the first time, we would not be changing it now, would we ?"


DoEvents
frmTaskMaint.Show

DoEvents
End Sub

Private Sub cmdDelete_Click()

On Error GoTo Error_H

If grdMain.Row < 1 Then
  MsgBox "Please select a task of the list to delete", vbInformation, "DeskMate"
  Exit Sub
End If

Dim Result
Result = MsgBox("Are you sure you want to delete : " & grdMain.TextMatrix(grdMain.Row, 3) & " ?", vbYesNo)
If Result <> vbYes Then
  Exit Sub
End If

If grdMain.Rows < 3 Then
  grdMain.AddItem Chr(9)
End If

'grdMain.RemoveItem grdMain.TextMatrix(grdMain.Row, 0)
grdMain.RemoveItem grdMain.Row

merlin.StopAll
merlin.Play "DoMagic1"
merlin.Play "DoMagic2"
merlin.Speak "It's gone !"

DoEvents

gbCheck = SaveFileFromGrid(gsFileName2, frmTasks.grdMain, True)

Exit Sub

Error_H:
  MsgBox Err.Description
End Sub



Private Sub cmdOk_Click()
Me.Visible = False
gbCheck = SaveFileFromGrid(gsFileName2, grdMain, True)
End Sub


Private Sub cmdPrint_Click()
gbCheck = frmPrint.PrintGrid("Tasks", grdMain)
End Sub

Private Sub Form_Load()

gsFileName2 = App.Path & "\" & "tskdata.txt"

gbCheck = StartGrid2(grdMain)

If Trim(Dir(gsFileName2)) <> "" Then
  gbCheck = OpenFileToGrid(gsFileName2, grdMain)
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)

gbCheck = SaveFileFromGrid(gsFileName2, grdMain, True)

End Sub

Private Sub mnuAddApp_Click()
Call cmdAdd_Click
End Sub

Private Sub mnuApp_Click()
Me.Visible = True
End Sub
Public Function RefreshGrid() As Boolean
Dim ii As Integer

gbCheck = StartGrid2(grdMain)

For ii = 0 To grdMain.Rows - 1
  grdMain.AddItem ii & Chr(9) & grdMain.TextMatrix(ii, 0) & Chr(9) & Format(grdMain.TextMatrix(ii, 1), "hh:mm") & Chr(9) & grdMain.TextMatrix(ii, 2) & Chr(9) & grdMain.TextMatrix(ii, 3)
Next ii

'Save the reminders
gbCheck = SaveFileFromGrid(gsFileName, grdMain)

grdMain.Col = 3
grdMain.Sort = flexSortGenericAscending
End Function


Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuHide_Click()
merlin.Hide
End Sub

Private Sub mnuShow_Click()
merlin.Show
End Sub

Private Sub grdMain_DblClick()
Call cmdChange_Click
End Sub
