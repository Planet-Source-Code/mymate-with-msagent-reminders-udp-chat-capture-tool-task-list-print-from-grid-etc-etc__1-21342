VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTaskMaint 
   Caption         =   "Task add/edit"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   Icon            =   "frmTaskMaint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Left            =   2400
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox txtSubject 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Text            =   "Subject"
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox txtDesc 
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   5535
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   615
      Left            =   4080
      Picture         =   "frmTaskMaint.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   2520
      Picture         =   "frmTaskMaint.frx":2AAC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24444930
      CurrentDate     =   36861.3333333333
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24444928
      CurrentDate     =   36861
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   0
      TabIndex        =   10
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Due Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Due Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   5535
   End
End
Attribute VB_Name = "frmTaskMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOk As Boolean

Private Sub cmdCancel_Click()
bOk = False
Unload Me
End Sub

Private Sub cmdOk_Click()
bOk = True
Unload Me
End Sub

Private Sub Form_Load()

cboStatus.Clear
cboStatus.AddItem "Not Started"
cboStatus.AddItem "Pending"
cboStatus.AddItem "Busy"
cboStatus.AddItem "Urgent"
cboStatus.AddItem "User Wait"
cboStatus.AddItem "Finished"
cboStatus.AddItem "Problems"
cboStatus.AddItem "In Testing"



If giType2 = 2 Then ' Change existing Reminder
  DTPicker1 = frmTasks.grdMain.TextMatrix(glRemKey2, 0)
  DTPicker2 = frmTasks.grdMain.TextMatrix(glRemKey2, 1)
  txtSubject = frmTasks.grdMain.TextMatrix(glRemKey2, 2)
  txtDesc = frmTasks.grdMain.TextMatrix(glRemKey2, 3)
  cboStatus.Text = frmTasks.grdMain.TextMatrix(glRemKey2, 4)
Else
  DTPicker1 = Now
  DTPicker2 = Now
End If

bOk = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

If bOk = True Then
   If giType2 = 1 Then ' New Reminder
      merlin.StopAll
      merlin.Play "write"
      frmTasks.grdMain.AddItem Format(DTPicker1.Value, "dd/mm/yyyy") & Chr(9) & Format(DTPicker2.Value, "hh:mm") & Chr(9) & txtSubject & Chr(9) & txtDesc & Chr(9) & cboStatus
      merlin.Speak "The new task was added to the list, now go and do it !"
   End If
   
   If giType2 = 2 Then ' Change existing Reminder
      If frmTasks.grdMain.Rows < 3 Then frmTasks.grdMain.AddItem Chr(9)
      frmTasks.grdMain.RemoveItem glRemKey2
      merlin.StopAll
      merlin.Play "write"
      frmTasks.grdMain.AddItem Format(DTPicker1.Value, "dd/mm/yyyy") & Chr(9) & Format(DTPicker2.Value, "hh:mm") & Chr(9) & txtSubject & Chr(9) & txtDesc & Chr(9) & cboStatus
      merlin.Speak "The task details was changed..."
   End If
   
End If

gbCheck = SaveFileFromGrid(gsFileName2, frmTasks.grdMain, True)


frmMymate.tmrIrri.Enabled = True

End Sub

