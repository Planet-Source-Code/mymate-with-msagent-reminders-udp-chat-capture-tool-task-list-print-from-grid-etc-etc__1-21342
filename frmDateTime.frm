VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRemindMaint 
   Caption         =   "Reminder add/edit"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   Icon            =   "frmDateTime.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   2520
      Picture         =   "frmDateTime.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   615
      Left            =   4080
      Picture         =   "frmDateTime.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox txtSubject 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Text            =   "Subject"
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox txtDesc 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   5535
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24510466
      CurrentDate     =   36861.3333333333
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24510464
      CurrentDate     =   36861
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select the Date"
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
      TabIndex        =   7
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select the Time"
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   1080
      Width           =   5535
   End
End
Attribute VB_Name = "frmRemindMaint"
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

If giType = 2 Then ' Change existing Reminder
  DTPicker1 = frmMymate.grdMain.TextMatrix(glRemKey, 0)
  DTPicker2 = frmMymate.grdMain.TextMatrix(glRemKey, 1)
  txtSubject = frmMymate.grdMain.TextMatrix(glRemKey, 2)
  txtDesc = frmMymate.grdMain.TextMatrix(glRemKey, 3)
Else
  DTPicker1 = Now
  DTPicker2 = Now + 5
End If

bOk = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

If bOk = True Then
   If giType = 1 Then ' New Reminder
      merlin.StopAll
      merlin.Play "write"
      frmMymate.grdMain.AddItem DTPicker1.Value & Chr(9) & DTPicker2.Value & Chr(9) & txtSubject & Chr(9) & txtDesc & Chr(9) & "Pending"
      merlin.Speak "The new reminder has been entered !"
   End If
   
   If giType = 2 Then ' Change existing Reminder
      frmMymate.grdMain.RemoveItem glRemKey
      merlin.StopAll
      merlin.Play "write"
      frmMymate.grdMain.AddItem DTPicker1.Value & Chr(9) & DTPicker2.Value & Chr(9) & txtSubject & Chr(9) & txtDesc & Chr(9) & "Pending"
      merlin.Speak "The deed is done !"
   End If
   
End If


gbCheck = frmMymate.RefreshGrid
frmMymate.tmrIrri.Enabled = True

End Sub
