VERSION 5.00
Begin VB.Form frmRemind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please Note"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   ClipControls    =   0   'False
   Icon            =   "frmRemind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   615
      Left            =   4200
      Picture         =   "frmRemind.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtTime 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Text            =   "Subject"
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox txtSubject 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Text            =   "Subject"
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox txtDesc 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Timer tmrAnnounce 
      Interval        =   20000
      Left            =   120
      Top             =   3360
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "You have a reminder !"
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
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   5535
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time"
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
      Left            =   120
      TabIndex        =   4
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
      Left            =   120
      TabIndex        =   3
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
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   5535
   End
End
Attribute VB_Name = "frmRemind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
    frmMymate.tmrIrri.Enabled = False
    

    merlin.StopAll
    merlin.Play "announce"
    merlin.Speak "Hearken all, hearken all ! A new reminder came to the ball."
    merlin.Play "announce"
    
    Me.txtTime = Format(Now, "hh:mm")
    Me.txtSubject = gsSubject
    Me.txtDesc = gsDescription
    Me.Caption = Me.Caption & ": " & gsSubject
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
merlin.StopAll
frmMymate.tmrIrri.Enabled = True
tmrAnnounce.Enabled = False
End Sub

Private Sub tmrAnnounce_Timer()
  frmMymate.tmrIrri.Enabled = False
  merlin.StopAll
  merlin.Play ("announce")
  merlin.Speak txtDesc
End Sub
