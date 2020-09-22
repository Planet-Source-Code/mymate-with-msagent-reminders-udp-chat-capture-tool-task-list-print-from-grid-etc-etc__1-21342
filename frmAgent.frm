VERSION 5.00
Begin VB.Form frmAgent 
   Caption         =   "Agent Selection"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   Icon            =   "frmAgent.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   615
      Left            =   4680
      Picture         =   "frmAgent.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.OptionButton optAgent 
         Caption         =   "Peedy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   4
         Top             =   4560
         Width           =   1695
      End
      Begin VB.OptionButton optAgent 
         Caption         =   "Genie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   3
         Top             =   2160
         Width           =   1695
      End
      Begin VB.OptionButton optAgent 
         Caption         =   "Robby"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   4560
         Width           =   1695
      End
      Begin VB.OptionButton optAgent 
         Caption         =   "Merlin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   1
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Image Image4 
         Height          =   1800
         Left            =   3840
         Picture         =   "frmAgent.frx":2AAC
         Top             =   2640
         Width           =   1800
      End
      Begin VB.Image Image3 
         Height          =   1800
         Left            =   600
         Picture         =   "frmAgent.frx":5E1F
         Top             =   2640
         Width           =   1800
      End
      Begin VB.Image Image2 
         Height          =   1800
         Left            =   3840
         Picture         =   "frmAgent.frx":9FBF
         Top             =   240
         Width           =   1800
      End
      Begin VB.Image Image1 
         Height          =   1800
         Left            =   600
         Picture         =   "frmAgent.frx":D67A
         Top             =   240
         Width           =   1800
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   6240
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   3120
         X2              =   3120
         Y1              =   120
         Y2              =   4920
      End
   End
End
Attribute VB_Name = "frmAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()

If optAgent(0).Value = True Then 'Merlin
  SaveSetting "MyMate", "Merlin", "Name", "Merlin.acs"
End If

If optAgent(1).Value = True Then 'Robby
  SaveSetting "MyMate", "Merlin", "Name", "Robby.acs"
End If

If optAgent(2).Value = True Then 'Genie
  SaveSetting "MyMate", "Merlin", "Name", "Genie.acs"
End If

If optAgent(3).Value = True Then 'Peedy
  SaveSetting "MyMate", "Merlin", "Name", "Peedy.acs"
End If



Unload Me

End Sub
