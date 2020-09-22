VERSION 5.00
Begin VB.Form frmAddr 
   Caption         =   "Address Form"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Address Book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   8
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txtIp 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   7
         Top             =   720
         Width           =   4215
      End
      Begin VB.ListBox lstAddr 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   6255
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6480
         Picture         =   "frmAddr.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6480
         Picture         =   "frmAddr.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exit the Program"
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "Replace"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   2
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Name"
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
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "IP / Computer name"
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
         TabIndex        =   9
         Top             =   720
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmAddr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ii As Integer
Dim adName As String * 20
Dim adIP As String * 20

Private Sub cmdAdd_Click()
adName = txtName
adIP = txtIp
Dim Result
Result = MsgBox("Are you sure you want to add" & _
vbCrLf & adName & adIP & vbCrLf & _
"to you Address Book ?", vbYesNo, "")

If Result = vbNo Then
    Exit Sub
End If
lstAddr.AddItem adName & adIP
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDel_Click()

adName = txtName
adIP = txtIp

Dim Result
Result = MsgBox("Are you sure you want to delete " & vbCrLf & vbCrLf & lstAddr, vbYesNo, "")

If Result = vbNo Then
    Exit Sub
End If
lstAddr.RemoveItem (lstAddr.ListIndex)
End Sub

Private Sub cmdExit_Click()
On Error Resume Next
Open "address.txt" For Output As #11 ' Open file.
    For ii = 0 To lstAddr.ListCount - 1
        Write #11, lstAddr.List(ii) ' Read line into variable.
        frmQMail.cmbAddress.AddItem (Trim(lstAddr.List(ii)))
    Next ii

Close #11    ' Close file.

Call frmQMail.LoadAddr

Unload Me

End Sub

Private Sub cmdReplace_Click()

adName = txtName
adIP = txtIp

Dim Result
Result = MsgBox("Are you sure you want to replace" _
& vbCrLf & lstAddr & vbCrLf & "with" & vbCrLf & _
txtName & txtIp, vbYesNo, "")

If Result = vbNo Then
    Exit Sub
End If
End Sub

Private Sub Form_Load()

For ii = 0 To frmQMail.cmbAddress.ListCount - 1
lstAddr.AddItem frmQMail.cmbAddress.List(ii)
Next ii

End Sub

Private Sub lstAddr_Click()

txtName = Mid(lstAddr, 1, 20)
txtIp = Mid(lstAddr, 21, 20)

End Sub


