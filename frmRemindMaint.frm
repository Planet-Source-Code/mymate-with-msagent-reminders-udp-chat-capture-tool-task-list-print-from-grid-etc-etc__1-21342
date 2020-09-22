VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMymate 
   Caption         =   "My Mate Merlin"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   Icon            =   "frmRemindMaint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   226
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   613
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox chkAlert 
      Caption         =   "Message Alert"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock sckMail 
      Left            =   5880
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrNextDay 
      Interval        =   60000
      Left            =   5880
      Top             =   2760
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   615
      Left            =   6360
      Picture         =   "frmRemindMaint.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Timer tmrAlarm 
      Interval        =   60000
      Left            =   8760
      Top             =   2760
   End
   Begin VB.Timer tmrIrri 
      Interval        =   40000
      Left            =   7680
      Top             =   2760
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Reminder"
      Height          =   615
      Left            =   2880
      Picture         =   "frmRemindMaint.frx":2AAC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change Reminder"
      Height          =   615
      Left            =   1440
      Picture         =   "frmRemindMaint.frx":524E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Reminder"
      Height          =   615
      Left            =   0
      Picture         =   "frmRemindMaint.frx":79F0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid grdMain 
      Height          =   1335
      Left            =   0
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2355
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      GridLines       =   3
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   615
      Left            =   7800
      Picture         =   "frmRemindMaint.frx":A192
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24444928
      CurrentDate     =   36861
   End
   Begin MSFlexGridLib.MSFlexGrid grdShow 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   360
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
   Begin VB.Timer tmrDisplay 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   7800
      Top             =   6600
   End
   Begin MSFlexGridLib.MSFlexGrid grdCurrent 
      Height          =   2295
      Left            =   0
      TabIndex        =   8
      Top             =   6600
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      FixedCols       =   3
      BackColor       =   -2147483639
      BackColorBkg    =   12632256
      HighLight       =   2
      BorderStyle     =   0
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "Reminders"
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
      TabIndex        =   13
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblType 
      Caption         =   "Reminders"
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
      TabIndex        =   10
      Top             =   2910
      Width           =   1695
   End
   Begin VB.Label lblFlexDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reminders qued for today"
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
      TabIndex        =   9
      Top             =   6120
      Width           =   9135
   End
   Begin VB.Label lblFlexDesc 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Active reminders for date"
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
      TabIndex        =   1
      Top             =   0
      Width           =   6135
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   7200
      Top             =   6600
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuApp 
         Caption         =   "Show Reminders"
      End
      Begin VB.Menu mnuAddApp 
         Caption         =   "Add Reminder"
      End
      Begin VB.Menu xx 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuShowTasks 
         Caption         =   "Show Tasks"
      End
      Begin VB.Menu xxxxx 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show Mate"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Hide Mate"
      End
      Begin VB.Menu mnuChange 
         Caption         =   "Change Agent"
      End
      Begin VB.Menu sdfgsdfg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMail 
         Caption         =   "Quick Mail"
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "Quick Capture"
      End
      Begin VB.Menu xxx 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit MyMate"
      End
   End
End
Attribute VB_Name = "frmMymate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colRems As New Collection


Private Sub Agent1_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)

merlin.Play "Surprised"
Me.PopupMenu mnuPopup, , Screen.Height / 2, Screen.Width / 2

End Sub

Private Sub Agent1_DragComplete(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
merlin.Speak "Stop horsing around !"
End Sub

Private Sub cmdAdd_Click()

merlin.StopAll

giType = 1 ' New Reminder
tmrIrri.Enabled = False

merlin.Play "Read"
merlin.Speak "Alright ! A new reminder. Well, get on with it !"


DoEvents
frmRemindMaint.Show
frmRemindMaint.DTPicker1.Value = DTPicker1.Value
DoEvents

End Sub

Private Sub cmdChange_Click()

If grdShow.Row < 1 Then
  MsgBox "Please select a reminder to change from the grid", vbInformation
  Exit Sub
End If

merlin.StopAll

giType = 2 ' Change Reminder
glRemKey = grdShow.TextMatrix(grdShow.Row, 0)
tmrIrri.Enabled = False

merlin.Play "Read"
merlin.Speak "Oh. If you did it right the first time, we would not be changing it now, would we ?"


DoEvents
frmRemindMaint.Show
frmRemindMaint.DTPicker1.Value = DTPicker1.Value
DoEvents
End Sub

Private Sub cmdDelete_Click()

On Error GoTo Error_H

If grdShow.Row < 1 Then
  MsgBox "Please select a reminder of the list to delete", vbInformation, "DeskMate"
  Exit Sub
End If

Dim Result
Result = MsgBox("Are you sure you want to delete : " & grdShow.TextMatrix(grdShow.Row, 3) & " ?", vbYesNo)
If Result <> vbYes Then
  Exit Sub
End If

grdMain.RemoveItem grdShow.TextMatrix(grdShow.Row, 0)

merlin.StopAll
merlin.Play "DoMagic1"
merlin.Play "DoMagic2"
merlin.Speak "It's gone !"

DoEvents

RefreshGrid
Exit Sub

Error_H:
  MsgBox Err.Description
End Sub



Private Sub cmdOk_Click()
Me.Visible = False
End Sub

Private Sub cmdPrint_Click()
gbCheck = frmPrint.PrintGrid("Reminders", grdShow)
End Sub

Private Sub DTPicker1_Change()
gbCheck = RefreshGrid
End Sub

Private Sub Form_Load()

'sckMail.LocalPort = 3060
'sckMail.Listen
sckMail.Protocol = sckUDPProtocol
sckMail.RemotePort = 3060 ' Port to connect to.
sckMail.LocalPort = 3060
sckMail.Close
DoEvents
sckMail.Bind 3060     ' Bind to the local port.

gsFileName = App.Path & "\" & "remdata.txt"

If Trim(Dir(gsFileName)) <> "" Then
  gbCheck = OpenFileToGrid(gsFileName, grdMain)
End If

gbCheck = frmMissed.CheckforOldReminders(grdMain)

grdMain.Cols = 5

DTPicker1 = Now

gbCheck = StartGrid(grdShow)
gbCheck = RefreshGrid
Init

Call LoadAgent


End Sub

Public Sub LoadAgent()
gsAgentName = GetSetting("MyMate", "Merlin", "Name", "Merlin.acs")
Agent1.Characters.Load gsAgentName, gsAgentName

Set merlin = Agent1.Characters(gsAgentName)
merlin.Width = merlin.Width / 1.5
merlin.Height = merlin.Height / 1.5
SavedWidth = merlin.Width
SAvedHeight = merlin.Height

merlin.Top = GetSetting("MyMate", "Merlin", "Top", 0)
merlin.Left = GetSetting("MyMate", "Merlin", "Left", 0)

merlin.Show
DoEvents
merlin.Play "surprised"
merlin.Speak "Huh ? What ? Oh !"
merlin.Play "announce"
merlin.Play "greet"
DoEvents

On Error Resume Next
merlin.Speak GetRandom(Greetings)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmTasks
merlin.StopAll
merlin.Play "greet"
merlin.Speak "Until later then."
DoEvents

gbCheck = SaveFileFromGrid(gsFileName, grdMain)
SaveSetting "MyMate", "Merlin", "Top", merlin.Top
SaveSetting "MyMate", "Merlin", "Left", merlin.Left
End Sub

Private Sub grdShow_DblClick()

Call cmdChange_Click

End Sub

Private Sub mnuAddApp_Click()
Call cmdAdd_Click
End Sub

Private Sub mnuAddTask_Click()
Me.Visible = False
frmTasks.Show
DoEvents
'Call frmTasks.cmdAdd_Click
End Sub

Private Sub mnuApp_Click()
Me.Visible = True
End Sub
Public Function RefreshGrid() As Boolean
Dim ii As Integer

gbCheck = StartGrid(grdShow)

For ii = 0 To grdMain.Rows - 1
  If Format(grdMain.TextMatrix(ii, 0), "dd/mm/yyyy") = Format(DTPicker1.Value, "dd/mm/yyyy") Then
    grdShow.AddItem ii & Chr(9) & grdMain.TextMatrix(ii, 0) & Chr(9) & Format(grdMain.TextMatrix(ii, 1), "hh:mm") & Chr(9) & grdMain.TextMatrix(ii, 2) & Chr(9) & grdMain.TextMatrix(ii, 3) & Chr(9) & grdMain.TextMatrix(ii, 4)
  End If
Next ii

For ii = 0 To grdShow.Rows - 1

  If grdShow.TextMatrix(ii, 4) = "Pending" Then
    grdShow.Row = ii
    grdShow.Col = 4
    grdShow.CellForeColor = vbGreen
  End If
  
Next ii

'Make sure the Current grid is kept up to date
gbCheck = GetCurrentGrid

'Save the reminders
gbCheck = SaveFileFromGrid(gsFileName, grdMain)

grdShow.Col = 2
grdShow.Sort = flexSortGenericAscending
grdShow.Refresh

End Function
Public Function GetCurrentGrid() As Boolean
Dim ii As Integer

gbCheck = StartGrid(grdCurrent)

For ii = 0 To grdMain.Rows - 1
  If Format(grdMain.TextMatrix(ii, 0), "dd/mm/yyyy") = Format(Now, "dd/mm/yyyy") Then
    grdCurrent.AddItem ii & Chr(9) & grdMain.TextMatrix(ii, 0) & Chr(9) & Format(grdMain.TextMatrix(ii, 1), "hh:mm") & Chr(9) & grdMain.TextMatrix(ii, 2) & Chr(9) & grdMain.TextMatrix(ii, 3) & Chr(9) & grdMain.TextMatrix(ii, 4)
  End If
Next ii

'sortFlex grdShow, 2, False, True, True, True
grdCurrent.Col = 3
grdCurrent.Sort = flexSortGenericDescending
End Function

Private Sub mnuCapture_Click()
frmCapture.Show
End Sub

Private Sub mnuChange_Click()
frmAgent.Show vbModal
frmMymate.Agent1.Characters.Unload gsAgentName
Call LoadAgent
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuHide_Click()
merlin.StopAll
merlin.Hide
DoEvents
End Sub

Private Sub mnuMail_Click()
frmQMail.Show
End Sub

Private Sub mnuShow_Click()
merlin.Show
End Sub

Private Sub mnuShowTasks_Click()
Me.Visible = False
frmTasks.Show
End Sub

Private Sub sckMail_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
    
    sckMail.GetData strData
    frmQMail.lstHist.AddItem strData
    frmQMail.lstHist.ListIndex = frmQMail.lstHist.ListCount - 1
    
    If chkAlert.Value = 1 Then
      merlin.Show
      merlin.StopAll
      merlin.Speak "New message sire !"
      frmTag.ShowThis ("New Message")
      DoEvents
    End If
    
    DoEvents
 
End Sub

Private Sub tmrAlarm_Timer()
'This is the timer that will check to see if an alarm is triggered.
Dim ii As Integer
Dim bb As Integer

For ii = 1 To grdCurrent.Rows - 1
  If Format(grdCurrent.TextMatrix(ii, 2), "hh:mm") = Format(Now, "hh:mm") Then
    grdCurrent.TextMatrix(ii, 5) = "Reminded"
    grdMain.TextMatrix(grdCurrent.TextMatrix(ii, 0), 4) = "Reminded"
    merlin.Show
    gsSubject = grdCurrent.TextMatrix(ii, 3)
    gsDescription = grdCurrent.TextMatrix(ii, 4)
    Dim frmReminder As New frmRemind
    colRems.Add frmReminder
    colRems(colRems.Count).Show
    Set frmReminder = Nothing
    frmTag.ShowThis ("Reminder")
  End If
Next ii

gbCheck = StartGrid(grdShow)
gbCheck = RefreshGrid

End Sub


Private Sub tmrIrri_Timer()
Dim iActions As Integer

DoEvents
iActions = Int(Rnd * 14)
gbCheck = RandomAction(iActions)
DoEvents

End Sub

Private Sub tmrNextDay_Timer()
  gbCheck = StartGrid(grdShow)
  gbCheck = RefreshGrid
End Sub
