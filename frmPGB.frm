VERSION 5.00
Begin VB.Form frmPGB 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5610
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAnimate 
      Interval        =   100
      Left            =   4575
      Top             =   240
   End
   Begin VB.PictureBox Picture1 
      Height          =   250
      Left            =   15
      ScaleHeight     =   195
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   945
      Width           =   5580
      Begin VB.CheckBox chkPrg 
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   200
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   5535
      End
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Busy ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   15
      TabIndex        =   4
      Top             =   0
      Width           =   5535
   End
   Begin VB.Image imgGo 
      Height          =   480
      Index           =   0
      Left            =   255
      Picture         =   "frmPGB.frx":0000
      Top             =   210
      Width           =   480
   End
   Begin VB.Image imgGo 
      Height          =   480
      Index           =   1
      Left            =   2655
      Picture         =   "frmPGB.frx":030A
      Top             =   210
      Width           =   480
   End
   Begin VB.Image imgEnd 
      Height          =   480
      Left            =   5040
      Picture         =   "frmPGB.frx":0614
      Top             =   240
      Width           =   480
   End
   Begin VB.Image imgStart 
      Height          =   480
      Left            =   15
      Picture         =   "frmPGB.frx":091E
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblPerc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0% Completed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   15
      TabIndex        =   3
      Top             =   1260
      Width           =   5535
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Processing ..."
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   735
      Width           =   5580
   End
End
Attribute VB_Name = "frmPGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Author : Renier Barnard (renier_barnard@santam.co.za)
'
' Date    : July 1999
'
' Description :
' This code will demonstrate how to make a simple but nice
' looking progress bar. It could be more simple (Using the line command)
' but this looks better. The form_click event will start the progress bar of.
' Try resizing the progress bar form. There is some code to demonstrate
' how to make something like this generic in size !
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Declare Function timeGetTime Lib "winmm.dll" () As Long
'Private Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const FLAGS = 1
Const HWND_TOPMOST = -1
Dim Aindex As Integer
Dim LastPos As Long
Dim lLastTime As Double
Dim tLastTime


Public Function Progress(Value, MaxValue, Optional HeaderX As String, Optional color As ColorConstants)
'' This is the actual progress bar function.

DoEvents
Dim Perc
Dim bb As Integer
Dim lTime As Double
Dim lTimeDiff As Double
Dim lTimeLeft As Double
Dim lTotalTime As Double
'Me.Show

'Get a color to do it in
If color = 0 Then color = vbBlack
color = vbBlack ' Override anyway

If MaxValue = 0 Then MaxValue = 1

'Display the header , if any was returned
If HeaderX <> "" Then
    lblHeader = HeaderX
Else
    lblHeader = "Busy Processing...Please wait"
End If

'Now work out the percentage (0-100) of where we currently are
Perc = (Value / MaxValue) * 100
If Perc < 0 Then Perc = 0
If Perc > 100 Then Perc = 100
Perc = Int(Perc)

'Do the time remaining calculation
If (Perc Mod 10) = 0 Or Perc = 0 Then ' Every 10 percent
        lTimeDiff = lTime - lLastTime
        lTime = Time - tLastTime
        If Perc = 0 Or Perc < 0 Then
            lTotalTime = ((100 / 1) * 2) * lTime
            lTimeLeft = (((100 / 1) * 2) * lTime) - (((100 / 100) * 2) * lTime)
        Else
            lTotalTime = ((100 / Perc) * 2) * lTime
            lTimeLeft = (((100 / Perc) * 2) * lTime) - (((100 / 100) * 2) * lTime)
        End If
        lblTime = "Time Remaining : " & Format((lTimeLeft), "hh:mm:ss")

End If
DoEvents

DoEvents
lblPerc.Caption = Int(Perc) & "% Completed" 'Just the Label Display
chkPrg.Width = Int(Perc)

DoEvents

End Function



Private Sub Form_Load()

'Set Me.Picture = mdiMain.Picture
DoEvents
tLastTime = Time

Const FLAGS = 1
Const HWND_TOPMOST = -1
Aindex = 0
LastPos = 720

Me.Width = 5910
Me.Height = 1545

Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2

'Sets form on always on top.
Dim Success As Integer
'Success% = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
                                                ' Change the "0's" above to position the window.

Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2
DoEvents

End Sub



Private Sub Form_Unload(Cancel As Integer)
DoEvents
End Sub

Private Sub tmrAnimate_Timer()
'This funtion will animate a couple of icons , just to show that something is busy hapening

DoEvents
LastPos = LastPos + 1


If LastPos > 2680 And LastPos < 3250 Then
    LastPos = 3160
    Aindex = 1
Else
    If LastPos > 5360 Then
        LastPos = 720
        Aindex = 0
    Else
        
    End If
End If

If Aindex = 1 Then
    imgGo(1).Visible = True
    imgGo(0).Visible = False
Else
    imgGo(1).Visible = False
    imgGo(0).Visible = True
End If

LastPos = LastPos + 200
imgGo(Aindex).Left = LastPos
DoEvents

End Sub


