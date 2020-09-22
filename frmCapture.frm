VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCapture 
   Caption         =   "Quick Capture"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4260
   Icon            =   "frmCapture.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3315
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   0
      Width           =   2955
      Begin VB.PictureBox Picture1 
         Height          =   2925
         Left            =   0
         ScaleHeight     =   2865
         ScaleWidth      =   2790
         TabIndex        =   1
         Top             =   210
         Width           =   2850
      End
   End
   Begin MSComDlg.CommonDialog comD 
      Left            =   3150
      Top             =   2835
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   3000
      TabIndex        =   2
      Top             =   0
      Width           =   1215
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   1024
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   2130
         Width           =   1024
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clear"
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   1815
         Width           =   1024
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print"
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   1500
         Width           =   1024
      End
      Begin VB.CommandButton cmdActive 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Active"
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   1185
         Width           =   1024
      End
      Begin VB.CommandButton cmdClient 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Client"
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   870
         Width           =   1024
      End
      Begin VB.CommandButton cmdForm 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Form"
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   555
         Width           =   1024
      End
      Begin VB.CommandButton cmdScreen 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Screen"
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1024
      End
   End
End
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
      
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim SaveFile As String
comD.Filter = "*.BMP"
  comD.ShowSave
  If comD.CancelError Then Exit Sub
  
  SaveFile = Trim(comD.FileName)
  
  If comD.FileName <> "" Then
    If UCase(Right(SaveFile, 3)) <> "BMP" Then
        SaveFile = SaveFile & ".bmp"
    End If
  End If
  
  If SaveFile <> "" Then
    SavePicture Picture1.Picture, SaveFile    ' Save picture to file.
  End If
End Sub

      '--------------------------------------------------------------------
      ' Capture the entire screen
      Private Sub cmdScreen_Click()
         Set Picture1.Picture = CaptureScreen()
      End Sub

      ' Capture the entire form including title and border
      Private Sub cmdForm_Click()
         Set Picture1.Picture = CaptureForm(Me)
      End Sub

      ' Capture the client area of the form
      Private Sub cmdClient_Click()
         Set Picture1.Picture = CaptureClient(Me)
      End Sub

      ' Capture the active window after two seconds
      Private Sub cmdActive_Click()
         MsgBox "Five seconds after you close this dialog " & _
            "the active window will be captured."

         ' Wait for two seconds
         Dim EndTime As Date
         EndTime = DateAdd("s", 5, Now)
         Do Until Now > EndTime
            DoEvents
         Loop

         Set Picture1.Picture = CaptureActiveWindow()

         ' Set focus back to form
         Me.SetFocus
      End Sub

      ' Print the current contents of the picture box
      Private Sub cmdPrint_Click()
         PrintPictureToFitPage Printer, Picture1.Picture
         Printer.EndDoc
      End Sub

      ' Clear out the picture box
      Private Sub cmdClear_Click()
         Set Picture1.Picture = Nothing
      End Sub

      Private Sub Form_Load()
         Picture1.AutoSize = True
      End Sub
      '--------------------------------------------------------------------




