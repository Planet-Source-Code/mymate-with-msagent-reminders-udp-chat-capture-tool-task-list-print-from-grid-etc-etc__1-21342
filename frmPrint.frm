VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmPrint 
   ClientHeight    =   7560
   ClientLeft      =   -2685
   ClientTop       =   -1095
   ClientWidth     =   9150
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   945
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   9090
      TabIndex        =   1
      Top             =   6600
      Width           =   9150
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   615
         Left            =   7875
         Picture         =   "frmPrint.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   615
         Left            =   5355
         Picture         =   "frmPrint.frx":2AAC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   105
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrintToFile 
         Caption         =   "Print To File"
         Height          =   615
         Left            =   6615
         Picture         =   "frmPrint.frx":2EEE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Print Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   0
         TabIndex        =   5
         Top             =   240
         Width           =   4815
      End
   End
   Begin SHDocVwCtl.WebBrowser Browse 
      Height          =   6540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      ExtentX         =   16113
      ExtentY         =   11536
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PageCount As Integer
Dim iType As Integer
Dim iRowCount As Integer
Dim iMaxRows As Integer
Dim txtHTML As String
Dim ii As Integer

Public Function PrintGrid(sHeader As String, GridX As Object) As Boolean
On Error Resume Next
txtHTML = ConvertGridToHTML(GridX, sHeader)

Kill App.Path & "/test.html"

'Now Rewrite the whole file
Open App.Path & "/test.html" For Output As #2 ' Open file for output."

        Print #2, txtHTML  'Save Record

Close #2    ' Close file.


For ii = 1 To 10000
DoEvents
Next ii

Browse.Navigate "file:" & App.Path & "/test.html"
Me.Refresh
Me.Show

End Function


Private Sub cmdCancel_Click()
Unload Me
End Sub



Private Sub cmdPrint_Click()
On Error Resume Next

'Browse.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
Browse.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
'Unload Me
End Sub

Private Sub cmdPrintToFile_Click()
On Error Resume Next
Browse.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER
'Unload Me
End Sub

Private Sub Form_Load()
Me.Top = 100
Me.Left = 100
End Sub

