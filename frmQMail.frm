VERSION 5.00
Begin VB.Form frmQMail 
   Caption         =   " Quick Mail"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   ControlBox      =   0   'False
   Icon            =   "frmQMail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   5160
      Width           =   6135
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear messages"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   315
         Left            =   4440
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Recipient"
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   4440
      Width           =   6135
      Begin VB.CommandButton cmdAddr 
         Caption         =   "Address Book"
         Height          =   315
         Left            =   4440
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbAddress 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Send Message"
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   6135
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send Message"
         Default         =   -1  'True
         Height          =   735
         Left            =   4440
         Picture         =   "frmQMail.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtSend 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.ListBox lstHist 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "frmQMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddr_Click()
frmAddr.Show
End Sub

Private Sub cmdSend_Click()

On Error GoTo Error_H
  With frmMymate

    lstHist.AddItem "To : " & Trim(Mid(cmbAddress, 1, 20)) & " : " & txtSend.Text
    lstHist.ListIndex = lstHist.ListCount - 1

    
    If Len(Trim(Trim(Mid(cmbAddress, 21)))) < 3 Then
      MsgBox "Please select a valid recipient from the address list"
      Exit Sub
    End If
    
    .sckMail.RemoteHost = Trim(Mid(cmbAddress, 21))
    .sckMail.SendData .sckMail.LocalHostName & " : " & txtSend.Text
    txtSend.Text = ""
  End With
  
Exit Sub

Error_H:
MsgBox Err.Description & vbCrLf & "On Message send"
End Sub

Private Sub cmdClose_Click()
  Me.Hide
End Sub

Private Sub Form_Load()

Call LoadAddr

End Sub


Private Sub lstHist_DblClick()
MsgBox lstHist.List(lstHist.ListIndex)
End Sub

Public Sub LoadAddr()

On Error Resume Next

Dim ii As Integer
Dim TextLine As String

If Dir("address.txt") = "" Then
  MsgBox "Your address book has no entries. Please insert some entries."
  Exit Sub
End If

cmbAddress.Clear

Open "address.txt" For Input As #10 ' Open file.
    Do While Not EOF(10) ' Loop until end of file.
        ii = ii + 1
        Line Input #10, TextLine ' Read line into variable
        cmbAddress.AddItem (Trim(Mid(TextLine, 2, Len(TextLine) - 2)))
    Loop
Close #10    ' Close file.

cmbAddress.ListIndex = 0
End Sub
