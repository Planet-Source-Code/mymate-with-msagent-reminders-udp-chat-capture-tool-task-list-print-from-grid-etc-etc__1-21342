Attribute VB_Name = "modFileFunctions"
Option Explicit


Public Function OpenFileToGrid(sFilename As String, GridX As MSFlexGrid) As Boolean

Dim ii As Integer
Dim sString As String
Dim fno As Integer
Dim lTotalLength As Long
Dim lLength As Long
frmPGB.Show


'obtain the next free file handle from the system
fno = FreeFile

' Open file to read in
Open sFilename For Input As #fno

    lTotalLength = (LOF(fno))

    Do While Not EOF(fno)
        'Start the Input of file
        Line Input #fno, sString
        
        If Len(sString) > 10 Then
          GridX.AddItem sString
        End If
        
        frmPGB.Progress lLength, lTotalLength, "Reading in Data"
        DoEvents
    Loop
Close #fno

Unload frmPGB
'MsgBox "Text Imported Successfull", vbInformation, "BloodLine Import"


End Function
Public Function SaveFileFromGrid(sFilename As String, GridX As MSFlexGrid, Optional NoHeader As Boolean) As Boolean


Dim fno As Integer
Dim fname As String

Dim ii As Integer
Dim bb As Integer

Dim TXTstring
Dim iCount As Integer

Dim StartFrom As Integer

If NoHeader = True Then
  StartFrom = 1
Else
  StartFrom = 0
End If

If Trim(sFilename) = "" Then Exit Function

'frmPGB.Show

'obtain the next free file handle from the system
fno = FreeFile

Open sFilename For Output As #fno 'Open the Text File

For ii = StartFrom To GridX.Rows - 1
   TXTstring = ""
   If Len(GridX.TextMatrix(ii, 1)) > 3 Then
      For bb = 0 To GridX.Cols - 1
         TXTstring = TXTstring & GridX.TextMatrix(ii, bb) & Chr(9)
      Next bb
      TXTstring = Mid(TXTstring, 1, Len(TXTstring) - 1) '& vbCrLf
      Print #fno, (TXTstring) ' Save the line to the Text File
   End If
Next ii

Close #fno 'Close the Text File
'Unload frmPGB

DoEvents

End Function

Public Function CountChar(StringX As String, CharX As String) As Long
'Count the occurance of a character in a string

    Dim ii As Long
    
    For ii = 1 To Len(StringX)
        If Mid(StringX, ii, 1) = CharX Then CountChar = CountChar + 1
    Next ii
    
    'CountChar = CountChar + 1 ' For my needs , it can never be 0
    
End Function
Public Function GetInbedValue(sCountString As String, LookChar As String, sValPos As Integer) As String

Dim IFoundCount As Integer
Dim iCounter As Integer
Dim sFoundValue As String
Dim iLastval As Integer

iLastval = 1
IFoundCount = 0
GetInbedValue = 0

For iCounter = 1 To Len(sCountString)
   If iCounter = Len(sCountString) Then
      If (IFoundCount + 1) = sValPos Then
         If IFoundCount = 0 Then
            GetInbedValue = Trim(sCountString)
         Else
            GetInbedValue = Mid(sCountString, iLastval + 1, (iCounter - (iLastval + 1) + 1))
            If Right(GetInbedValue, 1) = LookChar Then
                GetInbedValue = Left(GetInbedValue, Len(GetInbedValue) - 1)
            End If
         End If
      End If
   Else
      If Mid(sCountString, iCounter, 1) = LookChar Then
      
         IFoundCount = IFoundCount + 1
         If IFoundCount = sValPos Then
            If iLastval = 1 Then
               GetInbedValue = Mid(sCountString, iLastval, (iCounter - iLastval))
            Else
               GetInbedValue = Mid(sCountString, iLastval + 1, (iCounter - (iLastval + 1)))
            End If
         Else
            iLastval = iCounter
         End If
      End If
   End If
Next iCounter

End Function
