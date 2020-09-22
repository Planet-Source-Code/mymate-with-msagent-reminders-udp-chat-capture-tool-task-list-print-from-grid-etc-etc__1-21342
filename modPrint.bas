Attribute VB_Name = "modPrint"
Option Explicit
Global Progress

Public Function ConvertGridToHTML(GridX As Object, Heading As String) As String
Dim HTML As String
Dim ii As Integer
Dim bb As Integer

On Error Resume Next

Load frmPGB
frmPGB.Show

HTML = AddHeaderToHTML()

HTML = HTML & vbCrLf & "<hr><br>"

HTML = HTML & vbCrLf & "<TABLE BORDER=1 WIDTH=100% bgcolor='#c0c0c0'>"
HTML = HTML & "<FONT SIZE = +1>"
HTML = HTML & vbCrLf & Heading
HTML = HTML & "</FONT>"
HTML = HTML & "</TD></TR>" & vbCrLf
HTML = HTML & "</TABLE>" & vbCrLf

HTML = HTML & "<FONT SIZE = -1>"

HTML = HTML & vbCrLf & "<TABLE BORDER=1 WIDTH=100%>"
For ii = 0 To GridX.Rows - 1
Progress = frmPGB.Progress(ii, GridX.Rows - 1, "Generating HTML for Grid")
    If ii = 0 Then
        HTML = HTML & vbCrLf & "<TR bgcolor='#c0c0c0'>"
    Else
        HTML = HTML & vbCrLf & "<TR>"
    End If
    
    For bb = 0 To GridX.Cols - 1
        HTML = HTML & vbCrLf & "<TD>"
        If ii = 0 Then
            If Trim(GridX.TextMatrix(ii, bb)) = "" Then
                HTML = HTML & "<B>" & "." & "<B>"
            Else
                HTML = HTML & "<B>" & GridX.TextMatrix(ii, bb) & "<B>"
            End If
        Else
            If Trim(GridX.TextMatrix(ii, bb)) = "" Then
                HTML = HTML & "."
            Else
                HTML = HTML & GridX.TextMatrix(ii, bb)
            End If
        End If
        HTML = HTML & "</TD>"
    Next bb
    HTML = HTML & vbCrLf & "</TR>"
    
Next ii
HTML = HTML & vbCrLf & "</TABLE>"
HTML = HTML & "</FONT>"
HTML = HTML & vbCrLf & "<br><hr>"
ConvertGridToHTML = HTML

Unload frmPGB
End Function
Public Function AddHeaderToHTML() As String
Dim HTML As String
Dim ii As Integer
Dim bb As Integer

HTML = vbCrLf & "<TABLE BORDER=0 WIDTH=100% bgcolor='#c0c0c0'>"
HTML = HTML & "<FONT SIZE = +4 color='#00cc00'>"
HTML = HTML & "<center>"
HTML = HTML & vbCrLf & "My "

HTML = HTML & "</FONT>"
HTML = HTML & "<FONT SIZE = +3 color='#000000'>"
HTML = HTML & "Mate"
HTML = HTML & "</FONT>"
HTML = HTML & "<FONT SIZE = +4 color='#00cc00'>"

HTML = HTML & " !"
HTML = HTML & "</center>"
HTML = HTML & "</FONT>"
HTML = HTML & "</TD></TR>" & vbCrLf
HTML = HTML & "</TABLE>" & vbCrLf



AddHeaderToHTML = HTML
End Function


