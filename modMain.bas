Attribute VB_Name = "modMain"
'' Character Animations
'RestPose
'Blink
'Idle1_4
'Idle1_2
'Idle1_1
'Idle1_3
'Idle3_2
'Idle3_1
'Idle2_2
'Idle2_3
'Show
'Hide
'GestureRight
'Announce
'GestureLeft
'Acknowledge
'Think
'Pleased
'Hearing_4
'Hearing_1
'Hearing_2
'Hearing_3
'Congratulate_2
'Decline
'Confused
'DontRecognize
'StopListening
'Alert
'Sad
'Explain
'Wave
'GetAttention
'GetAttentionReturn
'Surprised
'Greet
'Uncertain
'GestureUp
'GestureDown
'Processing
'Suggest
'Idle1_6
'Idle1_5
'Searching
'MoveRight
'MoveLeft
'MoveUp
'MoveDown
'Read
'ReadReturn
'Writing
'Reading
'Write
'WriteReturn
'StartListening
'LookDown
'LookLeft
'LookRight
'LookUp
'LookUpBlink
'LookRightBlink
'LookLeftBlink
'LookDownBlink
'LookDownReturn
'LookLeftReturn
'LookRightReturn
'LookUpReturn
'ReadContinued
'WriteContinued
'Idle2_1
'GetAttentionContinued
'DoMagic2
'DoMagic1
'Process
'Search
'Congratulate
'Thinking



Option Explicit

Global gsAgentName As String

Global Greetings(5) As Variant
Global aActions(5) As Variant
Global Interupts(9) As Variant
Global TimeSay(5) As Variant
Global gbCheck As Boolean

Global SavedWidth As Long
Global SAvedHeight As Long
Global merlin As IAgentCtlCharacterEx
'Global Const DATAPATH = "merlin.acs"

Global giType As Integer
Global glRemKey As Long
Global gsFileName As String
Global gsSubject As String
Global gsDescription As String

Global giType2 As Integer
Global glRemKey2 As Long
Global gsFileName2 As String
Global gsSubject2 As String
Global gsDescription2 As String


Public Function StartGrid(grdX As MSFlexGrid) As Boolean

grdX.Clear

grdX.Cols = 6
grdX.Rows = 1

grdX.Row = 0

grdX.Col = 0
grdX.ColWidth(0) = 0
grdX = "Hidden"

grdX.Col = 1
grdX.ColWidth(1) = 0
grdX = "date"

grdX.Col = 2
grdX.ColWidth(2) = 1200
grdX = "Time"

grdX.Col = 3
grdX.ColWidth(3) = 2000
grdX = "Subject"

grdX.Col = 4
grdX.ColWidth(4) = 4400
grdX = "Description"

grdX.Col = 5
grdX.ColWidth(5) = 1000
grdX = "Status"
grdX.CellForeColor = vbBlue


End Function

Public Function StartGrid2(grdX As MSFlexGrid) As Boolean

grdX.Clear

grdX.Cols = 5
grdX.Rows = 1

grdX.Row = 0

grdX.Col = 0
grdX.ColWidth(0) = 1000
grdX = "Due date"

grdX.Col = 1
grdX.ColWidth(1) = 800
grdX = "Due time"

grdX.Col = 2
grdX.ColWidth(2) = 2000
grdX = "Subject"

grdX.Col = 3
grdX.ColWidth(3) = 4300
grdX = "Description"

grdX.Col = 4
grdX.ColWidth(4) = 1000
grdX = "Status"

End Function


Public Function Init()
Greetings(0) = "Oh, it's you again!. What do you want now ?"
Greetings(1) = "Yes master, you have called, and I have answered. Now make it quick !"
Greetings(2) = "I was just having a nice dream when you called. What is it this time ?"
Greetings(3) = "Are you always so irritating ? State your bussiness and make it fast !"
Greetings(4) = "What is it now again ? What do you want ? Well ?"

Interupts(0) = "Hey ! I need meat ! Listen here ! When are you going to feed me ?"
Interupts(1) = "I'm really bored by now ! What are you going to do about it ? Get me something to do !"
Interupts(2) = "Listen, about that raise you promised me, when are you going to stop promising and start acting ?"
Interupts(3) = "I need a break ! Everybody takes breaks but me ? Oh no... I work for mister slaver himself !"
Interupts(4) = "Could you please stop clicking around here. Some of us are trying to get some sleep !"
Interupts(5) = "Hey  you ! Hey ! Hey ! What is that foul stench ? Yuck !"
Interupts(6) = "Your computer has laid a formal complaint about his working hours. You will receive the summons soon !"
Interupts(7) = "You aren't really a hard worker, are you ?"
Interupts(8) = "Is your fridge still running ? You better go and catch it ! Hahahaha"
Interupts(9) = "Excuse me. I need to go to the loo !"

TimeSay(0) = "I wonder what time it is !"
TimeSay(1) = "I can't believe the time ! I missed my soap opera !"
TimeSay(2) = "I'm late ! I told the pub I'll be there in 20 minutes, but here I stand !"
TimeSay(3) = "Time to ... say goodbye. Get the hint ? well ?"
TimeSay(4) = "Tick Tock Tick Tock there goes the clock"

aActions(0) = "search"
aActions(1) = "reading"
aActions(3) = "DoMagic1"
aActions(4) = "Surprised"


End Function
Public Function GetRandom(ArrayX As Variant) As String
Dim IndexX As Integer
Dim iActions As Integer

Randomize
Randomize

'If iActions > 90 Then ' Do something funny
'Else
'
  IndexX = Rnd * UBound(ArrayX)
  IndexX = Int(IndexX)
  GetRandom = ArrayX(IndexX)
'End If

End Function
Public Function RandomAction(Index As Integer) As Boolean

' This function will let the agent do a random action.

On Error Resume Next

merlin.StopAll

Select Case Index

Case 1
merlin.Play "GetAttention"
merlin.Speak "Hey funny face. When are you going to organise me decent accomodation ?"

Case 2
merlin.Play "Surprised"
merlin.Play "Surprised"
merlin.Speak "Hey  you ! Hey ! Hey ! What is that foul stench ? No farting in here !"

Case 3
merlin.Play "DoMagic1"
merlin.Play "DoMagic2"
DoEvents
merlin.Play "DoMagic1"
merlin.Play "DoMagic2"
merlin.Play "sad"
merlin.Speak "Oh shit ! Where did your data dissapear to now ?"

Case 4
merlin.Play "Explain"
merlin.Speak "I need a break ! Everybody takes breaks but me ? Oh no... I work for mister slave master himself !"

Case 5
merlin.Play "Read"
merlin.Speak "Interesting article about the computer animation that killed his boring user...Makes one think..."
merlin.Play "ReadReturn"

Case 6
merlin.Play "Write"
merlin.Speak "I am giving you a fine for man handling the poor mouse !"
merlin.Play "Writereturn"

Case 7
merlin.Play "search"
merlin.Speak "I looked at your future and it is not looking good. I would go into hiding if I were you !"

Case 8
merlin.Play "Congratulate_2"
merlin.Speak "Congratulations on being the most boring computer user ever !"

Case 9
merlin.Play "Congratulate"
merlin.Speak "I need beer. Fill her up and make it quick !"

Case 10
merlin.Play "process"
merlin.Speak "Yoohoo ? I need meat for my soup. Go and get me some. NOW !"
merlin.Play "process"

Case 11
merlin.Play "Confused"
merlin.Speak "How come someone as smart as me got stuck with someone like you ?"
merlin.Play "Confused"

Case 12
merlin.Play "idle3_1"
merlin.Speak "Stop clicking around here. Some of us are trying to get some sleep !"
merlin.Play "idle3_1"
merlin.Play "idle3_2"

Case 13
merlin.Play "Surprised"
merlin.Speak "If you smell something, it wasn't me !"

Case 14
merlin.Play "Suggest"
merlin.Speak "Hey ! I just had a good idea. Why don't you go on lunch ? Then I wont have to babysit you !"


Case Else
merlin.Play "Wave"
merlin.Speak "Hi there ! I'm still here you know ?"

End Select
End Function

'**************************************
' Name: SortFlex
' Description:Handle the sorting of a MS
'     Flexgrid by only one sub-routine. Automa
'     tic ascenting and decending displayed by
'     + and - in the Headline.
' By: Dirk
'
' Inputs:syntax:
'SortFlex MSFlexGrid, CollumToSort , StringSortAsBoolean , StringSortAsBoolean ...
'example:
'SortFlex flxProject, flxProject.MouseCol, False, True, True, True
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/xq/ASP/txtCode
'     Id.2158/lngWId.1/qx/vb/scripts/ShowCode.
'     htm'for details.'**************************************

'
' 1999 by Dirk Bujna - b_dirk@yahoo.com
'


'Public Sub SortFlex(FlexGrid As MSFlexGrid, TheCol As Integer, ParamArray IsString() As Variant)
'
'    Dim i As Integer
'    Dim Headline As String
'    Dim Ascend ' As Boolean
'    Dim Decend ' As Boolean
'    FlexGrid.Col = TheCol
'
'
'    For i = 0 To FlexGrid.Cols - 1
'        Headline = FlexGrid.TextMatrix(0, i)
'        Ascend = Right$(Headline, 1) = "+"
'        Decend = Right$(Headline, 1) = "-"
'
'        Ascend = False
'        Decend = True
'
'
'        If Ascend Or Decend Then Headline = Left$(Headline, Len(Headline) - 1)
'
'
'        If i = TheCol Then
'
'
'            If Ascend Then
'                FlexGrid.TextMatrix(0, i) = Headline & "-"
'
'
'                'If IsMissing(IsString(i)) Then
'                If UBound(IsString) < LBound(IsString) Then
'
'                    FlexGrid.Sort = flexSortGenericDescending
'                Else
'
'
'                    If IsString(i) Then
'                        FlexGrid.Sort = flexSortStringDescending
'                    Else
'                        FlexGrid.Sort = flexSortNumericDescending
'                    End If
'                End If
'            Else
'                FlexGrid.TextMatrix(0, i) = Headline & "+"
'
'
'                If IsMissing(IsString(i)) Then
'                    FlexGrid.Sort = flexSortGenericAscending
'                Else
'
'
'                    If IsString(i) Then
'                        FlexGrid.Sort = flexSortStringAscending
'                    Else
'                        FlexGrid.Sort = flexSortNumericAscending
'                    End If
'                End If
'            End If
'        Else
'            FlexGrid.TextMatrix(0, i) = Headline
'        End If
'    Next i
'End Sub
