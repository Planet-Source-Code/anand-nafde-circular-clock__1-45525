VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Developed in Pure VB"
   ClientHeight    =   3615
   ClientLeft      =   75
   ClientTop       =   330
   ClientWidth     =   7635
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   3120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Just compile this code and execute
'You should be able to see a Analog clock
'where ever you move the mouse.
'Hope you like this clock developed in PURE VB
'with no API calls etc..etc..
'
'Please do not forget to give ratings to me
'if you like this clock :)
'
'Best,
'Anand P. Nafde
'<anandnafde@indiatimes.com>
'

Const PI As Single = 3.14159
Const DEVELOPER As String = "Anand"
Dim str As String
Dim start As Integer
Dim sX As Single, sY As Single

Private Sub Form_Load()
str = Format$(Now(), "dd Mmm yyyy")
On Error GoTo IconError
Icon = LoadPicture(App.Path & "\CLOCK05.ICO")

Exit Sub
IconError:
  MsgBox "Please keep the file 'CLOCK05.ICO' supplied with the ZIP" & vbCrLf & _
        " in the same folder where you are running this application.", vbInformation, "Icon file not found"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sX = X
sY = Y
End Sub

Private Sub Timer1_Timer()
Cls
Dim temp As Integer, iSec As Integer, iMin As Integer, iHour As Integer
Dim i As Integer, temp1 As Integer
Dim j As Single

Caption = "Clock Developed in Pure VB - " & Format$(Now(), "HH:MM:SS")
iSec = Second(Now())
iMin = Minute(Now())
iHour = Hour(Now())
If iHour > 12 Then iHour = iHour - 12

start = start - 1
If start < 1 Then start = 48
temp = start - 1

'Display circular Day
ForeColor = vbYellow
For i = 1 To Len(str)
  temp = temp + 1
  If temp > 48 Then temp = 1
  
  j = temp * (2 * PI / 48)
  CurrentX = sX + 1500 * Cos(j)
  CurrentY = sY + 1500 * Sin(j)
  Me.Print Mid$(str, i, 1)
Next i


'Display hours
ForeColor = vbWhite
temp = 0
For i = 1 To 12
  temp = temp + 1
  
  j = temp * (2 * PI / 12)
  CurrentX = sX + 1200 * Cos(j)
  CurrentY = sY + 1200 * Sin(j)
  temp1 = temp + 3
  If temp1 > 12 Then temp1 = temp1 - 12
  Me.Print temp1
Next i

'<<<<<<<<<<<<<< Display Hour hand >>>>>>>>>>>>>>>
iHour = iHour * 5 - 15
If iHour < 0 Then iHour = 60 + iHour
iHour = iHour + (iMin \ 12)
j = iHour * (2 * PI / 60)
CurrentX = sX + 600 * Cos(j)
CurrentY = sY + 600 * Sin(j)
Me.Line (sX, sY)-(CurrentX, CurrentY), vbRed

'<<<<<<<<<<<<< Display Minute hand >>>>>>>>>>>>>
iMin = iMin - 15
If iMin < 0 Then iMin = 60 + iMin
j = iMin * (2 * PI / 60)
CurrentX = sX + 850 * Cos(j)
CurrentY = sY + 850 * Sin(j)
Me.Line (sX, sY)-(CurrentX, CurrentY), vbBlue

'<<<<<<<<<<<<< Display Second hand >>>>>>>>>>>>>
iSec = iSec - 15
If iSec < 0 Then iSec = 60 + iSec
j = iSec * (2 * PI / 60)
CurrentX = sX + 1000 * Cos(j)
CurrentY = sY + 1000 * Sin(j)
Me.Line (sX, sY)-(CurrentX, CurrentY), vbWhite

FontSize = 12
FontBold = True
CurrentX = sX - (TextWidth(DEVELOPER) / 2)
CurrentY = sY
ForeColor = vbCyan
Me.Print DEVELOPER
FontSize = 10
FontBold = False
End Sub
